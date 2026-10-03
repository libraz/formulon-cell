import { addrKey } from '../engine/address.js';
import { makeRangeResolver } from '../engine/range-resolver.js';
import type { Addr, Range } from '../engine/types.js';
import type { CellPatch, CellSnapshot, WorkbookHandle } from '../engine/workbook-handle.js';
import { formatWithPending } from '../store/pending-format.js';
import type { SpreadsheetStore } from '../store/store.js';
import type { CellFormat, State } from '../store/types.js';
import { type CoercedInput, coerceInputForCell } from './coerce-input.js';
import type { History, HistoryDirection, HistoryEntry } from './history.js';
import {
  type BatchRejection,
  type CellBatchCommand,
  type CellBatchOperation,
  type CellChangeInput,
  type ChangeBatchResult,
  type InteractionOperation,
  type InteractionOrigin,
  type InteractionPolicy,
  isMutatingInteraction,
  isPolicyCellEligible,
  type OperationEffect,
  type OperationIntent,
  operationPermission,
  type PermissionDecision,
  rangeContainsAddr,
} from './interaction-policy.js';
import { isCellWritable, isSheetProtected } from './protection.js';
import { normalizeR1C1Formula } from './refs.js';
import { cellValueViolatesValidation, validateAgainst } from './validate.js';

const MAX_MATERIALIZED_CELLS = 100_000;
const IMPLEMENTED_OPERATIONS: ReadonlySet<InteractionOperation> = new Set([
  'valueEdit',
  'formulaEdit',
  'clear',
  'format',
  'paste',
  'fill',
  'moveCells',
  'print',
  'export',
]);
const CELL_OPERATIONS: ReadonlySet<CellBatchOperation> = new Set([
  'valueEdit',
  'formulaEdit',
  'clear',
  'paste',
  'fill',
  'moveCells',
]);

const MAX_XLSX_ROW = 1_048_575;
const MAX_XLSX_COL = 16_383;
const MAC_SUBTOTAL_COMMAND_ID = 'mac.data.subtotal';
const MAC_REMOVE_HYPERLINK_COMMAND_ID = 'mac.automate.removeHyperlinks';

const macSubtotalStructuralOperation = (
  intent: OperationIntent,
): 'insertRows' | 'deleteRows' | null => {
  if (intent.commandId !== MAC_SUBTOTAL_COMMAND_ID) return null;
  if (intent.operation === 'insertRows' && (intent.origin === 'ribbon' || intent.origin === 'redo'))
    return 'insertRows';
  if (intent.operation === 'deleteRows' && intent.origin === 'undo') return 'deleteRows';
  return null;
};

const macSubtotalStructuralRange = (intent: OperationIntent): Range | null => {
  if (macSubtotalStructuralOperation(intent) === null || intent.effects.length !== 2) return null;
  let workbookEffects = 0;
  let range: Range | undefined;
  for (const effect of intent.effects) {
    if (effect.kind === 'workbook') {
      workbookEffects += 1;
      continue;
    }
    if (effect.kind !== 'range' || effect.includesFormula !== undefined || range) return null;
    range = effect.range;
  }
  return workbookEffects === 1 && range ? range : null;
};

const ownsMacRemoveHyperlinkCommand = (intent: OperationIntent): boolean =>
  intent.commandId === MAC_REMOVE_HYPERLINK_COMMAND_ID;

const macRemoveHyperlinkCells = (intent: OperationIntent): readonly Addr[] | null => {
  if (!ownsMacRemoveHyperlinkCommand(intent)) return null;
  if (
    intent.operation !== 'hyperlink' ||
    (intent.origin !== 'ribbon' && intent.origin !== 'undo' && intent.origin !== 'redo') ||
    intent.effects.length !== 1
  )
    return null;
  const effect = intent.effects[0];
  if (effect?.kind !== 'cells' || effect.includesFormula !== undefined || effect.cells.length === 0)
    return null;
  const sheet = effect.cells[0]?.sheet;
  if (sheet === undefined || effect.cells.some((addr) => addr.sheet !== sheet)) return null;
  return effect.cells;
};

export interface InteractionControllerOptions {
  readonly store: SpreadsheetStore;
  readonly getWb: () => WorkbookHandle;
  readonly history: History;
  /**
   * Keep the controller's history authorization guard attached to the shared
   * stack. Ephemeral command adapters leave an existing mounted controller's
   * guard alone while still using the same history callbacks.
   */
  readonly manageHistoryGuard?: boolean;
  readonly getBounds?: () => Range | undefined;
  readonly onChanged?: (result: ChangeBatchResult) => void;
}

export interface ApplyChangesOptions {
  readonly history?: 'reset' | 'record';
  readonly origin?: string;
}

type PreparedChange = {
  readonly patch: CellPatch;
  readonly implicitFormat?: CellFormat['numFmt'];
  readonly input?: CoercedInput;
  /** A format staged by the editor/formula bar for this resolved anchor. */
  readonly pendingFormat?: Partial<CellFormat>;
};

type HistoryIntents = {
  readonly redo: OperationIntent;
  readonly undo: OperationIntent;
  readonly replayAuthorization?: {
    readonly undo: readonly OperationIntent[];
    readonly redo: readonly OperationIntent[];
  };
};

type HistoryPlanResult =
  | { readonly plan: HistoryIntents }
  | { readonly decision: PermissionDecision };

const snapshotToPatch = (snapshot: CellSnapshot): CellPatch => ({
  addr: snapshot.addr,
  value: snapshot.value,
  formula: snapshot.formula,
});

const originFromHost = (origin: string | undefined): InteractionOrigin => {
  switch (origin) {
    case 'editor':
    case 'formulaBar':
    case 'keyboard':
    case 'clipboard':
    case 'fillHandle':
    case 'contextMenu':
    case 'ribbon':
    case 'undo':
    case 'redo':
      return origin;
    default:
      return 'instanceApi';
  }
};

const sameStructuredValue = (left: unknown, right: unknown): boolean => {
  if (Object.is(left, right)) return true;
  if (left === null || right === null || typeof left !== 'object' || typeof right !== 'object')
    return false;
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right) || left.length !== right.length) return false;
    return left.every((value, index) => sameStructuredValue(value, right[index]));
  }
  const leftRecord = left as Record<string, unknown>;
  const rightRecord = right as Record<string, unknown>;
  const leftKeys = Object.keys(leftRecord).sort();
  const rightKeys = Object.keys(rightRecord).sort();
  if (leftKeys.length !== rightKeys.length) return false;
  return leftKeys.every(
    (key, index) =>
      key === rightKeys[index] && sameStructuredValue(leftRecord[key], rightRecord[key]),
  );
};

const mergeFormatPatch = (
  current: CellFormat | undefined,
  patch: Partial<CellFormat> | undefined,
): CellFormat | undefined => {
  if (!patch) return current;
  const next: CellFormat = { ...(current ?? {}), ...patch };
  if (patch.borders) next.borders = { ...(current?.borders ?? {}), ...patch.borders };
  return next;
};

const projectPreparedFormat = (
  current: CellFormat | undefined,
  change: PreparedChange,
): CellFormat | undefined => {
  let next = current;
  if (change.implicitFormat && (next?.numFmt === undefined || next.numFmt.kind === 'general')) {
    next = { ...(next ?? {}), numFmt: change.implicitFormat };
  }
  return mergeFormatPatch(next, change.pendingFormat);
};

const projectPreparedFormats = (
  before: ReadonlyMap<string, CellFormat>,
  prepared: readonly PreparedChange[],
): Map<string, CellFormat> => {
  const next = new Map(before);
  for (const change of prepared) {
    const key = addrKey(change.patch.addr);
    const projected = projectPreparedFormat(next.get(key), change);
    if (projected === undefined) next.delete(key);
    else next.set(key, projected);
  }
  return next;
};

/** Return only addresses whose matching pending patch changes the format after
 * implicit input coercion has been projected. Implicit formats alone retain
 * the value-edit authorization path used before pending-format integration. */
const effectivePendingFormatAddresses = (
  before: ReadonlyMap<string, CellFormat>,
  prepared: readonly PreparedChange[],
): readonly Addr[] => {
  const implicitBase = projectPreparedFormats(
    before,
    prepared.map((change) => ({ ...change, pendingFormat: undefined })),
  );
  const pendingPrepared = prepared
    .filter((change) => change.pendingFormat !== undefined)
    .map((change) => ({ ...change, implicitFormat: undefined }));
  const pendingApplied = projectPreparedFormats(implicitBase, pendingPrepared);
  return changedFormatAddresses(implicitBase, pendingApplied, pendingPrepared);
};

const changedFormatAddresses = (
  before: ReadonlyMap<string, CellFormat>,
  after: ReadonlyMap<string, CellFormat>,
  prepared: readonly PreparedChange[],
): readonly Addr[] => {
  const changed: Addr[] = [];
  const seen = new Set<string>();
  for (const change of prepared) {
    const key = addrKey(change.patch.addr);
    if (seen.has(key)) continue;
    seen.add(key);
    if (!sameStructuredValue(before.get(key), after.get(key))) changed.push(change.patch.addr);
  }
  return changed;
};

const unionAddresses = (...groups: readonly (readonly Addr[])[]): readonly Addr[] => {
  const result: Addr[] = [];
  const seen = new Set<string>();
  for (const group of groups) {
    for (const addr of group) {
      const key = addrKey(addr);
      if (seen.has(key)) continue;
      seen.add(key);
      result.push(addr);
    }
  }
  return result;
};

const effectCells = (effect: OperationEffect): readonly Addr[] | null => {
  if (effect.kind === 'cells')
    return effect.cells.length <= MAX_MATERIALIZED_CELLS ? effect.cells : null;
  if (effect.kind === 'workbook') return [];
  const { range } = effect;
  if (
    !Number.isInteger(range.sheet) ||
    !Number.isInteger(range.r0) ||
    !Number.isInteger(range.c0) ||
    !Number.isInteger(range.r1) ||
    !Number.isInteger(range.c1) ||
    range.sheet < 0 ||
    range.r0 < 0 ||
    range.c0 < 0 ||
    range.r1 < range.r0 ||
    range.c1 < range.c0
  )
    return null;
  const area = (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);
  if (!Number.isSafeInteger(area) || area < 0 || area > MAX_MATERIALIZED_CELLS) return null;
  const cells: Addr[] = [];
  for (let row = range.r0; row <= range.r1; row += 1) {
    for (let col = range.c0; col <= range.c1; col += 1)
      cells.push({ sheet: range.sheet, row, col });
  }
  return cells;
};

/** Return the merge covering a cell without importing the merge command module
 * (the command module itself records history and would create a needless
 * dependency cycle at this policy boundary). */
const mergeRangeAt = (state: State, addr: Addr): Range | null => {
  const key = addrKey(addr);
  const anchorKey = state.merges.byCell.get(key) ?? key;
  return state.merges.byAnchor.get(anchorKey) ?? null;
};

const mergeCellsFor = (state: State, addr: Addr): readonly Addr[] | null => {
  const merge = mergeRangeAt(state, addr);
  if (!merge) return [addr];
  const height = merge.r1 - merge.r0 + 1;
  const width = merge.c1 - merge.c0 + 1;
  const area = height * width;
  if (!Number.isSafeInteger(area) || height <= 0 || width <= 0 || area > MAX_MATERIALIZED_CELLS)
    return null;
  const cells: Addr[] = [];
  for (let row = merge.r0; row <= merge.r1; row += 1) {
    for (let col = merge.c0; col <= merge.c1; col += 1)
      cells.push({ sheet: merge.sheet, row, col });
  }
  return cells;
};

const mergeAnchorFor = (state: State, addr: Addr): Addr => {
  const merge = mergeRangeAt(state, addr);
  return merge ? { sheet: merge.sheet, row: merge.r0, col: merge.c0 } : addr;
};

const cellEffects = (
  cells: readonly Addr[],
  formulaCells: readonly Addr[] = [],
): readonly OperationEffect[] => {
  const effects: OperationEffect[] = [{ kind: 'cells', cells }];
  if (formulaCells.length > 0)
    effects.push({ kind: 'cells', cells: formulaCells, includesFormula: true });
  return effects;
};

export class InteractionController {
  private policyValue: InteractionPolicy | undefined;
  private readonly store: SpreadsheetStore;
  private readonly getWb: () => WorkbookHandle;
  private readonly history: History;
  private readonly manageHistoryGuard: boolean;
  private readonly getBounds?: () => Range | undefined;
  private readonly onChanged?: (result: ChangeBatchResult) => void;
  private readonly listeners = new Set<(result: ChangeBatchResult) => void>();
  private disposed = false;
  private revisionValue = 0;

  constructor(options: InteractionControllerOptions) {
    this.store = options.store;
    this.getWb = options.getWb;
    this.history = options.history;
    this.manageHistoryGuard = options.manageHistoryGuard !== false;
    this.getBounds = options.getBounds;
    this.onChanged = options.onChanged;
    if (this.manageHistoryGuard) {
      this.history.setGuard((entry, direction) => this.guardHistory(entry, direction));
    }
  }

  get policy(): InteractionPolicy | undefined {
    return this.policyValue;
  }

  get restricted(): boolean {
    return this.policyValue !== undefined;
  }

  get revision(): number {
    return this.revisionValue;
  }

  setPolicy(next?: InteractionPolicy): void {
    this.policyValue = next;
    if (this.manageHistoryGuard) {
      this.history.setGuard((entry, direction) => this.guardHistory(entry, direction));
    }
    // Policy changes invalidate any host-held plan/request that was based on
    // the previous authorization. Notify renderer subscribers without sending
    // a mutation result through the host's onChanged callback.
    this.revisionValue += 1;
    const update: ChangeBatchResult = {
      status: 'noop',
      applied: [],
      rejected: [],
      revision: this.revisionValue,
    };
    this.notifySubscribers(update);
  }

  canSelect(): PermissionDecision {
    if (this.policyValue?.selection !== false) return { allowed: true };
    return { allowed: false, code: 'operationDenied', reason: 'selection is disabled' };
  }

  canCopy(): PermissionDecision {
    if (this.policyValue?.copy !== false) return { allowed: true };
    return { allowed: false, code: 'operationDenied', reason: 'copy is disabled' };
  }

  canExecute(intent: OperationIntent): PermissionDecision {
    if (this.disposed) return { allowed: false, code: 'invalid', reason: 'controller is disposed' };
    const policy = this.policyValue;
    try {
      const wb = this.getWb();
      if (!wb) return { allowed: false, code: 'invalid', reason: 'workbook is unavailable' };
      if (ownsMacRemoveHyperlinkCommand(intent)) {
        if (!macRemoveHyperlinkCells(intent)) {
          return {
            allowed: false,
            code: 'invalid',
            reason: 'Remove hyperlinks authorization needs one nonempty same-sheet cell effect',
          };
        }
        if (!policy) return this.legacyDecision(intent, wb);
        return this.policyDecision(intent, policy, wb);
      }
      if (macSubtotalStructuralOperation(intent) !== null)
        return this.macSubtotalStructuralDecision(intent, policy, wb);
      if (!policy) return this.legacyDecision(intent, wb);
      return this.policyDecision(intent, policy, wb);
    } catch {
      return { allowed: false, code: 'invalid', reason: 'workbook is unavailable' };
    }
  }

  private macSubtotalStructuralDecision(
    intent: OperationIntent,
    policy: InteractionPolicy | undefined,
    wb: WorkbookHandle,
  ): PermissionDecision {
    const range = macSubtotalStructuralRange(intent);
    if (!range) {
      return {
        allowed: false,
        code: 'invalid',
        reason: 'Subtotal structural authorization needs one workbook and one range effect',
      };
    }
    const state = this.store.getState();
    if (range.sheet !== state.data.sheetIndex) {
      return {
        allowed: false,
        code: 'outOfBounds',
        reason: 'Subtotal structural authorization must target the active sheet',
      };
    }
    if (
      !Number.isInteger(range.sheet) ||
      !Number.isInteger(range.r0) ||
      !Number.isInteger(range.c0) ||
      !Number.isInteger(range.r1) ||
      !Number.isInteger(range.c1) ||
      range.r0 < 0 ||
      range.c0 < 0 ||
      range.r1 < range.r0 ||
      range.c1 < range.c0 ||
      !this.isValidAddress(wb, { sheet: range.sheet, row: range.r0, col: range.c0 }) ||
      !this.isValidAddress(wb, { sheet: range.sheet, row: range.r1, col: range.c1 })
    ) {
      return {
        allowed: false,
        code: 'outOfBounds',
        reason: 'Subtotal structural authorization range is outside the worksheet',
      };
    }
    if (isSheetProtected(state, range.sheet)) {
      return {
        allowed: false,
        code: 'protected',
        reason: 'Subtotal structural authorization targets a protected sheet',
      };
    }

    if (!policy) return { allowed: true };
    if (!operationPermission(policy, intent.operation)) {
      return {
        allowed: false,
        code:
          policy.readOnly && isMutatingInteraction(intent.operation)
            ? 'readOnly'
            : 'operationDenied',
        reason:
          policy.readOnly && isMutatingInteraction(intent.operation)
            ? 'instance policy is read-only'
            : `operation ${intent.operation} is denied`,
      };
    }
    if (policy.operations?.[intent.operation] !== true) {
      return {
        allowed: false,
        code: 'operationDenied',
        reason: `operation ${intent.operation} needs explicit structural permission`,
      };
    }
    if (policy.editable !== undefined) {
      return {
        allowed: false,
        code: 'operationDenied',
        reason: 'editable cell policy cannot authorize workbook row structure',
      };
    }
    const bounds = this.getBounds?.();
    if (
      bounds &&
      (bounds.sheet !== range.sheet ||
        bounds.r0 !== 0 ||
        bounds.c0 !== 0 ||
        bounds.r1 !== MAX_XLSX_ROW ||
        bounds.c1 !== MAX_XLSX_COL)
    ) {
      return {
        allowed: false,
        code: 'outOfBounds',
        reason: 'Subtotal structural authorization needs the full target worksheet bound',
      };
    }
    return this.applyRestriction(policy, intent, { allowed: true });
  }

  private policyDecision(
    intent: OperationIntent,
    policy: InteractionPolicy,
    wb: WorkbookHandle,
  ): PermissionDecision {
    if (!IMPLEMENTED_OPERATIONS.has(intent.operation) && !macRemoveHyperlinkCells(intent)) {
      return {
        allowed: false,
        code: 'unsupported',
        reason: `operation ${intent.operation} is not routed yet`,
      };
    }
    if (!operationPermission(policy, intent.operation)) {
      return {
        allowed: false,
        code:
          policy.readOnly && isMutatingInteraction(intent.operation)
            ? 'readOnly'
            : 'operationDenied',
        reason:
          policy.readOnly && isMutatingInteraction(intent.operation)
            ? 'instance policy is read-only'
            : `operation ${intent.operation} is denied`,
      };
    }

    const cells: Addr[] = [];
    const formulaCells: Addr[] = [];
    for (const effect of intent.effects) {
      const addresses = effectCells(effect);
      if (addresses === null) {
        return {
          allowed: false,
          code: 'unsupported',
          reason: 'affected range exceeds the authorization bound',
        };
      }
      cells.push(...addresses);
      if (effect.kind !== 'workbook' && effect.includesFormula) formulaCells.push(...addresses);
    }
    const state = this.store.getState();
    const authorizedCells: Addr[] = [];
    const authorizedKeys = new Set<string>();
    for (const requestedAddr of cells) {
      if (authorizedKeys.has(addrKey(requestedAddr))) continue;
      const merged = mergeCellsFor(state, requestedAddr);
      if (merged === null) {
        return {
          allowed: false,
          code: 'unsupported',
          addr: requestedAddr,
          reason: 'merged cell range exceeds the authorization bound',
        };
      }
      for (const expandedAddr of merged) {
        const key = addrKey(expandedAddr);
        if (authorizedKeys.has(key)) continue;
        authorizedKeys.add(key);
        authorizedCells.push(expandedAddr);
      }
    }
    for (const addr of authorizedCells) {
      const bounds = this.getBounds?.();
      if (bounds && !rangeContainsAddr(bounds, addr)) {
        return {
          allowed: false,
          code: 'outOfBounds',
          addr,
          reason: 'cell is outside the configured bounds',
        };
      }
      if (!this.isValidAddress(wb, addr)) {
        return {
          allowed: false,
          code: 'outOfBounds',
          addr,
          reason: 'cell address is outside the workbook',
        };
      }
      if (isMutatingInteraction(intent.operation)) {
        if (!isPolicyCellEligible(policy, addr, intent.operation, intent.origin)) {
          return {
            allowed: false,
            code: 'cellIneligible',
            addr,
            reason: 'cell is outside editable cells',
          };
        }
        if (!isCellWritable(state, addr)) {
          return { allowed: false, code: 'protected', addr, reason: 'cell is protected' };
        }
      }
    }
    if (formulaCells.length > 0 && intent.operation !== 'formulaEdit') {
      const formulaDecision = this.policyDecision(
        {
          ...intent,
          operation: 'formulaEdit',
          effects: [{ kind: 'cells', cells: formulaCells, includesFormula: true }],
        },
        policy,
        wb,
      );
      if (!formulaDecision.allowed) return formulaDecision;
    }
    const allowed: PermissionDecision = { allowed: true };
    return this.applyRestriction(policy, intent, allowed, authorizedCells[0] ?? cells[0]);
  }

  execute(command: CellBatchCommand): ChangeBatchResult {
    const initial: ChangeBatchResult = {
      status: 'noop',
      applied: [],
      rejected: [],
      revision: this.revisionValue,
    };
    if (this.disposed)
      return this.rejectResult([{ code: 'invalid', reason: 'controller is disposed' }]);
    if (command.type !== 'cellBatch' || !CELL_OPERATIONS.has(command.operation)) {
      return this.rejectResult([{ code: 'unsupported', reason: 'unsupported cell batch command' }]);
    }
    if (command.changes.length === 0) return initial;

    let wb: WorkbookHandle;
    try {
      wb = this.getWb();
    } catch {
      return this.rejectResult([{ code: 'invalid', reason: 'workbook is unavailable' }]);
    }
    const state = this.store.getState();
    const prepared: PreparedChange[] = [];
    const globalRejections: BatchRejection[] = [];
    let includesFormula = false;
    const pending = state.ui.pendingFormat;
    const mayApplyPendingFormat =
      (command.origin === 'editor' || command.origin === 'formulaBar') &&
      (command.operation === 'valueEdit' || command.operation === 'formulaEdit');
    const pendingAnchorKey =
      mayApplyPendingFormat && pending ? addrKey(mergeAnchorFor(state, pending.addr)) : undefined;
    for (const change of command.changes) {
      const preparedChange = this.prepareChange(state, change, command.operation, wb, true);
      if ('rejection' in preparedChange) {
        globalRejections.push(preparedChange.rejection);
        continue;
      }
      if (preparedChange.patch.formula !== null) includesFormula = true;
      if (pendingAnchorKey === addrKey(preparedChange.patch.addr) && pending) {
        prepared.push({
          ...preparedChange,
          pendingFormat: structuredClone(pending.format),
        });
      } else {
        prepared.push(preparedChange);
      }
    }
    if (globalRejections.length > 0) return this.rejectResult(globalRejections);

    const intent: OperationIntent = {
      operation: command.operation,
      origin: command.origin,
      commandId: command.commandId,
      effects: cellEffects(
        prepared.map((change) => change.patch.addr),
        prepared
          .filter((change) => change.patch.formula !== null)
          .map((change) => change.patch.addr),
      ),
    };
    const operationDecision = this.canExecute({ ...intent, effects: [{ kind: 'workbook' }] });
    if (!operationDecision.allowed) return this.rejectResult([this.asRejection(operationDecision)]);
    if (includesFormula && command.operation !== 'formulaEdit') {
      const formulaDecision = this.canExecute({
        operation: 'formulaEdit',
        origin: command.origin,
        commandId: command.commandId,
        effects: [{ kind: 'workbook' }],
      });
      if (!formulaDecision.allowed) return this.rejectResult([this.asRejection(formulaDecision)]);
    }

    const beforeFormats = new Map(state.format.formats);
    const pendingFormatAddresses = effectivePendingFormatAddresses(beforeFormats, prepared);
    const pendingFormatAddressKeys = new Set(pendingFormatAddresses.map(addrKey));
    if (pendingFormatAddresses.length > 0) {
      const formatGlobal = this.canExecute({
        operation: 'format',
        origin: command.origin,
        commandId: command.commandId,
        effects: [{ kind: 'workbook' }],
      });
      if (!formatGlobal.allowed) return this.rejectResult([this.asRejection(formatGlobal)]);
    }

    const denied = command.denied ?? this.policyValue?.batchDenied ?? 'reject';
    const eligible: PreparedChange[] = [];
    const rejected: BatchRejection[] = [];
    for (const change of prepared) {
      let decision = this.canExecute({
        ...intent,
        effects: cellEffects(
          [change.patch.addr],
          change.patch.formula !== null ? [change.patch.addr] : [],
        ),
      });
      if (decision.allowed && pendingFormatAddressKeys.has(addrKey(change.patch.addr))) {
        decision = this.canExecute({
          operation: 'format',
          origin: command.origin,
          commandId: command.commandId,
          effects: [{ kind: 'cells', cells: [change.patch.addr] }],
        });
      }
      if (decision.allowed) {
        eligible.push(change);
        continue;
      }
      const skippable = decision.code === 'cellIneligible' || decision.code === 'protected';
      if (denied === 'skipIneligible' && skippable) {
        rejected.push(this.asRejection(decision));
      } else {
        return this.rejectResult([this.asRejection(decision)]);
      }
    }
    if (eligible.length === 0) return this.noopResult(rejected);
    const committedIntent: OperationIntent = {
      ...intent,
      effects: cellEffects(
        eligible.map((change) => change.patch.addr),
        eligible
          .filter((change) => change.patch.formula !== null)
          .map((change) => change.patch.addr),
      ),
    };
    const historyPlan = this.planHistory(
      committedIntent,
      eligible,
      wb,
      false,
      this.policyValue === undefined && command.operation === 'clear',
    );
    if ('decision' in historyPlan)
      return this.rejectResult([this.asRejection(historyPlan.decision)]);
    return this.commitPrepared(eligible, rejected, true, historyPlan.plan);
  }

  applyChanges(
    changes: readonly CellChangeInput[],
    options: ApplyChangesOptions = {},
  ): ChangeBatchResult {
    if (this.disposed)
      return this.rejectResult([{ code: 'invalid', reason: 'controller is disposed' }]);
    if (changes.length === 0) return this.emptyResult();
    let wb: WorkbookHandle;
    try {
      wb = this.getWb();
    } catch {
      return this.rejectResult([{ code: 'invalid', reason: 'workbook is unavailable' }]);
    }
    const state = this.store.getState();
    const prepared: PreparedChange[] = [];
    const rejections: BatchRejection[] = [];
    let includesFormula = false;
    for (const change of changes) {
      // Host refreshes address the physical workbook and may intentionally
      // update cells outside the current navigation/viewport rectangle.
      const next = this.prepareChange(state, change, 'valueEdit', wb, false);
      if ('rejection' in next) rejections.push(next.rejection);
      else {
        prepared.push(next);
        includesFormula ||= next.patch.formula !== null;
      }
    }
    if (rejections.length > 0) return this.rejectResult(rejections);
    const operation: CellBatchOperation = includesFormula ? 'formulaEdit' : 'valueEdit';
    const origin = originFromHost(options.origin);
    const intent: OperationIntent = {
      operation,
      origin,
      effects: cellEffects(
        prepared.map((change) => change.patch.addr),
        prepared
          .filter((change) => change.patch.formula !== null)
          .map((change) => change.patch.addr),
      ),
    };
    let historyPlan: HistoryIntents | undefined;
    if (options.history === 'record') {
      const planned = this.planHistory(intent, prepared, wb, true);
      if ('decision' in planned) return this.rejectResult([this.asRejection(planned.decision)]);
      historyPlan = planned.plan;
    }
    const resetHistory = (options.history ?? 'reset') === 'reset';
    return this.commitPrepared(
      prepared,
      [],
      options.history === 'record',
      historyPlan,
      resetHistory ? () => this.history.clear() : undefined,
    );
  }

  subscribe(listener: (result: ChangeBatchResult) => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    if (this.manageHistoryGuard) this.history.setGuard(null);
    this.listeners.clear();
  }

  private prepareChange(
    state: State,
    change: CellChangeInput,
    operation: CellBatchOperation,
    wb: WorkbookHandle,
    respectNavigationBounds: boolean,
  ): PreparedChange | { rejection: BatchRejection } {
    const addr = change.addr;
    if (
      !Number.isInteger(addr.sheet) ||
      !Number.isInteger(addr.row) ||
      !Number.isInteger(addr.col)
    ) {
      return { rejection: { addr, code: 'invalid', reason: 'cell address must be integral' } };
    }
    if (addr.sheet < 0 || addr.row < 0 || addr.col < 0) {
      return { rejection: { addr, code: 'outOfBounds', reason: 'cell address is negative' } };
    }
    if (!this.isValidAddress(wb, addr)) {
      return { rejection: { addr, code: 'outOfBounds', reason: 'cell is outside the workbook' } };
    }
    const merged = mergeCellsFor(state, addr);
    if (merged === null) {
      return {
        rejection: { addr, code: 'unsupported', reason: 'merged cell range exceeds the limit' },
      };
    }
    const target = mergeAnchorFor(state, addr);
    if (!this.isValidAddress(wb, target)) {
      return {
        rejection: { addr, code: 'outOfBounds', reason: 'merge anchor is outside the workbook' },
      };
    }
    if (respectNavigationBounds) {
      const bounds = this.getBounds?.();
      if (bounds && merged.some((cell) => !rangeContainsAddr(bounds, cell))) {
        return {
          rejection: {
            addr,
            code: 'outOfBounds',
            reason: 'cell or its merged anchor is outside the configured bounds',
          },
        };
      }
    }
    if (operation === 'clear') {
      return { patch: { addr: target, value: { kind: 'blank' }, formula: null } };
    }
    if ('input' in change) {
      const input = coerceInputForCell(state, target, change.input);
      if (input.kind === 'number' && !Number.isFinite(input.value)) {
        return { rejection: { addr, code: 'invalid', reason: 'number value must be finite' } };
      }
      const prepared = this.patchFromCoerced(target, input);
      const validation = formatWithPending(state, target)?.validation;
      if (validation) {
        const outcome = validateAgainst(validation, input, makeRangeResolver(wb, target.sheet));
        if (!outcome.ok && outcome.severity === 'stop' && validation.showErrorMessage !== false) {
          return { rejection: { addr, code: 'invalid', reason: outcome.message } };
        }
      }
      return { ...prepared, input };
    }
    const formula = change.formula === undefined ? null : change.formula;
    if (formula !== null && typeof formula !== 'string') {
      return { rejection: { addr, code: 'invalid', reason: 'formula must be a string or null' } };
    }
    if (change.value.kind === 'number' && !Number.isFinite(change.value.value)) {
      return { rejection: { addr, code: 'invalid', reason: 'number value must be finite' } };
    }
    const validation = formatWithPending(state, target)?.validation;
    if (formula === null && validation) {
      const resolveRange = makeRangeResolver(wb, target.sheet);
      if (cellValueViolatesValidation(change.value, validation, resolveRange)) {
        return {
          rejection: {
            addr,
            code: 'invalid',
            reason: validation.errorMessage || 'value violates validation',
          },
        };
      }
    }
    return {
      patch: {
        addr: target,
        value: change.value,
        formula: formula === null ? null : normalizeR1C1Formula(formula, target),
      },
    };
  }

  private patchFromCoerced(addr: Addr, input: CoercedInput): PreparedChange {
    switch (input.kind) {
      case 'blank':
        return { patch: { addr, value: { kind: 'blank' }, formula: null }, input };
      case 'formula':
        return {
          patch: {
            addr,
            value: { kind: 'blank' },
            formula: normalizeR1C1Formula(input.text, addr),
          },
          input,
        };
      case 'number':
        return {
          patch: { addr, value: { kind: 'number', value: input.value }, formula: null },
          implicitFormat: input.implicitFormat,
          input,
        };
      case 'bool':
        return {
          patch: { addr, value: { kind: 'bool', value: input.value }, formula: null },
          input,
        };
      case 'text':
        return {
          patch: { addr, value: { kind: 'text', value: input.value }, formula: null },
          input,
        };
    }
  }

  /** Build the forward and inverse authorization intents before touching the
   * engine. Inverse formula permission is derived from the actual pre-write
   * snapshots as well as from the forward command, so a value write replacing
   * a prefilled formula cannot later undo through a denied formula path. */
  private planHistory(
    intent: OperationIntent,
    prepared: readonly PreparedChange[],
    wb: WorkbookHandle,
    requireInverseAuthorization: boolean,
    forwardAlreadyAuthorized = false,
  ): HistoryPlanResult {
    let hadFormula = false;
    let beforeFormulaCells: readonly Addr[] = [];
    try {
      const formulas = wb.cellFormulas(prepared.map((change) => change.patch.addr));
      beforeFormulaCells = prepared
        .filter((change) => formulas.get(addrKey(change.patch.addr)) !== null)
        .map((change) => change.patch.addr);
      hadFormula = beforeFormulaCells.length > 0;
    } catch (error) {
      return {
        decision: {
          allowed: false,
          code: 'invalid',
          reason: error instanceof Error ? error.message : 'could not inspect the inverse patch',
        },
      };
    }
    const cells = prepared.map((change) => change.patch.addr);
    const forwardFormulaCells = prepared
      .filter((change) => change.patch.formula !== null)
      .map((change) => change.patch.addr);
    const redo: OperationIntent = {
      ...intent,
      origin: 'redo',
      effects: cellEffects(cells, forwardFormulaCells),
    };
    const undo: OperationIntent = {
      ...intent,
      operation: hadFormula ? 'formulaEdit' : 'valueEdit',
      origin: 'undo',
      effects: cellEffects(cells, beforeFormulaCells),
    };
    const beforeFormats = new Map(this.store.getState().format.formats);
    const formatCells = effectivePendingFormatAddresses(beforeFormats, prepared);
    const replayAuthorization =
      formatCells.length > 0
        ? {
            undo: [
              undo,
              {
                ...intent,
                operation: 'format',
                origin: 'undo',
                effects: [{ kind: 'cells', cells: formatCells }],
              } satisfies OperationIntent,
            ],
            redo: [
              redo,
              {
                ...intent,
                operation: 'format',
                origin: 'redo',
                effects: [{ kind: 'cells', cells: formatCells }],
              } satisfies OperationIntent,
            ],
          }
        : undefined;
    if (!forwardAlreadyAuthorized) {
      const forward = this.preflightIntent(requireInverseAuthorization ? redo : intent);
      if (!forward.allowed) return { decision: forward };
    }
    if (requireInverseAuthorization) {
      const inverse = this.preflightIntent(undo);
      if (!inverse.allowed) return { decision: inverse };
    }
    return { plan: { redo, undo, replayAuthorization } };
  }

  /** Authorize a complete intent, including every cell effect and any formula
   * capability implied by a value/paste/fill command. */
  private preflightIntent(intent: OperationIntent): PermissionDecision {
    if (macSubtotalStructuralOperation(intent) !== null) return this.canExecute(intent);
    const global = this.canExecute({ ...intent, effects: [{ kind: 'workbook' }] });
    if (!global.allowed) return global;
    const cells: Addr[] = [];
    const formulaCells: Addr[] = [];
    for (const effect of intent.effects) {
      const addresses = effectCells(effect);
      if (addresses === null)
        return { allowed: false, code: 'unsupported', reason: 'invalid affected range' };
      cells.push(...addresses);
      if (effect.kind !== 'workbook' && effect.includesFormula) formulaCells.push(...addresses);
    }
    const formulaKeys = new Set(formulaCells.map(addrKey));
    for (const addr of cells) {
      const decision = this.canExecute({
        ...intent,
        effects: [
          {
            kind: 'cells',
            cells: [addr],
            includesFormula: formulaKeys.has(addrKey(addr)),
          },
        ],
      });
      if (!decision.allowed) return decision;
    }
    if (formulaCells.length > 0 && intent.operation !== 'formulaEdit') {
      return this.preflightIntent({
        ...intent,
        operation: 'formulaEdit',
        effects: [{ kind: 'cells', cells: formulaCells, includesFormula: true }],
      });
    }
    return { allowed: true };
  }

  private commitPrepared(
    prepared: readonly PreparedChange[],
    rejected: readonly BatchRejection[],
    recordHistory = true,
    historyPlan?: HistoryIntents,
    beforeNotify?: () => void,
  ): ChangeBatchResult {
    let wb: WorkbookHandle;
    try {
      wb = this.getWb();
      const beforeFormats = new Map(this.store.getState().format.formats);
      const atomic = wb.applyCellPatchAtomic(prepared.map((change) => change.patch));
      const afterFormats = this.applyPreparedFormats(beforeFormats, prepared);
      const changedFormats = changedFormatAddresses(beforeFormats, afterFormats, prepared);
      const applied = unionAddresses(atomic.changed, changedFormats);
      if (applied.length === 0) {
        beforeNotify?.();
        return this.noopResult(rejected);
      }
      if (recordHistory) {
        const intents = historyPlan;
        if (!intents) throw new Error('history plan missing');
        this.history.push({
          undo: () => {
            const current = this.getWb();
            current.applyCellPatchAtomic(atomic.before.map(snapshotToPatch));
            this.restoreFormats(beforeFormats);
          },
          redo: () => {
            const current = this.getWb();
            current.applyCellPatchAtomic(atomic.after.map(snapshotToPatch));
            this.restoreFormats(afterFormats);
          },
          intent: intents.redo,
          inverseIntent: intents.undo,
          replayAuthorization: intents.replayAuthorization,
        });
      }
      this.revisionValue += 1;
      const result: ChangeBatchResult = {
        status: 'applied',
        applied,
        rejected,
        revision: this.revisionValue,
      };
      beforeNotify?.();
      this.notifyChanged(result);
      return result;
    } catch (error) {
      const reason = error instanceof Error ? error.message : 'cell patch failed';
      return this.rejectResult([{ code: 'invalid', reason }]);
    }
  }

  private applyPreparedFormats(
    before: ReadonlyMap<string, CellFormat>,
    prepared: readonly PreparedChange[],
  ): ReadonlyMap<string, CellFormat> {
    const next = projectPreparedFormats(before, prepared);
    if (
      next.size !== before.size ||
      [...next].some(([key, value]) => !sameStructuredValue(before.get(key), value))
    ) {
      this.restoreFormats(next);
    }
    return next;
  }

  private restoreFormats(formats: ReadonlyMap<string, CellFormat>): void {
    try {
      this.store.setState((state) => ({
        ...state,
        format: { ...state.format, formats: new Map(formats) },
      }));
    } catch {
      // Zustand commits the state before notifying subscribers. A throwing
      // renderer observer must not turn an already committed engine patch
      // into a rejected batch or prevent history/revision publication.
    }
  }

  private guardHistory(entry: HistoryEntry, direction: HistoryDirection): boolean {
    const replayAuthorization = entry.replayAuthorization;
    if (replayAuthorization) {
      const intents = direction === 'undo' ? replayAuthorization.undo : replayAuthorization.redo;
      if (intents.length === 0) return false;
      let allowed = true;
      for (const intent of intents) {
        if (
          macSubtotalStructuralOperation(intent) !== null ||
          ownsMacRemoveHyperlinkCommand(intent)
        ) {
          if (!this.canExecute(intent).allowed) allowed = false;
        } else if (this.policyValue && !this.preflightIntent(intent).allowed) {
          allowed = false;
        }
      }
      return allowed;
    }
    const intent = direction === 'undo' ? entry.inverseIntent : entry.intent;
    if (!intent) return this.policyValue === undefined;
    if (macSubtotalStructuralOperation(intent) !== null || ownsMacRemoveHyperlinkCommand(intent))
      return this.canExecute(intent).allowed;
    if (!this.policyValue) return true;
    return this.preflightIntent(intent).allowed;
  }

  private legacyDecision(intent: OperationIntent, wb: WorkbookHandle): PermissionDecision {
    for (const effect of intent.effects) {
      const addresses = effectCells(effect);
      if (addresses === null) return { allowed: false, code: 'unsupported' };
      for (const addr of addresses) {
        if (!this.isValidAddress(wb, addr)) return { allowed: false, code: 'outOfBounds', addr };
        if (
          isMutatingInteraction(intent.operation) &&
          !isCellWritable(this.store.getState(), addr)
        ) {
          return { allowed: false, code: 'protected', addr };
        }
      }
    }
    return { allowed: true };
  }

  private applyRestriction(
    policy: InteractionPolicy,
    intent: OperationIntent,
    decision: PermissionDecision,
    addr?: Addr,
  ): PermissionDecision {
    if (!policy.restrict) return decision;
    try {
      const restricted = policy.restrict({
        intent,
        decision,
        addr,
        operation: intent.operation,
        origin: intent.origin,
        commandId: intent.commandId,
      });
      if (
        restricted &&
        typeof restricted === 'object' &&
        'then' in restricted &&
        typeof restricted.then === 'function'
      ) {
        void Promise.resolve(restricted).catch(() => {});
        return { allowed: false, code: 'invalid', reason: 'restriction hook must be synchronous' };
      }
      if (restricted === false)
        return { allowed: false, code: 'operationDenied', reason: 'restricted by host' };
      if (restricted && typeof restricted === 'object' && restricted.allowed === false)
        return restricted;
    } catch (error) {
      return {
        allowed: false,
        code: 'invalid',
        reason: error instanceof Error ? error.message : 'restriction callback failed',
      };
    }
    return decision;
  }

  private isValidAddress(wb: WorkbookHandle, addr: Addr): boolean {
    try {
      return (
        Number.isInteger(addr.sheet) &&
        Number.isInteger(addr.row) &&
        Number.isInteger(addr.col) &&
        addr.sheet >= 0 &&
        addr.row >= 0 &&
        addr.col >= 0 &&
        addr.row <= MAX_XLSX_ROW &&
        addr.col <= MAX_XLSX_COL &&
        addr.sheet < wb.sheetCount
      );
    } catch {
      return false;
    }
  }

  private asRejection(decision: PermissionDecision): BatchRejection {
    return decision.allowed
      ? { code: 'invalid', reason: 'unexpected allowed decision' }
      : { addr: decision.addr, code: decision.code, reason: decision.reason };
  }

  private emptyResult(): ChangeBatchResult {
    return this.noopResult([]);
  }

  private noopResult(rejected: readonly BatchRejection[]): ChangeBatchResult {
    const result: ChangeBatchResult = {
      status: 'noop',
      applied: [],
      rejected,
      revision: this.revisionValue,
    };
    this.notifySubscribers(result);
    return result;
  }

  private rejectResult(rejected: readonly BatchRejection[]): ChangeBatchResult {
    const result: ChangeBatchResult = {
      status: 'rejected',
      applied: [],
      rejected,
      revision: this.revisionValue,
    };
    this.notifySubscribers(result);
    return result;
  }

  private notifyChanged(result: ChangeBatchResult): void {
    try {
      this.onChanged?.(result);
    } catch {
      // Host event observers cannot turn a committed engine transaction into a
      // rejected result.
    }
    this.notifySubscribers(result);
  }

  private notifySubscribers(result: ChangeBatchResult): void {
    for (const listener of [...this.listeners]) {
      try {
        listener(result);
      } catch {
        // One renderer subscriber must not truncate the remaining observers.
      }
    }
  }
}

const controllers = new WeakMap<SpreadsheetStore, InteractionController>();

export function registerInteractionController(
  store: SpreadsheetStore,
  controller: InteractionController,
): () => void {
  controllers.set(store, controller);
  return () => {
    if (controllers.get(store) === controller) controllers.delete(store);
  };
}

export function interactionControllerFor(
  store: SpreadsheetStore,
): InteractionController | undefined {
  return controllers.get(store);
}
