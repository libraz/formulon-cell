import { addrKey, MAX_COL, MAX_ROW } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { rangeContainsAddr } from '../store/selection-geometry.js';
import type { SpreadsheetStore } from '../store/store.js';
import type { HistoryDirection, HistoryEntry } from './history.js';
import { effectCells, isWorkbookAddress, mergeCellsFor } from './interaction-effects.js';
import {
  type InteractionOperation,
  type InteractionPolicy,
  isMutatingInteraction,
  isPolicyCellEligible,
  type OperationIntent,
  operationPermission,
  type PermissionDecision,
} from './interaction-policy.js';
import { isCellWritable, isSheetProtected } from './protection.js';

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

type CanExecute = (intent: OperationIntent) => PermissionDecision;

export interface InteractionAuthorizerOptions {
  readonly store: SpreadsheetStore;
  readonly getWb: () => WorkbookHandle;
  readonly getBounds?: () => Range | undefined;
  readonly isDisposed: () => boolean;
}

/** Permission decisions for interaction intents: the instance policy, the
 * legacy protection/bounds gate, and the history replay guard. */
export class InteractionAuthorizer {
  policy: InteractionPolicy | undefined;
  private readonly store: SpreadsheetStore;
  private readonly getWb: () => WorkbookHandle;
  private readonly getBounds?: () => Range | undefined;
  private readonly isDisposed: () => boolean;

  constructor(options: InteractionAuthorizerOptions) {
    this.store = options.store;
    this.getWb = options.getWb;
    this.getBounds = options.getBounds;
    this.isDisposed = options.isDisposed;
  }

  canSelect(): PermissionDecision {
    if (this.policy?.selection !== false) return { allowed: true };
    return { allowed: false, code: 'operationDenied', reason: 'selection is disabled' };
  }

  canCopy(): PermissionDecision {
    if (this.policy?.copy !== false) return { allowed: true };
    return { allowed: false, code: 'operationDenied', reason: 'copy is disabled' };
  }

  canExecute(intent: OperationIntent): PermissionDecision {
    if (this.isDisposed())
      return { allowed: false, code: 'invalid', reason: 'controller is disposed' };
    const policy = this.policy;
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
      !isWorkbookAddress(wb, { sheet: range.sheet, row: range.r0, col: range.c0 }) ||
      !isWorkbookAddress(wb, { sheet: range.sheet, row: range.r1, col: range.c1 })
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
        bounds.r1 !== MAX_ROW ||
        bounds.c1 !== MAX_COL)
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
      if (!isWorkbookAddress(wb, addr)) {
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

  /** Authorize a complete intent, including every cell effect and any formula
   * capability implied by a value/paste/fill command. */
  preflightIntent(intent: OperationIntent, canExecute: CanExecute): PermissionDecision {
    if (macSubtotalStructuralOperation(intent) !== null) return canExecute(intent);
    const global = canExecute({ ...intent, effects: [{ kind: 'workbook' }] });
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
      const decision = canExecute({
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
      return this.preflightIntent(
        {
          ...intent,
          operation: 'formulaEdit',
          effects: [{ kind: 'cells', cells: formulaCells, includesFormula: true }],
        },
        canExecute,
      );
    }
    return { allowed: true };
  }

  guardHistory(entry: HistoryEntry, direction: HistoryDirection, canExecute: CanExecute): boolean {
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
          if (!canExecute(intent).allowed) allowed = false;
        } else if (this.policy && !this.preflightIntent(intent, canExecute).allowed) {
          allowed = false;
        }
      }
      return allowed;
    }
    const intent = direction === 'undo' ? entry.inverseIntent : entry.intent;
    if (!intent) return this.policy === undefined;
    if (macSubtotalStructuralOperation(intent) !== null || ownsMacRemoveHyperlinkCommand(intent))
      return canExecute(intent).allowed;
    if (!this.policy) return true;
    return this.preflightIntent(intent, canExecute).allowed;
  }

  private legacyDecision(intent: OperationIntent, wb: WorkbookHandle): PermissionDecision {
    for (const effect of intent.effects) {
      const addresses = effectCells(effect);
      if (addresses === null) return { allowed: false, code: 'unsupported' };
      for (const addr of addresses) {
        if (!isWorkbookAddress(wb, addr)) return { allowed: false, code: 'outOfBounds', addr };
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
}
