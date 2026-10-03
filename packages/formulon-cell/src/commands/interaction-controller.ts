import { addrKey } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { formatWithPending } from '../store/pending-format.js';
import { rangeContainsAddr } from '../store/selection-geometry.js';
import type { SpreadsheetStore } from '../store/store.js';
import type { CellFormat, State } from '../store/types.js';
import { type CoercedInput, coerceInputForCell, validateCoercedInput } from './coerce-input.js';
import type { History } from './history.js';
import { InteractionAuthorizer } from './interaction-authorizer.js';
import {
  cellEffects,
  isWorkbookAddress,
  mergeAnchorFor,
  mergeCellsFor,
} from './interaction-effects.js';
import type {
  BatchRejection,
  CellBatchCommand,
  CellBatchOperation,
  CellChangeInput,
  ChangeBatchResult,
  InteractionOrigin,
  InteractionPolicy,
  OperationIntent,
  PermissionDecision,
} from './interaction-policy.js';
import {
  changedFormatAddresses,
  effectivePendingFormatAddresses,
  type PreparedChange,
  projectPreparedFormats,
  sameStructuredValue,
  snapshotToPatch,
  unionAddresses,
} from './prepared-change.js';
import { normalizeR1C1Formula } from './refs.js';
import { coerceCellValue } from './validate.js';

const CELL_OPERATIONS: ReadonlySet<CellBatchOperation> = new Set([
  'valueEdit',
  'formulaEdit',
  'clear',
  'paste',
  'fill',
  'moveCells',
]);

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

export class InteractionController {
  private readonly authorizer: InteractionAuthorizer;
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
    this.authorizer = new InteractionAuthorizer({
      store: options.store,
      getWb: options.getWb,
      getBounds: options.getBounds,
      isDisposed: () => this.disposed,
    });
    if (this.manageHistoryGuard) this.installHistoryGuard();
  }

  get policy(): InteractionPolicy | undefined {
    return this.authorizer.policy;
  }

  get restricted(): boolean {
    return this.authorizer.policy !== undefined;
  }

  get revision(): number {
    return this.revisionValue;
  }

  canSelect(): PermissionDecision {
    return this.authorizer.canSelect();
  }

  canCopy(): PermissionDecision {
    return this.authorizer.canCopy();
  }

  canExecute(intent: OperationIntent): PermissionDecision {
    return this.authorizer.canExecute(intent);
  }

  setPolicy(next?: InteractionPolicy): void {
    this.authorizer.policy = next;
    if (this.manageHistoryGuard) this.installHistoryGuard();
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

    const denied = command.denied ?? this.authorizer.policy?.batchDenied ?? 'reject';
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
      this.authorizer.policy === undefined && command.operation === 'clear',
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
    if (!isWorkbookAddress(wb, addr)) {
      return { rejection: { addr, code: 'outOfBounds', reason: 'cell is outside the workbook' } };
    }
    const merged = mergeCellsFor(state, addr);
    if (merged === null) {
      return {
        rejection: { addr, code: 'unsupported', reason: 'merged cell range exceeds the limit' },
      };
    }
    const target = mergeAnchorFor(state, addr);
    if (!isWorkbookAddress(wb, target)) {
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
        const outcome = validateCoercedInput(wb, target, input, validation);
        if (!outcome.ok && outcome.severity === 'stop') {
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
    if (formula === null && validation && change.value.kind !== 'error') {
      const outcome = validateCoercedInput(wb, target, coerceCellValue(change.value), validation);
      if (!outcome.ok && outcome.severity === 'stop') {
        return { rejection: { addr, code: 'invalid', reason: outcome.message } };
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
      const forward = this.authorizer.preflightIntent(
        requireInverseAuthorization ? redo : intent,
        this.routeCanExecute,
      );
      if (!forward.allowed) return { decision: forward };
    }
    if (requireInverseAuthorization) {
      const inverse = this.authorizer.preflightIntent(undo, this.routeCanExecute);
      if (!inverse.allowed) return { decision: inverse };
    }
    return { plan: { redo, undo, replayAuthorization } };
  }

  // Authorizer-internal checks re-enter through the controller so a host that
  // overrides `canExecute` on the instance observes every authorization.
  private readonly routeCanExecute = (intent: OperationIntent): PermissionDecision =>
    this.canExecute(intent);

  private installHistoryGuard(): void {
    this.history.setGuard((entry, direction) =>
      this.authorizer.guardHistory(entry, direction, this.routeCanExecute),
    );
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
