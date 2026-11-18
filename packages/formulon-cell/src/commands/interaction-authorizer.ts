import { addrKey } from '../engine/address.js';
import { sheetHasCells } from '../engine/sheet-kind.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { rangeContainsAddr, rangeContainsRange } from '../store/selection-geometry.js';
import type { SpreadsheetStore } from '../store/store.js';
import type { State } from '../store/types.js';
import { AUTO_HYPERLINK_COMMAND_ID } from './auto-hyperlink.js';
import type { HistoryDirection, HistoryEntry } from './history.js';
import {
  bandRange,
  createWorkbookAddressCheck,
  effectCells,
  isWorkbookAddress,
  mergeCellsFor,
  unenumerableEffect,
} from './interaction-effects.js';
import {
  type InteractionOperation,
  type InteractionPolicy,
  isMutatingInteraction,
  isPolicyCellEligible,
  isPolicyRangeEligible,
  needsCells,
  type OperationEffect,
  type OperationIntent,
  operationDecision,
  type PermissionDecision,
} from './interaction-policy.js';
import { deny } from './permission-reasons.js';
import {
  isCellWritable,
  isPartialShift,
  isSheetProtected,
  noteChangePermitted,
  protectedSheetPermissions,
  protectionOption,
  refusedOnProtectedSheet,
  sheetPermits,
  tableChangePermitted,
  unlockedCellsOption,
  workbookStructurePermits,
} from './protection.js';
import { isThreadIntent } from './thread-intent.js';

const automaticHyperlinkCell = (intent: OperationIntent): Addr | null => {
  if (
    intent.commandId !== AUTO_HYPERLINK_COMMAND_ID ||
    (intent.operation !== 'format' && intent.operation !== 'hyperlink') ||
    (intent.origin !== 'editor' && intent.origin !== 'undo' && intent.origin !== 'redo') ||
    intent.effects.length !== 1
  ) {
    return null;
  }
  const effect = intent.effects[0];
  if (effect?.kind !== 'cells' || effect.cells.length !== 1 || effect.includesFormula === true) {
    return null;
  }
  const cell = effect.cells[0];
  return cell &&
    Number.isInteger(cell.sheet) &&
    Number.isInteger(cell.row) &&
    Number.isInteger(cell.col)
    ? cell
    : null;
};

/** Whether sheet protection lets a sheet-level `operation` change `sheet`. */
const sheetProtectionPermits = (
  state: State,
  operation: InteractionOperation,
  sheet: number,
): boolean => {
  // Unprotecting must work on a protected sheet, so protection never refuses itself.
  if (operation === 'protection') return true;
  // Protect Sheet leaves page setup (margins, print area, titles, header/footer) editable.
  if (operation === 'pageSetup') return true;
  // Turning AutoFilter on or off (the sheet-level filter change) is enabled by no Protect Sheet option.
  if (operation === 'filter') return sheetPermits(state, sheet);
  // A table change is decided by its command, before any effect is read.
  if (operation === 'table') return true;
  return sheetPermits(state, sheet, protectionOption(operation) ?? unlockedCellsOption(operation));
};

/** Whether sheet protection lets `operation` write `addr`. */
const protectionPermits = (state: State, operation: InteractionOperation, addr: Addr): boolean => {
  // A table change is decided by its command, whatever the lock of the cells it covers.
  if (operation === 'table') return true;
  // A note goes by Edit objects (or its shape's lock), not by its cell's lock.
  if (operation === 'comment') return noteChangePermitted(state, addr);
  const option = protectionOption(operation);
  if (option) return sheetPermits(state, addr.sheet, option);
  const unlockedOption = unlockedCellsOption(operation);
  if (unlockedOption && !sheetPermits(state, addr.sheet, unlockedOption)) return false;
  return isCellWritable(state, addr);
};

/** The sheets an intent writes; a workbook effect counts as the active sheet. */
const effectSheets = (intent: OperationIntent, activeSheet: number): Set<number> => {
  const sheets = new Set<number>();
  for (const effect of intent.effects) {
    if (effect.kind === 'sheet') sheets.add(effect.sheet);
    else if (effect.kind === 'range') sheets.add(effect.range.sheet);
    else if (effect.kind === 'cells') for (const cell of effect.cells) sheets.add(cell.sheet);
    else sheets.add(activeSheet);
  }
  return sheets;
};

const ownsAutomaticHyperlinkCommand = (intent: OperationIntent): boolean =>
  intent.commandId === AUTO_HYPERLINK_COMMAND_ID;

const hasEditablePredicate = (policy: InteractionPolicy | undefined): boolean => {
  const editable = policy?.editable;
  return (
    typeof editable === 'function' ||
    (editable !== undefined && !Array.isArray(editable) && 'predicate' in editable)
  );
};

type CanExecute = (intent: OperationIntent) => PermissionDecision;
type WorkbookAddressCheck = (addr: Addr) => boolean;

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
    return deny('operationDenied', 'selectionDisabled');
  }

  canCopy(): PermissionDecision {
    if (this.policy?.copy !== false) return { allowed: true };
    return deny('operationDenied', 'copyDisabled');
  }

  canExecute(intent: OperationIntent): PermissionDecision {
    if (this.isDisposed()) return deny('invalid', 'controllerDisposed');
    const policy = this.policy;
    try {
      const wb = this.getWb();
      if (!wb) return deny('invalid', 'workbookUnavailable');
      const cachedAddressCheck = createWorkbookAddressCheck(wb);
      // Host callbacks may change the workbook between per-cell checks; keep reads live.
      const isAddress =
        policy && (this.getBounds !== undefined || hasEditablePredicate(policy))
          ? (addr: Addr) => isWorkbookAddress(wb, addr)
          : cachedAddressCheck;
      if (ownsAutomaticHyperlinkCommand(intent)) {
        const addr = automaticHyperlinkCell(intent);
        if (!addr) {
          return deny('invalid', 'hyperlinkNeedsOneCell');
        }
        if (isSheetProtected(this.store.getState(), addr.sheet)) {
          const permissions = protectedSheetPermissions(this.store.getState(), addr.sheet);
          if (permissions?.formatCells !== true || permissions.insertHyperlinks !== true) {
            return deny('protected', 'hyperlinkNeedsPermission', { addr });
          }
        }
        if (!policy) return this.legacyDecision(intent, isAddress);
        return this.policyDecision(intent, policy, isAddress);
      }
      if (!policy) return this.legacyDecision(intent, isAddress);
      return this.policyDecision(intent, policy, isAddress);
    } catch {
      return deny('invalid', 'workbookUnavailable');
    }
  }

  private policyDecision(
    intent: OperationIntent,
    policy: InteractionPolicy,
    isAddress: WorkbookAddressCheck,
  ): PermissionDecision {
    const operation = operationDecision(policy, intent.operation);
    if (!operation.allowed) return operation;
    const structure = this.structureDecision(intent);
    if (!structure.allowed) return structure;

    const cells: Addr[] = [];
    const formulaCells: Addr[] = [];
    const formulaBands: OperationEffect[] = [];
    for (const effect of intent.effects) {
      if (effect.kind === 'sheet') {
        const decision = this.sheetDecision(intent, effect.sheet, policy, isAddress);
        if (!decision.allowed) return decision;
        continue;
      }
      const band = bandRange(effect);
      if (band) {
        const decision = this.bandDecision(intent, band, policy, isAddress);
        if (!decision.allowed) return decision;
        if (effect.kind !== 'workbook' && effect.includesFormula) formulaBands.push(effect);
        continue;
      }
      const addresses = effectCells(effect);
      if (addresses === null) return unenumerableEffect(effect);
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
        return deny('effectTooLarge', 'mergedRangeTooLarge', { addr: requestedAddr });
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
        return deny('outOfBounds', 'cellOutsideBounds', { addr });
      }
      if (!isAddress(addr)) {
        return deny('outOfBounds', 'cellOutsideWorkbook', { addr });
      }
      if (isMutatingInteraction(intent.operation)) {
        if (!isPolicyCellEligible(policy, addr, intent.operation, intent.origin, this.getWb())) {
          return deny('cellIneligible', 'cellNotEditable', { addr });
        }
        if (!isThreadIntent(intent) && !protectionPermits(state, intent.operation, addr)) {
          return deny('protected', 'cellProtected', { addr });
        }
      }
    }
    if (
      (formulaCells.length > 0 || formulaBands.length > 0) &&
      intent.operation !== 'formulaEdit'
    ) {
      const formulaDecision = this.policyDecision(
        {
          ...intent,
          operation: 'formulaEdit',
          effects: [
            ...(formulaCells.length > 0
              ? [{ kind: 'cells' as const, cells: formulaCells, includesFormula: true }]
              : []),
            ...formulaBands,
          ],
        },
        policy,
        isAddress,
      );
      if (!formulaDecision.allowed) return formulaDecision;
    }
    const allowed: PermissionDecision = { allowed: true };
    return this.applyRestriction(policy, intent, allowed, authorizedCells[0] ?? cells[0]);
  }

  /** Authorize a complete intent, including every cell effect and any formula
   * capability implied by a value/paste/fill command. */
  preflightIntent(intent: OperationIntent, canExecute: CanExecute): PermissionDecision {
    const global = canExecute({ ...intent, effects: [{ kind: 'workbook' }] });
    if (!global.allowed) return global;
    const cells: Addr[] = [];
    const formulaCells: Addr[] = [];
    for (const effect of intent.effects) {
      if (effect.kind === 'sheet' || bandRange(effect)) {
        const decision = canExecute({ ...intent, effects: [effect] });
        if (!decision.allowed) return decision;
        continue;
      }
      const addresses = effectCells(effect);
      if (addresses === null) return unenumerableEffect(effect);
      cells.push(...addresses);
      if ((effect.kind === 'cells' || effect.kind === 'range') && effect.includesFormula)
        formulaCells.push(...addresses);
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
    // Under a policy every intent is preflighted; without one it still meets bounds and protection.
    const replayable = (intent: OperationIntent): boolean =>
      this.policy && !ownsAutomaticHyperlinkCommand(intent)
        ? this.preflightIntent(intent, canExecute).allowed
        : canExecute(intent).allowed;
    const replayAuthorization = entry.replayAuthorization;
    if (replayAuthorization) {
      const intents = direction === 'undo' ? replayAuthorization.undo : replayAuthorization.redo;
      return intents.length > 0 && intents.every(replayable);
    }
    const intent = direction === 'undo' ? entry.inverseIntent : entry.intent;
    if (!intent) return this.policy === undefined;
    return replayable(intent);
  }

  /** Decide a range too large to enumerate (a row/column shift band) as a
   *  whole: bounds, workbook address, editable cells and sheet protection. */
  private bandDecision(
    intent: OperationIntent,
    range: Range,
    policy: InteractionPolicy | undefined,
    isAddress: WorkbookAddressCheck,
  ): PermissionDecision {
    const first: Addr = { sheet: range.sheet, row: range.r0, col: range.c0 };
    const last: Addr = { sheet: range.sheet, row: range.r1, col: range.c1 };
    const bounds = policy ? this.getBounds?.() : undefined;
    if (bounds && !rangeContainsRange(bounds, range)) {
      return deny('outOfBounds', 'rangeOutsideBounds', { addr: first });
    }
    if (!isAddress(first) || !isAddress(last)) {
      return deny('outOfBounds', 'rangeOutsideWorkbook', { addr: first });
    }
    if (!isMutatingInteraction(intent.operation)) return { allowed: true };
    if (policy && !isPolicyRangeEligible(policy, range, this.getWb())) {
      return deny('cellIneligible', 'rangeNotEditable', { addr: first });
    }
    if (
      intent.operation !== 'table' &&
      !sheetPermits(this.store.getState(), range.sheet, protectionOption(intent.operation))
    )
      return deny('protected', 'sheetProtected', { addr: first });
    return { allowed: true };
  }

  /** Decide a sheet-level effect: the sheet must exist and be the bounds' sheet,
   *  and sheet protection must let the operation through. No cell is checked. */
  private sheetDecision(
    intent: OperationIntent,
    sheet: number,
    policy: InteractionPolicy | undefined,
    isAddress: WorkbookAddressCheck,
  ): PermissionDecision {
    const bounds = policy ? this.getBounds?.() : undefined;
    if (bounds && bounds.sheet !== sheet) {
      return deny('outOfBounds', 'sheetOutsideBounds');
    }
    if (!isAddress({ sheet, row: 0, col: 0 })) {
      return deny('outOfBounds', 'sheetNotInWorkbook');
    }
    if (!isMutatingInteraction(intent.operation)) return { allowed: true };
    if (!sheetProtectionPermits(this.store.getState(), intent.operation, sheet))
      return deny('protected', 'sheetProtected');
    return { allowed: true };
  }

  /** Protection's decision before any effect is read: workbook-structure protection governs the
   *  sheet operations, a protected sheet refuses some commands and every cell shift whatever its
   *  options, and a table change there is decided by its command alone. */
  private structureDecision(intent: OperationIntent): PermissionDecision {
    const state = this.store.getState();
    if (isThreadIntent(intent)) return { allowed: true };
    if (!workbookStructurePermits(state, intent.operation))
      return deny('protected', 'workbookStructureProtected');
    for (const sheet of effectSheets(intent, state.data.sheetIndex)) {
      if (needsCells(intent.operation) && !sheetHasCells(state, sheet))
        return deny('operationDenied', 'sheetHasNoCells');
      if (
        refusedOnProtectedSheet(state, sheet, intent.commandId) ||
        (intent.operation === 'table' && !tableChangePermitted(state, sheet))
      )
        return deny('protected', 'sheetProtected');
    }
    for (const effect of intent.effects) {
      if (
        effect.kind === 'range' &&
        isPartialShift(intent.operation, effect.range) &&
        isSheetProtected(state, effect.range.sheet)
      )
        return deny('protected', 'sheetProtected');
    }
    return { allowed: true };
  }

  private legacyDecision(
    intent: OperationIntent,
    isAddress: WorkbookAddressCheck,
  ): PermissionDecision {
    const structure = this.structureDecision(intent);
    if (!structure.allowed) return structure;
    for (const effect of intent.effects) {
      if (effect.kind === 'sheet') {
        const decision = this.sheetDecision(intent, effect.sheet, undefined, isAddress);
        if (!decision.allowed) return decision;
        continue;
      }
      const band = bandRange(effect);
      if (band) {
        const decision = this.bandDecision(intent, band, undefined, isAddress);
        if (!decision.allowed) return decision;
        continue;
      }
      // No host callback runs per cell without a policy, so a cell list of any length is checked.
      const addresses = effect.kind === 'cells' ? effect.cells : effectCells(effect);
      if (addresses === null) return unenumerableEffect(effect);
      for (const addr of addresses) {
        if (!isAddress(addr)) return deny('outOfBounds', 'cellOutsideWorkbook', { addr });
        if (
          isMutatingInteraction(intent.operation) &&
          !isThreadIntent(intent) &&
          !protectionPermits(this.store.getState(), intent.operation, addr)
        ) {
          return deny('protected', 'cellProtected', { addr });
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
        return deny('invalid', 'restrictionHookAsync');
      }
      if (restricted === false) return deny('operationDenied', 'restrictedByHost');
      if (restricted && typeof restricted === 'object' && restricted.allowed === false) {
        // A key survives only on the refusal it describes; a host rewording drops it.
        const same = !decision.allowed && restricted.reason === decision.reason;
        const { reasonKey: _hostKey, ...own } = restricted;
        return !decision.allowed && decision.reasonKey && same
          ? { ...own, reasonKey: decision.reasonKey }
          : own;
      }
    } catch (error) {
      return error instanceof Error
        ? { allowed: false, code: 'invalid', reason: error.message }
        : deny('invalid', 'restrictionCallbackFailed');
    }
    return decision;
  }
}
