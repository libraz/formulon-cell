import { addrKey } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { rangeContainsAddr, rangeContainsRange } from '../store/selection-geometry.js';
import type { SpreadsheetStore } from '../store/store.js';
import type { SheetProtectionPermissions, State } from '../store/types.js';
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
  type OperationEffect,
  type OperationIntent,
  operationDecision,
  type PermissionDecision,
} from './interaction-policy.js';
import {
  isCellWritable,
  isSheetProtected,
  protectedSheetPermissions,
  sheetPermits,
  unlockedCellsOption,
} from './protection.js';

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

/** Operations a Protect Sheet option lets through on a protected sheet, locked cells included.
 *  Every other mutation needs an unlocked cell. */
const PROTECTION_OPTION: Partial<Record<InteractionOperation, keyof SheetProtectionPermissions>> = {
  insertRows: 'insertRows',
  deleteRows: 'deleteRows',
  insertColumns: 'insertColumns',
  deleteColumns: 'deleteColumns',
  pivot: 'pivotTables',
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
  return sheetPermits(state, sheet, PROTECTION_OPTION[operation] ?? unlockedCellsOption(operation));
};

/** Whether sheet protection lets `operation` write `addr`. */
const protectionPermits = (state: State, operation: InteractionOperation, addr: Addr): boolean => {
  const option = PROTECTION_OPTION[operation];
  if (option) return sheetPermits(state, addr.sheet, option);
  const unlockedOption = unlockedCellsOption(operation);
  if (unlockedOption && !sheetPermits(state, addr.sheet, unlockedOption)) return false;
  return isCellWritable(state, addr);
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
      const cachedAddressCheck = createWorkbookAddressCheck(wb);
      // Host callbacks may change the workbook between per-cell checks; keep reads live.
      const isAddress =
        policy && (this.getBounds !== undefined || hasEditablePredicate(policy))
          ? (addr: Addr) => isWorkbookAddress(wb, addr)
          : cachedAddressCheck;
      if (ownsAutomaticHyperlinkCommand(intent)) {
        const addr = automaticHyperlinkCell(intent);
        if (!addr) {
          return {
            allowed: false,
            code: 'invalid',
            reason: 'Automatic hyperlink authorization needs one cell and no formula effect',
          };
        }
        if (isSheetProtected(this.store.getState(), addr.sheet)) {
          const permissions = protectedSheetPermissions(this.store.getState(), addr.sheet);
          if (permissions?.formatCells !== true || permissions.insertHyperlinks !== true) {
            return {
              allowed: false,
              code: 'protected',
              addr,
              reason: 'Automatic hyperlink needs protected-sheet format and hyperlink permission',
            };
          }
        }
        if (!policy) return this.legacyDecision(intent, isAddress);
        return this.policyDecision(intent, policy, isAddress);
      }
      if (!policy) return this.legacyDecision(intent, isAddress);
      return this.policyDecision(intent, policy, isAddress);
    } catch {
      return { allowed: false, code: 'invalid', reason: 'workbook is unavailable' };
    }
  }

  private policyDecision(
    intent: OperationIntent,
    policy: InteractionPolicy,
    isAddress: WorkbookAddressCheck,
  ): PermissionDecision {
    const operation = operationDecision(policy, intent.operation);
    if (!operation.allowed) return operation;

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
        return {
          allowed: false,
          code: 'effectTooLarge',
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
      if (!isAddress(addr)) {
        return {
          allowed: false,
          code: 'outOfBounds',
          addr,
          reason: 'cell address is outside the workbook',
        };
      }
      if (isMutatingInteraction(intent.operation)) {
        if (!isPolicyCellEligible(policy, addr, intent.operation, intent.origin, this.getWb())) {
          return {
            allowed: false,
            code: 'cellIneligible',
            addr,
            reason: 'cell is outside editable cells',
          };
        }
        if (!protectionPermits(state, intent.operation, addr)) {
          return { allowed: false, code: 'protected', addr, reason: 'cell is protected' };
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
      return {
        allowed: false,
        code: 'outOfBounds',
        addr: first,
        reason: 'range is outside the configured bounds',
      };
    }
    if (!isAddress(first) || !isAddress(last)) {
      return {
        allowed: false,
        code: 'outOfBounds',
        addr: first,
        reason: 'range is outside the workbook',
      };
    }
    if (!isMutatingInteraction(intent.operation)) return { allowed: true };
    if (policy && !isPolicyRangeEligible(policy, range, this.getWb())) {
      return {
        allowed: false,
        code: 'cellIneligible',
        addr: first,
        reason: 'range is outside editable cells',
      };
    }
    if (!sheetPermits(this.store.getState(), range.sheet, PROTECTION_OPTION[intent.operation]))
      return { allowed: false, code: 'protected', addr: first, reason: 'sheet is protected' };
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
      return {
        allowed: false,
        code: 'outOfBounds',
        reason: 'sheet is outside the configured bounds',
      };
    }
    if (!isAddress({ sheet, row: 0, col: 0 })) {
      return { allowed: false, code: 'outOfBounds', reason: 'sheet is not in the workbook' };
    }
    if (!isMutatingInteraction(intent.operation)) return { allowed: true };
    if (!sheetProtectionPermits(this.store.getState(), intent.operation, sheet))
      return { allowed: false, code: 'protected', reason: 'sheet is protected' };
    return { allowed: true };
  }

  private legacyDecision(
    intent: OperationIntent,
    isAddress: WorkbookAddressCheck,
  ): PermissionDecision {
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
        if (!isAddress(addr)) return { allowed: false, code: 'outOfBounds', addr };
        if (
          isMutatingInteraction(intent.operation) &&
          !protectionPermits(this.store.getState(), intent.operation, addr)
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
