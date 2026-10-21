import { addrKey } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { selectionContainsAddr, selectionCoversRange } from '../store/selection-geometry.js';
import type { SpreadsheetStore, State } from '../store/store.js';
import { History } from './history.js';
import { InteractionController, interactionControllerFor } from './interaction-controller.js';
import {
  type ChangeBatchResult,
  type InteractionOrigin,
  operationPermission,
} from './interaction-policy.js';
import { alertPivotEdit, pivotGuardAlert, planPivotSelection } from './pivot-guard.js';
import { deletePivotTables } from './pivot-table.js';

/** The part of the interaction controller used by clear-contents callers. */
export type ClearContentsCommands = Pick<InteractionController, 'execute' | 'policy'> & {
  readonly revision?: number;
};

export interface ClearSelectedContentsOptions {
  readonly store: SpreadsheetStore;
  readonly workbook: WorkbookHandle;
  readonly history?: History | null;
  readonly commands?: ClearContentsCommands;
  readonly origin?: InteractionOrigin;
  readonly commandId?: string;
}

const isContentCell = (cell: { value: { kind: string }; formula: string | null }): boolean =>
  cell.formula !== null || cell.value.kind !== 'blank';

const rangeForMergeAt = (state: State, addr: Addr): Range | null => {
  const key = addrKey(addr);
  const anchorKey = state.merges.byCell.get(key) ?? key;
  return state.merges.byAnchor.get(anchorKey) ?? null;
};

/**
 * Collect the physical, non-blank cells addressed by the complete selection.
 *
 * This deliberately walks the workbook's sparse physical index once. It does
 * not materialize a selected rectangle, which keeps full-column and full-sheet
 * selections bounded by the number of cells that can actually be cleared.
 */
export const collectSelectedContentAddresses = (
  store: SpreadsheetStore,
  workbook: WorkbookHandle,
): Addr[] => {
  const state = store.getState();
  const selection = state.selection;
  const primarySheet = selection.range.sheet;
  const addresses: Addr[] = [];
  const seen = new Set<string>();

  for (const cell of workbook.physicalCells(primarySheet)) {
    if (!isContentCell(cell)) continue;

    const merge = rangeForMergeAt(state, cell.addr);
    let selected = false;
    let target = cell.addr;
    if (merge) {
      // A merged cell is one logical value. Clearing it is safe only when the
      // union covers the complete merge rectangle; a hole must preserve it.
      selected = selectionCoversRange(selection, merge);
      target = { sheet: merge.sheet, row: merge.r0, col: merge.c0 };
    } else {
      selected = selectionContainsAddr(selection, cell.addr);
    }
    if (!selected) continue;

    const key = addrKey(target);
    if (seen.has(key)) continue;
    seen.add(key);
    addresses.push(target);
  }
  return addresses;
};

const emptyResult = (revision = 0): ChangeBatchResult => ({
  status: 'noop',
  applied: [],
  rejected: [],
  revision,
});

/**
 * Clear the selected contents through the mounted interaction controller.
 * When called outside a mounted instance, the helper creates an ephemeral
 * controller so the resulting clear still has one atomic apply/undo entry
 * without replacing the mounted history guard.
 */
export const clearSelectedContents = (options: ClearSelectedContentsOptions): ChangeBatchResult => {
  const { store, workbook } = options;
  const commands = options.commands ?? interactionControllerFor(store);
  const pivotPlan = planPivotSelection(workbook, store.getState().selection);
  // A host policy that restricts which cells are editable cannot vouch for a
  // whole report, so only an unrestricted clear may remove one.
  const mayDeletePivots =
    !commands?.policy ||
    (operationPermission(commands.policy, 'clear') && !commands.policy.editable);
  const rejectPivotEdit = (): ChangeBatchResult => {
    alertPivotEdit(workbook);
    return {
      status: 'rejected',
      applied: [],
      rejected: [{ code: 'protected', reason: pivotGuardAlert(workbook).message }],
      revision: commands?.revision ?? 0,
    };
  };
  if (pivotPlan.partial || (pivotPlan.whole.length > 0 && !mayDeletePivots)) {
    return rejectPivotEdit();
  }
  // Deleting a whole report and clearing the rest of the selection undo together.
  const history = options.history && !options.history.isReplaying() ? options.history : null;
  const txn = pivotPlan.whole.length > 0 ? history?.begin() : undefined;
  if (pivotPlan.whole.length > 0) {
    const removed = deletePivotTables(
      { store, workbook, history },
      store.getState().selection.range.sheet,
      pivotPlan.whole,
    );
    if (!removed) {
      if (history && txn) history.end(txn);
      return rejectPivotEdit();
    }
  }
  const cleared = clearCollectedContents(options, commands);
  if (history && txn) {
    if (cleared.status === 'rejected') history.abort(txn);
    else history.end(txn);
  }
  return cleared.status === 'noop' && pivotPlan.whole.length > 0
    ? { ...cleared, status: 'applied' }
    : cleared;
};

const clearCollectedContents = (
  options: ClearSelectedContentsOptions,
  commands: ClearContentsCommands | undefined,
): ChangeBatchResult => {
  const { store, workbook } = options;
  const addresses = collectSelectedContentAddresses(store, workbook);
  if (addresses.length === 0) return emptyResult(commands?.revision ?? 0);

  let resolvedCommands: ClearContentsCommands | undefined = commands;
  let ephemeral: InteractionController | undefined;
  if (!resolvedCommands) {
    ephemeral = new InteractionController({
      store,
      getWb: () => workbook,
      history: options.history ?? new History(),
      manageHistoryGuard: false,
    });
    resolvedCommands = ephemeral;
  }

  try {
    return resolvedCommands.execute({
      type: 'cellBatch',
      operation: 'clear',
      origin: options.origin ?? 'keyboard',
      commandId: options.commandId ?? 'clear-contents',
      changes: addresses.map((addr) => ({ addr, value: { kind: 'blank' as const } })),
      // A host policy is an atomic authorization boundary. Legacy/no-policy
      // calls preserve the historical writable-only partial clear behavior.
      denied: resolvedCommands.policy ? 'reject' : 'skipIneligible',
    });
  } finally {
    ephemeral?.dispose();
  }
};
