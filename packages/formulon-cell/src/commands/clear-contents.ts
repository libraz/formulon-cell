import { addrKey } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { selectionContainsAddr, selectionCoversRange } from '../store/selection-geometry.js';
import type { SpreadsheetStore, State } from '../store/store.js';
import { deleteDataTables } from './data-table.js';
import {
  alertDataTableEdit,
  dataTableGuardAlert,
  planDataTableSelection,
} from './data-table-guard.js';
import { rowHiddenByFilter } from './filter-visible.js';
import { History, recordingHistory, transactIn } from './history.js';
import { InteractionController, interactionControllerFor } from './interaction-controller.js';
import {
  appliedBatch,
  boundingRange,
  type ChangeBatchResult,
  type InteractionOrigin,
  operationPermission,
  refusedBatch,
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
    // Rows an AutoFilter hides keep their contents, as in the desktop app.
    if (rowHiddenByFilter(state, primarySheet, cell.addr.row)) continue;

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

/** The rectangle around the selected ranges, or null when they span sheets. */
export const selectionBounds = (state: State): Range | null => {
  const { range, extraRanges } = state.selection;
  return boundingRange(
    [range, ...(extraRanges ?? [])].flatMap((r) => [
      { sheet: r.sheet, row: r.r0, col: r.c0 },
      { sheet: r.sheet, row: r.r1, col: r.c1 },
    ]),
  );
};

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
  const mayDeleteObjects =
    !commands?.policy ||
    (operationPermission(commands.policy, 'clear') && !commands.policy.editable);
  const rejectPivotEdit = (): ChangeBatchResult => {
    alertPivotEdit(workbook);
    return refusedBatch({ code: 'protected', reason: pivotGuardAlert(workbook).message });
  };
  if (pivotPlan.partial || (pivotPlan.whole.length > 0 && !mayDeleteObjects)) {
    return rejectPivotEdit();
  }
  const dataTablePlan = planDataTableSelection(workbook, store.getState().selection);
  if (dataTablePlan.partial || (dataTablePlan.whole.length > 0 && !mayDeleteObjects)) {
    alertDataTableEdit(workbook);
    return refusedBatch({ code: 'protected', reason: dataTableGuardAlert(workbook).message });
  }
  const removesObjects = pivotPlan.whole.length > 0 || dataTablePlan.whole.length > 0;
  // Deleting a whole report or data table and clearing the rest of the
  // selection undo together.
  const history = recordingHistory(options.history);
  const cleared = transactIn(
    removesObjects ? history : null,
    (): ChangeBatchResult => {
      if (pivotPlan.whole.length > 0) {
        const removed = deletePivotTables(
          {
            store,
            workbook,
            history,
            origin: options.origin ?? 'keyboard',
            commandId: options.commandId ?? 'clear-contents',
          },
          store.getState().selection.range.sheet,
          pivotPlan.whole,
        );
        if (!removed) return rejectPivotEdit();
      }
      if (dataTablePlan.whole.length > 0) {
        deleteDataTables({ store, workbook, history }, dataTablePlan.whole);
      }
      return clearCollectedContents(options, commands);
    },
    { keep: (result) => result.status !== 'refused' },
  );
  // The clear targets the selection, whichever of its cells held content.
  return cleared.status === 'applied'
    ? { ...cleared, range: selectionBounds(store.getState()) }
    : cleared;
};

const clearCollectedContents = (
  options: ClearSelectedContentsOptions,
  commands: ClearContentsCommands | undefined,
): ChangeBatchResult => {
  const { store, workbook } = options;
  const addresses = collectSelectedContentAddresses(store, workbook);
  if (addresses.length === 0) return appliedBatch([], null);

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
    });
  } finally {
    ephemeral?.dispose();
  }
};
