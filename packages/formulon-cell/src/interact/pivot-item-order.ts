import type { History } from '../commands/history.js';
import { refreshPivotReportsAfterEdit } from '../commands/pivot-table.js';
import { addrKey } from '../engine/address.js';
import { PivotAxis, PivotSortOrder } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore, type State } from '../store/store.js';
import { commitPivotChange } from './pivot-field-settings-dialog-apply.js';

export interface PivotItemOrderDeps {
  store: SpreadsheetStore;
  workbook: WorkbookHandle;
  history?: History | null;
}

/** The item rows of a report with one row field, as the grid shows them. */
export interface PivotItemSlots {
  sheet: number;
  pivotIndex: number;
  /** Pivot field index of the row field. */
  fieldIdx: number;
  /** Column holding the item labels, and the report's last column. */
  col: number;
  c1: number;
  firstRow: number;
  lastRow: number;
}

/** The item rows of the report whose item label sits at `(row, col)`, or null
 *  when that cell is not an item of a single-row-field, no-column-field report. */
export function pivotItemSlotsAt(
  state: State,
  wb: WorkbookHandle,
  sheet: number,
  row: number,
  col: number,
): PivotItemSlots | null {
  if (!wb.capabilities.pivotTables) return null;
  for (const { pivotIndex, range } of wb.pivotBounds(sheet)) {
    if (col !== range.c0) continue;
    const definition = wb.getPivotDefinition(sheet, pivotIndex);
    const fieldIdx = definition?.rowFieldOrder[0];
    if (!definition || fieldIdx === undefined) continue;
    if (definition.rowFieldOrder.length !== 1 || definition.colFieldOrder.length > 0) continue;
    const filterRows = definition.fields.filter((f) => f.axis === PivotAxis.Page).length;
    const firstRow = range.r0 + (filterRows > 0 ? filterRows + 1 : 0) + 1;
    const lastRow = range.r1 - (definition.grandTotals.cols ? 1 : 0);
    if (row < firstRow || row > lastRow) continue;
    if (definition.fields[fieldIdx]?.items.some((item) => !item.visible)) return null;
    if (state.data.sheetIndex !== sheet) return null;
    return { sheet, pivotIndex, fieldIdx, col, c1: range.c1, firstRow, lastRow };
  }
  return null;
}

/** Where the item on `fromRow` ends up when it is released on `afterRow`: right
 *  below that row, with the header row above the first item meaning the top. */
export function pivotItemLandingRow(slots: PivotItemSlots, fromRow: number, afterRow: number) {
  const after = Math.min(Math.max(afterRow, slots.firstRow - 1), slots.lastRow);
  return after < fromRow ? after + 1 : after;
}

/** Move the item on `fromRow` to just below `afterRow` as one undo step by
 *  switching the field to a manual order. Returns false when nothing moved. */
export function movePivotItem(
  deps: PivotItemOrderDeps,
  slots: PivotItemSlots,
  fromRow: number,
  afterRow: number,
): boolean {
  const { store, workbook: wb } = deps;
  if (fromRow < slots.firstRow || fromRow > slots.lastRow) return false;
  const landing = pivotItemLandingRow(slots, fromRow, afterRow);
  if (landing === fromRow) return false;
  const cells = store.getState().data.cells;
  const labels: string[] = [];
  for (let row = slots.firstRow; row <= slots.lastRow; row += 1) {
    const value = cells.get(addrKey({ sheet: slots.sheet, row, col: slots.col }))?.value;
    if (value?.kind === 'text' && value.value !== '') labels.push(value.value);
    else if (value?.kind === 'number') labels.push(String(value.value));
    else return false;
  }
  const [moved] = labels.splice(fromRow - slots.firstRow, 1);
  if (moved === undefined) return false;
  labels.splice(landing - slots.firstRow, 0, moved);

  const ctx = {
    wb,
    history: deps.history ?? null,
    store,
    origin: 'editor' as const,
    commandId: 'pivotMoveItem',
    onChanged: () => {
      refreshPivotReportsAfterEdit(store, wb);
    },
  };
  const ok = commitPivotChange(
    ctx,
    slots.sheet,
    slots.pivotIndex,
    () =>
      wb.setPivotFieldSort(slots.sheet, slots.pivotIndex, slots.fieldIdx, PivotSortOrder.Manual) &&
      wb.clearPivotFieldItems(slots.sheet, slots.pivotIndex, slots.fieldIdx) &&
      labels.every((label) =>
        wb.addPivotFieldItem(slots.sheet, slots.pivotIndex, slots.fieldIdx, label, true),
      ),
  );
  if (!ok) return false;
  mutators.setActive(store, { sheet: slots.sheet, row: slots.firstRow, col: slots.col });
  mutators.setRange(store, {
    sheet: slots.sheet,
    r0: slots.firstRow,
    c0: slots.col,
    r1: slots.lastRow,
    c1: slots.col,
  });
  return true;
}
