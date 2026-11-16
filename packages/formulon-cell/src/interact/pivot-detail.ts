import { pivotFieldFilterModel } from '../commands/pivot-filter.js';
import { PIVOT_CELL_KIND, pivotItemText } from '../engine/pivot-cell-kind.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { pivotCellEntryAt } from './pivot-cell-at.js';
import { commitPivotChange, type PivotCommitContext } from './pivot-field-settings-dialog-apply.js';

/** A row label of a report: its field, its item and the field's place among the row fields. */
export interface PivotRowItemRef {
  sheet: number;
  pivotIndex: number;
  /** Pivot field index of the label's field. */
  fieldIdx: number;
  /** Index of the label's item in that field; -1 when the label names no stored item. */
  itemIdx: number;
  /** The field's position in the row fields; the last one is the innermost. */
  level: number;
  rowFieldOrder: readonly number[];
  /** The item is expanded. */
  showDetail: boolean;
}

/** The row label of a report at `addr`, or null when the cell is not one. */
export function pivotRowItemAt(wb: WorkbookHandle, addr: Addr): PivotRowItemRef | null {
  if (!wb.capabilities.pivotDefinitionRead) return null;
  const entry = pivotCellEntryAt(wb, addr);
  if (!entry) return null;
  if (entry.kind !== PIVOT_CELL_KIND.RowLabel && entry.kind !== PIVOT_CELL_KIND.RowSubtotal) {
    return null;
  }
  const definition = wb.getPivotDefinition(addr.sheet, entry.pivotIndex);
  if (!definition) return null;
  const { rowFieldOrder } = definition;
  const level = rowFieldOrder.findIndex(
    (i) => definition.fields[i]?.sourceName === entry.fieldName,
  );
  const fieldIdx = rowFieldOrder[level];
  const field = fieldIdx === undefined ? undefined : definition.fields[fieldIdx];
  if (fieldIdx === undefined || !field) return null;
  const text = pivotItemText(entry.value);
  const model =
    field.items.length > 0
      ? pivotFieldFilterModel(wb, addr.sheet, entry.pivotIndex, fieldIdx)
      : null;
  const itemIdx = model ? model.items.findIndex((item) => item.value === text) : -1;
  return {
    sheet: addr.sheet,
    pivotIndex: entry.pivotIndex,
    fieldIdx,
    itemIdx,
    level,
    rowFieldOrder,
    showDetail: field.items[itemIdx]?.showDetail ?? true,
  };
}

/** The label is on the innermost row field, so there is no detail under it to show. */
export const isInnermostRowItem = (ref: PivotRowItemRef): boolean =>
  ref.level === ref.rowFieldOrder.length - 1;

/** Expands a collapsed item or collapses an expanded one, as its button does; one undo step. */
export function togglePivotItemDetail(ctx: PivotCommitContext, ref: PivotRowItemRef): boolean {
  if (ref.itemIdx < 0) return false;
  return commitPivotChange(ctx, ref.sheet, ref.pivotIndex, () =>
    ctx.wb.setPivotFieldItemShowDetail(
      ref.sheet,
      ref.pivotIndex,
      ref.fieldIdx,
      ref.itemIdx,
      !ref.showDetail,
    ),
  );
}

/** Expands every item of the label's field; one undo step. The innermost field has nothing
 *  to expand into: the caller offers the Show Detail picker for it instead. */
export function expandPivotField(ctx: PivotCommitContext, ref: PivotRowItemRef): boolean {
  if (isInnermostRowItem(ref)) return false;
  return commitPivotChange(ctx, ref.sheet, ref.pivotIndex, () =>
    ctx.wb.setPivotFieldShowDetail(ref.sheet, ref.pivotIndex, ref.fieldIdx, true),
  );
}

/** Collapses every item of the label's field, or of the field above it when the label is on
 *  the innermost field; one undo step. A report with one row field has nothing to collapse. */
export function collapsePivotField(ctx: PivotCommitContext, ref: PivotRowItemRef): boolean {
  const level = isInnermostRowItem(ref) ? ref.level - 1 : ref.level;
  const fieldIdx = ref.rowFieldOrder[level];
  if (fieldIdx === undefined) return false;
  return commitPivotChange(ctx, ref.sheet, ref.pivotIndex, () =>
    ctx.wb.setPivotFieldShowDetail(ref.sheet, ref.pivotIndex, fieldIdx, false),
  );
}
