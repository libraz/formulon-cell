import { PivotAxis, type PivotReportLayout } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { clearPivotFieldFilter, pivotFieldFilterModel } from './pivot-filter.js';

/**
 * Report-wide PivotTable layout edits behind the Design and Analyze tabs.
 * Each function writes the engine and reports success; callers wrap it in
 * one undo step (`commitPivotChange`) so the report re-lays out once.
 */

/** A report addressed on its sheet. */
export interface PivotRef {
  sheetIndex: number;
  pivotIndex: number;
}

/** What the Design and Analyze menus show as checked or available. */
export interface PivotLayoutState {
  grandTotals: { rows: boolean; cols: boolean };
  layout: PivotReportLayout | null;
  /** Subtotals at the top of their groups for every row and column field;
   *  null when the fields disagree or there are none. */
  subtotalTop: boolean | null;
  /** Some field is on an axis or in Values. */
  hasFields: boolean;
  /** Some field has hidden items or a label / value filter. */
  hasFilter: boolean;
}

const axisFields = (wb: WorkbookHandle, ref: PivotRef) =>
  (wb.getPivotDefinition(ref.sheetIndex, ref.pivotIndex)?.fields ?? []).filter(
    (field) => field.axis === PivotAxis.Row || field.axis === PivotAxis.Col,
  );

export function pivotLayoutState(wb: WorkbookHandle, ref: PivotRef): PivotLayoutState | null {
  const definition = wb.getPivotDefinition(ref.sheetIndex, ref.pivotIndex);
  if (!definition) return null;
  const onAxes = axisFields(wb, ref);
  const tops = new Set(onAxes.map((field) => field.subtotalTop));
  return {
    grandTotals: { ...definition.grandTotals },
    layout: definition.layout,
    subtotalTop: tops.size === 1 ? [...tops][0] === true : null,
    hasFields:
      definition.dataFields.length > 0 || definition.fields.some((field) => field.axis !== null),
    hasFilter: definition.fields.some(
      (field) =>
        pivotFieldFilterModel(wb, ref.sheetIndex, ref.pivotIndex, field.index)?.filtered === true,
    ),
  };
}

export const applyPivotGrandTotals = (
  wb: WorkbookHandle,
  ref: PivotRef,
  totals: { rows: boolean; cols: boolean },
): boolean => wb.setPivotTableGrandTotals(ref.sheetIndex, ref.pivotIndex, totals.rows, totals.cols);

export const applyPivotReportLayout = (
  wb: WorkbookHandle,
  ref: PivotRef,
  layout: PivotReportLayout,
): boolean => wb.setPivotReportLayout(ref.sheetIndex, ref.pivotIndex, layout);

/** Subtotals at the top or the bottom of their groups, for every row and column field. */
export const applyPivotSubtotalPosition = (
  wb: WorkbookHandle,
  ref: PivotRef,
  top: boolean,
): boolean =>
  axisFields(wb, ref).every((field) =>
    wb.setPivotFieldSubtotalTop(ref.sheetIndex, ref.pivotIndex, field.index, top),
  );

/** Clear All: every field leaves the report, which stays as an empty placeholder. */
export function clearPivotFields(wb: WorkbookHandle, ref: PivotRef): boolean {
  const { sheetIndex, pivotIndex } = ref;
  const definition = wb.getPivotDefinition(sheetIndex, pivotIndex);
  if (!definition) return false;
  let ok =
    wb.clearPivotFilters(sheetIndex, pivotIndex) && wb.clearPivotDataFields(sheetIndex, pivotIndex);
  for (const field of definition.fields) {
    if (field.axis !== null)
      ok = wb.setPivotFieldAxis(sheetIndex, pivotIndex, field.index, null) && ok;
  }
  return ok;
}

/** Clear Filters: every field shows all its items, with no label or value filter. */
export function clearPivotFilters(wb: WorkbookHandle, ref: PivotRef): boolean {
  const { sheetIndex, pivotIndex } = ref;
  const definition = wb.getPivotDefinition(sheetIndex, pivotIndex);
  if (!definition) return false;
  return definition.fields.every(
    (field) =>
      pivotFieldFilterModel(wb, sheetIndex, pivotIndex, field.index)?.filtered !== true ||
      clearPivotFieldFilter(wb, sheetIndex, pivotIndex, field.index),
  );
}
