import type { PivotDefinition } from '../engine/pivot-definition.js';
import { PivotAxis } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';

/** Rebuild a report on its original pivot cache from a stored definition.
 *  Pivot field `i` binds to cache field `i`, so every field is re-added in
 *  order; a field that sat on no axis is taken back off its axis. The engine
 *  cannot express a field with neither the automatic subtotal nor custom ones,
 *  so that state returns as automatic. Returns the new pivot index, or -1 on
 *  failure. */
export function recreatePivotTable(
  wb: WorkbookHandle,
  sheet: number,
  anchor: { row: number; col: number },
  definition: PivotDefinition,
): number {
  const pivotIndex = wb.createPivotTable(sheet, definition.name, definition.cacheId, anchor);
  if (pivotIndex < 0) return -1;
  const rebuild = (): boolean => {
    const { grandTotals, layout } = definition;
    if (!wb.setPivotTableGrandTotals(sheet, pivotIndex, grandTotals.rows, grandTotals.cols)) {
      return false;
    }
    if (layout !== null && !wb.setPivotReportLayout(sheet, pivotIndex, layout)) return false;
    for (const field of definition.fields) {
      const added = wb.addPivotField(sheet, pivotIndex, {
        sourceName: field.sourceName,
        ...(field.customName ? { customName: field.customName } : {}),
        axis: field.axis ?? PivotAxis.Value,
        subtotalTop: field.subtotalTop,
        ...(field.numberFormat ? { numberFormat: field.numberFormat } : {}),
      });
      if (added !== field.index) return false;
      if (field.axis === null && !wb.setPivotFieldAxis(sheet, pivotIndex, added, null)) {
        return false;
      }
      for (const fn of field.subtotalFns) {
        if (!wb.addPivotFieldSubtotalFn(sheet, pivotIndex, added, fn)) return false;
      }
      if (!wb.setPivotFieldSort(sheet, pivotIndex, added, field.sortOrder, field.sortByField)) {
        return false;
      }
      // A group field comes back with the grouping of the field it groups.
      if (field.groupBaseIndex === null && field.dateGroup) {
        const { granularity, calendar, ...bounds } = field.dateGroup;
        if (!wb.setPivotFieldDateGroup(sheet, pivotIndex, added, granularity, calendar, bounds)) {
          return false;
        }
      }
      if (
        field.groupBaseIndex === null &&
        field.numericGroup &&
        !wb.setPivotFieldNumericGroup(sheet, pivotIndex, added, field.numericGroup)
      ) {
        return false;
      }
      if (field.groupBaseIndex !== null) {
        // A group field's items come back with the grouping; only their visibility is the report's.
        const hidden = field.items.flatMap((item, i) => (item.visible ? [] : [i]));
        if (!hidden.every((i) => wb.setPivotFieldItemVisible(sheet, pivotIndex, added, i, false))) {
          return false;
        }
        continue;
      }
      // A field lists its items when it joins an axis; the saved list replaces them, order included.
      if (!wb.clearPivotFieldItems(sheet, pivotIndex, added)) return false;
      for (const item of field.items) {
        const placed =
          item.cacheIndex >= 0 && wb.capabilities.pivotItemByCacheIndex
            ? wb.addPivotFieldItemAt(sheet, pivotIndex, added, item.cacheIndex, item.visible)
            : wb.addPivotFieldItem(sheet, pivotIndex, added, item.label, item.visible);
        if (!placed) return false;
      }
    }
    // The data fields go in first: a field order places the Values field only once there are two.
    return (
      definition.dataFields.every((spec) => wb.addPivotDataField(sheet, pivotIndex, spec) >= 0) &&
      wb.setPivotRowFieldOrder(sheet, pivotIndex, definition.rowFieldOrder) &&
      wb.setPivotColFieldOrder(sheet, pivotIndex, definition.colFieldOrder) &&
      definition.filters.every((filter) => wb.addPivotFilter(sheet, pivotIndex, filter)) &&
      (definition.style === null ||
        wb.setPivotStyleInfo(sheet, pivotIndex, {
          ...definition.style,
          name: definition.style.name || undefined,
        }))
    );
  };
  if (rebuild()) return pivotIndex;
  wb.removePivotTable(sheet, pivotIndex);
  return -1;
}
