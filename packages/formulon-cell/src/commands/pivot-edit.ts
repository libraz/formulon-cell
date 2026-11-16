import {
  type PivotAggregation,
  PivotAxis,
  type PivotDataFieldSpec,
  type PivotFilterSpec,
  type PivotReportLayout,
} from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { PivotDefinition } from '../engine/workbook-handle-pivot.js';
import { sameJsonValue } from './prepared-change.js';

/** One field's settings as submitted from the pivot edit form. */
export interface PivotEditField {
  fieldIndex: number;
  /** Null takes the field out of the report. */
  axis: PivotAxis | null;
  /** Custom subtotal functions to write (empty for the automatic subtotal), or
   *  omitted to leave the field's subtotals untouched. */
  subtotalFns?: readonly PivotAggregation[];
  /** Used when `axis` is Value. */
  aggregation: PivotAggregation;
  /** Used when `axis` is Value; empty means no explicit format. */
  numberFormat: string;
  /** Used when `axis` is Page and `filterChecklist` is null: visible items by name. */
  filterItems: readonly string[];
  /** Used when `axis` is Page: every known item with its visibility, or null
   *  when the field's items are not known and `filterItems` applies. */
  filterChecklist: readonly { value: string; checked: boolean; cacheIndex?: number }[] | null;
}

export interface PivotEdit {
  /** Skip the name, anchor, grand-total and layout writes. */
  fieldListOnly: boolean;
  name: string;
  anchor: { row: number; col: number };
  rowGrandTotals: boolean;
  colGrandTotals: boolean;
  layout: PivotReportLayout;
  fields: readonly PivotEditField[];
  /** Replacement for the pivot's filters, or null to leave them untouched. */
  filters: readonly PivotFilterSpec[] | null;
  /** The data fields as read back from the engine. When given, edits keep each
   *  field's name, show-as and order, and are written as one replacement list;
   *  when omitted, value fields overwrite the existing data fields by position. */
  dataFields?: readonly PivotDataFieldSpec[];
}

/** Key-order-independent form of a data-field list, for comparing it. */
const canonicalSpecs = (specs: readonly PivotDataFieldSpec[]): string =>
  JSON.stringify(specs.map((spec) => Object.entries(spec).sort(([a], [b]) => (a < b ? -1 : 1))));

/** The data-field list after applying the form's value-field settings: a
 *  field's first data field takes the edited aggregation and format, data
 *  fields of fields moved off the Value axis (or out of the report) drop out,
 *  and newly valued fields append. A name survives unless the aggregation changed, in which
 *  case the engine names it afresh. */
function editedDataFields(
  current: readonly PivotDataFieldSpec[],
  fields: readonly PivotEditField[],
): PivotDataFieldSpec[] {
  const byIndex = new Map(fields.map((field) => [field.fieldIndex, field]));
  const edited = new Set<number>();
  const next: PivotDataFieldSpec[] = [];
  for (const dataField of current) {
    const field = byIndex.get(dataField.fieldIndex);
    if (field && field.axis !== PivotAxis.Value) continue;
    if (!field || edited.has(dataField.fieldIndex)) {
      next.push(dataField);
      continue;
    }
    edited.add(dataField.fieldIndex);
    const { numberFormat: _format, name, ...rest } = dataField;
    next.push({
      ...rest,
      ...(field.aggregation === dataField.aggregation && name ? { name } : {}),
      aggregation: field.aggregation,
      ...(field.numberFormat.length > 0 ? { numberFormat: field.numberFormat } : {}),
    });
  }
  for (const field of fields) {
    if (field.axis !== PivotAxis.Value || edited.has(field.fieldIndex)) continue;
    next.push({
      fieldIndex: field.fieldIndex,
      aggregation: field.aggregation,
      ...(field.numberFormat.length > 0 ? { numberFormat: field.numberFormat } : {}),
    });
  }
  return next;
}

/** Rewrite a field's item list with each item's visibility: the one path the
 *  pane, the field dialogs and the filter panel hide and show items through.
 *  True only when every write succeeds. */
export function setPivotFieldItemsChecked(
  wb: WorkbookHandle,
  sheetIndex: number,
  pivotIndex: number,
  fieldIndex: number,
  checklist: readonly { value: string; checked: boolean; cacheIndex?: number }[],
): boolean {
  return keepingCollapsedItems(wb, sheetIndex, pivotIndex, fieldIndex, () =>
    addPivotFieldItems(wb, sheetIndex, pivotIndex, fieldIndex, checklist),
  );
}

/** An item's identity across a rewrite of its field: its cache index, else its label. */
const itemKey = (item: { label: string; cacheIndex: number }): string =>
  item.cacheIndex >= 0 ? `#${item.cacheIndex}` : `=${item.label}`;

/** Runs `rewrite`, which clears and re-adds a field's items, then collapses again the items
 *  that were collapsed before: re-added items come back expanded. */
export function keepingCollapsedItems(
  wb: WorkbookHandle,
  sheetIndex: number,
  pivotIndex: number,
  fieldIndex: number,
  rewrite: () => boolean,
): boolean {
  const before = wb.getPivotField(sheetIndex, pivotIndex, fieldIndex)?.items ?? [];
  const collapsed = new Set(before.filter((i) => !i.showDetail).map(itemKey));
  if (!rewrite()) return false;
  if (collapsed.size === 0) return true;
  const after = wb.getPivotField(sheetIndex, pivotIndex, fieldIndex)?.items ?? [];
  return after.every(
    (item, itemIdx) =>
      !collapsed.has(itemKey(item)) ||
      wb.setPivotFieldItemShowDetail(sheetIndex, pivotIndex, fieldIndex, itemIdx, false),
  );
}

function addPivotFieldItems(
  wb: WorkbookHandle,
  sheetIndex: number,
  pivotIndex: number,
  fieldIndex: number,
  checklist: readonly { value: string; checked: boolean; cacheIndex?: number }[],
): boolean {
  if (!wb.clearPivotFieldItems(sheetIndex, pivotIndex, fieldIndex)) return false;
  return checklist.every((item) => {
    // A cache index states the item exactly, including the blank member;
    // without one, a field whose items were read off the projected layout,
    // the label is all there is.
    if (item.cacheIndex !== undefined) {
      if (
        wb.addPivotFieldItemAt(sheetIndex, pivotIndex, fieldIndex, item.cacheIndex, item.checked)
      ) {
        return true;
      }
      // An engine without the by-index form cannot express a blank member at
      // all; skip it rather than adding an item that filters nothing.
      if (!item.value) return true;
    }
    return wb.addPivotFieldItem(sheetIndex, pivotIndex, fieldIndex, item.value, item.checked);
  });
}

/** Write a pivot edit back to the engine. Every step runs even when an earlier
 *  one fails; returns true only when all of them succeed. */
export function applyPivotEdit(
  wb: WorkbookHandle,
  pivot: {
    sheetIndex: number;
    pivotIndex: number;
    rows: number;
    cols: number;
  },
  edit: PivotEdit,
): boolean {
  const { sheetIndex, pivotIndex } = pivot;
  const { fieldListOnly } = edit;
  const renamed = fieldListOnly || wb.renamePivotTable(sheetIndex, pivotIndex, edit.name);
  const moved =
    fieldListOnly ||
    wb.setPivotTableAnchor(sheetIndex, pivotIndex, {
      row: edit.anchor.row,
      col: edit.anchor.col,
      rows: pivot.rows,
      cols: pivot.cols,
    });
  const totaled =
    fieldListOnly ||
    wb.setPivotTableGrandTotals(sheetIndex, pivotIndex, edit.rowGrandTotals, edit.colGrandTotals);
  const layoutUpdated =
    fieldListOnly || wb.setPivotReportLayout(sheetIndex, pivotIndex, edit.layout);
  // Taking a field out of the report makes the engine drop its data fields
  // and filters; a field already out is left alone so nothing else is touched.
  const fieldsUpdated = edit.fields.every((field) => {
    if (
      field.axis === null &&
      wb.getPivotField(sheetIndex, pivotIndex, field.fieldIndex)?.axis === null
    ) {
      return true;
    }
    return wb.setPivotFieldAxis(sheetIndex, pivotIndex, field.fieldIndex, field.axis);
  });
  const subtotalsUpdated = edit.fields.every((field) => {
    if (!field.subtotalFns) return true;
    return (
      wb.clearPivotFieldSubtotalFns(sheetIndex, pivotIndex, field.fieldIndex) &&
      field.subtotalFns.every((fn) =>
        wb.addPivotFieldSubtotalFn(sheetIndex, pivotIndex, field.fieldIndex, fn),
      )
    );
  });
  const rowFieldOrder = edit.fields
    .filter((field) => field.axis === PivotAxis.Row)
    .map((field) => field.fieldIndex);
  const colFieldOrder = edit.fields
    .filter((field) => field.axis === PivotAxis.Col)
    .map((field) => field.fieldIndex);
  const axisOrdersUpdated =
    fieldsUpdated &&
    wb.setPivotRowFieldOrder(sheetIndex, pivotIndex, rowFieldOrder) &&
    wb.setPivotColFieldOrder(sheetIndex, pivotIndex, colFieldOrder);
  const dataFieldCount = wb.pivotDataFieldCount(sheetIndex, pivotIndex);
  let nextDataFieldIndex = 0;
  const replaceDataFields = (current: readonly PivotDataFieldSpec[]): boolean => {
    const next = editedDataFields(current, edit.fields);
    if (canonicalSpecs(next) === canonicalSpecs(current)) return true;
    return replacePivotDataFields(wb, sheetIndex, pivotIndex, next);
  };
  const valueFieldsUpdated = edit.dataFields
    ? replaceDataFields(edit.dataFields)
    : edit.fields.every((field) => {
        if (field.axis !== PivotAxis.Value) return true;
        const spec: PivotDataFieldSpec = {
          fieldIndex: field.fieldIndex,
          aggregation: field.aggregation,
          ...(field.numberFormat.length > 0 ? { numberFormat: field.numberFormat } : {}),
        };
        const dataFieldIndex = nextDataFieldIndex;
        nextDataFieldIndex += 1;
        if (dataFieldIndex < dataFieldCount) {
          return wb.setPivotDataField(sheetIndex, pivotIndex, dataFieldIndex, spec);
        }
        return wb.addPivotDataField(sheetIndex, pivotIndex, spec) >= 0;
      });
  const filterItemsUpdated = edit.fields.every((field) => {
    if (field.axis !== PivotAxis.Page) return true;
    if (field.filterChecklist) {
      return setPivotFieldItemsChecked(
        wb,
        sheetIndex,
        pivotIndex,
        field.fieldIndex,
        field.filterChecklist,
      );
    }
    if (!wb.clearPivotFieldItems(sheetIndex, pivotIndex, field.fieldIndex)) return false;
    return field.filterItems.every((item) =>
      wb.addPivotFieldItem(sheetIndex, pivotIndex, field.fieldIndex, item, true),
    );
  });
  const pivotFiltersUpdated =
    edit.filters === null ||
    (wb.clearPivotFilters(sheetIndex, pivotIndex) &&
      edit.filters.every((filter) => wb.addPivotFilter(sheetIndex, pivotIndex, filter)));
  return (
    renamed &&
    moved &&
    totaled &&
    layoutUpdated &&
    fieldsUpdated &&
    axisOrdersUpdated &&
    subtotalsUpdated &&
    valueFieldsUpdated &&
    filterItemsUpdated &&
    pivotFiltersUpdated
  );
}

/** Replace a report's value fields with `specs`, in order. False when a write fails. */
export function replacePivotDataFields(
  wb: WorkbookHandle,
  sheetIndex: number,
  pivotIndex: number,
  specs: readonly PivotDataFieldSpec[],
): boolean {
  return (
    wb.clearPivotDataFields(sheetIndex, pivotIndex) &&
    specs.every((spec) => wb.addPivotDataField(sheetIndex, pivotIndex, spec) >= 0)
  );
}

/** Write a field layout read back with `getPivotDefinition` onto the same
 *  report: axes, axis orders, subtotals, items, sort, date and numeric grouping, data
 *  fields and filters. Parts already equal are left alone. Returns true only
 *  when every write succeeds. */
export function restorePivotFieldLayout(
  wb: WorkbookHandle,
  sheetIndex: number,
  pivotIndex: number,
  definition: PivotDefinition,
): boolean {
  const grouped = wb.getPivotDefinition(sheetIndex, pivotIndex);
  if (!grouped) return false;
  const same = sameJsonValue;
  let ok = true;
  // Grouping lives on the cache and adds or removes group fields, so it is restored first, on the
  // fields it groups.
  for (const field of definition.fields) {
    if (field.groupBaseIndex !== null) continue;
    const now = grouped.fields[field.index];
    if (!now) {
      ok = false;
      continue;
    }
    if (!same(now.dateGroup, field.dateGroup)) {
      if (field.dateGroup) {
        const { granularity, calendar, ...bounds } = field.dateGroup;
        ok =
          wb.setPivotFieldDateGroup(
            sheetIndex,
            pivotIndex,
            field.index,
            granularity,
            calendar,
            bounds,
          ) && ok;
      } else {
        ok = wb.clearPivotFieldDateGroup(sheetIndex, pivotIndex, field.index) && ok;
      }
    }
    if (!same(now.numericGroup, field.numericGroup)) {
      ok =
        wb.setPivotFieldNumericGroup(sheetIndex, pivotIndex, field.index, field.numericGroup) && ok;
    }
  }
  const current = wb.getPivotDefinition(sheetIndex, pivotIndex);
  if (!current) return false;
  for (const field of definition.fields) {
    const now = current.fields[field.index];
    if (!now) {
      ok = false;
      continue;
    }
    if (now.axis !== field.axis) {
      ok = wb.setPivotFieldAxis(sheetIndex, pivotIndex, field.index, field.axis) && ok;
    }
    if (!same(now.subtotalFns, field.subtotalFns)) {
      ok =
        wb.clearPivotFieldSubtotalFns(sheetIndex, pivotIndex, field.index) &&
        field.subtotalFns.every((fn) =>
          wb.addPivotFieldSubtotalFn(sheetIndex, pivotIndex, field.index, fn),
        ) &&
        ok;
    }
    const membership = (items: PivotDefinition['fields'][number]['items']) =>
      items.map(({ showDetail: _detail, ...item }) => item);
    if (!same(membership(now.items), membership(field.items))) {
      ok =
        wb.clearPivotFieldItems(sheetIndex, pivotIndex, field.index) &&
        field.items.every((item) =>
          item.cacheIndex >= 0 && wb.capabilities.pivotItemByCacheIndex
            ? wb.addPivotFieldItemAt(
                sheetIndex,
                pivotIndex,
                field.index,
                item.cacheIndex,
                item.visible,
              )
            : wb.addPivotFieldItem(sheetIndex, pivotIndex, field.index, item.label, item.visible),
        ) &&
        ok;
    }
    // Re-added items come back expanded, so collapsed ones are collapsed again.
    const detailNow = wb.getPivotField(sheetIndex, pivotIndex, field.index)?.items ?? [];
    for (const [itemIdx, item] of field.items.entries()) {
      if ((detailNow[itemIdx]?.showDetail ?? true) === item.showDetail) continue;
      ok =
        wb.setPivotFieldItemShowDetail(
          sheetIndex,
          pivotIndex,
          field.index,
          itemIdx,
          item.showDetail,
        ) && ok;
    }
    if (now.sortOrder !== field.sortOrder || now.sortByField !== field.sortByField) {
      ok =
        wb.setPivotFieldSort(
          sheetIndex,
          pivotIndex,
          field.index,
          field.sortOrder,
          field.sortByField,
        ) && ok;
    }
    if (now.subtotalTop !== field.subtotalTop) {
      ok =
        wb.setPivotFieldSubtotalTop(sheetIndex, pivotIndex, field.index, field.subtotalTop) && ok;
    }
  }
  if (!same(current.grandTotals, definition.grandTotals)) {
    ok =
      wb.setPivotTableGrandTotals(
        sheetIndex,
        pivotIndex,
        definition.grandTotals.rows,
        definition.grandTotals.cols,
      ) && ok;
  }
  if (definition.layout !== null && current.layout !== definition.layout) {
    ok = wb.setPivotReportLayout(sheetIndex, pivotIndex, definition.layout) && ok;
  }
  ok =
    wb.setPivotRowFieldOrder(sheetIndex, pivotIndex, definition.rowFieldOrder) &&
    wb.setPivotColFieldOrder(sheetIndex, pivotIndex, definition.colFieldOrder) &&
    ok;
  // Axis changes may have dropped data fields and filters, so compare afresh.
  const afterAxes = wb.getPivotDefinition(sheetIndex, pivotIndex);
  if (
    !afterAxes ||
    canonicalSpecs(afterAxes.dataFields) !== canonicalSpecs(definition.dataFields)
  ) {
    ok = replacePivotDataFields(wb, sheetIndex, pivotIndex, definition.dataFields) && ok;
  }
  if (!afterAxes || !same(afterAxes.filters, definition.filters)) {
    ok =
      wb.clearPivotFilters(sheetIndex, pivotIndex) &&
      definition.filters.every((filter) => wb.addPivotFilter(sheetIndex, pivotIndex, filter)) &&
      ok;
  }
  return ok;
}
