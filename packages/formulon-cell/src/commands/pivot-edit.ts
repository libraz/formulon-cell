import {
  type PivotAggregation,
  PivotAxis,
  type PivotDataFieldSpec,
  type PivotFilterSpec,
  type PivotReportLayout,
} from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';

/** One field's settings as submitted from the pivot edit form. */
export interface PivotEditField {
  fieldIndex: number;
  axis: PivotAxis;
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
 *  fields of fields moved off the Value axis drop out, and newly valued
 *  fields append. A name survives unless the aggregation changed, in which
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
  const fieldsUpdated = edit.fields.every((field) =>
    wb.setPivotFieldAxis(sheetIndex, pivotIndex, field.fieldIndex, field.axis),
  );
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
    return (
      wb.clearPivotDataFields(sheetIndex, pivotIndex) &&
      next.every((spec) => wb.addPivotDataField(sheetIndex, pivotIndex, spec) >= 0)
    );
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
    if (!wb.clearPivotFieldItems(sheetIndex, pivotIndex, field.fieldIndex)) return false;
    if (field.filterChecklist) {
      return field.filterChecklist.every((item) => {
        // A cache index states the item exactly, including the blank
        // member; without one — a field whose items were read off the
        // projected layout — the label is all there is.
        if (item.cacheIndex !== undefined) {
          if (
            wb.addPivotFieldItemAt(
              sheetIndex,
              pivotIndex,
              field.fieldIndex,
              item.cacheIndex,
              item.checked,
            )
          ) {
            return true;
          }
          // An engine without the by-index form cannot express a blank
          // member at all; skip it rather than adding an item that
          // filters nothing.
          if (!item.value) return true;
        }
        return wb.addPivotFieldItem(
          sheetIndex,
          pivotIndex,
          field.fieldIndex,
          item.value,
          item.checked,
        );
      });
    }
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
    valueFieldsUpdated &&
    filterItemsUpdated &&
    pivotFiltersUpdated
  );
}
