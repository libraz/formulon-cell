import { pivotFieldFiltered } from '../engine/pivot-style.js';
import type { PivotAxis, PivotFilterSpec, PivotSortOrder } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { writePivotFieldItems } from './pivot-edit.js';

/** One item of a field's checklist. */
export interface PivotFilterItem {
  value: string;
  checked: boolean;
  /** Index into the cache field's shared items, when known. */
  cacheIndex?: number;
}

/** What a field's filter panel shows. */
export interface PivotFieldFilterModel {
  fieldIndex: number;
  /** The field's caption. */
  label: string;
  sourceName: string;
  axis: PivotAxis | null;
  items: PivotFilterItem[];
  /** Some item is hidden or a label / value filter applies. */
  filtered: boolean;
  /** The field's label / value filter, if any. */
  condition: PivotFilterSpec | null;
  sortOrder: PivotSortOrder;
  /** Empty when the field sorts by its own labels, else a data field caption. */
  sortByField: string;
  /** Data field captions the field can sort by. */
  dataFields: string[];
}

/** A cache field's item names in cache-index order. */
function cacheItemNames(wb: WorkbookHandle, cacheId: number, cacheFieldIndex: number): string[] {
  const shared = cacheId >= 0 ? wb.pivotCacheSharedItems(cacheId, cacheFieldIndex) : [];
  return shared.map((value) =>
    value.kind === 'text'
      ? value.value
      : value.kind === 'number' || value.kind === 'bool'
        ? String(value.value)
        : '',
  );
}

/** The filter state of a report field; null when the definition cannot be read. */
export function pivotFieldFilterModel(
  wb: WorkbookHandle,
  sheet: number,
  pivotIndex: number,
  fieldIndex: number,
): PivotFieldFilterModel | null {
  const definition = wb.getPivotDefinition(sheet, pivotIndex);
  const field = definition?.fields[fieldIndex];
  if (!definition || !field) return null;
  let items: PivotFilterItem[];
  if (field.items.length > 0) {
    // Items written by cache index read back without a name.
    const names = field.items.some((item) => !item.label && item.cacheIndex >= 0)
      ? cacheItemNames(wb, definition.cacheId, field.cacheFieldIndex)
      : [];
    items = field.items.map((item) => ({
      value: item.label || (item.cacheIndex >= 0 ? (names[item.cacheIndex] ?? '') : ''),
      checked: item.visible,
      ...(item.cacheIndex >= 0 ? { cacheIndex: item.cacheIndex } : {}),
    }));
  } else {
    // A field whose items were never written lists the cache's members, all shown.
    const pivot = wb
      .getPivotTables()
      .find((p) => p.sheetIndex === sheet && p.pivotIndex === pivotIndex);
    const labels = pivot?.fieldItems[field.sourceName] ?? [];
    const indexes = pivot?.fieldItemIndexes?.[field.sourceName] ?? [];
    items = labels.map((value, i) => ({
      value,
      checked: true,
      ...(indexes[i] !== undefined ? { cacheIndex: indexes[i] } : {}),
    }));
  }
  return {
    fieldIndex,
    label: field.customName || field.sourceName,
    sourceName: field.sourceName,
    axis: field.axis,
    items,
    filtered: pivotFieldFiltered(definition, fieldIndex),
    condition: definition.filters.find((f) => f.fieldName === field.sourceName) ?? null,
    sortOrder: field.sortOrder,
    sortByField: field.sortByField,
    dataFields: definition.dataFields.map((d) => d.name ?? '').filter((name) => name !== ''),
  };
}

/** Show the checked items of a field and hide the rest. */
export function setPivotFieldItemsChecked(
  wb: WorkbookHandle,
  sheet: number,
  pivotIndex: number,
  fieldIndex: number,
  items: readonly PivotFilterItem[],
): boolean {
  return writePivotFieldItems(wb, sheet, pivotIndex, fieldIndex, items);
}

/** Replace the label / value filter of a field; null removes it. Filters on
 *  other fields are written back as they were. */
export function setPivotFieldCondition(
  wb: WorkbookHandle,
  sheet: number,
  pivotIndex: number,
  fieldIndex: number,
  spec: PivotFilterSpec | null,
): boolean {
  const definition = wb.getPivotDefinition(sheet, pivotIndex);
  const field = definition?.fields[fieldIndex];
  if (!definition || !field) return false;
  const others = definition.filters.filter((f) => f.fieldName !== field.sourceName);
  return (
    wb.clearPivotFilters(sheet, pivotIndex) &&
    [...others, ...(spec ? [spec] : [])].every((f) => wb.addPivotFilter(sheet, pivotIndex, f))
  );
}

/** Show every item of a field and drop its label / value filter. */
export function clearPivotFieldFilter(
  wb: WorkbookHandle,
  sheet: number,
  pivotIndex: number,
  fieldIndex: number,
): boolean {
  const model = pivotFieldFilterModel(wb, sheet, pivotIndex, fieldIndex);
  if (!model) return false;
  const shown = setPivotFieldItemsChecked(
    wb,
    sheet,
    pivotIndex,
    fieldIndex,
    model.items.map((item) => ({ ...item, checked: true })),
  );
  const unfiltered =
    model.condition === null || setPivotFieldCondition(wb, sheet, pivotIndex, fieldIndex, null);
  return shown && unfiltered;
}
