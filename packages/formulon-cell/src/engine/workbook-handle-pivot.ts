import type { PivotMutationWorkbook } from './pivot-mutation.js';
import type {
  Addr,
  CellValue,
  EngineCapabilities,
  PivotAggregation,
  PivotCalendar,
  PivotCell,
  PivotDataFieldSpec,
  PivotDateGrouping,
  PivotFieldSpec,
  PivotFilterSpec,
  Workbook,
} from './types.js';
import { PivotAxis, PivotFilterType, PivotFilterValueKind } from './types.js';
import { fromEngineValue } from './value.js';
import type { WorkbookHandle } from './workbook-handle.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  capabilities: WorkbookHandle['capabilities'];
  assertAlive(): void;
};
type MutablePivotFilterSpecDraft = {
  -readonly [K in keyof PivotFilterSpec]?: PivotFilterSpec[K];
};

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandlePivotMethods {}
}

function internals(handle: unknown): WorkbookHandleInternals {
  return handle as WorkbookHandleInternals;
}

function assertAlive(handle: unknown): void {
  internals(handle).assertAlive();
}

function pivotWb(handle: unknown): PivotMutationWorkbook {
  return internals(handle).wb as PivotMutationWorkbook;
}

export abstract class WorkbookHandlePivotMethods {
  declare readonly capabilities: EngineCapabilities;
  declare readonly sheetCount: number;
  abstract getValue(addr: Addr): CellValue;

  /** Iterate over evaluated PivotTable layout cells on a sheet. The engine
   *  returns sparse cells; blanks are skipped so existing empty-grid behavior
   *  remains unchanged. */
  *pivotCells(sheet: number): Generator<{
    addr: Addr;
    value: CellValue;
    formula: string | null;
    kind: number;
    numberFormat: string;
    pivotIndex: number;
  }> {
    assertAlive(this);
    if (!this.capabilities.pivotTables) return;
    const n = pivotWb(this).pivotCount(sheet);
    for (let i = 0; i < n; i += 1) {
      const layout = pivotWb(this).pivotLayout(sheet, i);
      if (!layout.status.ok) continue;
      for (const cell of layout.cells) {
        const value = fromEngineValue(cell.value);
        if (value.kind === 'blank') continue;
        yield pivotCellEntry(sheet, i, cell, value);
      }
    }
  }

  /** Snapshot of projected PivotTable layouts. This is read-only metadata:
   *  the current engine can evaluate loaded PivotTables into grid cells but
   *  does not expose authoring/editing of the PivotTable definition. */
  getPivotTables(): {
    sheetIndex: number;
    pivotIndex: number;
    top: number;
    left: number;
    rows: number;
    cols: number;
    cells: number;
    fields: string[];
    fieldItems: Record<string, string[]>;
    pivotFilters?: readonly PivotFilterSpec[];
  }[] {
    assertAlive(this);
    if (!this.capabilities.pivotTables) return [];
    const out: {
      sheetIndex: number;
      pivotIndex: number;
      top: number;
      left: number;
      rows: number;
      cols: number;
      cells: number;
      fields: string[];
      fieldItems: Record<string, string[]>;
      pivotFilters?: readonly PivotFilterSpec[];
    }[] = [];
    for (let sheet = 0; sheet < this.sheetCount; sheet += 1) {
      const n = pivotWb(this).pivotCount(sheet);
      for (let pivotIndex = 0; pivotIndex < n; pivotIndex += 1) {
        const layout = pivotWb(this).pivotLayout(sheet, pivotIndex);
        if (!layout.status.ok) continue;
        const fields = new Set<string>();
        const fieldItems = new Map<string, Set<string>>();
        for (const cell of layout.cells) {
          if (!cell.fieldName) continue;
          fields.add(cell.fieldName);
          const value = fromEngineValue(cell.value);
          const label = pivotCellItemLabel(value);
          if (!label) continue;
          const items = fieldItems.get(cell.fieldName) ?? new Set<string>();
          items.add(label);
          fieldItems.set(cell.fieldName, items);
        }
        const cacheFieldItems = pivotCacheSharedItemLabels(
          this,
          fields,
          this.pivotTableCacheId(sheet, pivotIndex),
        );
        for (const [field, labels] of cacheFieldItems) {
          if (labels.length === 0) continue;
          fieldItems.set(field, new Set(labels));
        }
        const pivotFilters = pivotFilterSpecs(this, sheet, pivotIndex);
        out.push({
          sheetIndex: sheet,
          pivotIndex,
          top: layout.top,
          left: layout.left,
          rows: layout.rows,
          cols: layout.cols,
          cells: layout.cells.length,
          fields: [...fields],
          fieldItems: Object.fromEntries(
            [...fieldItems.entries()].map(([field, items]) => [field, [...items]]),
          ),
          ...(pivotFilters.length > 0 ? { pivotFilters } : {}),
        });
      }
    }
    return out;
  }

  pivotCacheCount(): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return 0;
    return pivotWb(this).pivotCacheCount();
  }

  pivotCacheIds(): number[] {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return [];
    const out: number[] = [];
    const n = pivotWb(this).pivotCacheCount();
    for (let i = 0; i < n; i += 1) {
      const r = pivotWb(this).pivotCacheIdAt(i);
      if (r.status.ok) out.push(r.index);
    }
    return out;
  }

  createPivotCache(requestedId = 0): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return -1;
    const r = pivotWb(this).pivotCacheCreate(requestedId);
    return r.status.ok ? r.index : -1;
  }

  removePivotCache(cacheId: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotCacheRemove(cacheId).ok;
  }

  pivotCacheFieldCount(cacheId: number): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return 0;
    return pivotWb(this).pivotCacheFieldCount(cacheId);
  }

  pivotCacheFieldNames(cacheId: number): string[] {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return [];
    const out: string[] = [];
    const n = pivotWb(this).pivotCacheFieldCount(cacheId);
    for (let i = 0; i < n; i += 1) {
      const r = pivotWb(this).pivotCacheFieldName(cacheId, i);
      out.push(r.status.ok ? r.value : '');
    }
    return out;
  }

  pivotCacheSharedItems(cacheId: number, fieldIdx: number): CellValue[] {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return [];
    const wb = pivotWb(this);
    if (!wb.pivotCacheFieldSharedItemCount || !wb.pivotCacheFieldSharedItemValue) return [];
    const out: CellValue[] = [];
    const n = wb.pivotCacheFieldSharedItemCount(cacheId, fieldIdx);
    for (let i = 0; i < n; i += 1) {
      const r = wb.pivotCacheFieldSharedItemValue(cacheId, fieldIdx, i);
      if (r.status.ok) out.push(fromEngineValue(r.value));
    }
    return out;
  }

  addPivotCacheField(cacheId: number, name: string): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return -1;
    const r = pivotWb(this).pivotCacheFieldAdd(cacheId, name);
    return r.status.ok ? r.index : -1;
  }

  clearPivotCacheFields(cacheId: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotCacheFieldClear(cacheId).ok;
  }

  addPivotCacheSharedItem(cacheId: number, fieldIdx: number, value: CellValue): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    if (value.kind === 'number') {
      return pivotWb(this).pivotCacheFieldAddSharedItemNumber(cacheId, fieldIdx, value.value).ok;
    }
    if (value.kind === 'text') {
      return pivotWb(this).pivotCacheFieldAddSharedItemText(cacheId, fieldIdx, value.value).ok;
    }
    if (value.kind === 'bool') {
      return pivotWb(this).pivotCacheFieldAddSharedItemBool(cacheId, fieldIdx, value.value).ok;
    }
    if (value.kind === 'blank')
      return pivotWb(this).pivotCacheFieldAddSharedItemBlank(cacheId, fieldIdx).ok;
    return false;
  }

  clearPivotCacheSharedItems(cacheId: number, fieldIdx: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotCacheFieldClearSharedItems(cacheId, fieldIdx).ok;
  }

  addPivotCacheRecord(cacheId: number): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return -1;
    const r = pivotWb(this).pivotCacheRecordAdd(cacheId);
    return r.status.ok ? r.index : -1;
  }

  clearPivotCacheRecords(cacheId: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotCacheRecordClear(cacheId).ok;
  }

  setPivotCacheRecordValue(
    cacheId: number,
    recordIdx: number,
    fieldIdx: number,
    value: CellValue,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    if (value.kind === 'number') {
      return pivotWb(this).pivotCacheRecordSetNumber(cacheId, recordIdx, fieldIdx, value.value).ok;
    }
    if (value.kind === 'text') {
      return pivotWb(this).pivotCacheRecordSetText(cacheId, recordIdx, fieldIdx, value.value).ok;
    }
    if (value.kind === 'bool') {
      return pivotWb(this).pivotCacheRecordSetBool(cacheId, recordIdx, fieldIdx, value.value).ok;
    }
    if (value.kind === 'blank')
      return pivotWb(this).pivotCacheRecordSetBlank(cacheId, recordIdx, fieldIdx).ok;
    return pivotWb(this).pivotCacheRecordSetError(cacheId, recordIdx, fieldIdx, value.code).ok;
  }

  createPivotTable(
    sheet: number,
    name: string,
    cacheId: number,
    anchor: { row: number; col: number },
  ): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return -1;
    const r = pivotWb(this).pivotCreate(sheet, name, cacheId, anchor.row, anchor.col);
    return r.status.ok ? r.index : -1;
  }

  pivotTableCacheId(sheet: number, pivotIdx: number): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return -1;
    const wb = pivotWb(this);
    if (!wb.pivotCacheId) return -1;
    const r = wb.pivotCacheId(sheet, pivotIdx);
    return r.status.ok ? r.index : -1;
  }

  removePivotTable(sheet: number, pivotIdx: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotRemove(sheet, pivotIdx).ok;
  }

  renamePivotTable(sheet: number, pivotIdx: number, name: string): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotSetName(sheet, pivotIdx, name).ok;
  }

  setPivotTableAnchor(
    sheet: number,
    pivotIdx: number,
    anchor: { row: number; col: number; rows: number; cols: number },
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotSetAnchor(
      sheet,
      pivotIdx,
      anchor.row,
      anchor.col,
      anchor.rows,
      anchor.cols,
    ).ok;
  }

  setPivotTableGrandTotals(
    sheet: number,
    pivotIdx: number,
    rowsEnabled: boolean,
    colsEnabled: boolean,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotSetGrandTotals(sheet, pivotIdx, rowsEnabled, colsEnabled).ok;
  }

  pivotFieldCount(sheet: number, pivotIdx: number): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return 0;
    return pivotWb(this).pivotFieldCount(sheet, pivotIdx);
  }

  addPivotField(sheet: number, pivotIdx: number, spec: PivotFieldSpec): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return -1;
    const r = pivotWb(this).pivotFieldAdd(sheet, pivotIdx, spec);
    return r.status.ok ? r.index : -1;
  }

  clearPivotFields(sheet: number, pivotIdx: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotFieldClear(sheet, pivotIdx).ok;
  }

  setPivotFieldAxis(sheet: number, pivotIdx: number, fieldIdx: number, axis: PivotAxis): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotFieldSetAxis(sheet, pivotIdx, fieldIdx, axis).ok;
  }

  setPivotFieldSort(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    ascending: boolean,
    byField = '',
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotFieldSetSort(sheet, pivotIdx, fieldIdx, ascending, byField).ok;
  }

  setPivotFieldSubtotalTop(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    top: boolean,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotFieldSetSubtotalTop(sheet, pivotIdx, fieldIdx, top).ok;
  }

  addPivotFieldAggregation(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    agg: PivotAggregation,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotFieldAddAggregation(sheet, pivotIdx, fieldIdx, agg).ok;
  }

  clearPivotFieldAggregations(sheet: number, pivotIdx: number, fieldIdx: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotFieldClearAggregations(sheet, pivotIdx, fieldIdx).ok;
  }

  addPivotFieldItem(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    name: string,
    visible: boolean,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotFieldAddItem(sheet, pivotIdx, fieldIdx, name, visible).ok;
  }

  clearPivotFieldItems(sheet: number, pivotIdx: number, fieldIdx: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotFieldClearItems(sheet, pivotIdx, fieldIdx).ok;
  }

  setPivotFieldItemVisible(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    itemIdx: number,
    visible: boolean,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotFieldSetItemVisible(sheet, pivotIdx, fieldIdx, itemIdx, visible).ok;
  }

  addPivotFieldSubtotalFn(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    agg: PivotAggregation,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotFieldAddSubtotalFn(sheet, pivotIdx, fieldIdx, agg).ok;
  }

  clearPivotFieldSubtotalFns(sheet: number, pivotIdx: number, fieldIdx: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotFieldClearSubtotalFns(sheet, pivotIdx, fieldIdx).ok;
  }

  setPivotFieldDateGroup(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    granularity: PivotDateGrouping,
    calendar: PivotCalendar,
    bounds: { startYear?: number; endYear?: number } = {},
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotFieldSetDateGroup(
      sheet,
      pivotIdx,
      fieldIdx,
      granularity,
      calendar,
      bounds.startYear ?? -1,
      bounds.endYear ?? -1,
    ).ok;
  }

  clearPivotFieldDateGroup(sheet: number, pivotIdx: number, fieldIdx: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotFieldClearDateGroup(sheet, pivotIdx, fieldIdx).ok;
  }

  setPivotFieldNumberFormat(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    format: string,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotFieldSetNumberFormat(sheet, pivotIdx, fieldIdx, format).ok;
  }

  setPivotRowFieldOrder(sheet: number, pivotIdx: number, indices: readonly number[]): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotSetRowFieldOrder(sheet, pivotIdx, indices).ok;
  }

  setPivotColFieldOrder(sheet: number, pivotIdx: number, indices: readonly number[]): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotSetColFieldOrder(sheet, pivotIdx, indices).ok;
  }

  pivotDataFieldCount(sheet: number, pivotIdx: number): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return 0;
    return pivotWb(this).pivotDataFieldCount(sheet, pivotIdx);
  }

  addPivotDataField(sheet: number, pivotIdx: number, spec: PivotDataFieldSpec): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return -1;
    const r = pivotWb(this).pivotDataFieldAdd(sheet, pivotIdx, spec);
    return r.status.ok ? r.index : -1;
  }

  setPivotDataField(
    sheet: number,
    pivotIdx: number,
    dataFieldIdx: number,
    spec: PivotDataFieldSpec,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotDataFieldSet(sheet, pivotIdx, dataFieldIdx, spec).ok;
  }

  clearPivotDataFields(sheet: number, pivotIdx: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotDataFieldClear(sheet, pivotIdx).ok;
  }

  pivotFilterCount(sheet: number, pivotIdx: number): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return 0;
    return pivotWb(this).pivotFilterCount(sheet, pivotIdx);
  }

  addPivotFilter(sheet: number, pivotIdx: number, spec: PivotFilterSpec): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotFilterAdd(sheet, pivotIdx, spec).ok;
  }

  clearPivotFilters(sheet: number, pivotIdx: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotFilterClear(sheet, pivotIdx).ok;
  }

  removePivotFilter(sheet: number, pivotIdx: number, filterIdx: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotFilterRemoveAt(sheet, pivotIdx, filterIdx).ok;
  }
}

function pivotCellItemLabel(value: CellValue): string {
  if (value.kind === 'text') return value.value.trim();
  if (value.kind === 'number') return String(value.value);
  if (value.kind === 'bool') return value.value ? 'TRUE' : 'FALSE';
  if (value.kind === 'error') return `#${value.code}`;
  return '';
}

function pivotCacheSharedItemLabels(
  handle: WorkbookHandlePivotMethods,
  fields: ReadonlySet<string>,
  preferredCacheId: number,
): Map<string, string[]> {
  const out = new Map<string, string[]>();
  if (!handle.capabilities.pivotTableMutate) return out;
  const cacheIds = preferredCacheId >= 0 ? [preferredCacheId] : handle.pivotCacheIds();
  for (const cacheId of cacheIds) {
    for (const [fieldIdx, fieldName] of handle.pivotCacheFieldNames(cacheId).entries()) {
      if (!fields.has(fieldName) || out.has(fieldName)) continue;
      const labels = handle
        .pivotCacheSharedItems(cacheId, fieldIdx)
        .map(pivotCellItemLabel)
        .filter((label) => label.length > 0);
      if (labels.length > 0) out.set(fieldName, [...new Set(labels)]);
    }
  }
  return out;
}

function pivotFilterSpecs(
  handle: WorkbookHandlePivotMethods,
  sheet: number,
  pivotIndex: number,
): PivotFilterSpec[] {
  const wb = pivotWb(handle);
  if (typeof wb.pivotFilterCount !== 'function') return [];
  const count = wb.pivotFilterCount(sheet, pivotIndex);
  if (!Number.isFinite(count) || count <= 0) return [];
  const out: PivotFilterSpec[] = [];
  for (let filterIndex = 0; filterIndex < count; filterIndex += 1) {
    const direct = wb.pivotFilterSpec?.(sheet, pivotIndex, filterIndex);
    const spec =
      sanitizePivotFilterSpec(direct?.status.ok ? direct.spec : undefined) ??
      readPivotFilterSpec(wb, sheet, pivotIndex, filterIndex);
    if (spec) out.push(spec);
  }
  return out;
}

function readPivotFilterSpec(
  wb: PivotMutationWorkbook,
  sheet: number,
  pivotIndex: number,
  filterIndex: number,
): PivotFilterSpec | null {
  const axis = pivotFilterNumber(wb.pivotFilterAxis, sheet, pivotIndex, filterIndex);
  const fieldName = pivotFilterString(wb.pivotFilterFieldName, sheet, pivotIndex, filterIndex);
  const type = pivotFilterNumber(wb.pivotFilterType, sheet, pivotIndex, filterIndex);
  if (
    !isPivotAxis(axis) ||
    fieldName === undefined ||
    fieldName.trim().length === 0 ||
    !isPivotFilterType(type)
  ) {
    return null;
  }
  const spec: MutablePivotFilterSpecDraft = {
    axis,
    fieldName,
    type,
    valueInt: pivotFilterNumber(wb.pivotFilterValueInt, sheet, pivotIndex, filterIndex),
    valueDouble: pivotFilterNumber(wb.pivotFilterValueDouble, sheet, pivotIndex, filterIndex),
    valueText: pivotFilterString(wb.pivotFilterValueText, sheet, pivotIndex, filterIndex),
    valueHighInt: pivotFilterNumber(wb.pivotFilterValueHighInt, sheet, pivotIndex, filterIndex),
    valueHighDouble: pivotFilterNumber(
      wb.pivotFilterValueHighDouble,
      sheet,
      pivotIndex,
      filterIndex,
    ),
  };
  const valueKind = pivotFilterNumber(wb.pivotFilterValueKind, sheet, pivotIndex, filterIndex);
  const valueHighKind = pivotFilterNumber(
    wb.pivotFilterValueHighKind,
    sheet,
    pivotIndex,
    filterIndex,
  );
  if (isPivotFilterValueKind(valueKind)) spec.valueKind = valueKind;
  if (isPivotFilterValueKind(valueHighKind)) spec.valueHighKind = valueHighKind;
  return sanitizePivotFilterSpec(spec);
}

function sanitizePivotFilterSpec(
  spec: Partial<PivotFilterSpec> | undefined | null,
): PivotFilterSpec | null {
  if (!spec) return null;
  if (
    !isPivotAxis(spec.axis) ||
    typeof spec.fieldName !== 'string' ||
    spec.fieldName.trim().length === 0 ||
    !isPivotFilterType(spec.type)
  ) {
    return null;
  }
  return {
    axis: spec.axis,
    fieldName: spec.fieldName.trim(),
    type: spec.type,
    ...(isPivotFilterValueKind(spec.valueKind) ? { valueKind: spec.valueKind } : {}),
    ...(Number.isFinite(spec.valueInt) ? { valueInt: spec.valueInt } : {}),
    ...(Number.isFinite(spec.valueDouble) ? { valueDouble: spec.valueDouble } : {}),
    ...(typeof spec.valueText === 'string' ? { valueText: spec.valueText } : {}),
    ...(isPivotFilterValueKind(spec.valueHighKind) ? { valueHighKind: spec.valueHighKind } : {}),
    ...(Number.isFinite(spec.valueHighInt) ? { valueHighInt: spec.valueHighInt } : {}),
    ...(Number.isFinite(spec.valueHighDouble) ? { valueHighDouble: spec.valueHighDouble } : {}),
  };
}

function pivotFilterNumber(
  reader:
    | ((
        sheet: number,
        pivotIndex: number,
        filterIndex: number,
      ) => { status: { ok: boolean }; value: number })
    | undefined,
  sheet: number,
  pivotIndex: number,
  filterIndex: number,
): number | undefined {
  const result = reader?.(sheet, pivotIndex, filterIndex);
  return result?.status.ok && Number.isFinite(result.value) ? result.value : undefined;
}

function pivotFilterString(
  reader:
    | ((
        sheet: number,
        pivotIndex: number,
        filterIndex: number,
      ) => { status: { ok: boolean }; value: string })
    | undefined,
  sheet: number,
  pivotIndex: number,
  filterIndex: number,
): string | undefined {
  const result = reader?.(sheet, pivotIndex, filterIndex);
  return result?.status.ok && typeof result.value === 'string' ? result.value : undefined;
}

function isPivotAxis(value: unknown): value is PivotAxis {
  return typeof value === 'number' && Object.values(PivotAxis).includes(value as PivotAxis);
}

function isPivotFilterType(value: unknown): value is PivotFilterType {
  return (
    typeof value === 'number' && Object.values(PivotFilterType).includes(value as PivotFilterType)
  );
}

function isPivotFilterValueKind(value: unknown): value is PivotFilterValueKind {
  return (
    typeof value === 'number' &&
    Object.values(PivotFilterValueKind).includes(value as PivotFilterValueKind)
  );
}

function pivotCellEntry(
  sheet: number,
  pivotIndex: number,
  cell: PivotCell,
  value: CellValue,
): {
  addr: Addr;
  value: CellValue;
  formula: string | null;
  kind: number;
  numberFormat: string;
  pivotIndex: number;
} {
  return {
    addr: { sheet, row: cell.row, col: cell.col },
    value,
    formula: null,
    kind: cell.kind,
    numberFormat: cell.numberFormat,
    pivotIndex,
  };
}

export function installPivotMethods(target: WorkbookHandleCtor): void {
  for (const key of Object.getOwnPropertyNames(WorkbookHandlePivotMethods.prototype)) {
    if (key === 'constructor') continue;
    const descriptor = Object.getOwnPropertyDescriptor(WorkbookHandlePivotMethods.prototype, key);
    if (!descriptor) continue;
    Object.defineProperty(target.prototype, key, descriptor);
  }
}
