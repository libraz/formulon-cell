import { PIVOT_CELL_KIND } from './pivot-cell-kind.js';
import type { PivotMutationWorkbook } from './pivot-mutation.js';
import type {
  Addr,
  CellValue,
  EngineCapabilities,
  PivotCalendar,
  PivotCell,
  PivotDataFieldSpec,
  PivotDateGrouping,
  PivotFieldSpec,
  PivotFilterSpec,
  PivotReportLayout,
  PivotWorksheetSource,
  Range,
  Workbook,
} from './types.js';
import {
  type PivotAggregation,
  PivotAxis,
  PivotFilterType,
  PivotFilterValueKind,
  PivotShowValuesAs,
  type PivotSortOrder,
} from './types.js';
import { fromEngineValue } from './value.js';
import type { WorkbookHandle } from './workbook-handle.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  capabilities: WorkbookHandle['capabilities'];
  assertAlive(): void;
};
/** One member of a pivot field, paired with the cache index that addresses it.
 *  The blank member's `label` is empty — it has no label of its own, and the
 *  index is the only way to name it. */
export interface PivotFieldItem {
  readonly label: string;
  readonly cacheIndex: number;
}

/** One member of a pivot field's item list, as stored on the report. */
export interface PivotItemDefinition {
  /** The item's own label; empty for an item bound only by cache index. */
  readonly label: string;
  /** Index into the bound cache field's shared items, -1 for a label-only item. */
  readonly cacheIndex: number;
  readonly visible: boolean;
}

/** A pivot field as stored on the report, whichever axis (if any) it sits on. */
export interface PivotFieldDefinition {
  readonly index: number;
  readonly sourceName: string;
  readonly customName: string;
  /** Null for a field that sits on no axis. */
  readonly axis: PivotAxis | null;
  readonly subtotalTop: boolean;
  /** Number format code, empty for none. */
  readonly numberFormat: string;
  readonly cacheFieldIndex: number;
  readonly sortOrder: PivotSortOrder;
  readonly sortByField: string;
  readonly dateGroup: {
    readonly granularity: PivotDateGrouping;
    readonly calendar: PivotCalendar;
    readonly startYear: number;
    readonly endYear: number;
    readonly intervalDays: number;
    readonly startSerial: number;
    readonly endSerial: number;
  } | null;
  readonly items: readonly PivotItemDefinition[];
}

/** The stored definition of a PivotTable, read back from the engine. */
export interface PivotDefinition {
  readonly name: string;
  readonly cacheId: number;
  readonly grandTotals: { readonly rows: boolean; readonly cols: boolean };
  readonly layout: PivotReportLayout | null;
  /** Pivot field indices of the row / column axes, in display order. */
  readonly rowFieldOrder: readonly number[];
  readonly colFieldOrder: readonly number[];
  readonly fields: readonly PivotFieldDefinition[];
  /** Value fields in display order; `fieldIndex` is a pivot field index. */
  readonly dataFields: readonly PivotDataFieldSpec[];
  /** Filters added in this session; a filter block read from a file is not listed. */
  readonly filters: readonly PivotFilterSpec[];
}

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandlePivotMethods {}
}

function internals(handle: unknown): WorkbookHandleInternals {
  return handle as WorkbookHandleInternals;
}

function assertAlive(handle: unknown): void {
  internals(handle).assertAlive();
}

/** Cache index of `fieldName`. When the engine cannot tie the report to its
 *  cache, a name found in exactly one cache stands in. */
function cacheFieldIndex(handle: WorkbookHandlePivotMethods, cacheId: number, fieldName: string) {
  if (cacheId >= 0) return handle.pivotCacheFieldNames(cacheId).indexOf(fieldName);
  let found = -1;
  let matches = 0;
  for (const id of handle.pivotCacheIds()) {
    const index = handle.pivotCacheFieldNames(id).indexOf(fieldName);
    if (index < 0) continue;
    found = index;
    matches += 1;
  }
  return matches === 1 ? found : -1;
}

function pivotWb(handle: unknown): PivotMutationWorkbook {
  return internals(handle).wb as PivotMutationWorkbook;
}

/** Unwrap a 0.12 numeric result without turning an engine failure into a
 * legitimate zero count. */
function unwrapNumberResult(result: { status: { ok: boolean }; value: number }): number | null {
  if (!result.status.ok || !Number.isFinite(result.value) || result.value < 0) return null;
  return result.value;
}

export abstract class WorkbookHandlePivotMethods {
  declare readonly capabilities: EngineCapabilities;
  declare readonly sheetCount: number;
  abstract getValue(addr: Addr): CellValue;

  /** Whether any PivotTable on `range`'s sheet occupies a cell of `range`. */
  pivotOverlapsRange(range: Range): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTables) return false;
    const n = unwrapNumberResult(pivotWb(this).pivotCount(range.sheet));
    if (n === null) return false;
    for (let i = 0; i < n; i += 1) {
      const layout = pivotWb(this).pivotLayout(range.sheet, i);
      if (!layout.status.ok) continue;
      if (
        range.r0 < layout.top + layout.rows &&
        range.r1 >= layout.top &&
        range.c0 < layout.left + layout.cols &&
        range.c1 >= layout.left
      ) {
        return true;
      }
    }
    return false;
  }

  /** Rectangle each PivotTable of `sheet` occupies, by pivot index. */
  pivotBounds(sheet: number): { pivotIndex: number; range: Range }[] {
    assertAlive(this);
    if (!this.capabilities.pivotTables) return [];
    const n = unwrapNumberResult(pivotWb(this).pivotCount(sheet));
    if (n === null) return [];
    const out: { pivotIndex: number; range: Range }[] = [];
    for (let pivotIndex = 0; pivotIndex < n; pivotIndex += 1) {
      const layout = pivotWb(this).pivotLayout(sheet, pivotIndex);
      if (!layout.status.ok || layout.rows <= 0 || layout.cols <= 0) continue;
      out.push({
        pivotIndex,
        range: {
          sheet,
          r0: layout.top,
          c0: layout.left,
          r1: layout.top + layout.rows - 1,
          c1: layout.left + layout.cols - 1,
        },
      });
    }
    return out;
  }

  /** The PivotTable layout cell at `addr`, or null outside every report.
   *  `fieldIndex` is the cache index of the cell's source field, -1 when the
   *  cell belongs to no source field (captions, data-field names). */
  pivotCellInfo(addr: Addr): {
    pivotIndex: number;
    kind: number;
    fieldName: string;
    fieldIndex: number;
    /** Cache bound to the report, -1 when the engine cannot say. */
    cacheId: number;
    isValue: boolean;
  } | null {
    assertAlive(this);
    if (!this.capabilities.pivotTables) return null;
    const n = unwrapNumberResult(pivotWb(this).pivotCount(addr.sheet));
    if (n === null) return null;
    for (let pivotIndex = 0; pivotIndex < n; pivotIndex += 1) {
      const layout = pivotWb(this).pivotLayout(addr.sheet, pivotIndex);
      if (!layout.status.ok) continue;
      if (
        addr.row < layout.top ||
        addr.row >= layout.top + layout.rows ||
        addr.col < layout.left ||
        addr.col >= layout.left + layout.cols
      ) {
        continue;
      }
      const cell = layout.cells.find((c) => c.row === addr.row && c.col === addr.col);
      const kind = cell?.kind ?? PIVOT_CELL_KIND.Blank;
      const fieldName = cell?.fieldName ?? '';
      const mutable = this.capabilities.pivotTableMutate;
      const cacheId = mutable ? this.pivotTableCacheId(addr.sheet, pivotIndex) : -1;
      return {
        pivotIndex,
        kind,
        fieldName,
        fieldIndex: mutable && fieldName ? cacheFieldIndex(this, cacheId, fieldName) : -1,
        cacheId,
        isValue:
          kind === PIVOT_CELL_KIND.Data ||
          (!!cell && fromEngineValue(cell.value).kind === 'number'),
      };
    }
    return null;
  }

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
    const n = unwrapNumberResult(pivotWb(this).pivotCount(sheet));
    if (n === null) return;
    for (let i = 0; i < n; i += 1) {
      const layout = pivotWb(this).pivotLayout(sheet, i);
      if (!layout.status.ok) continue;
      for (const cell of layout.cells) {
        const value = fromEngineValue(cell.value);
        if (value.kind === 'blank') continue;
        yield pivotCellEntry(this, sheet, i, cell, value);
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
    /** Cache index of each entry in `fieldItems`, positionally aligned. A
     *  field whose items were inferred from the projected layout rather than
     *  read out of the cache has no entry, since those labels address no
     *  shared item. */
    fieldItemIndexes?: Record<string, number[]>;
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
      fieldItemIndexes?: Record<string, number[]>;
      pivotFilters?: readonly PivotFilterSpec[];
    }[] = [];
    for (let sheet = 0; sheet < this.sheetCount; sheet += 1) {
      const n = unwrapNumberResult(pivotWb(this).pivotCount(sheet));
      if (n === null) continue;
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
        // The cache is the better source: it carries every member of the
        // field, including ones no projected cell happens to show, and the
        // index each is addressed by.
        const fieldItemIndexes = new Map<string, number[]>();
        for (const [field, items] of cacheFieldItems) {
          if (items.length === 0) continue;
          fieldItems.set(field, new Set(items.map((item) => item.label)));
          fieldItemIndexes.set(
            field,
            items.map((item) => item.cacheIndex),
          );
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
          fieldItemIndexes: Object.fromEntries(fieldItemIndexes),
          ...(pivotFilters.length > 0 ? { pivotFilters } : {}),
        });
      }
    }
    return out;
  }

  pivotCacheCount(): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return 0;
    return unwrapNumberResult(pivotWb(this).pivotCacheCount()) ?? 0;
  }

  pivotCacheIds(): number[] {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return [];
    const out: number[] = [];
    const n = unwrapNumberResult(pivotWb(this).pivotCacheCount());
    if (n === null) return out;
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

  getPivotCacheWorksheetSource(cacheId: number): PivotWorksheetSource | null {
    assertAlive(this);
    if (!this.capabilities.pivotCacheSource) return null;
    const r = pivotWb(this).pivotCacheGetWorksheetSource(cacheId);
    if (!r.status.ok) return null;
    return {
      present: r.present,
      ...(r.ref ? { ref: r.ref } : {}),
      ...(r.sheet ? { sheet: r.sheet } : {}),
      ...(r.name ? { name: r.name } : {}),
    };
  }

  setPivotCacheWorksheetSource(cacheId: number, source: PivotWorksheetSource): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotCacheSource) return false;
    return pivotWb(this).pivotCacheSetWorksheetSource(cacheId, source).ok;
  }

  pivotCacheFieldCount(cacheId: number): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return 0;
    return unwrapNumberResult(pivotWb(this).pivotCacheFieldCount(cacheId)) ?? 0;
  }

  pivotCacheFieldNames(cacheId: number): string[] {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return [];
    const out: string[] = [];
    const n = unwrapNumberResult(pivotWb(this).pivotCacheFieldCount(cacheId));
    if (n === null) return out;
    for (let i = 0; i < n; i += 1) {
      const r = pivotWb(this).pivotCacheFieldName(cacheId, i);
      out.push(r.status.ok ? r.value : '');
    }
    return out;
  }

  /** Shared items of a pivot cache field, in the index order `<item x="N">`
   *  addresses. A value the engine refuses to read back is kept as a blank so
   *  a caller can still use the array position as that cache index. */
  pivotCacheSharedItems(cacheId: number, fieldIdx: number): CellValue[] {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return [];
    const wb = pivotWb(this);
    if (!wb.pivotCacheFieldSharedItemCount || !wb.pivotCacheFieldSharedItemValue) return [];
    const out: CellValue[] = [];
    const n = unwrapNumberResult(wb.pivotCacheFieldSharedItemCount(cacheId, fieldIdx));
    if (n === null) return out;
    for (let i = 0; i < n; i += 1) {
      const r = wb.pivotCacheFieldSharedItemValue(cacheId, fieldIdx, i);
      out.push(r.status.ok ? fromEngineValue(r.value) : { kind: 'blank' });
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
    return pivotWb(this).pivotCacheFieldAddSharedItemError(cacheId, fieldIdx, value.code).ok;
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

  /** The PivotTable's stored name, or null when it cannot be read. */
  pivotTableName(sheet: number, pivotIdx: number): string | null {
    assertAlive(this);
    if (!this.capabilities.pivotDefinitionRead) return null;
    const r = pivotWb(this).pivotName(sheet, pivotIdx);
    return r.status.ok ? r.value : null;
  }

  getPivotTableGrandTotals(
    sheet: number,
    pivotIdx: number,
  ): { rows: boolean; cols: boolean } | null {
    assertAlive(this);
    if (!this.capabilities.pivotDefinitionRead) return null;
    const r = pivotWb(this).pivotGetGrandTotals(sheet, pivotIdx);
    return r.status.ok ? { rows: r.rows, cols: r.cols } : null;
  }

  /** Pivot field indices on the row axis, in display order; null when unreadable. */
  pivotRowFieldOrder(sheet: number, pivotIdx: number): number[] | null {
    assertAlive(this);
    if (!this.capabilities.pivotDefinitionRead) return null;
    const r = pivotWb(this).pivotRowFieldOrder(sheet, pivotIdx);
    return r.status.ok ? [...r.indices] : null;
  }

  /** Pivot field indices on the column axis, in display order; null when unreadable. */
  pivotColFieldOrder(sheet: number, pivotIdx: number): number[] | null {
    assertAlive(this);
    if (!this.capabilities.pivotDefinitionRead) return null;
    const r = pivotWb(this).pivotColFieldOrder(sheet, pivotIdx);
    return r.status.ok ? [...r.indices] : null;
  }

  getPivotField(sheet: number, pivotIdx: number, fieldIdx: number): PivotFieldDefinition | null {
    assertAlive(this);
    if (!this.capabilities.pivotDefinitionRead) return null;
    const wb = pivotWb(this);
    const r = wb.pivotFieldAt(sheet, pivotIdx, fieldIdx);
    if (!r.status.ok) return null;
    const items: PivotItemDefinition[] = [];
    const itemCount = unwrapNumberResult(wb.pivotFieldItemCount(sheet, pivotIdx, fieldIdx)) ?? 0;
    for (let i = 0; i < itemCount; i += 1) {
      const item = wb.pivotFieldItemAt(sheet, pivotIdx, fieldIdx, i);
      if (!item.status.ok) continue;
      items.push({ label: item.name, cacheIndex: item.cacheIndex, visible: item.visible });
    }
    return {
      index: fieldIdx,
      sourceName: r.sourceName,
      customName: r.customName,
      axis: isPivotAxis(r.axis) ? r.axis : null,
      subtotalTop: r.subtotalTop,
      numberFormat: pivotNumberFormatCode(this, r.numberFormat),
      cacheFieldIndex: r.cacheFieldIndex,
      sortOrder: r.sortOrder as PivotSortOrder,
      sortByField: r.sortByField,
      dateGroup: r.hasDateGroup
        ? {
            granularity: r.dateGranularity as PivotDateGrouping,
            calendar: r.dateCalendar as PivotCalendar,
            startYear: r.dateStartYear,
            endYear: r.dateEndYear,
            intervalDays: r.dateIntervalDays,
            startSerial: r.dateStartSerial,
            endSerial: r.dateEndSerial,
          }
        : null,
      items,
    };
  }

  getPivotDataField(
    sheet: number,
    pivotIdx: number,
    dataFieldIdx: number,
  ): PivotDataFieldSpec | null {
    assertAlive(this);
    if (!this.capabilities.pivotDefinitionRead) return null;
    const r = pivotWb(this).pivotDataFieldAt(sheet, pivotIdx, dataFieldIdx);
    if (!r.status.ok) return null;
    const numberFormat = pivotNumberFormatCode(this, r.numberFormat);
    return {
      name: r.name,
      fieldIndex: r.fieldIndex,
      aggregation: r.aggregation as PivotAggregation,
      ...(numberFormat ? { numberFormat } : {}),
      ...(r.showAs !== PivotShowValuesAs.Normal ? { showValuesAs: r.showAs } : {}),
      ...(r.showAsBaseField >= 0 ? { showAsBaseField: r.showAsBaseField } : {}),
      ...(r.showAsBaseItem >= 0 ? { showAsBaseItem: r.showAsBaseItem } : {}),
    };
  }

  /** The whole stored definition of a PivotTable, or null when the engine
   *  cannot read it back (older engine, or a pivot that does not exist). */
  getPivotDefinition(sheet: number, pivotIdx: number): PivotDefinition | null {
    assertAlive(this);
    if (!this.capabilities.pivotDefinitionRead) return null;
    const name = this.pivotTableName(sheet, pivotIdx);
    const grandTotals = this.getPivotTableGrandTotals(sheet, pivotIdx);
    const rowFieldOrder = this.pivotRowFieldOrder(sheet, pivotIdx);
    const colFieldOrder = this.pivotColFieldOrder(sheet, pivotIdx);
    const cacheId = this.pivotTableCacheId(sheet, pivotIdx);
    if (name === null || !grandTotals || !rowFieldOrder || !colFieldOrder || cacheId < 0) {
      return null;
    }
    const fields: PivotFieldDefinition[] = [];
    const fieldCount = this.pivotFieldCount(sheet, pivotIdx);
    for (let i = 0; i < fieldCount; i += 1) {
      const field = this.getPivotField(sheet, pivotIdx, i);
      if (!field) return null;
      fields.push(field);
    }
    const dataFields: PivotDataFieldSpec[] = [];
    const dataFieldCount = this.pivotDataFieldCount(sheet, pivotIdx);
    for (let i = 0; i < dataFieldCount; i += 1) {
      const dataField = this.getPivotDataField(sheet, pivotIdx, i);
      if (!dataField) return null;
      dataFields.push(dataField);
    }
    return {
      name,
      cacheId,
      grandTotals,
      layout: this.getPivotReportLayout(sheet, pivotIdx),
      rowFieldOrder,
      colFieldOrder,
      fields,
      dataFields,
      filters: pivotFilterSpecs(this, sheet, pivotIdx),
    };
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

  getPivotReportLayout(sheet: number, pivotIdx: number): PivotReportLayout | null {
    assertAlive(this);
    if (!this.capabilities.pivotReportLayout) return null;
    const r = pivotWb(this).pivotGetLayout(sheet, pivotIdx);
    return r.status.ok ? r.layout : null;
  }

  setPivotReportLayout(sheet: number, pivotIdx: number, layout: PivotReportLayout): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotReportLayout) return false;
    return pivotWb(this).pivotSetLayout(sheet, pivotIdx, layout).ok;
  }

  pivotFieldCount(sheet: number, pivotIdx: number): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return 0;
    return unwrapNumberResult(pivotWb(this).pivotFieldCount(sheet, pivotIdx)) ?? 0;
  }

  addPivotField(sheet: number, pivotIdx: number, spec: PivotFieldSpec): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return -1;
    const engineSpec = toEnginePivotFieldSpec(this, spec);
    if (!engineSpec) return -1;
    const r = pivotWb(this).pivotFieldAdd(sheet, pivotIdx, engineSpec);
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
    order: PivotSortOrder,
    byField = '',
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotFieldSetSort(sheet, pivotIdx, fieldIdx, order, byField).ok;
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

  /** Append a manual-filter item addressed by its rendered label. The item
   *  carries no cache binding, so the filter engine matches source records by
   *  comparing their label against `name`; an empty `name` therefore names
   *  nothing. Use `addPivotFieldItemAt` for the blank member. */
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

  /** Append a manual-filter item addressed by its position in the bound cache
   *  field's shared items — the index space OOXML `<item x="N">` uses. This is
   *  the only form that can express the blank member. Returns false on an
   *  engine that has only the by-label form. */
  addPivotFieldItemAt(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    cacheIndex: number,
    visible: boolean,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotItemByCacheIndex) return false;
    const add = pivotWb(this).pivotFieldAddItemAt;
    if (typeof add !== 'function') return false;
    return add.call(pivotWb(this), sheet, pivotIdx, fieldIdx, cacheIndex, visible).ok;
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
    bounds: {
      startYear?: number;
      endYear?: number;
      intervalDays?: number;
      startSerial?: number;
      endSerial?: number;
    } = {},
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
      bounds.intervalDays ?? 1,
      bounds.startSerial ?? -1,
      bounds.endSerial ?? -1,
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
    const numFmtId = registerPivotNumberFormat(this, format);
    if (numFmtId === null) return false;
    return pivotWb(this).pivotFieldSetNumberFormat(sheet, pivotIdx, fieldIdx, numFmtId).ok;
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
    return unwrapNumberResult(pivotWb(this).pivotDataFieldCount(sheet, pivotIdx)) ?? 0;
  }

  addPivotDataField(sheet: number, pivotIdx: number, spec: PivotDataFieldSpec): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return -1;
    const engineSpec = toEnginePivotDataFieldSpec(this, spec);
    if (!engineSpec) return -1;
    const r = pivotWb(this).pivotDataFieldAdd(sheet, pivotIdx, engineSpec);
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
    const engineSpec = toEnginePivotDataFieldSpec(this, spec);
    if (!engineSpec) return false;
    return pivotWb(this).pivotDataFieldSet(sheet, pivotIdx, dataFieldIdx, engineSpec).ok;
  }

  clearPivotDataFields(sheet: number, pivotIdx: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return pivotWb(this).pivotDataFieldClear(sheet, pivotIdx).ok;
  }

  pivotFilterCount(sheet: number, pivotIdx: number): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return 0;
    return unwrapNumberResult(pivotWb(this).pivotFilterCount(sheet, pivotIdx)) ?? 0;
  }

  addPivotFilter(sheet: number, pivotIdx: number, spec: PivotFilterSpec): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    if (!isPivotFilterType(spec.type)) return false;
    const engineSpec = {
      axis: spec.axis,
      fieldName: spec.fieldName,
      type: spec.type as unknown as import('@libraz/formulon').PivotFilterType,
      ...(spec.dataFieldIndex !== undefined ? { dataFieldIndex: spec.dataFieldIndex } : {}),
      ...(spec.valueKind !== undefined ? { valueKind: spec.valueKind } : {}),
      ...(spec.valueInt !== undefined ? { valueInt: spec.valueInt } : {}),
      ...(spec.valueDouble !== undefined ? { valueDouble: spec.valueDouble } : {}),
      ...(spec.valueText !== undefined ? { valueText: spec.valueText } : {}),
      ...(spec.valueHighKind !== undefined ? { valueHighKind: spec.valueHighKind } : {}),
      ...(spec.valueHighInt !== undefined ? { valueHighInt: spec.valueHighInt } : {}),
      ...(spec.valueHighDouble !== undefined ? { valueHighDouble: spec.valueHighDouble } : {}),
    };
    return pivotWb(this).pivotFilterAdd(sheet, pivotIdx, engineSpec).ok;
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

/**
 * Shared items of every cache field a pivot names, each paired with the index
 * `<item x="N">` addresses it by.
 *
 * The blank member is kept. It carries no label of its own, so it can only be
 * named by that index — a filter spelled from an empty label matches nothing —
 * and dropping it here is what left it unreachable from the filter UI.
 *
 * Two shared items can render to the same label (`1` and `"1"`); the first
 * wins, so the list a user sees has no visible duplicates and every entry
 * still resolves to a real cache index.
 */
function pivotCacheSharedItemLabels(
  handle: WorkbookHandlePivotMethods,
  fields: ReadonlySet<string>,
  preferredCacheId: number,
): Map<string, PivotFieldItem[]> {
  const out = new Map<string, PivotFieldItem[]>();
  if (!handle.capabilities.pivotTableMutate) return out;
  const cacheIds = preferredCacheId >= 0 ? [preferredCacheId] : handle.pivotCacheIds();
  for (const cacheId of cacheIds) {
    for (const [fieldIdx, fieldName] of handle.pivotCacheFieldNames(cacheId).entries()) {
      if (!fields.has(fieldName) || out.has(fieldName)) continue;
      const seen = new Set<string>();
      const items: PivotFieldItem[] = [];
      for (const [cacheIndex, value] of handle.pivotCacheSharedItems(cacheId, fieldIdx).entries()) {
        const label = pivotCellItemLabel(value);
        if (seen.has(label)) continue;
        seen.add(label);
        items.push({ label, cacheIndex });
      }
      if (items.length > 0) out.set(fieldName, items);
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
  const count = unwrapNumberResult(wb.pivotFilterCount(sheet, pivotIndex));
  if (count === null || count === 0) return [];
  const out: PivotFilterSpec[] = [];
  for (let filterIndex = 0; filterIndex < count; filterIndex += 1) {
    // 0.12 exposes one coherent readback envelope. The old granular readers
    // were never part of the upstream Workbook surface and could mix values
    // from failed calls into an apparently valid filter.
    const direct = wb.pivotFilterAt(sheet, pivotIndex, filterIndex);
    const spec = direct.status.ok
      ? sanitizePivotFilterSpec(direct as unknown as Partial<PivotFilterSpec>)
      : null;
    if (spec) out.push(spec);
  }
  return out;
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
  const dataFieldIndex = spec.dataFieldIndex;
  return {
    axis: spec.axis,
    fieldName: spec.fieldName.trim(),
    type: spec.type,
    ...(typeof dataFieldIndex === 'number' && Number.isInteger(dataFieldIndex) && dataFieldIndex > 0
      ? { dataFieldIndex }
      : {}),
    ...filterPayload(spec.valueKind, spec.valueInt, spec.valueDouble, spec.valueText),
    ...filterPayload(
      spec.valueHighKind,
      spec.valueHighInt,
      spec.valueHighDouble,
      spec.valueHighText,
      'valueHigh',
    ),
  };
}

function filterPayload(
  kind: PivotFilterValueKind | undefined,
  intValue: number | undefined,
  doubleValue: number | undefined,
  textValue: string | undefined,
  prefix = 'value',
): Partial<PivotFilterSpec> {
  if (isPivotFilterValueKind(kind)) {
    if (kind === PivotFilterValueKind.None) return {};
    if (kind === PivotFilterValueKind.Int) {
      return {
        [`${prefix}Kind`]: kind,
        [`${prefix}Int`]: Number.isFinite(intValue) ? intValue : undefined,
      } as Partial<PivotFilterSpec>;
    }
    if (kind === PivotFilterValueKind.Double) {
      return {
        [`${prefix}Kind`]: kind,
        [`${prefix}Double`]: Number.isFinite(doubleValue) ? doubleValue : undefined,
      } as Partial<PivotFilterSpec>;
    }
    return {
      [`${prefix}Kind`]: kind,
      [`${prefix}Text`]: typeof textValue === 'string' ? textValue : undefined,
    } as Partial<PivotFilterSpec>;
  }
  // Keep hand-authored public specs that omit the discriminator usable.
  return {
    ...(Number.isFinite(intValue) ? { [`${prefix}Int`]: intValue } : {}),
    ...(Number.isFinite(doubleValue) ? { [`${prefix}Double`]: doubleValue } : {}),
    ...(typeof textValue === 'string' ? { [`${prefix}Text`]: textValue } : {}),
  } as Partial<PivotFilterSpec>;
}

function isPivotAxis(value: unknown): value is PivotAxis {
  return typeof value === 'number' && Object.values(PivotAxis).includes(value as PivotAxis);
}

type PivotNumberFormatApi = {
  addNumFmtCode?: (formatCode: string) => number;
};

/** Translate a public format code into the decimal id expected by 0.12. */
function registerPivotNumberFormat(handle: unknown, format: string): string | null {
  const code = format.trim();
  if (code.length === 0) return '';
  const addNumFmtCode = (handle as PivotNumberFormatApi).addNumFmtCode;
  if (typeof addNumFmtCode !== 'function') return null;
  const id = addNumFmtCode.call(handle, code);
  return Number.isInteger(id) && id >= 0 ? String(id) : null;
}

function toEnginePivotFieldSpec(
  handle: unknown,
  spec: PivotFieldSpec,
): {
  sourceName: string;
  customName?: string;
  axis: number;
  subtotalTop?: boolean;
  numberFormat?: string;
} | null {
  let numberFormat: string | undefined;
  if (spec.numberFormat !== undefined) {
    const registered = registerPivotNumberFormat(handle, spec.numberFormat);
    if (registered === null) return null;
    numberFormat = registered || undefined;
  }
  return {
    sourceName: spec.sourceName,
    ...(spec.customName !== undefined ? { customName: spec.customName } : {}),
    axis: spec.axis,
    ...(spec.subtotalTop !== undefined ? { subtotalTop: spec.subtotalTop } : {}),
    ...(numberFormat !== undefined ? { numberFormat } : {}),
  };
}

function toEnginePivotDataFieldSpec(
  handle: unknown,
  spec: PivotDataFieldSpec,
): {
  name: string;
  fieldIndex: number;
  aggregation: number;
  numberFormat?: string;
  showAs?: number;
  showAsBaseField?: number;
  showAsBaseItem?: number;
} | null {
  let numberFormat: string | undefined;
  if (spec.numberFormat !== undefined) {
    const registered = registerPivotNumberFormat(handle, spec.numberFormat);
    if (registered === null) return null;
    numberFormat = registered || undefined;
  }
  return {
    // An empty name has the engine name the field from the workbook profile.
    name: spec.name?.trim() ?? '',
    fieldIndex: spec.fieldIndex,
    aggregation: spec.aggregation,
    ...(numberFormat !== undefined ? { numberFormat } : {}),
    ...(spec.showValuesAs !== undefined ? { showAs: spec.showValuesAs } : {}),
    ...(spec.showAsBaseField !== undefined ? { showAsBaseField: spec.showAsBaseField } : {}),
    ...(spec.showAsBaseItem !== undefined ? { showAsBaseItem: spec.showAsBaseItem } : {}),
  };
}

function isPivotFilterType(value: unknown): value is PivotFilterType {
  return (
    typeof value === 'number' &&
    Number.isInteger(value) &&
    value >= PivotFilterType.ValueTop10 &&
    value <= PivotFilterType.LabelDate
  );
}

function isPivotFilterValueKind(value: unknown): value is PivotFilterValueKind {
  return (
    typeof value === 'number' &&
    Object.values(PivotFilterValueKind).includes(value as PivotFilterValueKind)
  );
}

function pivotCellEntry(
  handle: WorkbookHandlePivotMethods,
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
    numberFormat: pivotNumberFormatCode(handle, cell.numberFormat),
    pivotIndex,
  };
}

/** PivotCell.numberFormat is a decimal numFmtId in formulon 0.12. The cell
 * layer exposes a format code to the renderer/store. Keep a non-numeric value
 * intact for lightweight hosts that already hand us decoded test data. */
function pivotNumberFormatCode(handle: unknown, raw: string): string {
  const value = raw.trim();
  if (!value) return '';
  const id = Number(value);
  if (!Number.isInteger(id) || id < 0) return raw;
  const getNumFmtCode = (handle as { getNumFmtCode?: (numFmtId: number) => string | null })
    .getNumFmtCode;
  if (typeof getNumFmtCode !== 'function') return raw;
  return getNumFmtCode.call(handle, id) ?? '';
}

export function installPivotMethods(target: WorkbookHandleCtor): void {
  for (const key of Object.getOwnPropertyNames(WorkbookHandlePivotMethods.prototype)) {
    if (key === 'constructor') continue;
    const descriptor = Object.getOwnPropertyDescriptor(WorkbookHandlePivotMethods.prototype, key);
    if (!descriptor) continue;
    Object.defineProperty(target.prototype, key, descriptor);
  }
}
