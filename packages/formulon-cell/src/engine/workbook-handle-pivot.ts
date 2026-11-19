import type {
  CivilTime,
  CivilTimeInput,
  Formulon,
  Pivot,
  PivotCache,
  PivotCacheField,
  PivotCacheOptions,
  PivotCacheOptionsPatch,
  PivotOptions,
  PivotOptionsPatch,
  PivotWorksheetSourceInput,
} from '@libraz/formulon';
import { StatusCode } from '@libraz/formulon';
import { rangeOfRect } from './address.js';
import {
  isEmptyPivotLayout,
  PIVOT_CELL_KIND,
  PIVOT_PLACEHOLDER_COLS,
  PIVOT_PLACEHOLDER_ROWS,
  type PivotCellExpansion,
} from './pivot-cell-kind.js';
import type {
  PivotCellInfo,
  PivotDefinition,
  PivotFieldDefinition,
  PivotFieldItem,
  PivotItemDefinition,
  PivotNumericGroupSpec,
  PivotSpill,
  PivotStyleDefinition,
} from './pivot-definition.js';
import {
  DATE_GROUPING_PERIOD,
  dateGroupingOf,
  fromNativeValue,
  pivotCellItemLabel,
  toCacheValue,
  toDataFieldInput,
  toFilterInput,
  toFilterSpec,
} from './pivot-engine-codec.js';
import { isFormulonError, orFallback } from './result.js';
import { type SpillOccupantSource, spillBlockers, spillOccupancy } from './spill-blockers.js';
import type {
  Addr,
  CellValue,
  EngineCapabilities,
  PivotCalendar,
  PivotDataFieldSpec,
  PivotFieldSpec,
  PivotFilterSpec,
  PivotReportLayout,
  PivotWorksheetSource,
  Range,
  Workbook,
} from './types.js';
import {
  type PivotAggregation,
  type PivotAxis,
  type PivotDateGrouping,
  PivotShowValuesAs,
  type PivotSortOrder,
} from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  module: Formulon;
  wb: Workbook;
  capabilities: WorkbookHandle['capabilities'];
  assertAlive(): void;
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

function pivotOf(handle: unknown, sheet: number, pivotIdx: number): Pivot {
  return rawWb(handle).sheet(sheet).pivot(pivotIdx);
}

function cacheOf(handle: unknown, cacheId: number): PivotCache {
  return rawWb(handle).pivotCache(cacheId);
}

/** Number of reports on `sheet`; null when the sheet cannot be read. */
function pivotCount(handle: unknown, sheet: number): number | null {
  return orFallback(null, () => rawWb(handle).sheet(sheet).listPivots().length);
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

function rawWb(handle: unknown): Workbook {
  return internals(handle).wb;
}

/** One layout cell with its value converted for the cell layer. */
interface LayoutCell {
  row: number;
  col: number;
  value: CellValue;
  kind: number;
  depth: number;
  fieldName: string;
  numberFormat: string;
  /** The label's expand/collapse button; `null` for a cell without one. */
  expansion: PivotCellExpansion | null;
}

/** A report's layout as an anchor and extent; an extent of 0 spans no cell. */
interface RawPivotLayout {
  top: number;
  left: number;
  rows: number;
  cols: number;
  cells: readonly LayoutCell[];
}

/** The engine's layout of a report, or null when it cannot be read. */
function readLayout(handle: unknown, sheet: number, pivotIndex: number): RawPivotLayout | null {
  const layout = orFallback(null, () => pivotOf(handle, sheet, pivotIndex).getLayout());
  if (!layout) return null;
  const cells = layout.cells.map((c) => ({
    row: c.row,
    col: c.col,
    value: fromNativeValue(c.value),
    kind: c.kind,
    depth: c.depth,
    fieldName: c.fieldName ?? '',
    numberFormat: c.numberFormat ?? '',
    expansion: c.expansion,
  }));
  const r = layout.range;
  if (!r) return { top: 0, left: 0, rows: 0, cols: 0, cells };
  return {
    top: r.firstRow,
    left: r.firstCol,
    rows: r.lastRow - r.firstRow + 1,
    cols: r.lastCol - r.firstCol + 1,
    cells,
  };
}

/** The engine's layout of a report. A report with no fields occupies its
 *  placeholder area rather than the engine's caption extent, and a blocked
 *  report only its #SPILL! anchor; every extent consumer reads it from here. */
function effectiveLayout(
  handle: WorkbookHandlePivotMethods,
  sheet: number,
  pivotIndex: number,
): RawPivotLayout | null {
  const layout = readLayout(handle, sheet, pivotIndex);
  if (!layout || layout.rows <= 0 || layout.cols <= 0) return layout;
  if (orFallback(null, () => pivotOf(handle, sheet, pivotIndex).getSpill())?.blocked) {
    const anchor = { sheet, row: layout.top, col: layout.left };
    const cell: LayoutCell = {
      row: anchor.row,
      col: anchor.col,
      value: handle.getValue(anchor),
      kind: PIVOT_CELL_KIND.Blank,
      depth: 0,
      fieldName: '',
      numberFormat: '',
      expansion: null,
    };
    return { ...layout, rows: 1, cols: 1, cells: [cell] };
  }
  const empty = isEmptyPivotLayout(layout.cells.filter((c) => c.value.kind !== 'blank'));
  return empty ? { ...layout, rows: PIVOT_PLACEHOLDER_ROWS, cols: PIVOT_PLACEHOLDER_COLS } : layout;
}

export abstract class WorkbookHandlePivotMethods {
  declare readonly capabilities: EngineCapabilities;
  declare readonly sheetCount: number;
  abstract getValue(addr: Addr): CellValue;
  abstract physicalCells(
    sheet: number,
  ): Iterable<{ addr: Addr; value: CellValue; formula: string | null }>;

  abstract cellFormula(addr: Addr): string | null;

  /** Reports on `sheet` held at #SPILL!: by cells in the way, or by running past the sheet edge,
   *  in which case `needed` is the part inside the grid. */
  pivotSpills(sheet: number): PivotSpill[] {
    assertAlive(this);
    if (!this.capabilities.pivotTables) return [];
    const n = pivotCount(this, sheet);
    if (n === null) return [];
    const out: PivotSpill[] = [];
    for (let pivotIndex = 0; pivotIndex < n; pivotIndex += 1) {
      const spill = orFallback(null, () => pivotOf(this, sheet, pivotIndex).getSpill());
      if (!spill?.blocked) continue;
      const layout = readLayout(this, sheet, pivotIndex);
      const { range } = spill;
      const needed: Range = rangeOfRect(sheet, range);
      const anchor: Addr = layout
        ? { sheet, row: layout.top, col: layout.left }
        : { sheet, row: needed.r0, col: needed.c0 };
      const occupied = spillOccupancy(
        this as unknown as SpillOccupantSource,
        sheet,
        (row, col) => {
          const addr = { sheet, row, col };
          return { value: this.getValue(addr), formula: this.cellFormula(addr) };
        },
        pivotIndex,
      );
      out.push({ pivotIndex, anchor, needed, blockers: spillBlockers(needed, anchor, occupied) });
    }
    return out;
  }
  abstract getPivotStyleInfo(sheet: number, pivotIndex: number): PivotStyleDefinition | null;

  /** Whether any PivotTable on `range`'s sheet occupies a cell of `range`. */
  pivotOverlapsRange(range: Range): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTables) return false;
    rawWb(this);
    const n = pivotCount(this, range.sheet);
    if (n === null) return false;
    for (let i = 0; i < n; i += 1) {
      const layout = effectiveLayout(this, range.sheet, i);
      if (!layout) continue;
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
    rawWb(this);
    const n = pivotCount(this, sheet);
    if (n === null) return [];
    const out: { pivotIndex: number; range: Range }[] = [];
    for (let pivotIndex = 0; pivotIndex < n; pivotIndex += 1) {
      const layout = effectiveLayout(this, sheet, pivotIndex);
      if (!layout || layout.rows <= 0 || layout.cols <= 0) continue;
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
  pivotCellInfo(addr: Addr): PivotCellInfo | null {
    assertAlive(this);
    if (!this.capabilities.pivotTables) return null;
    rawWb(this);
    const n = pivotCount(this, addr.sheet);
    if (n === null) return null;
    for (let pivotIndex = 0; pivotIndex < n; pivotIndex += 1) {
      const layout = effectiveLayout(this, addr.sheet, pivotIndex);
      if (!layout) continue;
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
        isValue: kind === PIVOT_CELL_KIND.Data,
      };
    }
    return null;
  }

  /** Blank the sheet's own cells at `cells` without journaling them: values a
   *  report cached in the sheet and has since moved away from. The caller
   *  recalculates. */
  clearPivotCachedCells(sheet: number, cells: readonly { row: number; col: number }[]): void {
    assertAlive(this);
    const ws = internals(this).wb.sheet(sheet);
    for (const { row, col } of cells) ws.setBlank(row, col);
  }

  /** Iterate over evaluated PivotTable layout cells on a sheet. The engine
   *  returns sparse cells; blanks are skipped so existing empty-grid behavior
   *  remains unchanged. */
  *pivotCells(sheet: number): Generator<{
    addr: Addr;
    value: CellValue;
    formula: string | null;
    kind: number;
    /** Header nesting depth; 0 for data cells and the outermost field's items. */
    depth: number;
    /** Source field of a label, data field of a value; empty for captions. */
    fieldName: string;
    numberFormat: string;
    pivotIndex: number;
    /** A label's expand/collapse button; `null` for a cell without one. */
    expansion: PivotCellExpansion | null;
  }> {
    assertAlive(this);
    if (!this.capabilities.pivotTables) return;
    rawWb(this);
    const n = pivotCount(this, sheet);
    if (n === null) return;
    for (let i = 0; i < n; i += 1) {
      const layout = effectiveLayout(this, sheet, i);
      if (!layout) continue;
      const entries = [];
      for (const cell of layout.cells) {
        if (cell.value.kind === 'blank') continue;
        entries.push(pivotCellEntry(sheet, i, cell));
      }
      // The engine captions a report with no fields; the sheet shows it empty.
      if (isEmptyPivotLayout(entries)) continue;
      yield* entries;
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
    rawWb(this);
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
      const n = pivotCount(this, sheet);
      if (n === null) continue;
      for (let pivotIndex = 0; pivotIndex < n; pivotIndex += 1) {
        const layout = effectiveLayout(this, sheet, pivotIndex);
        if (!layout) continue;
        const fields = new Set<string>();
        const fieldItems = new Map<string, Set<string>>();
        for (const cell of layout.cells) {
          if (!cell.fieldName) continue;
          fields.add(cell.fieldName);
          const label = pivotCellItemLabel(cell.value);
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
    return orFallback(0, () => rawWb(this).listPivotCaches().length);
  }

  pivotCacheIds(): number[] {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return [];
    return orFallback([], () =>
      rawWb(this)
        .listPivotCaches()
        .map((cache) => cache.cacheId),
    );
  }

  createPivotCache(requestedId = 0): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return -1;
    // 0 asks the engine to assign the id.
    return orFallback(-1, () => rawWb(this).addPivotCache(requestedId || null).cacheId);
  }

  removePivotCache(cacheId: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      rawWb(this).removePivotCache(cacheId);
      return true;
    });
  }

  getPivotCacheWorksheetSource(cacheId: number): PivotWorksheetSource | null {
    assertAlive(this);
    if (!this.capabilities.pivotCacheSource) return null;
    return orFallback(null, () => cacheOf(this, cacheId).getWorksheetSource());
  }

  setPivotCacheWorksheetSource(cacheId: number, source: PivotWorksheetSourceInput | null): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotCacheSource) return false;
    return orFallback(false, () => {
      cacheOf(this, cacheId).setWorksheetSource(source);
      return true;
    });
  }

  pivotCacheFieldCount(cacheId: number): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return 0;
    return orFallback(0, () => cacheOf(this, cacheId).listFields().length);
  }

  /** The fields of a pivot cache with the value kinds their source column holds. */
  pivotCacheFields(cacheId: number): readonly PivotCacheField[] {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return [];
    return orFallback([], () => cacheOf(this, cacheId).listFields());
  }

  /** Rebuilds a cache from its worksheet source and re-projects its reports, as Refresh does.
   *  Returns the engine's status code when it refuses, else null. */
  refreshPivotCache(cacheId: number): number | null {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return StatusCode.NotImplemented;
    try {
      cacheOf(this, cacheId).refresh();
      return null;
    } catch (e) {
      if (!isFormulonError(e)) throw e;
      return e.code;
    }
  }

  pivotCacheFieldNames(cacheId: number): string[] {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return [];
    return orFallback([], () =>
      cacheOf(this, cacheId)
        .listFields()
        .map((field) => field.name),
    );
  }

  /** Shared items of a pivot cache field, in the index order `<item x="N">`
   *  addresses. */
  pivotCacheSharedItems(cacheId: number, fieldIdx: number): CellValue[] {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return [];
    return orFallback([], () =>
      cacheOf(this, cacheId)
        .listFieldSharedItems(fieldIdx)
        .map((item) => fromNativeValue(item.value)),
    );
  }

  addPivotCacheField(cacheId: number, name: string): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return -1;
    return orFallback(-1, () => cacheOf(this, cacheId).addField(name));
  }

  clearPivotCacheFields(cacheId: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      cacheOf(this, cacheId).clearFields();
      return true;
    });
  }

  addPivotCacheSharedItem(cacheId: number, fieldIdx: number, value: CellValue): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      cacheOf(this, cacheId).addFieldSharedItem(fieldIdx, toCacheValue(value));
      return true;
    });
  }

  clearPivotCacheSharedItems(cacheId: number, fieldIdx: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      cacheOf(this, cacheId).clearFieldSharedItems(fieldIdx);
      return true;
    });
  }

  addPivotCacheRecord(cacheId: number): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return -1;
    return orFallback(-1, () => cacheOf(this, cacheId).addRecord());
  }

  clearPivotCacheRecords(cacheId: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      cacheOf(this, cacheId).clearRecords();
      return true;
    });
  }

  setPivotCacheRecordValue(
    cacheId: number,
    recordIdx: number,
    fieldIdx: number,
    value: CellValue,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      cacheOf(this, cacheId).setRecordValue(recordIdx, fieldIdx, toCacheValue(value));
      return true;
    });
  }

  createPivotTable(
    sheet: number,
    name: string,
    cacheId: number,
    anchor: { row: number; col: number },
  ): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return -1;
    return orFallback(
      -1,
      () =>
        rawWb(this)
          .sheet(sheet)
          .addPivot({ name, cacheId, anchor: { row: anchor.row, col: anchor.col } }).index,
    );
  }

  pivotTableCacheId(sheet: number, pivotIdx: number): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return -1;
    return orFallback(-1, () => pivotOf(this, sheet, pivotIdx).cache().cacheId);
  }

  /** The PivotTable's stored name, or null when it cannot be read. */
  pivotTableName(sheet: number, pivotIdx: number): string | null {
    assertAlive(this);
    if (!this.capabilities.pivotDefinitionRead) return null;
    return orFallback(null, () => pivotOf(this, sheet, pivotIdx).name);
  }

  getPivotTableGrandTotals(
    sheet: number,
    pivotIdx: number,
  ): { rows: boolean; cols: boolean } | null {
    assertAlive(this);
    if (!this.capabilities.pivotDefinitionRead) return null;
    return orFallback(null, () => {
      const totals = pivotOf(this, sheet, pivotIdx).getGrandTotals();
      return { rows: totals.rows, cols: totals.columns };
    });
  }

  /** Pivot field indices on the row axis, in display order; null when unreadable. */
  pivotRowFieldOrder(sheet: number, pivotIdx: number): number[] | null {
    assertAlive(this);
    if (!this.capabilities.pivotDefinitionRead) return null;
    return orFallback(null, () => [...pivotOf(this, sheet, pivotIdx).getRowFieldOrder()]);
  }

  /** Pivot field indices on the column axis, in display order; null when unreadable. */
  pivotColFieldOrder(sheet: number, pivotIdx: number): number[] | null {
    assertAlive(this);
    if (!this.capabilities.pivotDefinitionRead) return null;
    return orFallback(null, () => [...pivotOf(this, sheet, pivotIdx).getColumnFieldOrder()]);
  }

  getPivotField(sheet: number, pivotIdx: number, fieldIdx: number): PivotFieldDefinition | null {
    assertAlive(this);
    if (!this.capabilities.pivotDefinitionRead) return null;
    return orFallback(null, () => {
      const pivot = pivotOf(this, sheet, pivotIdx);
      const r = pivot.getField(fieldIdx);
      const items: PivotItemDefinition[] = pivot.listFieldItems(fieldIdx).map((item) => ({
        label: item.name ?? '',
        cacheIndex: item.cacheIndex ?? -1,
        visible: item.visible,
        showDetail: item.showDetail,
      }));
      const dateGroup = pivot.getFieldDateGroup(fieldIdx);
      const granularity = dateGroup === null ? null : dateGroupingOf(dateGroup.periods);
      const numericGroup = pivot.getFieldNumericGroup(fieldIdx);
      const serialOf = (date: CivilTime): number =>
        internals(this).module.convertCivilToSerial(date, {
          date1904: internals(this).wb.getDate1904(),
        });
      return {
        index: fieldIdx,
        sourceName: r.sourceName,
        customName: r.customName ?? '',
        axis: r.axis,
        subtotalTop: r.subtotalTop,
        defaultSubtotal: r.defaultSubtotal,
        subtotalFns: [...r.subtotalFunctions],
        numberFormat: r.numberFormat ?? '',
        cacheFieldIndex: r.cacheFieldIndex ?? -1,
        sortOrder: r.sortOrder,
        sortByField: r.sortByField ?? '',
        dateGroup:
          dateGroup === null || granularity === null
            ? null
            : {
                granularity,
                periods: dateGroup.periods,
                calendar: dateGroup.calendar,
                startYear: -1,
                endYear: -1,
                intervalDays: dateGroup.intervalDays,
                startSerial: dateGroup.autoStart ? -1 : serialOf(dateGroup.start),
                endSerial: dateGroup.autoEnd ? -1 : serialOf(dateGroup.end),
              },
        groupBaseIndex: r.groupBaseIndex,
        numericGroup:
          numericGroup === null
            ? null
            : {
                interval: numericGroup.interval,
                start: numericGroup.autoStart ? null : numericGroup.start,
                end: numericGroup.autoEnd ? null : numericGroup.end,
              },
        items,
        showDetail: r.showDetail,
      };
    });
  }

  getPivotDataField(
    sheet: number,
    pivotIdx: number,
    dataFieldIdx: number,
  ): PivotDataFieldSpec | null {
    assertAlive(this);
    if (!this.capabilities.pivotDefinitionRead) return null;
    return orFallback(null, () => {
      const r = pivotOf(this, sheet, pivotIdx).getDataField(dataFieldIdx);
      return {
        name: r.name,
        fieldIndex: r.fieldIndex,
        aggregation: r.aggregation,
        ...(r.numberFormat ? { numberFormat: r.numberFormat } : {}),
        ...(r.showAs !== PivotShowValuesAs.Normal ? { showValuesAs: r.showAs } : {}),
        ...(r.showAsBaseField !== null ? { showAsBaseField: r.showAsBaseField } : {}),
        ...(r.showAsBaseItem !== null ? { showAsBaseItem: r.showAsBaseItem } : {}),
      };
    });
  }

  /** The whole stored definition of a PivotTable, or null when the engine
   *  cannot read it back (a pivot that does not exist). */
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
      style: this.getPivotStyleInfo(sheet, pivotIdx),
    };
  }

  removePivotTable(sheet: number, pivotIdx: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      rawWb(this).sheet(sheet).removePivot(pivotIdx);
      return true;
    });
  }

  renamePivotTable(sheet: number, pivotIdx: number, name: string): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).rename(name);
      return true;
    });
  }

  /** Move a report's top-left cell to `anchor`; the report's extent follows its layout. */
  setPivotTableAnchor(
    sheet: number,
    pivotIdx: number,
    anchor: { row: number; col: number; rows: number; cols: number },
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).setAnchor({ row: anchor.row, col: anchor.col });
      return true;
    });
  }

  setPivotTableGrandTotals(
    sheet: number,
    pivotIdx: number,
    rowsEnabled: boolean,
    colsEnabled: boolean,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).updateGrandTotals({ rows: rowsEnabled, columns: colsEnabled });
      return true;
    });
  }

  /** Whether the report fits its column widths when it updates; true when unreadable. */
  pivotAutofitColumns(sheet: number, pivotIdx: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotDefinitionRead) return true;
    return orFallback(true, () => pivotOf(this, sheet, pivotIdx).getOptions().autoFormatting);
  }

  /** The report's PivotTable Options; null when unreadable. */
  pivotOptions(sheet: number, pivotIdx: number): PivotOptions | null {
    assertAlive(this);
    if (!this.capabilities.pivotDefinitionRead) return null;
    return orFallback(null, () => pivotOf(this, sheet, pivotIdx).getOptions());
  }

  /** Writes the members of `patch` to the report's options; an omitted member keeps its value. */
  updatePivotOptions(sheet: number, pivotIdx: number, patch: PivotOptionsPatch): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).updateOptions(patch);
      return true;
    });
  }

  /** The options of the cache the report reads; null when unreadable. */
  pivotCacheOptions(sheet: number, pivotIdx: number): PivotCacheOptions | null {
    assertAlive(this);
    if (!this.capabilities.pivotDefinitionRead) return null;
    return orFallback(null, () => pivotOf(this, sheet, pivotIdx).cache().getOptions());
  }

  /** Writes the members of `patch` to the options of the cache the report reads. */
  updatePivotCacheOptions(sheet: number, pivotIdx: number, patch: PivotCacheOptionsPatch): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).cache().updateOptions(patch);
      return true;
    });
  }

  getPivotReportLayout(sheet: number, pivotIdx: number): PivotReportLayout | null {
    assertAlive(this);
    if (!this.capabilities.pivotReportLayout) return null;
    return orFallback(null, () => pivotOf(this, sheet, pivotIdx).getReportLayout());
  }

  setPivotReportLayout(sheet: number, pivotIdx: number, layout: PivotReportLayout): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotReportLayout) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).setReportLayout(layout);
      return true;
    });
  }

  pivotFieldCount(sheet: number, pivotIdx: number): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return 0;
    return orFallback(0, () => pivotOf(this, sheet, pivotIdx).listFields().length);
  }

  /** Place the pivot field mirroring cache field `spec.sourceName` on `spec.axis`
   *  with the given settings. Returns its field index, or -1 on failure. */
  addPivotField(sheet: number, pivotIdx: number, spec: PivotFieldSpec): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return -1;
    return orFallback(-1, () => {
      const pivot = pivotOf(this, sheet, pivotIdx);
      const index = pivot.listFields().findIndex((f) => f.sourceName === spec.sourceName);
      if (index < 0) return -1;
      pivot.setFieldAxis(index, spec.axis);
      pivot.updateField(index, {
        ...(spec.customName !== undefined ? { customName: spec.customName || null } : {}),
        ...(spec.subtotalTop !== undefined ? { subtotalTop: spec.subtotalTop } : {}),
        ...(spec.numberFormat !== undefined ? { numberFormat: spec.numberFormat || null } : {}),
      });
      return index;
    });
  }

  clearPivotFields(sheet: number, pivotIdx: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).clearFields();
      return true;
    });
  }

  /** Move a field onto `axis`; null takes it off every axis, dropping its data
   *  fields and the filters naming it. */
  setPivotFieldAxis(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    axis: PivotAxis | null,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).setFieldAxis(fieldIdx, axis);
      return true;
    });
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
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).updateField(fieldIdx, {
        sortOrder: order,
        sortByField: byField || null,
      });
      return true;
    });
  }

  setPivotFieldSubtotalTop(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    top: boolean,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).updateField(fieldIdx, { subtotalTop: top });
      return true;
    });
  }

  /** Append a manual-filter item addressed by its rendered label. The filter
   *  engine matches source records by comparing their label against `name`;
   *  an empty `name` therefore names nothing. Use `addPivotFieldItemAt` for
   *  the blank member. */
  addPivotFieldItem(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    name: string,
    visible: boolean,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).addFieldItem(fieldIdx, name, { visible });
      return true;
    });
  }

  /** Append a manual-filter item addressed by its position in the bound cache
   *  field's shared items — the index space OOXML `<item x="N">` uses. This is
   *  the only form that can express the blank member. */
  addPivotFieldItemAt(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    cacheIndex: number,
    visible: boolean,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotItemByCacheIndex) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).addFieldItem(fieldIdx, { cacheIndex }, { visible });
      return true;
    });
  }

  clearPivotFieldItems(sheet: number, pivotIdx: number, fieldIdx: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).clearFieldItems(fieldIdx);
      return true;
    });
  }

  /** Expands (`show`) or collapses one item of a row or column field. */
  setPivotFieldItemShowDetail(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    itemIdx: number,
    show: boolean,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).updateFieldItem(fieldIdx, itemIdx, { showDetail: show });
      return true;
    });
  }

  /** Expands (`show`) or collapses every item of a row or column field. */
  setPivotFieldShowDetail(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    show: boolean,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).updateField(fieldIdx, { showDetail: show });
      return true;
    });
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
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).updateFieldItem(fieldIdx, itemIdx, { visible });
      return true;
    });
  }

  addPivotFieldSubtotalFn(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    agg: PivotAggregation,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      const pivot = pivotOf(this, sheet, pivotIdx);
      const current = pivot.getField(fieldIdx).subtotalFunctions;
      pivot.setFieldSubtotalFunctions(fieldIdx, [...current, agg]);
      return true;
    });
  }

  clearPivotFieldSubtotalFns(sheet: number, pivotIdx: number, fieldIdx: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).setFieldSubtotalFunctions(fieldIdx, []);
      return true;
    });
  }

  setPivotFieldDateGroup(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    granularity: PivotDateGrouping,
    calendar: PivotCalendar,
    bounds: {
      /** Every period to group by; the granularity's own period when omitted. */
      periods?: number;
      startYear?: number;
      endYear?: number;
      intervalDays?: number;
      startSerial?: number;
      endSerial?: number;
    } = {},
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      const start = dateBound(this, bounds.startSerial, bounds.startYear, 'start');
      const end = dateBound(this, bounds.endSerial, bounds.endYear, 'end');
      pivotOf(this, sheet, pivotIdx).setFieldDateGroup(fieldIdx, {
        periods: bounds.periods ?? DATE_GROUPING_PERIOD[granularity],
        calendar,
        ...(bounds.intervalDays !== undefined ? { intervalDays: bounds.intervalDays } : {}),
        ...(start ? { start } : {}),
        ...(end ? { end } : {}),
      });
      return true;
    });
  }

  clearPivotFieldDateGroup(sheet: number, pivotIdx: number, fieldIdx: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).setFieldDateGroup(fieldIdx, null);
      return true;
    });
  }

  /** Groups a number field into ranges, or ungroups it with `null`. */
  setPivotFieldNumericGroup(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    group: PivotNumericGroupSpec | null,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).setFieldNumericGroup(
        fieldIdx,
        group === null
          ? null
          : {
              interval: group.interval,
              ...(group.start !== null ? { start: group.start } : {}),
              ...(group.end !== null ? { end: group.end } : {}),
            },
      );
      return true;
    });
  }

  setPivotFieldNumberFormat(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    format: string,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).updateField(fieldIdx, {
        numberFormat: format.trim() || null,
      });
      return true;
    });
  }

  setPivotRowFieldOrder(sheet: number, pivotIdx: number, indices: readonly number[]): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).setRowFieldOrder(indices);
      return true;
    });
  }

  setPivotColFieldOrder(sheet: number, pivotIdx: number, indices: readonly number[]): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).setColumnFieldOrder(indices);
      return true;
    });
  }

  pivotDataFieldCount(sheet: number, pivotIdx: number): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return 0;
    return orFallback(0, () => pivotOf(this, sheet, pivotIdx).listDataFields().length);
  }

  addPivotDataField(sheet: number, pivotIdx: number, spec: PivotDataFieldSpec): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return -1;
    return orFallback(-1, () =>
      pivotOf(this, sheet, pivotIdx).addDataField(toDataFieldInput(spec)),
    );
  }

  setPivotDataField(
    sheet: number,
    pivotIdx: number,
    dataFieldIdx: number,
    spec: PivotDataFieldSpec,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).setDataField(dataFieldIdx, toDataFieldInput(spec));
      return true;
    });
  }

  clearPivotDataFields(sheet: number, pivotIdx: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).clearDataFields();
      return true;
    });
  }

  pivotFilterCount(sheet: number, pivotIdx: number): number {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return 0;
    return orFallback(0, () => pivotOf(this, sheet, pivotIdx).listFilters().length);
  }

  addPivotFilter(sheet: number, pivotIdx: number, spec: PivotFilterSpec): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).addFilter(toFilterInput(spec));
      return true;
    });
  }

  clearPivotFilters(sheet: number, pivotIdx: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).clearFilters();
      return true;
    });
  }

  removePivotFilter(sheet: number, pivotIdx: number, filterIdx: number): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotTableMutate) return false;
    return orFallback(false, () => {
      pivotOf(this, sheet, pivotIdx).removeFilter(filterIdx);
      return true;
    });
  }
}

/** A date group bound from a serial, else from a year; null leaves it automatic. */
function dateBound(
  handle: unknown,
  serial: number | undefined,
  year: number | undefined,
  edge: 'start' | 'end',
): CivilTimeInput | null {
  if (serial !== undefined && serial >= 0) {
    const date1904 = internals(handle).wb.getDate1904();
    return internals(handle).module.convertSerialToCivil(serial, { date1904 });
  }
  if (year !== undefined && year >= 0) {
    return edge === 'start' ? { year, month: 1, day: 1 } : { year, month: 12, day: 31 };
  }
  return null;
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
  const read = orFallback(null, () => {
    const pivot = pivotOf(handle, sheet, pivotIndex);
    const filters = pivot.listFilters();
    return filters.length === 0 ? null : { filters, fields: pivot.listFields() };
  });
  if (!read) return [];
  const out: PivotFilterSpec[] = [];
  for (const filter of read.filters) {
    const field = read.fields[filter.fieldIndex];
    const spec = field ? toFilterSpec(filter, field.axis, field.sourceName) : null;
    if (spec) out.push(spec);
  }
  return out;
}

function pivotCellEntry(
  sheet: number,
  pivotIndex: number,
  cell: LayoutCell,
): {
  addr: Addr;
  value: CellValue;
  formula: string | null;
  kind: number;
  depth: number;
  fieldName: string;
  numberFormat: string;
  pivotIndex: number;
  expansion: PivotCellExpansion | null;
} {
  return {
    addr: { sheet, row: cell.row, col: cell.col },
    value: cell.value,
    formula: null,
    kind: cell.kind,
    depth: cell.depth,
    fieldName: cell.fieldName,
    numberFormat: cell.numberFormat,
    pivotIndex,
    expansion: cell.expansion,
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
