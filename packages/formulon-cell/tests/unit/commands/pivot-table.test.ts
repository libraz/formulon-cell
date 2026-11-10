import { type PivotWorksheetSourceInput, StatusCode } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import {
  createPivotTableFromRange,
  executeRibbonPivotTableAction,
  refreshPivotCacheFromRange,
  refreshPivotTable,
  refreshPivotTableFromRange,
} from '../../../src/commands/pivot-table.js';
import { formatA1Range } from '../../../src/engine/address.js';
import {
  type CellValue,
  PIVOT_SHOW_AS_BASE_PREVIOUS,
  PivotAggregation,
  PivotAxis,
  PivotFilterType,
  PivotFilterValueKind,
  PivotShowValuesAs,
  PivotSortOrder,
} from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const blank: CellValue = { kind: 'blank' };
const text = (value: string): CellValue => ({ kind: 'text', value });
const num = (value: number): CellValue => ({ kind: 'number', value });

const makeWb = () => {
  const values = new Map<string, CellValue>([
    ['0:0', text('Region')],
    ['0:1', text('Product')],
    ['0:2', text('Sales')],
    ['1:0', text('East')],
    ['1:1', text('Desk')],
    ['1:2', num(12)],
    ['2:0', text('West')],
    ['2:1', text('Chair')],
    ['2:2', num(8)],
  ]);
  const calls: string[] = [];
  let sourceRows = 0;
  const wb = {
    capabilities: { pivotTableMutate: true, pivotCacheSource: true },
    sheetCount: 1,
    sheetName: (sheet: number) => (sheet === 0 ? 'Sheet1' : `Sheet${sheet + 1}`),
    getValue: ({ row, col }: { row: number; col: number }) => values.get(`${row}:${col}`) ?? blank,
    createPivotCache: () => {
      calls.push('cache');
      return 4;
    },
    removePivotCache: () => {
      calls.push('remove-cache');
      return true;
    },
    setPivotCacheWorksheetSource: (_cache: number, source: PivotWorksheetSourceInput | null) => {
      const range = typeof source?.range === 'object' ? source.range : null;
      const ref = range
        ? formatA1Range(
            {
              r0: range.rect.firstRow,
              c0: range.rect.firstCol,
              r1: range.rect.lastRow,
              c1: range.rect.lastCol,
            },
            { collapse: false },
          )
        : '';
      calls.push(
        `cache-source:${source !== null}:${range ? `Sheet${range.sheet + 1}` : ''}:${ref}`,
      );
      sourceRows = range ? range.rect.lastRow - range.rect.firstRow + 1 : 0;
      return true;
    },
    pivotBounds: () => [],
    // The engine refuses a source past its record cell limit.
    refreshPivotCache: () => {
      calls.push('refresh');
      return sourceRows > 100_000 ? StatusCode.SecurityResourceLimit : null;
    },
    pivotCacheFields: () => [
      { name: 'Region', containsNumber: false },
      { name: 'Product', containsNumber: false },
      { name: 'Sales', containsNumber: true },
    ],
    pivotCacheSharedItems: (_cache: number, field: number) =>
      [
        [text('East'), text('West')],
        [text('Desk'), text('Chair')],
        [num(12), num(8)],
      ][field] ?? [],
    pivotTableCacheId: (_sheet: number, pivotIndex: number) => (pivotIndex === 2 ? 4 : -1),
    pivotFieldCount: () => 0,
    pivotTableName: () => 'PivotTable1',
    getPivotTables: () => [
      {
        sheetIndex: 0,
        pivotIndex: 2,
        top: 5,
        left: 0,
        rows: 3,
        cols: 3,
        cells: 9,
        fields: ['Region', 'Product', 'Sales'],
      },
    ],
    cells: () => [],
    getPivotCacheWorksheetSource: (_cache: number) => ({
      range: { sheet: 0, rect: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 2 } },
      name: null,
    }),
    createPivotTable: (_sheet: number, name: string, cacheId: number, anchor: unknown) => {
      calls.push(`pivot:${name}:${cacheId}:${JSON.stringify(anchor)}`);
      return 2;
    },
    removePivotTable: () => true,
    setPivotTableGrandTotals: (_sheet: number, _pivot: number, rows: boolean, cols: boolean) => {
      calls.push(`grand:${rows}:${cols}`);
      return true;
    },
    setPivotFieldAxis: (_sheet: number, _pivot: number, field: number, axis: number | null) => {
      calls.push(`pivot-field-axis:${field}:${axis}`);
      return true;
    },
    addPivotField: (
      _sheet: number,
      _pivot: number,
      spec: { sourceName: string; axis: number; subtotalTop?: boolean },
    ) => {
      calls.push(`pivot-field:${spec.sourceName}:${spec.axis}`);
      if (spec.subtotalTop === false) calls.push(`subtotal-top:${spec.sourceName}:false`);
      return calls.filter((c) => c.startsWith('pivot-field:')).length - 1;
    },
    addPivotFieldItem: (
      _sheet: number,
      _pivot: number,
      fieldIdx: number,
      name: string,
      visible: boolean,
    ) => {
      calls.push(`pivot-item:${fieldIdx}:${name}:${visible}`);
      return true;
    },
    clearPivotFieldItems: (_sheet: number, _pivot: number, fieldIdx: number) => {
      calls.push(`clear-items:${fieldIdx}`);
      return true;
    },
    addPivotFilter: (_sheet: number, _pivot: number, spec: { fieldName: string; type: number }) => {
      calls.push(`pivot-filter:${spec.fieldName}:${spec.type}`);
      return true;
    },
    setPivotFieldSort: (
      _sheet: number,
      _pivot: number,
      fieldIdx: number,
      order: PivotSortOrder,
      byField: string,
    ) => {
      calls.push(`sort:${fieldIdx}:${order}:${byField}`);
      return true;
    },
    addPivotDataField: (
      _sheet: number,
      _pivot: number,
      spec: { name?: string; fieldIndex: number; aggregation: number; numberFormat?: string },
    ) => {
      calls.push(
        [
          'data-field',
          spec.name ?? '',
          String(spec.fieldIndex),
          String(spec.aggregation),
          spec.numberFormat ?? '',
          String((spec as { showValuesAs?: number }).showValuesAs ?? ''),
        ].join(':'),
      );
      const base = spec as { showAsBaseField?: number; showAsBaseItem?: number };
      if (base.showAsBaseField !== undefined || base.showAsBaseItem !== undefined) {
        calls.push(
          `data-field-base:${spec.name ?? ''}:${base.showAsBaseField}:${base.showAsBaseItem}`,
        );
      }
      return 0;
    },
  } as unknown as WorkbookHandle;
  return { wb, calls };
};

describe('pivot-table command helpers', () => {
  it('creates a cache, records, pivot fields, and a data field from a range', () => {
    const { wb, calls } = makeWb();
    const result = createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
      destination: { sheet: 0, row: 5, col: 0 },
      name: 'SalesPivot',
      rowField: 'Region',
      columnField: 'Product',
      valueField: 'Sales',
      aggregation: PivotAggregation.Sum,
      showRowGrandTotals: false,
      rowSort: 'asc',
      columnSort: 'desc',
      rowSubtotalTop: false,
      valueNumberFormat: '#,##0',
    });

    expect(result).toEqual({ ok: true, cacheId: 4, pivotIndex: 2, sharedCache: false });
    expect(calls).toContain('cache-source:true:Sheet1:A1:C3');
    expect(calls).toContain('refresh');
    expect(calls).toContain('pivot-field:Region:0');
    expect(calls).toContain('grand:false:true');
    expect(calls).toContain('subtotal-top:Region:false');
    expect(calls).toContain(`sort:0:${PivotSortOrder.Ascending}:`);
    expect(calls).toContain(`sort:1:${PivotSortOrder.Descending}:`);
    expect(calls).toContain('pivot-field:Product:1');
    expect(calls).toContain('pivot-field:Sales:2');
    expect(calls).toContain('data-field::2:0:#,##0:');
  });

  it('projects source row and column labels through the real WASM engine', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.isStub).toBe(false);
      expect(wb.capabilities.pivotTableMutate).toBe(true);
      wb.setText({ sheet: 0, row: 0, col: 0 }, 'Region');
      wb.setText({ sheet: 0, row: 0, col: 1 }, 'Product');
      wb.setText({ sheet: 0, row: 0, col: 2 }, 'Sales');
      wb.setText({ sheet: 0, row: 1, col: 0 }, 'East');
      wb.setText({ sheet: 0, row: 1, col: 1 }, 'Desk');
      wb.setNumber({ sheet: 0, row: 1, col: 2 }, 12);
      wb.setText({ sheet: 0, row: 2, col: 0 }, 'West');
      wb.setText({ sheet: 0, row: 2, col: 1 }, 'Chair');
      wb.setNumber({ sheet: 0, row: 2, col: 2 }, 8);

      const result = createPivotTableFromRange(wb, {
        source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
        destination: { sheet: 0, row: 5, col: 0 },
        rowField: 'Region',
        columnField: 'Product',
        valueField: 'Sales',
        aggregation: PivotAggregation.Sum,
      });
      expect(result).toMatchObject({ ok: true });
      if (!result.ok) return;

      const projected = [...wb.pivotCells(0)];
      const labels = projected
        .filter((cell) => cell.value.kind === 'text')
        .map((cell) => (cell.value.kind === 'text' ? cell.value.value : ''));
      const numbers = projected
        .filter((cell) => cell.value.kind === 'number')
        .map((cell) => (cell.value.kind === 'number' ? cell.value.value : NaN));
      expect(labels).toEqual(expect.arrayContaining(['East', 'West', 'Desk', 'Chair']));
      expect(numbers).toEqual(expect.arrayContaining([12, 8, 20]));
    } finally {
      wb.dispose();
    }
  });

  it('uses the selected aggregation in the real data-field name', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.isStub).toBe(false);
      expect(wb.capabilities.pivotTableMutate).toBe(true);
      wb.setText({ sheet: 0, row: 0, col: 0 }, 'Region');
      wb.setText({ sheet: 0, row: 0, col: 1 }, 'Sales');
      wb.setText({ sheet: 0, row: 1, col: 0 }, 'East');
      wb.setNumber({ sheet: 0, row: 1, col: 1 }, 2);
      wb.setText({ sheet: 0, row: 2, col: 0 }, 'East');
      wb.setNumber({ sheet: 0, row: 2, col: 1 }, 3);

      const result = createPivotTableFromRange(wb, {
        source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
        destination: { sheet: 0, row: 5, col: 0 },
        rowField: 'Region',
        valueField: 'Sales',
        aggregation: PivotAggregation.Product,
      });
      expect(result).toMatchObject({ ok: true });
      if (!result.ok) return;

      const projected = [...wb.pivotCells(0)];
      const labels = projected
        .filter((cell) => cell.value.kind === 'text')
        .map((cell) => (cell.value.kind === 'text' ? cell.value.value : ''));
      const numbers = projected
        .filter((cell) => cell.value.kind === 'number')
        .map((cell) => (cell.value.kind === 'number' ? cell.value.value : NaN));
      expect(labels).toContain('Product of Sales');
      expect(numbers).toContain(6);
    } finally {
      wb.dispose();
    }
  });

  it('refuses a source the engine finds too large and removes the cache', () => {
    const { wb, calls } = makeWb();
    const result = createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 1048575, c1: 2 },
      destination: { sheet: 0, row: 5, col: 0 },
      rowField: 'Region',
      valueField: 'Sales',
    });

    expect(result).toEqual({ ok: false, reason: 'invalid-range', step: 'refresh' });
    expect(calls.at(-1)).toBe('remove-cache');
    expect(calls.some((c) => c.startsWith('pivot:'))).toBe(false);
  });

  it('adds a page/filter field when requested', () => {
    const { wb, calls } = makeWb();
    const result = createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
      destination: { sheet: 0, row: 5, col: 0 },
      rowField: 'Region',
      filterField: 'Product',
      valueField: 'Sales',
    });

    expect(result.ok).toBe(true);
    expect(calls).toContain('pivot-field:Product:3');
  });

  it('adds page/filter field item visibility when requested', () => {
    const { wb, calls } = makeWb();
    const result = createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
      destination: { sheet: 0, row: 5, col: 0 },
      rowField: 'Region',
      filterField: 'Product',
      filterItems: [
        { fieldName: 'Product', itemName: 'Desk', visible: true },
        { fieldName: 'Product', itemName: 'Chair', visible: false },
      ],
      valueField: 'Sales',
    });

    expect(result.ok).toBe(true);
    expect(calls).toContain('pivot-field:Product:3');
    expect(calls.indexOf('clear-items:1')).toBeGreaterThanOrEqual(0);
    expect(calls.indexOf('clear-items:1')).toBeLessThan(calls.indexOf('pivot-item:1:Desk:true'));
    expect(calls).toContain('pivot-item:1:Chair:false');
  });

  it('adds page/filter condition specs when requested', () => {
    const { wb, calls } = makeWb();
    const result = createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
      destination: { sheet: 0, row: 5, col: 0 },
      rowField: 'Region',
      filterField: 'Product',
      pivotFilters: [
        {
          axis: PivotAxis.Page,
          fieldName: 'Product',
          type: PivotFilterType.LabelContains,
          valueKind: PivotFilterValueKind.Text,
          valueText: 'Desk',
        },
      ],
      valueField: 'Sales',
    });

    expect(result.ok).toBe(true);
    expect(calls).toContain('pivot-filter:Product:3');
  });

  it('adds multiple page/filter fields when requested', () => {
    const { wb, calls } = makeWb();
    const result = createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
      destination: { sheet: 0, row: 5, col: 0 },
      rowField: 'Region',
      filterField: 'Product',
      filterFields: ['Product', 'Sales'],
      valueField: 'Sales',
    });

    expect(result.ok).toBe(true);
    expect(calls).toContain('pivot-field:Product:3');
    expect(calls).toContain('pivot-field:Sales:3');
  });

  it('creates multiple value/data fields when requested', () => {
    const { wb, calls } = makeWb();
    const result = createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
      destination: { sheet: 0, row: 5, col: 0 },
      rowField: 'Region',
      valueField: 'Sales',
      valueFields: ['Sales', 'Product'],
      aggregation: PivotAggregation.Count,
    });

    expect(result.ok).toBe(true);
    expect(calls).toContain('pivot-field:Sales:2');
    expect(calls).toContain('pivot-field:Product:2');
    expect(calls).toContain('data-field::2:1::');
    expect(calls).toContain('data-field::1:1::');
  });

  it('applies value field aggregation and number format per field', () => {
    const { wb, calls } = makeWb();
    const result = createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
      destination: { sheet: 0, row: 5, col: 0 },
      rowField: 'Region',
      valueField: 'Sales',
      valueFields: ['Sales', 'Product'],
      aggregation: PivotAggregation.Sum,
      valueNumberFormat: '#,##0',
      valueFieldSettings: [
        {
          fieldName: 'Sales',
          aggregation: PivotAggregation.Average,
          numberFormat: '#,##0.00',
          showValuesAs: PivotShowValuesAs.PercentOfTotal,
        },
        { fieldName: 'Product', aggregation: PivotAggregation.Count },
      ],
    });

    expect(result.ok).toBe(true);
    expect(calls).toContain('data-field::2:2:#,##0.00:3');
    expect(calls).toContain('data-field::1:1:#,##0:');
    expect(calls).not.toContain('data-field::1:2:#,##0.00:3');
  });

  it('maps the show-as base field to its pivot field without listing its items', () => {
    const { wb, calls } = makeWb();
    const result = createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
      destination: { sheet: 0, row: 5, col: 0 },
      rowField: 'Region',
      valueField: 'Sales',
      valueFieldSettings: [
        {
          fieldName: 'Sales',
          showValuesAs: PivotShowValuesAs.PercentOf,
          showAsBaseField: 'Region',
          showAsBaseItem: 1,
        },
      ],
    });

    expect(result.ok).toBe(true);
    // The field lists its items when it joins the row axis.
    expect(calls.some((c) => c.startsWith('pivot-item:'))).toBe(false);
    expect(calls).toContain('data-field-base::0:1');
  });

  it('forwards a sentinel base item without listing items and rejects a non-axis base field', () => {
    const { wb, calls } = makeWb();
    const source = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 };
    const destination = { sheet: 0, row: 5, col: 0 };
    const ok = createPivotTableFromRange(wb, {
      source,
      destination,
      rowField: 'Region',
      valueField: 'Sales',
      valueFieldSettings: [
        {
          fieldName: 'Sales',
          showValuesAs: PivotShowValuesAs.DifferenceFrom,
          showAsBaseField: 'Region',
          showAsBaseItem: PIVOT_SHOW_AS_BASE_PREVIOUS,
        },
      ],
    });
    expect(ok.ok).toBe(true);
    expect(calls.some((c) => c.startsWith('pivot-item:'))).toBe(false);
    expect(calls).toContain(`data-field-base::0:${PIVOT_SHOW_AS_BASE_PREVIOUS}`);

    expect(
      createPivotTableFromRange(wb, {
        source,
        destination,
        rowField: 'Region',
        valueField: 'Sales',
        valueFieldSettings: [
          {
            fieldName: 'Sales',
            showValuesAs: PivotShowValuesAs.RankAscending,
            showAsBaseField: 'Product',
          },
        ],
      }),
    ).toEqual({ ok: false, reason: 'invalid-field' });
  });

  it('refreshes an existing pivot cache from a matching source range', () => {
    const { wb, calls } = makeWb();
    const result = refreshPivotCacheFromRange(wb, {
      cacheId: 4,
      source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
    });

    expect(result).toEqual({ ok: true, cacheId: 4 });
    expect(calls).toEqual(['cache-source:true:Sheet1:A1:C3', 'refresh']);
  });

  it('reports a refresh the engine refuses as too large as an invalid range', () => {
    const { wb } = makeWb();
    const result = refreshPivotCacheFromRange(wb, {
      cacheId: 4,
      source: { sheet: 0, r0: 0, c0: 0, r1: 1048575, c1: 2 },
    });

    expect(result).toEqual({ ok: false, reason: 'invalid-range', step: 'refresh' });
  });

  it('refreshes an existing pivot table by resolving its cache id', () => {
    const { wb, calls } = makeWb();
    const result = refreshPivotTableFromRange(wb, {
      sheet: 0,
      pivotIndex: 2,
      source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
    });

    expect(result).toEqual({ ok: true, cacheId: 4 });
    expect(calls).toContain('refresh');
  });

  it('refreshes an existing pivot table from stored worksheet source metadata', () => {
    const { wb, calls } = makeWb();
    const result = refreshPivotTable(wb, {
      sheet: 0,
      pivotIndex: 2,
    });

    expect(result).toEqual({ ok: true, cacheId: 4 });
    expect(calls).toContain('refresh');
  });

  it('rejects stored-source pivot refresh when metadata is unavailable', () => {
    const { wb, calls } = makeWb();
    wb.getPivotCacheWorksheetSource = () => null;
    const result = refreshPivotTable(wb, {
      sheet: 0,
      pivotIndex: 2,
    });

    expect(result).toEqual({ ok: false, reason: 'invalid-range' });
    expect(calls).not.toContain('refresh');
  });

  it('rejects pivot table refresh when the pivot cache id is unavailable', () => {
    const { wb, calls } = makeWb();
    const result = refreshPivotTableFromRange(wb, {
      sheet: 0,
      pivotIndex: 99,
      source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
    });

    expect(result).toEqual({ ok: false, reason: 'invalid-pivot' });
    expect(calls).not.toContain('refresh');
  });

  it('rejects duplicate row and column fields and leaves no cache behind', () => {
    const { wb, calls } = makeWb();
    const result = createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
      destination: { sheet: 0, row: 5, col: 0 },
      rowField: 'Region',
      columnField: 'Region',
      valueField: 'Sales',
    });
    expect(result).toEqual({ ok: false, reason: 'invalid-field' });
    expect(calls.at(-1)).toBe('remove-cache');
    expect(calls.some((c) => c.startsWith('pivot:'))).toBe(false);
  });

  it('reports unsupported engines before mutating', () => {
    const { wb } = makeWb();
    Object.defineProperty(wb, 'capabilities', { value: { pivotTableMutate: false } });
    const result = createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
      destination: { sheet: 0, row: 5, col: 0 },
      rowField: 'Region',
      valueField: 'Sales',
    });
    expect(result).toEqual({ ok: false, reason: 'unsupported' });
  });

  it('keeps Recommended PivotTables as a report instead of silently creating a pivot', () => {
    const { wb, calls } = makeWb();
    const store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 });

    const result = executeRibbonPivotTableAction({
      origin: 'ribbon' as const,
      commandId: 'test',
      store,
      workbook: wb,
      action: 'recommended',
      strings: {
        pivotTable: 'PivotTable',
        pivotTableNewSheet: 'New Worksheet',
        pivotTableRefreshData: 'Refresh Data',
        pivotTableRefreshUnavailable: 'Select a PivotTable before refreshing.',
        recommendedPivotTables: 'Recommended PivotTables',
        pivotAuthoringDetail: 'Recommended PivotTables are not implemented.',
        workbookStructureProtectedBlocked: 'Workbook structure is protected.',
      },
    });

    expect(result).toEqual({
      kind: 'report',
      report: {
        title: 'Recommended PivotTables',
        items: [
          {
            severity: 'info',
            label: 'PivotTable',
            detail: 'Recommended PivotTables are not implemented.',
          },
        ],
      },
    });
    expect(calls).toEqual([]);
  });

  it('refreshes the active PivotTable from stored source metadata', () => {
    const { wb, calls } = makeWb();
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 5, col: 1 });

    const result = executeRibbonPivotTableAction({
      origin: 'ribbon' as const,
      commandId: 'test',
      store,
      workbook: wb,
      action: 'refresh',
      strings: {
        pivotTable: 'PivotTable',
        pivotTableNewSheet: 'New Worksheet',
        pivotTableRefreshData: 'Refresh Data',
        pivotTableRefreshUnavailable: 'Select a PivotTable before refreshing.',
        recommendedPivotTables: 'Recommended PivotTables',
        pivotAuthoringDetail: 'Recommended PivotTables are not implemented.',
        workbookStructureProtectedBlocked: 'Workbook structure is protected.',
      },
    });

    expect(result).toEqual({ kind: 'refreshed', sheet: 0 });
    expect(calls).toContain('refresh');
    expect(calls).toContain('cache-source:true:Sheet1:A1:C3');
  });

  it('reports Refresh Data when the active cell is outside a PivotTable', () => {
    const { wb, calls } = makeWb();
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });

    const result = executeRibbonPivotTableAction({
      origin: 'ribbon' as const,
      commandId: 'test',
      store,
      workbook: wb,
      action: 'refresh',
      strings: {
        pivotTable: 'PivotTable',
        pivotTableNewSheet: 'New Worksheet',
        pivotTableRefreshData: 'Refresh Data',
        pivotTableRefreshUnavailable: 'Select a PivotTable before refreshing.',
        recommendedPivotTables: 'Recommended PivotTables',
        pivotAuthoringDetail: 'Recommended PivotTables are not implemented.',
        workbookStructureProtectedBlocked: 'Workbook structure is protected.',
      },
    });

    expect(result).toEqual({
      kind: 'report',
      report: {
        title: 'Refresh Data',
        items: [
          {
            severity: 'warning',
            label: 'PivotTable',
            detail: 'Select a PivotTable before refreshing.',
          },
        ],
      },
    });
    expect(calls).toEqual([]);
  });
});
