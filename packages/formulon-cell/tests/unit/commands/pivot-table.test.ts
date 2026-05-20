import { describe, expect, it } from 'vitest';
import {
  createPivotTableFromRange,
  executeRibbonPivotTableAction,
  inferPivotFieldItems,
  inferPivotSourceFields,
} from '../../../src/commands/pivot-table.js';
import {
  type CellValue,
  PivotAggregation,
  PivotAxis,
  PivotFilterType,
  PivotFilterValueKind,
} from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
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
  const wb = {
    capabilities: { pivotTableMutate: true },
    getValue: ({ row, col }: { row: number; col: number }) => values.get(`${row}:${col}`) ?? blank,
    createPivotCache: () => {
      calls.push('cache');
      return 4;
    },
    removePivotCache: () => true,
    addPivotCacheField: (_cache: number, name: string) => {
      calls.push(`cache-field:${name}`);
      return calls.filter((c) => c.startsWith('cache-field:')).length - 1;
    },
    addPivotCacheSharedItem: (_cache: number, field: number, value: CellValue) => {
      calls.push(`shared:${field}:${value.kind}`);
      return true;
    },
    addPivotCacheRecord: () => {
      calls.push('record');
      return calls.filter((c) => c === 'record').length - 1;
    },
    setPivotCacheRecordValue: (_cache: number, record: number, field: number, value: CellValue) => {
      calls.push(`value:${record}:${field}:${value.kind}`);
      return true;
    },
    createPivotTable: (_sheet: number, name: string, cacheId: number, anchor: unknown) => {
      calls.push(`pivot:${name}:${cacheId}:${JSON.stringify(anchor)}`);
      return 2;
    },
    removePivotTable: () => true,
    setPivotTableGrandTotals: (_sheet: number, _pivot: number, rows: boolean, cols: boolean) => {
      calls.push(`grand:${rows}:${cols}`);
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
    addPivotFilter: (_sheet: number, _pivot: number, spec: { fieldName: string; type: number }) => {
      calls.push(`pivot-filter:${spec.fieldName}:${spec.type}`);
      return true;
    },
    setPivotFieldSort: (
      _sheet: number,
      _pivot: number,
      fieldIdx: number,
      ascending: boolean,
      byField: string,
    ) => {
      calls.push(`sort:${fieldIdx}:${ascending}:${byField}`);
      return true;
    },
    addPivotDataField: (
      _sheet: number,
      _pivot: number,
      spec: { name: string; fieldIndex: number; aggregation: number; numberFormat?: string },
    ) => {
      calls.push(
        `data-field:${spec.name}:${spec.fieldIndex}:${spec.aggregation}:${spec.numberFormat ?? ''}`,
      );
      return 0;
    },
  } as unknown as WorkbookHandle;
  return { wb, calls };
};

describe('pivot-table command helpers', () => {
  it('infers fields from the header row and counts numeric values', () => {
    const { wb } = makeWb();
    expect(inferPivotSourceFields(wb, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 })).toEqual([
      { name: 'Region', index: 0, numericCount: 0 },
      { name: 'Product', index: 1, numericCount: 0 },
      { name: 'Sales', index: 2, numericCount: 2 },
    ]);
  });

  it('infers unique PivotTable field items from source values', () => {
    const { wb } = makeWb();
    expect(inferPivotFieldItems(wb, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 }, 'Region')).toEqual([
      'East',
      'West',
    ]);
  });

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

    expect(result).toEqual({ ok: true, cacheId: 4, pivotIndex: 2 });
    expect(calls).toContain('cache-field:Region');
    expect(calls).toContain('shared:2:number');
    expect(calls).toContain('record');
    expect(calls).toContain('pivot-field:Region:0');
    expect(calls).toContain('grand:false:true');
    expect(calls).toContain('subtotal-top:Region:false');
    expect(calls).toContain('sort:0:true:');
    expect(calls).toContain('sort:1:false:');
    expect(calls).toContain('pivot-field:Product:1');
    expect(calls).toContain('pivot-field:Sales:2');
    expect(calls).toContain('data-field:Sum of Sales:2:0:#,##0');
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
    expect(calls).toContain('pivot-item:1:Desk:true');
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
    expect(calls).toContain('data-field:Count of Sales:2:1:');
    expect(calls).toContain('data-field:Count of Product:1:1:');
  });

  it('rejects duplicate row and column fields before creating a cache', () => {
    const { wb, calls } = makeWb();
    const result = createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
      destination: { sheet: 0, row: 5, col: 0 },
      rowField: 'Region',
      columnField: 'Region',
      valueField: 'Sales',
    });
    expect(result).toEqual({ ok: false, reason: 'invalid-field' });
    expect(calls).toEqual([]);
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
      store,
      workbook: wb,
      action: 'recommended',
      strings: {
        pivotTable: 'PivotTable',
        pivotTableNewSheet: 'New Worksheet',
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
});
