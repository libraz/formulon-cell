import type { Formulon } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import type { Workbook } from '../../../src/engine/types.js';
import {
  PivotAggregation,
  PivotAxis,
  PivotCalendar,
  PivotDateGrouping,
  PivotFilterKind,
  PivotReportLayout,
  PivotSortOrder,
  ValueKind,
} from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const engineError = (message: string): Error =>
  Object.assign(new Error(message), { name: 'FormulonError', code: 1 });

const field = (sourceName: string, axis: number | null = null) => ({
  sourceName,
  customName: null,
  axis,
  subtotalTop: true,
  numberFormat: null,
  cacheFieldIndex: null,
  sortOrder: PivotSortOrder.Manual,
  sortByField: null,
  datePeriods: 0,
  dateCalendar: PivotCalendar.Gregorian,
  groupBaseIndex: null,
  defaultSubtotal: true,
  insertBlankRow: false,
  fillDownLabels: false,
  showAll: false,
  showDetail: true,
  subtotalFunctions: [],
});

interface FakeOverrides {
  workbook?: Record<string, unknown>;
  sheet?: Record<string, unknown>;
  pivot?: Record<string, unknown>;
  cache?: Record<string, unknown>;
}

const makeHandle = (overrides: FakeOverrides = {}): WorkbookHandle => {
  const cache = (cacheId: number) => ({
    cacheId,
    getWorksheetSource: () => ({
      range: { sheet: 0, rect: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 2 } },
      name: null,
    }),
    setWorksheetSource: () => undefined,
    listFields: () => [
      { name: 'Region', formula: null },
      { name: 'Sales', formula: null },
    ],
    addField: (_name: string) => 2,
    clearFields: () => undefined,
    listFieldSharedItems: (fieldIdx: number) =>
      fieldIdx === 0
        ? [0, 1].map((itemIdx) => ({
            value: cacheId === 7 ? (itemIdx === 0 ? 'East' : 'West') : 'Wrong cache',
            isDate: false,
          }))
        : [],
    addFieldSharedItem: () => 0,
    clearFieldSharedItems: () => undefined,
    getRecordCount: () => 0,
    addRecord: () => 0,
    clearRecords: () => undefined,
    setRecordValue: () => undefined,
    ...overrides.cache,
  });
  const pivot = {
    name: 'PivotTable1',
    index: 0,
    getLayout: () => ({
      range: { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 },
      cells: [
        {
          row: 0,
          col: 0,
          value: 'pivot header',
          kind: 0,
          depth: 0,
          fieldName: 'Region',
          numberFormat: null,
        },
        { row: 1, col: 0, value: null, kind: 7, depth: 0, fieldName: null, numberFormat: null },
        { row: 1, col: 1, value: 42, kind: 3, depth: 0, fieldName: 'Sales', numberFormat: '#,##0' },
      ],
    }),
    cache: () => cache(7),
    getSpill: () => null,
    rename: () => undefined,
    setAnchor: () => undefined,
    getGrandTotals: () => ({ rows: true, columns: true }),
    updateGrandTotals: () => undefined,
    getReportLayout: () => PivotReportLayout.Tabular,
    setReportLayout: () => undefined,
    listFields: () => [field('Region')],
    getField: () => field('Region'),
    updateField: () => undefined,
    clearFields: () => undefined,
    setFieldAxis: () => undefined,
    listFieldItems: () => [],
    updateFieldItem: () => undefined,
    addFieldItem: () => 0,
    clearFieldItems: () => undefined,
    setFieldSubtotalFunctions: () => undefined,
    setFieldDateGroup: () => undefined,
    getRowFieldOrder: () => [],
    setRowFieldOrder: () => undefined,
    getColumnFieldOrder: () => [],
    setColumnFieldOrder: () => undefined,
    listDataFields: () => [],
    addDataField: () => 0,
    setDataField: () => undefined,
    clearDataFields: () => undefined,
    listFilters: () => [],
    addFilter: () => 0,
    removeFilter: () => undefined,
    clearFilters: () => undefined,
    ...overrides.pivot,
  };
  const sheet = {
    listCells: () => [{ row: 0, col: 0, kind: ValueKind.Text, value: 'cached', formula: null }],
    setBlank: () => undefined,
    listPivots: () => [pivot],
    pivot: () => pivot,
    addPivot: (_input: {
      name: string;
      cacheId: number;
      anchor: { row: number; col: number };
    }) => ({
      index: 3,
    }),
    removePivot: () => undefined,
    ...overrides.sheet,
  };
  const wb = {
    sheetCount: 1,
    sheet: () => sheet,
    getDate1904: () => false,
    getPinnedNow: () => null,
    listPivotCaches: () => [cache(7), cache(9)],
    pivotCache: (cacheId: number) => cache(cacheId),
    addPivotCache: (requestedId: number | null) => cache(requestedId ?? 8),
    removePivotCache: () => undefined,
    ...overrides.workbook,
  } as unknown as Workbook;

  // A serial passes through so a test can read the bound it became.
  const module = {
    getVersion: () => 'test',
    convertSerialToCivil: (serial: number) => ({ year: 2000, month: 1, day: 1, serial }),
  } as unknown as Formulon;
  const Ctor = WorkbookHandle as unknown as new (module: Formulon, wb: Workbook) => WorkbookHandle;
  return new Ctor(module, wb);
};

describe('WorkbookHandle PivotTable projection', () => {
  it('projects pivot layout cells after physical cells and skips blank projected cells', () => {
    const wb = makeHandle();

    expect(wb.capabilities.pivotTables).toBe(true);
    expect([...wb.physicalCells(0)]).toEqual([
      {
        addr: { sheet: 0, row: 0, col: 0 },
        value: { kind: 'text', value: 'cached' },
        formula: null,
      },
    ]);
    expect([...wb.cells(0)]).toEqual([
      {
        addr: { sheet: 0, row: 0, col: 0 },
        value: { kind: 'text', value: 'cached' },
        formula: null,
      },
      {
        addr: { sheet: 0, row: 0, col: 0 },
        value: { kind: 'text', value: 'pivot header' },
        formula: null,
        kind: 0,
        depth: 0,
        fieldName: 'Region',
        numberFormat: '',
        pivotIndex: 0,
      },
      {
        addr: { sheet: 0, row: 1, col: 1 },
        value: { kind: 'number', value: 42 },
        formula: null,
        kind: 3,
        depth: 0,
        fieldName: 'Sales',
        numberFormat: '#,##0',
        pivotIndex: 0,
      },
    ]);
  });

  it('summarizes projected PivotTable layouts for object inspectors', () => {
    const wb = makeHandle();

    expect(wb.getPivotTables()).toEqual([
      {
        sheetIndex: 0,
        pivotIndex: 0,
        top: 0,
        left: 0,
        rows: 2,
        cols: 2,
        cells: 3,
        fields: ['Region', 'Sales'],
        fieldItems: {
          Region: ['East', 'West'],
          Sales: ['42'],
        },
        // Only the cache-backed field carries indices; `Sales` was inferred
        // from the projected layout, which addresses no shared item.
        fieldItemIndexes: {
          Region: [0, 1],
        },
      },
    ]);
  });

  it('adds PivotCache shared error items without aborting cache construction', () => {
    const calls: string[] = [];
    const wb = makeHandle({
      cache: {
        addFieldSharedItem: (_field: number, value: { kind: number; errorCode: number }) => {
          calls.push(`error:${value.errorCode}`);
          return 0;
        },
      },
    });

    expect(wb.addPivotCacheSharedItem(7, 0, { kind: 'error', code: 7, text: '#DIV/0!' })).toBe(
      true,
    );
    expect(calls).toEqual(['error:7']);
  });

  it('includes readable PivotTable filter specs in object summaries when engine hooks exist', () => {
    const wb = makeHandle({
      pivot: {
        listFields: () => [field(' Region ', PivotAxis.Page), field('Date', PivotAxis.Page)],
        listFilters: () => [
          {
            fieldIndex: 0,
            kind: PivotFilterKind.LabelContains,
            dataFieldIndex: 0,
            value: 'East',
            valueHigh: null,
          },
          {
            fieldIndex: 1,
            kind: PivotFilterKind.DateBetween,
            dataFieldIndex: 0,
            value: 46143,
            valueHigh: 46173,
          },
          {
            fieldIndex: 1,
            kind: PivotFilterKind.ThisMonth,
            dataFieldIndex: 0,
            value: null,
            valueHigh: null,
          },
        ],
      },
    });

    expect(wb.getPivotTables()[0]?.pivotFilters).toEqual([
      {
        axis: PivotAxis.Page,
        fieldName: 'Region',
        kind: PivotFilterKind.LabelContains,
        value: 'East',
      },
      {
        axis: PivotAxis.Page,
        fieldName: 'Date',
        kind: PivotFilterKind.DateBetween,
        value: 46143,
        valueHigh: 46173,
      },
      { axis: PivotAxis.Page, fieldName: 'Date', kind: PivotFilterKind.ThisMonth },
    ]);
  });

  it('wraps low-level PivotCache and PivotTable mutation APIs', () => {
    const wb = makeHandle();

    expect(wb.capabilities.pivotTableMutate).toBe(true);
    expect(wb.pivotCacheIds()).toEqual([7, 9]);
    expect(wb.createPivotCache()).toBe(8);
    expect(wb.getPivotCacheWorksheetSource(8)).toEqual({
      range: { sheet: 0, rect: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 2 } },
      name: null,
    });
    expect(wb.setPivotCacheWorksheetSource(8, { range: 'B2:D10', sheet: 'Sheet1' })).toBe(true);
    expect(wb.addPivotCacheField(8, 'Channel')).toBe(2);
    expect(wb.pivotCacheFieldNames(8)).toEqual(['Region', 'Sales']);
    expect(wb.pivotTableCacheId(0, 3)).toBe(7);
    expect(wb.pivotCacheSharedItems(7, 0)).toEqual([
      { kind: 'text', value: 'East' },
      { kind: 'text', value: 'West' },
    ]);
    expect(wb.addPivotCacheRecord(8)).toBe(0);
    expect(
      wb.setPivotCacheRecordValue(8, 0, 1, {
        kind: 'number',
        value: 42,
      }),
    ).toBe(true);
    expect(wb.createPivotTable(0, 'Pivot1', 8, { row: 4, col: 1 })).toBe(3);
    expect(wb.getPivotReportLayout(0, 3)).toBe(PivotReportLayout.Tabular);
    expect(wb.setPivotReportLayout(0, 3, PivotReportLayout.Outline)).toBe(true);
    expect(wb.setPivotFieldSort(0, 3, 0, PivotSortOrder.Ascending, 'Sales')).toBe(true);
    expect(wb.setPivotFieldSubtotalTop(0, 3, 0, false)).toBe(true);
    expect(wb.addPivotFieldItem(0, 3, 0, 'East', true)).toBe(true);
    expect(wb.clearPivotFieldItems(0, 3, 0)).toBe(true);
    expect(wb.setPivotFieldItemVisible(0, 3, 0, 1, false)).toBe(true);
    expect(wb.addPivotFieldSubtotalFn(0, 3, 0, PivotAggregation.Sum)).toBe(true);
    expect(wb.clearPivotFieldSubtotalFns(0, 3, 0)).toBe(true);
    expect(
      wb.setPivotFieldDateGroup(0, 3, 0, PivotDateGrouping.Month, PivotCalendar.Gregorian, {
        startYear: 2024,
        endYear: 2026,
      }),
    ).toBe(true);
    expect(wb.clearPivotFieldDateGroup(0, 3, 0)).toBe(true);
    expect(wb.setPivotFieldNumberFormat(0, 3, 0, '#,##0')).toBe(true);
  });

  it('keeps failed pivot counts neutral instead of treating the value as valid', () => {
    const wb = makeHandle({
      sheet: {
        listPivots: () => {
          throw engineError('invalid pivot');
        },
      },
      workbook: {
        listPivotCaches: () => {
          throw engineError('invalid pivot');
        },
      },
    });

    expect([...wb.pivotCells(0)]).toEqual([]);
    expect(wb.pivotCacheCount()).toBe(0);
    expect(wb.pivotCacheIds()).toEqual([]);
  });

  it('passes date-group and data-field payloads through the adapter', () => {
    let dateArgs: unknown[] = [];
    let dataSpec: Record<string, unknown> | undefined;
    const wb = makeHandle({
      pivot: {
        setFieldDateGroup: (
          _field: number,
          group: { intervalDays: number; start: { serial: number }; end: { serial: number } },
        ) => {
          dateArgs = [group.intervalDays, group.start.serial, group.end.serial];
        },
        addDataField: (spec: Record<string, unknown>) => {
          dataSpec = spec;
          return 0;
        },
      },
    });

    expect(
      wb.setPivotFieldDateGroup(0, 3, 0, PivotDateGrouping.Days, PivotCalendar.Gregorian, {
        intervalDays: 7,
        startSerial: 45000,
        endSerial: 46000,
      }),
    ).toBe(true);
    expect(dateArgs.slice(-3)).toEqual([7, 45000, 46000]);

    expect(
      wb.addPivotDataField(0, 3, {
        fieldIndex: 1,
        aggregation: PivotAggregation.Average,
        numberFormat: '#,##0.00',
        showValuesAs: 3,
      }),
    ).toBe(0);
    expect(dataSpec).toMatchObject({
      name: '',
      numberFormat: '#,##0.00',
      showAs: 3,
    });
    expect(dataSpec).not.toHaveProperty('showValuesAs');
  });
});
