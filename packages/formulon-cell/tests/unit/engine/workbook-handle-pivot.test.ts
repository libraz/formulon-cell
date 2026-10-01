import { describe, expect, it } from 'vitest';
import type { FormulonModule, Value, Workbook } from '../../../src/engine/types.js';
import {
  PivotAggregation,
  PivotAxis,
  PivotCalendar,
  PivotDateGrouping,
  PivotFilterType,
  PivotFilterValueKind,
  PivotReportLayout,
  ValueKind,
} from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const ok = { ok: true, code: 0, message: '' };
const numberResult = (value: number) => ({ status: ok, value });

const textValue = (text: string): Value => ({
  kind: ValueKind.Text,
  number: 0,
  boolean: 0,
  text,
  errorCode: 0,
});

const numberValue = (number: number): Value => ({
  kind: ValueKind.Number,
  number,
  boolean: 0,
  text: '',
  errorCode: 0,
});

const blankValue = (): Value => ({
  kind: ValueKind.Blank,
  number: 0,
  boolean: 0,
  text: '',
  errorCode: 0,
});

const makeHandle = (overrides: Record<string, unknown> = {}): WorkbookHandle => {
  const wb = {
    sheetCount: () => numberResult(1),
    cellCount: (_sheet: number) => numberResult(1),
    cellAt: (_sheet: number, _idx: number) => ({
      status: ok,
      row: 0,
      col: 0,
      value: textValue('cached'),
      formula: null,
    }),
    pivotCount: (_sheet: number) => numberResult(1),
    pivotLayout: (_sheet: number, _pivotIndex: number) => ({
      status: ok,
      top: 0,
      left: 0,
      rows: 2,
      cols: 2,
      cells: [
        {
          row: 0,
          col: 0,
          value: textValue('pivot header'),
          kind: 0,
          depth: 0,
          fieldName: 'Region',
          numberFormat: '',
        },
        {
          row: 1,
          col: 0,
          value: blankValue(),
          kind: 7,
          depth: 0,
          fieldName: '',
          numberFormat: '',
        },
        {
          row: 1,
          col: 1,
          value: numberValue(42),
          kind: 3,
          depth: 0,
          fieldName: 'Sales',
          numberFormat: '4',
        },
      ],
    }),
    pivotCacheCount: () => numberResult(2),
    pivotCacheIdAt: (idx: number) => ({ status: ok, index: idx === 0 ? 7 : 9 }),
    pivotCacheCreate: (requestedId: number) => {
      return { status: ok, index: requestedId || 8 };
    },
    pivotCacheRemove: (_cacheId: number) => ok,
    pivotCacheGetWorksheetSource: (_cacheId: number) => ({
      status: ok,
      present: true,
      ref: 'A1:C3',
      sheet: 'Data',
      name: '',
    }),
    pivotCacheSetWorksheetSource: () => ok,
    getNumFmt: (numFmtId: number) => ({
      status: ok,
      numFmtId,
      formatCode: numFmtId === 4 ? '#,##0' : 'General',
    }),
    addNumFmt: (_formatCode: string) => ({ status: ok, numFmtId: 164 }),
    getCellXfIndex: () => ({ status: ok, index: 0 }),
    setCellXfIndex: () => ok,
    getCellXf: () => ({
      status: ok,
      fontIndex: 0,
      fillIndex: 0,
      borderIndex: 0,
      numFmtId: 0,
      horizontalAlign: 0,
      verticalAlign: 2,
      wrapText: false,
      justifyLastLine: false,
    }),
    addFont: () => ({ status: ok, index: 0 }),
    addFill: () => ({ status: ok, index: 0 }),
    addBorder: () => ({ status: ok, index: 0 }),
    addXf: () => ({ status: ok, index: 0 }),
    getFont: () => ({ status: ok, name: 'Calibri', size: 11 }),
    getFill: () => ({ status: ok, pattern: 0 }),
    getBorder: () => ({ status: ok }),
    pivotCacheFieldCount: (_cacheId: number) => numberResult(2),
    pivotCacheFieldName: (_cacheId: number, fieldIdx: number) => ({
      status: ok,
      value: fieldIdx === 0 ? 'Region' : 'Sales',
    }),
    pivotCacheFieldAdd: (_cacheId: number, name: string) => {
      void name;
      return { status: ok, index: 2 };
    },
    pivotCacheFieldClear: (_cacheId: number) => ok,
    pivotCacheFieldSharedItemCount: (_cacheId: number, fieldIdx: number) =>
      numberResult(fieldIdx === 0 ? 2 : 0),
    pivotCacheFieldSharedItemValue: (cacheId: number, _fieldIdx: number, itemIdx: number) => ({
      status: ok,
      value: textValue(cacheId === 7 ? (itemIdx === 0 ? 'East' : 'West') : 'Wrong cache'),
    }),
    pivotCacheFieldAddSharedItemNumber: () => ok,
    pivotCacheFieldAddSharedItemText: () => ok,
    pivotCacheFieldAddSharedItemBool: () => ok,
    pivotCacheFieldAddSharedItemBlank: () => ok,
    pivotCacheFieldClearSharedItems: () => ok,
    pivotCacheRecordCount: () => numberResult(0),
    pivotCacheRecordAdd: () => ({ status: ok, index: 0 }),
    pivotCacheRecordClear: () => ok,
    pivotCacheRecordSetNumber: () => ok,
    pivotCacheRecordSetText: () => ok,
    pivotCacheRecordSetBool: () => ok,
    pivotCacheRecordSetBlank: () => ok,
    pivotCacheRecordSetError: () => ok,
    pivotCreate: (_sheet: number, name: string, cacheId: number, row: number, col: number) => {
      void name;
      void cacheId;
      void row;
      void col;
      return { status: ok, index: 3 };
    },
    pivotCacheId: () => ({ status: ok, index: 7 }),
    pivotRemove: () => ok,
    pivotSetName: () => ok,
    pivotSetAnchor: () => ok,
    pivotSetGrandTotals: () => ok,
    pivotGetLayout: () => ({ status: ok, layout: PivotReportLayout.Tabular }),
    pivotSetLayout: () => ok,
    pivotFieldCount: () => numberResult(0),
    pivotFieldAdd: () => ({ status: ok, index: 0 }),
    pivotFieldClear: () => ok,
    pivotFieldSetAxis: () => ok,
    pivotFieldSetSort: () => ok,
    pivotFieldSetSubtotalTop: () => ok,
    pivotFieldAddItem: () => ok,
    pivotFieldClearItems: () => ok,
    pivotFieldSetItemVisible: () => ok,
    pivotFieldAddSubtotalFn: () => ok,
    pivotFieldClearSubtotalFns: () => ok,
    pivotFieldSetDateGroup: () => ok,
    pivotFieldClearDateGroup: () => ok,
    pivotFieldSetNumberFormat: () => ok,
    pivotSetRowFieldOrder: () => ok,
    pivotSetColFieldOrder: () => ok,
    pivotDataFieldCount: () => numberResult(0),
    pivotDataFieldAdd: () => ({ status: ok, index: 0 }),
    pivotDataFieldClear: () => ok,
    pivotDataFieldSet: () => ok,
    pivotFilterCount: () => numberResult(0),
    pivotFilterAt: () => ({
      status: { ok: false, code: 1, message: 'no filter' },
      axis: PivotAxis.Page,
      fieldName: '',
      type: PivotFilterType.ValueTop10,
      dataFieldIndex: 0,
      valueKind: PivotFilterValueKind.None,
      valueInt: 0,
      valueDouble: 0,
      valueText: '',
      valueHighKind: PivotFilterValueKind.None,
      valueHighInt: 0,
      valueHighDouble: 0,
    }),
    pivotFilterAdd: () => ok,
    pivotFilterClear: () => ok,
    pivotFilterRemoveAt: () => ok,
    ...overrides,
  } as unknown as Workbook;

  const module = { versionString: () => 'test' } as unknown as FormulonModule;
  const Ctor = WorkbookHandle as unknown as new (
    module: FormulonModule,
    wb: Workbook,
  ) => WorkbookHandle;
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
        numberFormat: '',
        pivotIndex: 0,
      },
      {
        addr: { sheet: 0, row: 1, col: 1 },
        value: { kind: 'number', value: 42 },
        formula: null,
        kind: 3,
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
      pivotCacheFieldAddSharedItemError: (_cacheId: number, _fieldIdx: number, code: number) => {
        calls.push(`error:${code}`);
        return ok;
      },
    });

    expect(wb.addPivotCacheSharedItem(7, 0, { kind: 'error', code: 7, text: '#DIV/0!' })).toBe(
      true,
    );
    expect(calls).toEqual(['error:7']);
  });

  it('includes readable PivotTable filter specs in object summaries when engine hooks exist', () => {
    const wb = makeHandle({
      pivotFilterCount: () => numberResult(3),
      pivotFilterAt: (_sheet: number, _pivot: number, filterIdx: number) => {
        if (filterIdx === 0) {
          return {
            status: ok,
            axis: PivotAxis.Page,
            fieldName: ' Region ',
            type: PivotFilterType.LabelContains,
            dataFieldIndex: 0,
            valueKind: PivotFilterValueKind.Text,
            valueInt: 0,
            valueDouble: 0,
            valueText: 'East',
            valueHighKind: PivotFilterValueKind.None,
            valueHighInt: 0,
            valueHighDouble: 0,
          };
        }
        if (filterIdx === 1) {
          return {
            status: ok,
            axis: PivotAxis.Page,
            fieldName: 'Date',
            type: PivotFilterType.LabelDate,
            dataFieldIndex: 0,
            valueKind: PivotFilterValueKind.Text,
            valueInt: 0,
            valueDouble: 0,
            valueText: '2026-05-01',
            valueHighKind: PivotFilterValueKind.None,
            valueHighInt: 0,
            valueHighDouble: 0,
          };
        }
        return {
          status: { ok: false, code: 1, message: 'fallback' },
          axis: PivotAxis.Page,
          fieldName: 'Ignored',
          type: PivotFilterType.LabelBeginsWith,
          dataFieldIndex: 0,
          valueKind: PivotFilterValueKind.None,
          valueInt: 0,
          valueDouble: 0,
          valueText: '',
          valueHighKind: PivotFilterValueKind.None,
          valueHighInt: 0,
          valueHighDouble: 0,
        };
      },
    });

    expect(wb.getPivotTables()[0]?.pivotFilters).toEqual([
      {
        axis: PivotAxis.Page,
        fieldName: 'Region',
        type: PivotFilterType.LabelContains,
        valueKind: PivotFilterValueKind.Text,
        valueText: 'East',
      },
      {
        axis: PivotAxis.Page,
        fieldName: 'Date',
        type: PivotFilterType.LabelDate,
        valueKind: PivotFilterValueKind.Text,
        valueText: '2026-05-01',
      },
    ]);
  });

  it('wraps low-level PivotCache and PivotTable mutation APIs', () => {
    const wb = makeHandle();

    expect(wb.capabilities.pivotTableMutate).toBe(true);
    expect(wb.pivotCacheIds()).toEqual([7, 9]);
    expect(wb.createPivotCache()).toBe(8);
    expect(wb.getPivotCacheWorksheetSource(8)).toEqual({
      present: true,
      ref: 'A1:C3',
      sheet: 'Data',
    });
    expect(
      wb.setPivotCacheWorksheetSource(8, {
        present: true,
        ref: 'B2:D10',
        sheet: 'Sheet1',
      }),
    ).toBe(true);
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
    expect(wb.setPivotFieldSort(0, 3, 0, true, 'Sales')).toBe(true);
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
    const failed = { ok: false, code: 1, message: 'invalid pivot' };
    const wb = makeHandle({
      pivotCount: () => ({ status: failed, value: 99 }),
      pivotCacheCount: () => ({ status: failed, value: 99 }),
    });

    expect([...wb.pivotCells(0)]).toEqual([]);
    expect(wb.pivotCacheCount()).toBe(0);
    expect(wb.pivotCacheIds()).toEqual([]);
  });

  it('passes 0.12 date-group and data-field payloads through the adapter', () => {
    let dateArgs: unknown[] = [];
    let dataSpec: Record<string, unknown> | undefined;
    const wb = makeHandle({
      pivotFieldSetDateGroup: (...args: unknown[]) => {
        dateArgs = args;
        return ok;
      },
      pivotDataFieldAdd: (_sheet: number, _pivot: number, spec: Record<string, unknown>) => {
        dataSpec = spec;
        return { status: ok, index: 0 };
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
      name: 'Average of field 1',
      numberFormat: '164',
      showAs: 3,
    });
    expect(dataSpec).not.toHaveProperty('showValuesAs');
  });
});
