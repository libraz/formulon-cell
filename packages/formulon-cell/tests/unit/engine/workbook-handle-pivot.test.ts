import { describe, expect, it } from 'vitest';
import type { FormulonModule, Value, Workbook } from '../../../src/engine/types.js';
import {
  PivotAggregation,
  PivotCalendar,
  PivotDateGrouping,
  ValueKind,
} from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const ok = { ok: true, code: 0, message: '' };

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

const makeHandle = (): WorkbookHandle => {
  const wb = {
    sheetCount: () => 1,
    cellCount: (_sheet: number) => 1,
    cellAt: (_sheet: number, _idx: number) => ({
      status: ok,
      row: 0,
      col: 0,
      value: textValue('cached'),
      formula: null,
    }),
    pivotCount: (_sheet: number) => 1,
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
          numberFormat: '#,##0',
        },
      ],
    }),
    pivotCacheCount: () => 1,
    pivotCacheIdAt: (_idx: number) => ({ status: ok, index: 7 }),
    pivotCacheCreate: (requestedId: number) => {
      return { status: ok, index: requestedId || 8 };
    },
    pivotCacheRemove: (_cacheId: number) => ok,
    pivotCacheFieldCount: (_cacheId: number) => 2,
    pivotCacheFieldName: (_cacheId: number, fieldIdx: number) => ({
      status: ok,
      value: fieldIdx === 0 ? 'Region' : 'Sales',
    }),
    pivotCacheFieldAdd: (_cacheId: number, name: string) => {
      void name;
      return { status: ok, index: 2 };
    },
    pivotCacheFieldClear: (_cacheId: number) => ok,
    pivotCacheFieldSharedItemCount: () => 0,
    pivotCacheFieldAddSharedItemNumber: () => ok,
    pivotCacheFieldAddSharedItemText: () => ok,
    pivotCacheFieldAddSharedItemBool: () => ok,
    pivotCacheFieldAddSharedItemBlank: () => ok,
    pivotCacheFieldClearSharedItems: () => ok,
    pivotCacheRecordCount: () => 0,
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
    pivotRemove: () => ok,
    pivotSetName: () => ok,
    pivotSetAnchor: () => ok,
    pivotSetGrandTotals: () => ok,
    pivotFieldCount: () => 0,
    pivotFieldAdd: () => ({ status: ok, index: 0 }),
    pivotFieldClear: () => ok,
    pivotFieldSetAxis: () => ok,
    pivotFieldSetSort: () => ok,
    pivotFieldSetSubtotalTop: () => ok,
    pivotFieldAddAggregation: () => ok,
    pivotFieldClearAggregations: () => ok,
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
    pivotDataFieldCount: () => 0,
    pivotDataFieldAdd: () => ({ status: ok, index: 0 }),
    pivotDataFieldClear: () => ok,
    pivotDataFieldSet: () => ok,
    pivotFilterCount: () => 0,
    pivotFilterAdd: () => ok,
    pivotFilterClear: () => ok,
    pivotFilterRemoveAt: () => ok,
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
      },
      {
        addr: { sheet: 0, row: 1, col: 1 },
        value: { kind: 'number', value: 42 },
        formula: null,
        kind: 3,
        numberFormat: '#,##0',
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
      },
    ]);
  });

  it('wraps low-level PivotCache and PivotTable mutation APIs', () => {
    const wb = makeHandle();

    expect(wb.capabilities.pivotTableMutate).toBe(true);
    expect(wb.pivotCacheIds()).toEqual([7]);
    expect(wb.createPivotCache()).toBe(8);
    expect(wb.addPivotCacheField(8, 'Channel')).toBe(2);
    expect(wb.pivotCacheFieldNames(8)).toEqual(['Region', 'Sales']);
    expect(wb.addPivotCacheRecord(8)).toBe(0);
    expect(
      wb.setPivotCacheRecordValue(8, 0, 1, {
        kind: 'number',
        value: 42,
      }),
    ).toBe(true);
    expect(wb.createPivotTable(0, 'Pivot1', 8, { row: 4, col: 1 })).toBe(3);
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
});
