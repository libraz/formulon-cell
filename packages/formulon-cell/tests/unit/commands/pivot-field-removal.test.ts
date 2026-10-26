import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { applyPivotEdit, type PivotEdit } from '../../../src/commands/pivot-edit.js';
import { createPivotTableFromRange, deletePivotTables } from '../../../src/commands/pivot-table.js';
import {
  PivotAggregation,
  PivotAxis,
  PivotFilterType,
  PivotFilterValueKind,
  PivotReportLayout,
} from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const SHEET = 0;

const setup = async (opts: { valueFields?: string[]; filterFields?: string[] } = {}) => {
  const wb = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Region', 'Product', 'Sales', 'Qty', 'Note'],
    ['East', 'Desk', 12, 1, 'a'],
    ['East', 'Chair', 8, 2, 'b'],
    ['West', 'Desk', 5, 3, 'c'],
  ];
  for (const [r, row] of rows.entries()) {
    for (const [c, value] of row.entries()) {
      const addr = { sheet: SHEET, row: r, col: c };
      if (typeof value === 'number') wb.setNumber(addr, value);
      else wb.setText(addr, value);
    }
  }
  const created = createPivotTableFromRange(wb, {
    source: { sheet: SHEET, r0: 0, c0: 0, r1: 3, c1: 4 },
    destination: { sheet: SHEET, row: 8, col: 1 },
    name: 'Report',
    rowField: 'Region',
    valueField: opts.valueFields?.[0] ?? 'Sales',
    ...(opts.valueFields ? { valueFields: opts.valueFields } : {}),
    ...(opts.filterFields ? { filterFields: opts.filterFields } : {}),
  });
  expect(created.ok).toBe(true);
  return wb;
};

const axes = (wb: WorkbookHandle) =>
  wb.getPivotDefinition(SHEET, 0)?.fields.map((field) => [field.sourceName, field.axis]);

const editOf = (over: Partial<PivotEdit>): PivotEdit => ({
  fieldListOnly: true,
  name: 'Report',
  anchor: { row: 8, col: 1 },
  rowGrandTotals: true,
  colGrandTotals: true,
  layout: PivotReportLayout.Compact,
  fields: [],
  filters: null,
  ...over,
});

const fieldEdit = (fieldIndex: number, axis: PivotAxis | null) => ({
  fieldIndex,
  axis,
  aggregation: PivotAggregation.Sum,
  numberFormat: '',
  filterItems: [],
  filterChecklist: null,
});

describe('pivot field removal', () => {
  it('creates a report whose unused fields sit on no axis', async () => {
    const wb = await setup();
    expect(axes(wb)).toEqual([
      ['Region', PivotAxis.Row],
      ['Product', null],
      ['Sales', PivotAxis.Value],
      ['Qty', null],
      ['Note', null],
    ]);
    expect(wb.getPivotDefinition(SHEET, 0)?.dataFields.map((d) => d.fieldIndex)).toEqual([2]);
    wb.dispose();
  });

  it('keeps the remaining data fields when a field is unticked, through save and load', async () => {
    const wb = await setup({ valueFields: ['Sales', 'Qty'] });
    const before = wb.getPivotDefinition(SHEET, 0);
    expect(before?.dataFields.map((d) => d.fieldIndex)).toEqual([2, 3]);
    const ok = applyPivotEdit(
      wb,
      { sheetIndex: SHEET, pivotIndex: 0, rows: 4, cols: 3 },
      editOf({
        dataFields: before?.dataFields,
        fields: [
          fieldEdit(0, PivotAxis.Row),
          fieldEdit(2, null),
          { ...fieldEdit(3, PivotAxis.Value), aggregation: PivotAggregation.Sum },
        ],
      }),
    );
    expect(ok).toBe(true);
    const check = (handle: WorkbookHandle) => {
      const definition = handle.getPivotDefinition(SHEET, 0);
      expect(definition?.fields[2]?.axis).toBeNull();
      expect(definition?.fields[3]?.axis).toBe(PivotAxis.Value);
      expect(definition?.dataFields.map((d) => [d.name, d.fieldIndex])).toEqual([
        ['Sum of Qty', 3],
      ]);
      expect(definition?.rowFieldOrder).toEqual([0]);
    };
    check(wb);
    const loaded = await WorkbookHandle.loadBytes(wb.save());
    check(loaded);
    wb.dispose();
    loaded.dispose();
  });

  it('leaves a field that is already out of the report alone', async () => {
    const wb = await setup({ filterFields: ['Product'] });
    wb.addPivotFilter(SHEET, 0, {
      axis: PivotAxis.Page,
      fieldName: 'Product',
      type: PivotFilterType.LabelContains,
      valueKind: PivotFilterValueKind.Text,
      valueText: 'De',
    });
    const ok = applyPivotEdit(
      wb,
      { sheetIndex: SHEET, pivotIndex: 0, rows: 4, cols: 3 },
      editOf({
        fields: [
          fieldEdit(0, PivotAxis.Row),
          fieldEdit(1, PivotAxis.Page),
          fieldEdit(2, PivotAxis.Value),
          fieldEdit(3, null),
          fieldEdit(4, null),
        ],
      }),
    );
    expect(ok).toBe(true);
    expect(wb.getPivotDefinition(SHEET, 0)?.filters).toHaveLength(1);
    wb.dispose();
  });

  it('writes custom subtotal functions and reads them back', async () => {
    const wb = await setup();
    expect(wb.getPivotField(SHEET, 0, 0)).toMatchObject({ defaultSubtotal: true, subtotalFns: [] });
    const ok = applyPivotEdit(
      wb,
      { sheetIndex: SHEET, pivotIndex: 0, rows: 4, cols: 3 },
      editOf({
        fields: [
          {
            ...fieldEdit(0, PivotAxis.Row),
            subtotalFns: [PivotAggregation.Sum, PivotAggregation.Count],
          },
        ],
      }),
    );
    expect(ok).toBe(true);
    expect(wb.getPivotField(SHEET, 0, 0)).toMatchObject({
      defaultSubtotal: false,
      subtotalFns: [PivotAggregation.Sum, PivotAggregation.Count],
    });
    applyPivotEdit(
      wb,
      { sheetIndex: SHEET, pivotIndex: 0, rows: 4, cols: 3 },
      editOf({ fields: [{ ...fieldEdit(0, PivotAxis.Row), subtotalFns: [] }] }),
    );
    expect(wb.getPivotField(SHEET, 0, 0)).toMatchObject({ defaultSubtotal: true, subtotalFns: [] });
    wb.dispose();
  });

  it('restores off-axis fields, subtotals, filters and style when a delete is undone', async () => {
    const wb = await setup({ valueFields: ['Sales', 'Qty'], filterFields: ['Product'] });
    wb.addPivotFieldSubtotalFn(SHEET, 0, 0, PivotAggregation.Sum);
    wb.addPivotFieldSubtotalFn(SHEET, 0, 0, PivotAggregation.Max);
    wb.setPivotFieldSubtotalTop(SHEET, 0, 0, false);
    wb.addPivotFilter(SHEET, 0, {
      axis: PivotAxis.Page,
      fieldName: 'Product',
      type: PivotFilterType.LabelContains,
      valueKind: PivotFilterValueKind.Text,
      valueText: 'De',
    });
    applyPivotEdit(
      wb,
      { sheetIndex: SHEET, pivotIndex: 0, rows: 4, cols: 3 },
      editOf({ fields: [fieldEdit(3, null)] }),
    );
    const raw = (wb as unknown as { wb: import('@libraz/formulon').Workbook }).wb;
    raw.pivotSetStyleInfo(SHEET, 0, { name: 'PivotStyleMedium9', showRowStripes: true });

    const before = wb.getPivotDefinition(SHEET, 0);
    expect(before?.fields[3]?.axis).toBeNull();
    expect(before?.filters).toHaveLength(1);
    expect(before?.style?.name).toBe('PivotStyleMedium9');
    expect(before?.dataFields.map((d) => d.fieldIndex)).toEqual([2]);

    const store = createSpreadsheetStore();
    mutators.replaceCells(store, wb.cells(SHEET));
    const history = new History();
    const range = wb.pivotBounds(SHEET)[0]?.range;
    if (!range) throw new Error('report has no extent');
    expect(
      deletePivotTables({ store, workbook: wb, history }, SHEET, [{ pivotIndex: 0, range }]),
    ).toBe(true);
    expect(wb.getPivotTables()).toEqual([]);
    expect(history.undo()).toBe(true);
    expect(wb.getPivotDefinition(SHEET, 0)).toEqual(before);
    wb.dispose();
  });
});
