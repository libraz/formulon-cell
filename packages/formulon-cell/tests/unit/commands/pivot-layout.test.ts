import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  pivotFieldFilterModel,
  setPivotFieldItemsChecked,
} from '../../../src/commands/pivot-filter.js';
import {
  applyPivotGrandTotals,
  applyPivotReportLayout,
  applyPivotSubtotalPosition,
  clearPivotFields,
  clearPivotFilters,
  pivotLayoutState,
} from '../../../src/commands/pivot-layout.js';
import { PivotAxis, PivotReportLayout } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { commitPivotChange } from '../../../src/interact/pivot-field-settings-dialog-apply.js';
import { loadFixture } from '../engine/pivot-excel-parity/fixture.js';

// ja_c0_rows.xlsx: PivotTable1 at A3:C7 on the first sheet, Region on rows,
// two value fields, grand totals on.
const pivot = { sheetIndex: 0, pivotIndex: 0 };

describe('PivotTable layout commands', () => {
  let wb: WorkbookHandle;
  let history: History;
  const commit = (apply: () => boolean): boolean =>
    commitPivotChange({ wb, history, onChanged: () => {} }, 0, 0, apply);
  const definition = () => wb.getPivotDefinition(0, 0);

  beforeEach(async () => {
    wb = await loadFixture('ja_c0_rows.xlsx');
    history = new History();
  });

  afterEach(() => {
    wb.dispose();
  });

  it('reads the state the Design menus check', () => {
    expect(pivotLayoutState(wb, pivot)).toMatchObject({
      grandTotals: { rows: true, cols: true },
      subtotalTop: true,
      hasFields: true,
      hasFilter: false,
    });
  });

  it('switches grand totals as one undo step', () => {
    expect(commit(() => applyPivotGrandTotals(wb, pivot, { rows: false, cols: true }))).toBe(true);
    expect(definition()?.grandTotals).toEqual({ rows: false, cols: true });
    expect(history.undo()).toBe(true);
    expect(definition()?.grandTotals).toEqual({ rows: true, cols: true });
    expect(history.redo()).toBe(true);
    expect(definition()?.grandTotals).toEqual({ rows: false, cols: true });
  });

  it('changes the report layout as one undo step', () => {
    const before = definition()?.layout;
    expect(commit(() => applyPivotReportLayout(wb, pivot, PivotReportLayout.Tabular))).toBe(true);
    expect(definition()?.layout).toBe(PivotReportLayout.Tabular);
    expect(history.undo()).toBe(true);
    expect(definition()?.layout).toBe(before);
  });

  it('moves every axis field subtotal to the bottom as one undo step', () => {
    expect(commit(() => applyPivotSubtotalPosition(wb, pivot, false))).toBe(true);
    const axisFields = () =>
      (definition()?.fields ?? []).filter(
        (f) => f.axis === PivotAxis.Row || f.axis === PivotAxis.Col,
      );
    expect(axisFields().every((f) => !f.subtotalTop)).toBe(true);
    expect(pivotLayoutState(wb, pivot)?.subtotalTop).toBe(false);
    expect(history.undo()).toBe(true);
    expect(axisFields().every((f) => f.subtotalTop)).toBe(true);
  });

  it('clears every field, and undo puts them back', () => {
    expect(commit(() => clearPivotFields(wb, pivot))).toBe(true);
    expect(definition()?.dataFields).toEqual([]);
    expect(definition()?.rowFieldOrder).toEqual([]);
    expect(pivotLayoutState(wb, pivot)?.hasFields).toBe(false);
    expect(history.undo()).toBe(true);
    expect(definition()?.dataFields).toHaveLength(2);
  });

  it('clears item filters on every field', () => {
    const region = definition()?.fields.find((f) => f.axis === PivotAxis.Row);
    if (!region) throw new Error('no row field');
    const model = pivotFieldFilterModel(wb, 0, 0, region.index);
    if (!model) throw new Error('no filter model');
    setPivotFieldItemsChecked(
      wb,
      0,
      0,
      region.index,
      model.items.map((item) => ({ ...item, checked: item.value !== 'North' })),
    );
    expect(pivotLayoutState(wb, pivot)?.hasFilter).toBe(true);
    expect(commit(() => clearPivotFilters(wb, pivot))).toBe(true);
    expect(pivotLayoutState(wb, pivot)?.hasFilter).toBe(false);
  });
});
