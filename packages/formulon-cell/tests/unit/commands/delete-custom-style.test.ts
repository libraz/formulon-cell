import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  customPivotTableStyleId,
  customTableStyleId,
  formatAsTable,
} from '../../../src/commands/format-as-table.js';
import { updatePivotStyleInfo } from '../../../src/commands/pivot-style-info.js';
import { applyTableStyle, deleteCustomStyle } from '../../../src/commands/table-style-actions.js';
import { mutators } from '../../../src/store/store.js';
import type { MountedStubSheet } from '../../test-utils/mount.js';
import { scaledTimeout } from '../../timeouts.js';
import { mountPivotSheet } from '../mount/toolbar/fixtures.js';

vi.setConfig({ testTimeout: scaledTimeout(20_000) });

// Excel measured through the object model: the users of each kind a deleted style was flagged for
// take that kind's workbook default, and a deleted default resets first.
describe('deleteCustomStyle', () => {
  let sheet: MountedStubSheet;
  const pivot = { sheetIndex: 0, pivotIndex: 0 };
  const elements = { headerRow: { fill: '#336699', bold: true } };

  beforeEach(async () => {
    sheet = await mountPivotSheet();
    const wb = sheet.workbook;
    for (const [row, text] of ['Head', 'a', 'b'].entries()) {
      wb.setText({ sheet: 0, row: 20 + row, col: 0 }, text);
    }
    mutators.replaceCells(sheet.instance.store, wb.cells(0));
    formatAsTable(
      sheet.instance.store,
      { sheet: 0, r0: 20, c0: 0, r1: 22, c1: 0 },
      { workbook: wb },
    );
  });

  afterEach(() => sheet.dispose());

  const ctx = () => ({ store: sheet.instance.store, history: sheet.instance.history });
  const table = () => sheet.instance.store.getState().tables.tables[0];
  const pivotStyle = () => sheet.workbook.getPivotStyleInfo(0, 0)?.name;
  const defaults = () => sheet.workbook.getTableStyleDefaults();

  /** Adds the custom style `name` for the given kinds and puts the table and the PivotTable in it. */
  const useStyle = (name: string, kinds: { table?: boolean; pivot?: boolean }): void => {
    const store = sheet.instance.store;
    if (kinds.table) {
      mutators.upsertCustomTableStyle(store, {
        id: customTableStyleId(name),
        label: name,
        elements,
      });
      const current = table();
      if (!current) throw new Error('no table');
      applyTableStyle(
        { ...ctx(), origin: 'ribbon', commandId: 'test' },
        current,
        customTableStyleId(name),
      );
    }
    if (kinds.pivot) {
      mutators.upsertCustomPivotTableStyle(store, {
        id: customPivotTableStyleId(name),
        label: name,
        elements,
      });
      updatePivotStyleInfo({ ...ctx(), workbook: sheet.workbook }, pivot, { name });
    }
  };

  it('moves tables to the unchanged default table style', () => {
    useStyle('Brand', { table: true });
    expect(table()?.styleName).toBe('Brand');
    expect(deleteCustomStyle(ctx(), sheet.workbook, 'Brand')).toEqual({ ok: true });
    expect(table()?.styleName).toBe('TableStyleMedium2');
    expect(sheet.instance.store.getState().tables.customTableStyles).toEqual([]);
  });

  it('moves tables to a changed default table style', () => {
    sheet.workbook.setTableStyleDefaults({ tableStyle: 'TableStyleLight1' });
    useStyle('Brand', { table: true });
    deleteCustomStyle(ctx(), sheet.workbook, 'Brand');
    expect(table()?.styleName).toBe('TableStyleLight1');
  });

  it('resets a deleted default table style before its tables take it', () => {
    useStyle('Brand', { table: true });
    sheet.workbook.setTableStyleDefaults({ tableStyle: 'Brand' });
    deleteCustomStyle(ctx(), sheet.workbook, 'Brand');
    expect(table()?.styleName).toBe('TableStyleMedium2');
    expect(defaults()?.tableStyle || 'TableStyleMedium2').toBe('TableStyleMedium2');
  });

  it('moves PivotTables to the unchanged default PivotTable style', () => {
    useStyle('Brand', { pivot: true });
    expect(pivotStyle()).toBe('Brand');
    deleteCustomStyle(ctx(), sheet.workbook, 'Brand');
    expect(pivotStyle()).toBe('PivotStyleLight16');
    expect(sheet.instance.store.getState().tables.customPivotTableStyles).toEqual([]);
  });

  it('moves PivotTables to a changed default PivotTable style', () => {
    sheet.workbook.setTableStyleDefaults({ pivotStyle: 'PivotStyleMedium9' });
    useStyle('Brand', { pivot: true });
    deleteCustomStyle(ctx(), sheet.workbook, 'Brand');
    expect(pivotStyle()).toBe('PivotStyleMedium9');
  });

  it('resets a deleted default PivotTable style before its PivotTables take it', () => {
    useStyle('Brand', { pivot: true });
    sheet.workbook.setTableStyleDefaults({ pivotStyle: 'Brand' });
    deleteCustomStyle(ctx(), sheet.workbook, 'Brand');
    expect(pivotStyle()).toBe('PivotStyleLight16');
  });

  it('moves each kind of a style flagged for both to its own default', () => {
    sheet.workbook.setTableStyleDefaults({
      tableStyle: 'TableStyleLight1',
      pivotStyle: 'PivotStyleMedium9',
    });
    useStyle('Brand', { table: true, pivot: true });
    deleteCustomStyle(ctx(), sheet.workbook, 'Brand');
    expect(table()?.styleName).toBe('TableStyleLight1');
    expect(pivotStyle()).toBe('PivotStyleMedium9');
  });

  it('resets both defaults when a style flagged for both was the default of both', () => {
    useStyle('Brand', { table: true, pivot: true });
    sheet.workbook.setTableStyleDefaults({ tableStyle: 'Brand', pivotStyle: 'Brand' });
    deleteCustomStyle(ctx(), sheet.workbook, 'Brand');
    expect(table()?.styleName).toBe('TableStyleMedium2');
    expect(pivotStyle()).toBe('PivotStyleLight16');
  });

  it('undoes the delete, the moves and the default reset in one step', () => {
    useStyle('Brand', { table: true, pivot: true });
    sheet.workbook.setTableStyleDefaults({ tableStyle: 'Brand', pivotStyle: 'Brand' });
    deleteCustomStyle(ctx(), sheet.workbook, 'Brand');
    sheet.instance.history.undo();
    expect(table()?.styleName).toBe('Brand');
    expect(pivotStyle()).toBe('Brand');
    expect(defaults()).toMatchObject({ tableStyle: 'Brand', pivotStyle: 'Brand' });
    sheet.instance.history.redo();
    expect(table()?.styleName).toBe('TableStyleMedium2');
    expect(pivotStyle()).toBe('PivotStyleLight16');
  });

  it('refuses a name no custom style has', () => {
    expect(deleteCustomStyle(ctx(), sheet.workbook, 'TableStyleMedium2')).toEqual({
      ok: false,
      reason: 'notFound',
    });
  });
});
