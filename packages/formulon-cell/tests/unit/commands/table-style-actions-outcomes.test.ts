import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  createPivotTableStyleFromActivePivot,
  customTableStyleId,
  formatAsTable,
  updateTableOverlay,
} from '../../../src/commands/format-as-table.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import {
  applyPivotStyle,
  applyTableStyle,
  clearTableStyle,
  createCustomStyle,
  deleteCustomStyle,
  modifyCustomStyle,
} from '../../../src/commands/table-style-actions.js';
import { addrKey } from '../../../src/engine/address.js';
import { mutators } from '../../../src/store/store.js';
import type { MountedStubSheet } from '../../test-utils/mount.js';
import { mountPivotSheet } from '../mount/toolbar/fixtures.js';

vi.setConfig({ testTimeout: 20_000 });

const elements = { headerRow: { fill: '#336699', bold: true } };
const RANGE = { sheet: 0, r0: 20, c0: 0, r1: 22, c1: 0 };

describe('table style commands', () => {
  let sheet: MountedStubSheet;
  beforeEach(async () => {
    sheet = await mountPivotSheet();
    for (const [row, text] of ['Head', 'a', 'b'].entries()) {
      sheet.workbook.setText({ sheet: 0, row: 20 + row, col: 0 }, text);
    }
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    formatAsTable(sheet.instance.store, RANGE, { workbook: sheet.workbook });
  });
  afterEach(() => sheet.dispose());

  const store = () => sheet.instance.store;
  const history = () => sheet.instance.history;
  const ctx = () => ({ store: store(), history: history() });
  const entry = () => ({ ...ctx(), origin: 'ribbon' as const, commandId: 'test' });
  const table = () => {
    const found = store().getState().tables.tables[0];
    if (!found) throw new Error('no table');
    return found;
  };
  // An unset default reads as '' once cleared and as the engine's own default before.
  const UNSET = { tableStyle: 'TableStyleMedium2', pivotStyle: 'PivotStyleLight16' };
  const defaults = () => {
    const read = sheet.workbook.getTableStyleDefaults();
    return (
      read && {
        tableStyle: read.tableStyle || UNSET.tableStyle,
        pivotStyle: read.pivotStyle || UNSET.pivotStyle,
      }
    );
  };
  const value = (name: string, setAsDefault: boolean) => ({
    name,
    elements,
    stripeSizes: {},
    setAsDefault,
  });

  describe('results of a change that does nothing', () => {
    it('reports a refused restyle and clear instead of success', () => {
      const stale = { ...table(), id: 'gone' };
      expect(applyTableStyle(entry(), stale, 'TableStyleLight1')).toBe(false);
      expect(clearTableStyle(entry(), stale)).toBe(false);
    });

    it('keeps the cell formatting when Apply and Clear Formatting is refused', () => {
      const cell = { sheet: 0, row: 21, col: 0 };
      mutators.setCellFormat(store(), cell, { fill: '#ff0000' });
      const stale = { ...table(), id: 'gone' };
      expect(applyTableStyle(entry(), stale, 'TableStyleLight1', { clearFormatting: true })).toBe(
        false,
      );
      expect(store().getState().format.formats.get(addrKey(cell))?.fill).toBe('#ff0000');
    });

    it('keeps a PivotTable custom assignment when the built-in style cannot be written', () => {
      createPivotTableStyleFromActivePivot(store(), history(), 'Brand', { elements });
      const at = { sheetIndex: 0, pivotIndex: 99 };
      const assignment = { ...at, styleId: 'custom-pivot-table:Brand' };
      mutators.upsertPivotTableStyle(store(), assignment);
      expect(applyPivotStyle(ctx(), sheet.workbook, at, 'PivotStyleLight1')).toBe(false);
      expect(store().getState().tables.pivotTableStyles).toEqual([assignment]);
    });
  });

  describe('workbook defaults of a custom style', () => {
    it('undoes and redoes the default set by New together with the style', () => {
      expect(createCustomStyle(ctx(), sheet.workbook, 'table', value('Brand', true))).toBe(true);
      expect(defaults()?.tableStyle).toBe('Brand');
      expect(history().undo()).toBe(true);
      expect(defaults()?.tableStyle).toBe(UNSET.tableStyle);
      expect(store().getState().tables.customTableStyles).toEqual([]);
      expect(history().redo()).toBe(true);
      expect(defaults()?.tableStyle).toBe('Brand');
    });

    it('undoes the default a Modify sets', () => {
      createCustomStyle(ctx(), sheet.workbook, 'table', value('Brand', false));
      const next = { label: 'Brand', elements };
      modifyCustomStyle(ctx(), 'Brand', next, {
        workbook: sheet.workbook,
        kind: 'table',
        setAsDefault: true,
      });
      expect(defaults()?.tableStyle).toBe('Brand');
      expect(history().undo()).toBe(true);
      expect(defaults()?.tableStyle).toBe(UNSET.tableStyle);
    });

    it('clears the default when Modify unticks Set As Default', () => {
      createCustomStyle(ctx(), sheet.workbook, 'table', value('Brand', true));
      modifyCustomStyle(
        ctx(),
        'Brand',
        { label: 'Brand', elements },
        { workbook: sheet.workbook, kind: 'table', setAsDefault: false },
      );
      expect(defaults()?.tableStyle).toBe(UNSET.tableStyle);
      expect(history().undo()).toBe(true);
      expect(defaults()?.tableStyle).toBe('Brand');
    });

    it('moves the defaults of both kinds to the new name when a style is renamed', () => {
      createTableStyleAndPivot('Brand');
      sheet.workbook.setTableStyleDefaults({ tableStyle: 'Brand', pivotStyle: 'Brand' });
      modifyCustomStyle(
        ctx(),
        'Brand',
        { label: 'Renamed', elements },
        { workbook: sheet.workbook, kind: 'table', setAsDefault: true },
      );
      expect(defaults()).toEqual({ tableStyle: 'Renamed', pivotStyle: 'Renamed' });
      expect(history().undo()).toBe(true);
      expect(defaults()).toEqual({ tableStyle: 'Brand', pivotStyle: 'Brand' });
    });

    const createTableStyleAndPivot = (name: string): void => {
      createCustomStyle(ctx(), null, 'table', value(name, false));
      createPivotTableStyleFromActivePivot(store(), history(), name, { elements });
    };
  });

  describe('styles on a protected sheet', () => {
    const useBrand = (): void => {
      createCustomStyle(ctx(), null, 'table', value('Brand', false));
      applyTableStyle(entry(), table(), customTableStyleId('Brand'));
      setProtectedSheet(store(), 0, true);
    };
    const REFUSED = { ok: false, reason: 'workbookHasProtectedSheet' };

    // Excel for Mac: with any sheet of the workbook protected, Delete and Modify of a custom
    // style change nothing, whether or not the style is used and wherever its users are.
    it('refuses to delete or modify a style a protected sheet table uses', () => {
      useBrand();
      expect(deleteCustomStyle(ctx(), sheet.workbook, 'Brand')).toEqual(REFUSED);
      expect(modifyCustomStyle(ctx(), 'Brand', { label: 'Brand 2', elements })).toEqual(REFUSED);
      expect(table().styleName).toBe('Brand');
      expect(
        store()
          .getState()
          .tables.customTableStyles?.map((s) => s.label),
      ).toEqual(['Brand']);
      setProtectedSheet(store(), 0, false);
      expect(deleteCustomStyle(ctx(), sheet.workbook, 'Brand')).toEqual({ ok: true });
    });

    it('refuses to delete or modify an unused style while another sheet is protected', () => {
      createCustomStyle(ctx(), null, 'table', value('Other', false));
      createPivotTableStyleFromActivePivot(store(), history(), 'Pivot Brand', { elements });
      setProtectedSheet(store(), 1, true);
      expect(deleteCustomStyle(ctx(), sheet.workbook, 'Other')).toEqual(REFUSED);
      expect(modifyCustomStyle(ctx(), 'Other', { label: 'O2', elements })).toEqual(REFUSED);
      expect(deleteCustomStyle(ctx(), sheet.workbook, 'Pivot Brand')).toEqual(REFUSED);
      expect(modifyCustomStyle(ctx(), 'Pivot Brand', { label: 'P2', elements })).toEqual(REFUSED);
      expect(
        store()
          .getState()
          .tables.customTableStyles?.map((s) => s.label),
      ).toEqual(['Other']);
      expect(
        store()
          .getState()
          .tables.customPivotTableStyles?.map((s) => s.label),
      ).toEqual(['Pivot Brand']);
    });

    // Delete and Modify check their arguments first, as the engine does: an unknown style or an
    // unusable new name is reported before the protection refusal.
    it('reports a bad argument before the protection refusal, for Delete and Modify alike', () => {
      useBrand();
      createCustomStyle(ctx(), null, 'table', value('Taken', false));
      expect(deleteCustomStyle(ctx(), sheet.workbook, 'Missing')).toEqual({
        ok: false,
        reason: 'notFound',
      });
      expect(modifyCustomStyle(ctx(), 'Missing', { label: 'M2', elements })).toEqual({
        ok: false,
        reason: 'notFound',
      });
      for (const label of ['Taken', '  ']) {
        expect(modifyCustomStyle(ctx(), 'Brand', { label, elements })).toEqual({
          ok: false,
          reason: 'invalidName',
        });
      }
    });

    it('still duplicates a style while a sheet is protected', () => {
      useBrand();
      expect(createCustomStyle(ctx(), null, 'table', value('Brand 2', false))).toBe(true);
    });

    it('still refuses every other table change on a protected sheet', () => {
      useBrand();
      const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
      expect(updateTableOverlay(store(), table().id, { showHeader: false })).toBeNull();
      expect(table().showHeader).toBe(true);
      warn.mockRestore();
    });
  });
});
