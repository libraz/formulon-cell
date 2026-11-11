import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  createTableStyleFromActiveTable,
  customTableStyleId,
  formatAsTable,
} from '../../../src/commands/format-as-table.js';
import {
  activeStyleTarget,
  type TableStyleKind,
} from '../../../src/commands/table-style-actions.js';
import { addrKey } from '../../../src/engine/address.js';
import { mutators } from '../../../src/store/store.js';
import type { MountedStubSheet } from '../../test-utils/mount.js';
import { mountPivotSheet } from '../mount/toolbar/fixtures.js';

vi.setConfig({ testTimeout: 20_000 });

const TABLE_CELL = { sheet: 0, row: 21, col: 0 };
const PIVOT_CELL = { sheet: 0, row: 9, col: 4 };

// Both galleries read the current style and apply a pick through one per-kind target.
describe('the style target under the active cell', () => {
  let sheet: MountedStubSheet;
  beforeEach(async () => {
    sheet = await mountPivotSheet();
    for (const [row, text] of ['Head', 'a', 'b'].entries()) {
      sheet.workbook.setText({ sheet: 0, row: 20 + row, col: 0 }, text);
    }
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    formatAsTable(
      sheet.instance.store,
      { sheet: 0, r0: 20, c0: 0, r1: 22, c1: 0 },
      {
        workbook: sheet.workbook,
      },
    );
  });
  afterEach(() => sheet.dispose());

  const target = (kind: TableStyleKind) =>
    activeStyleTarget(
      {
        store: sheet.instance.store,
        history: sheet.instance.history,
        workbook: sheet.workbook,
        origin: 'ribbon',
        commandId: kind === 'table' ? 'tableQuickStyles' : 'pivotQuickStyles',
      },
      kind,
    );
  const activate = (addr: typeof TABLE_CELL) => mutators.setActive(sheet.instance.store, addr);

  it('is null outside its kind', () => {
    activate({ sheet: 0, row: 30, col: 9 });
    expect(target('table')).toBeNull();
    expect(target('pivot')).toBeNull();
    activate(TABLE_CELL);
    expect(target('pivot')).toBeNull();
    activate(PIVOT_CELL);
    expect(target('table')).toBeNull();
  });

  it('reads, applies and clears a table style, a custom one by its gallery id', () => {
    activate(TABLE_CELL);
    expect(target('table')?.kind).toBe('table');
    expect(target('table')?.apply('TableStyleLight1')).toBe(true);
    expect(target('table')?.currentId).toBe('TableStyleLight1');
    createTableStyleFromActiveTable(sheet.instance.store, sheet.instance.history, 'Brand', {
      elements: { headerRow: { bold: true } },
    });
    const brand = customTableStyleId('Brand');
    expect(target('table')?.apply(brand)).toBe(true);
    expect(target('table')?.currentId).toBe(brand);
    expect(target('table')?.clear()).toBe(true);
    expect(target('table')?.currentId).toBe('');
  });

  it('reads, applies and clears a PivotTable style', () => {
    activate(PIVOT_CELL);
    expect(target('pivot')?.apply('PivotStyleMedium9')).toBe(true);
    expect(target('pivot')?.currentId).toBe('PivotStyleMedium9');
    expect(target('pivot')?.clear()).toBe(true);
    expect(target('pivot')?.currentId).toBe('');
  });

  it('keeps a PivotTable cell format when Apply and Clear Formatting is refused', () => {
    activate(PIVOT_CELL);
    mutators.setCellFormat(sheet.instance.store, PIVOT_CELL, { fill: '#ff0000' });
    vi.spyOn(sheet.workbook, 'setPivotStyleInfo').mockReturnValue(false);
    expect(target('pivot')?.apply('PivotStyleMedium9', { clearFormatting: true })).toBe(false);
    expect(sheet.instance.store.getState().format.formats.get(addrKey(PIVOT_CELL))?.fill).toBe(
      '#ff0000',
    );
  });

  it('clears a PivotTable cell format with the style as one undo step', () => {
    activate(PIVOT_CELL);
    mutators.setCellFormat(sheet.instance.store, PIVOT_CELL, { fill: '#ff0000' });
    expect(target('pivot')?.apply('PivotStyleMedium9', { clearFormatting: true })).toBe(true);
    const fill = () =>
      sheet.instance.store.getState().format.formats.get(addrKey(PIVOT_CELL))?.fill;
    expect(fill()).toBeUndefined();
    expect(sheet.instance.history.undo()).toBe(true);
    expect(fill()).toBe('#ff0000');
    expect(target('pivot')?.currentId).not.toBe('PivotStyleMedium9');
  });
});
