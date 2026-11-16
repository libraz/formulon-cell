import { describe, expect, it } from 'vitest';
import { en } from '../../../../src/i18n/strings/en.js';
import {
  buildCellEntries,
  buildColEntries,
  buildMacCellEntries,
  buildPivotEntries,
  buildRowEntries,
  type MenuEntry,
  type PivotMenuTarget,
} from '../../../../src/interact/context-menu-spec.js';

const pivotTarget: PivotMenuTarget = {
  cell: 'item',
  fieldName: 'Region',
  itemName: 'East',
  dataFieldName: 'Sum of Amt',
  itemOnColumns: false,
  subtotalOn: false,
  filtered: false,
  itemPosition: 0,
  itemCount: 3,
  fieldPosition: 0,
  fieldCount: 1,
  grouped: false,
  groupable: false,
  dateField: false,
  canRefresh: true,
  canOpenFieldList: true,
};

const hints = (entries: readonly MenuEntry[]): Record<string, string> => {
  const out: Record<string, string> = {};
  for (const entry of entries) {
    if (entry.kind === 'item' && entry.hint !== undefined) out[entry.id] = entry.hint;
    if (entry.kind === 'submenu') Object.assign(out, hints(entry.children));
  }
  return out;
};

describe('context-menu shortcut hints', () => {
  const windowsMenus: Record<string, MenuEntry[]> = {
    cell: buildCellEntries(en, false),
    cellWithNote: buildCellEntries(en, true),
    row: buildRowEntries(en),
    col: buildColEntries(en),
    pivot: buildPivotEntries(en, pivotTarget, 'default'),
  };

  it.each(Object.entries(windowsMenus))('has no Mac glyph on the Windows %s menu', (_, entries) => {
    const all = Object.values(hints(entries));
    expect(all.length).toBeGreaterThan(0);
    for (const hint of all) expect(hint).not.toMatch(/[⌘⌥⇧⌃^]/);
  });

  it('spells the Windows hints as the ribbon shortcut formatter does', () => {
    expect(hints(windowsMenus.cell ?? [])).toMatchObject({
      cut: 'Ctrl+X',
      copy: 'Ctrl+C',
      paste: 'Ctrl+V',
      formatCells: 'Ctrl+1',
      insertHyperlink: 'Ctrl+K',
      selectAll: 'Ctrl+A',
      insertComment: 'Shift+F2',
    });
    expect(hints(windowsMenus.pivot ?? [])).toMatchObject({
      copy: 'Ctrl+C',
      formatCells: 'Ctrl+1',
    });
  });

  it('keeps the Mac glyphs', () => {
    expect(hints(buildPivotEntries(en, pivotTarget, 'mac'))).toMatchObject({
      copy: '⌘C',
      formatCells: '⌘1',
    });
    expect(hints(buildMacCellEntries(en, true))).toMatchObject({
      cut: '⌘X',
      formatCells: '⌘1',
    });
  });
});
