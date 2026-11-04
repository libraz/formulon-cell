import { describe, expect, it } from 'vitest';
import {
  ACCESSIBILITY_SECTIONS,
  checkAccessibility,
  contrastRatio,
  suggestedTextColors,
} from '../../../src/commands/accessibility-check.js';
import { formatAsTable } from '../../../src/commands/format-as-table.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('contrastRatio', () => {
  it('follows the WCAG relative-luminance formula', () => {
    expect(contrastRatio('#000000', '#ffffff')).toBeCloseTo(21, 5);
    expect(contrastRatio('#ffffff', '#ffffff')).toBeCloseTo(1, 5);
    // The measured failing pair: red text on pink.
    expect(contrastRatio('#ff0000', '#ffc0cb')).toBeLessThan(4.5);
  });

  it('suggests text colours that reach the AA and AAA ratios on the same fill', () => {
    const [aa, aaa] = suggestedTextColors('#ff0000', '#ffc0cb');
    expect(contrastRatio(aa as string, '#ffc0cb')).toBeGreaterThanOrEqual(4.5);
    expect(contrastRatio(aaa as string, '#ffc0cb')).toBeGreaterThanOrEqual(7);
    expect(aa).not.toBe(aaa);
  });
});

describe('checkAccessibility', () => {
  it('lists the desktop sections and rules in order', () => {
    expect(ACCESSIBILITY_SECTIONS.map((s) => [s.id, s.rules])).toEqual([
      ['colorContrast', ['contrast', 'redFormat']],
      ['media', ['altText']],
      ['tables', ['tableHeader', 'mergedCells']],
      ['structure', ['defaultSheetName']],
      ['access', ['restrictedAccess']],
    ]);
  });

  it('flags low-contrast text and a header-less table, as the desktop did', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const store = createSpreadsheetStore();
    const at = (row: number, col: number) => ({ sheet: 0, row, col });
    // A1:C1 merged title, A3 red on pink, a table A7:B8 without its header row.
    wb.setText(at(0, 0), 'merged title');
    mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
    wb.setText(at(2, 0), 'warning');
    mutators.setCellFormat(store, at(2, 0), { color: '#ff0000', fill: '#ffc0cb' });
    wb.setText(at(3, 0), 'plain');
    for (const [r, v] of [
      [6, 1],
      [7, 2],
    ] as const) {
      wb.setNumber(at(r, 0), v);
      wb.setNumber(at(r, 1), v * 10);
    }
    formatAsTable(store, { sheet: 0, r0: 6, c0: 0, r1: 7, c1: 1 }, { showHeader: false });

    const found = checkAccessibility(store.getState(), wb);
    expect(found.contrast.map((f) => f.range)).toEqual([{ sheet: 0, r0: 2, c0: 0, r1: 2, c1: 0 }]);
    expect(found.contrast[0]?.foreground).toBe('#ff0000');
    expect(found.contrast[0]?.background).toBe('#ffc0cb');
    expect(found.tableHeader.map((f) => f.range)).toEqual([
      { sheet: 0, r0: 6, c0: 0, r1: 7, c1: 1 },
    ]);
    // Measured as passing: a merge outside a table, a lone "Sheet1", red text itself.
    expect(found.mergedCells).toEqual([]);
    expect(found.defaultSheetName).toEqual([]);
    expect(found.redFormat).toEqual([]);
    expect(found.altText).toEqual([]);
    expect(found.restrictedAccess).toEqual([]);
    wb.dispose();
  });
});
