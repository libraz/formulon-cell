import { describe, expect, it } from 'vitest';

import {
  defaultTableOverlay,
  type TableOverlay,
} from '../../../../src/commands/format-as-table.js';
import { tableCellFormat } from '../../../../src/render/grid/table-format.js';

function table(style: TableOverlay['style'], over: Partial<TableOverlay> = {}): TableOverlay {
  return {
    ...defaultTableOverlay('t1', { sheet: 0, r0: 0, c0: 0, r1: 5, c1: 3 }),
    style,
    ...over,
  };
}

describe('render/grid/table-format', () => {
  describe('header row', () => {
    it('applies bold + dark-fill + white-text for the dark style', () => {
      const fmt = tableCellFormat(table('dark'), 0, 1);
      expect(fmt).toEqual({ fill: '#325575', color: '#ffffff', bold: true });
    });

    it('applies a light fill + ink text for the light style', () => {
      const fmt = tableCellFormat(table('light'), 0, 1);
      expect(fmt).toEqual({ fill: '#d1e3f3', color: '#1f1f1f', bold: true });
    });

    it('returns undefined for the header row when showHeader is off', () => {
      const fmt = tableCellFormat(table('medium', { showHeader: false }), 0, 1);
      expect(fmt).toBeUndefined();
    });
  });

  describe('total row', () => {
    it('applies bold + total-fill when showTotal is on', () => {
      const fmt = tableCellFormat(table('medium', { showTotal: true }), 5, 1);
      expect(fmt?.bold).toBe(true);
      // Medium total row reuses the header swatch, which equals the base hue.
      expect(fmt?.fill).toBe('#5b9bd5');
    });

    it('does not apply the total style when showTotal is off', () => {
      const fmt = tableCellFormat(table('medium'), 5, 1);
      // Falls through to either banded or undefined; bold must not be true.
      expect(fmt?.bold ?? false).toBe(false);
    });
  });

  describe('banded rows', () => {
    it('zebra-fills odd data rows', () => {
      const t = table('medium');
      // header at row 0 → data starts at 1; data row 0 (sheet row 1) is "even"
      // — no band; data row 1 (sheet row 2) is "odd" — band applies.
      expect(tableCellFormat(t, 1, 1)).toBeUndefined();
      expect(tableCellFormat(t, 2, 1)?.fill).toBe('#e5eff8');
    });

    it('respects banded=false', () => {
      const t = table('medium', { banded: false });
      expect(tableCellFormat(t, 2, 1)).toBeUndefined();
    });
  });

  describe('out of range', () => {
    it('returns undefined for cells outside the table', () => {
      const t = table('medium', { range: { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 3 } });
      expect(tableCellFormat(t, 0, 0)).toBeUndefined();
      expect(tableCellFormat(t, 4, 4)).toBeUndefined();
    });
  });

  describe('first / last column emphasis', () => {
    it('bolds the first column when firstCol is on', () => {
      const t = table('medium', { firstCol: true, banded: false });
      // First column, non-header data row.
      expect(tableCellFormat(t, 1, 0)).toEqual({ bold: true });
      // Other columns stay plain.
      expect(tableCellFormat(t, 1, 1)).toBeUndefined();
    });

    it('keeps the band fill while still bolding the first column', () => {
      const t = table('medium', { firstCol: true, banded: true });
      const fmt = tableCellFormat(t, 2, 0);
      expect(fmt?.bold).toBe(true);
      expect(fmt?.fill).toBe('#e5eff8');
    });

    it('bolds the last column when lastCol is on', () => {
      const t = table('medium', { lastCol: true, banded: false });
      expect(tableCellFormat(t, 1, 3)).toEqual({ bold: true });
      expect(tableCellFormat(t, 1, 2)).toBeUndefined();
    });
  });
});
