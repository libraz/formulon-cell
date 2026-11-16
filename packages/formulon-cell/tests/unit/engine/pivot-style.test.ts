import { describe, expect, it } from 'vitest';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { findPivotTableAtCell } from '../../../src/engine/passthrough-sync.js';
import { PIVOT_CELL_KIND as K } from '../../../src/engine/pivot-cell-kind.js';
import {
  buildPivotStyleLayer,
  builtinPivotStyleElements,
  type PivotLayoutCell,
  pivotRowOutline,
  pivotStyleRegions,
} from '../../../src/engine/pivot-style.js';
import {
  composeStyleElements,
  PIVOT_ELEMENT_ORDER,
} from '../../../src/engine/table-style-elements.js';
import { PivotAggregation, PivotAxis } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

/** clrScheme of the measured workbook (Office 2007-2010 theme): accent1 #4f81bd. */
const THEME = [
  0xff000000, 0xffffffff, 0xff1f497d, 0xffeeece1, 0xff4f81bd, 0xffc0504d, 0xff9bbb59, 0xff8064a2,
  0xff4bacc6, 0xfff79646, 0xff0000ff, 0xff800080,
];

// Measured on PivotStyleLight16 (sRGB): header / grand-total fill (220,229,241)
// within one level of the P3 fit (the theme tint gives 230), rules (149,179,215).
const BAND = '#dce6f1';
const RULE = { style: 'thin', color: '#95b3d7' };

const INFO = { showRowHeaders: true, showColHeaders: true, showRowStripes: false };

const cell = (row: number, col: number, kind: number, depth = 0): PivotLayoutCell => ({
  row,
  col,
  kind,
  depth,
});

/** s1: rows Region, values Sum of Amt, at A3:B7. */
const S1_RANGE = { sheet: 0, r0: 2, c0: 0, r1: 6, c1: 1 };
const S1 = [
  cell(2, 0, K.Header),
  cell(2, 1, K.Header),
  ...[3, 4, 5].flatMap((r) => [cell(r, 0, K.RowLabel), cell(r, 1, K.Data)]),
  cell(6, 0, K.GrandTotal),
  cell(6, 1, K.GrandTotal),
];

/** s2: rows Product > Region, subtotals at the top, at A3:B11. */
const S2_RANGE = { sheet: 0, r0: 2, c0: 0, r1: 10, c1: 1 };
const S2 = [
  cell(2, 0, K.Header),
  cell(2, 1, K.Header),
  cell(3, 0, K.RowSubtotal),
  cell(3, 1, K.RowSubtotal),
  ...[4, 5, 6].flatMap((r) => [cell(r, 0, K.RowLabel, 1), cell(r, 1, K.Data)]),
  cell(7, 0, K.RowSubtotal),
  cell(7, 1, K.RowSubtotal),
  ...[8, 9].flatMap((r) => [cell(r, 0, K.RowLabel, 1), cell(r, 1, K.Data)]),
  cell(10, 0, K.GrandTotal),
  cell(10, 1, K.GrandTotal),
];

const resolve = (range: typeof S1_RANGE, cells: PivotLayoutCell[]) => {
  const elements = builtinPivotStyleElements('PivotStyleLight16', THEME);
  if (!elements) throw new Error('no elements');
  const { regions, buttons } = pivotStyleRegions(range, cells, INFO);
  return {
    at: (row: number, col: number) =>
      composeStyleElements(elements, PIVOT_ELEMENT_ORDER, regions, row, col),
    buttons,
  };
};

describe('PivotStyleLight16', () => {
  it('fills the header row, bold, with a rule along its bottom', () => {
    const { at } = resolve(S1_RANGE, S1);
    for (const col of [0, 1]) {
      expect(at(2, col)).toEqual({ fill: BAND, bold: true, borders: { bottom: RULE } });
    }
  });

  it('leaves item rows plain', () => {
    const { at } = resolve(S1_RANGE, S1);
    for (const row of [3, 4, 5]) {
      expect(at(row, 0)).toBeUndefined();
      expect(at(row, 1)).toBeUndefined();
    }
  });

  it('fills the grand total row, bold, with a rule along its top', () => {
    const { at } = resolve(S1_RANGE, S1);
    expect(at(6, 1)).toEqual({ fill: BAND, bold: true, borders: { top: RULE } });
  });

  it('makes outer items bold with a rule above them; inner items stay plain', () => {
    const { at } = resolve(S2_RANGE, S2);
    for (const row of [3, 7]) {
      expect(at(row, 0)).toEqual({ bold: true, borders: { top: RULE } });
      expect(at(row, 1)).toEqual({ bold: true, borders: { top: RULE } });
    }
    for (const row of [4, 5, 6, 8, 9]) expect(at(row, 0)).toBeUndefined();
  });

  it('puts the field button on the Row Labels caption only', () => {
    expect(resolve(S1_RANGE, S1).buttons).toEqual([{ row: 2, col: 0, slot: 'rows' }]);
  });

  it('tints the other members of the family from their theme colour', () => {
    expect(builtinPivotStyleElements('PivotStyleLight17', THEME)?.headerRow?.fill).toBe('#f2dcdb');
    expect(builtinPivotStyleElements('Custom', THEME)).toBeNull();
  });
});

describe('pivotRowOutline', () => {
  it('pushes outer items one step in with a collapse button, inner items two', () => {
    const outline = new Map(pivotRowOutline(S2_RANGE, S2).map((l) => [l.row, l.outline] as const));
    expect(outline.get(3)).toEqual({ steps: 1, collapse: true });
    expect(outline.get(4)).toEqual({ steps: 2, collapse: false });
    expect(outline.get(7)).toEqual({ steps: 1, collapse: true });
    expect(outline.has(2)).toBe(false);
    expect(outline.has(10)).toBe(false);
  });

  it('leaves single-field items at the cell padding', () => {
    for (const l of pivotRowOutline(S1_RANGE, S1)) {
      expect(l.outline).toEqual({ steps: 0, collapse: false });
    }
  });
});

describe('buildPivotStyleLayer', () => {
  it('styles a report created in session from its layout', async () => {
    const wb = await WorkbookHandle.createDefault();
    const rows: (string | number)[][] = [
      ['Region', 'Product', 'Amt'],
      ['East', 'A', 500],
      ['North', 'A', 100],
      ['North', 'B', 200],
    ];
    for (const [r, row] of rows.entries()) {
      for (const [c, v] of row.entries()) {
        const addr = { sheet: 0, row: r, col: c };
        if (typeof v === 'number') wb.setNumber(addr, v);
        else wb.setText(addr, v);
      }
    }
    const created = createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 2 },
      destination: { sheet: 0, row: 6, col: 0 },
      rowField: 'Product',
      valueField: 'Amt',
      aggregation: PivotAggregation.Sum,
    });
    expect(created.ok).toBe(true);
    wb.addPivotField(0, 0, { sourceName: 'Region', axis: PivotAxis.Row });
    const layer = buildPivotStyleLayer(wb, 0, {});
    expect(layer.ranges).toEqual([{ sheet: 0, r0: 6, c0: 0, r1: 12, c1: 1 }]);
    const accent = builtinPivotStyleElements('PivotStyleLight16', wb.getTheme()?.colors);
    expect(layer.formatAt(6, 0)).toEqual(accent?.headerRow);
    // A (subtotal at the top), then its two regions, then B and its region.
    expect(layer.formatAt(7, 1)).toEqual(accent?.firstRowSubheading);
    expect(layer.formatAt(8, 0)).toBeUndefined();
    expect(layer.formatAt(10, 0)).toEqual(accent?.firstRowSubheading);
    expect(layer.formatAt(12, 1)).toEqual(accent?.totalRow);
    expect([...layer.buttons]).toEqual(['0:6:0']);
    wb.dispose();
  });
});

describe('report options in the style layer', () => {
  const twoRowFieldReport = async (): Promise<WorkbookHandle> => {
    const wb = await WorkbookHandle.createDefault();
    const rows: (string | number)[][] = [
      ['Region', 'Product', 'Amt'],
      ['East', 'A', 500],
      ['North', 'A', 100],
      ['North', 'B', 200],
    ];
    for (const [r, row] of rows.entries()) {
      for (const [c, v] of row.entries()) {
        const addr = { sheet: 0, row: r, col: c };
        if (typeof v === 'number') wb.setNumber(addr, v);
        else wb.setText(addr, v);
      }
    }
    createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 2 },
      destination: { sheet: 0, row: 6, col: 0 },
      rowField: 'Product',
      valueField: 'Amt',
      aggregation: PivotAggregation.Sum,
    });
    wb.addPivotField(0, 0, { sourceName: 'Region', axis: PivotAxis.Row });
    return wb;
  };

  it('draws no collapse buttons with +/- Buttons off, moving the labels left by the button', async () => {
    const wb = await twoRowFieldReport();
    expect(buildPivotStyleLayer(wb, 0, {}).outlineAt(7, 0)).toEqual({ steps: 1, collapse: true });
    expect(wb.updatePivotOptions(0, 0, { showDrill: false })).toBe(true);
    const layer = buildPivotStyleLayer(wb, 0, {});
    expect(layer.outlineAt(7, 0)).toEqual({ steps: 0, collapse: false });
    expect(layer.outlineAt(8, 0)).toEqual({ steps: 1, collapse: false });
    expect([...layer.buttons]).toEqual(['0:6:0']);
    wb.dispose();
  });
});

describe('empty report placeholder', () => {
  it('reserves 3 columns x 18 rows from the anchor and drops the style', async () => {
    const wb = await WorkbookHandle.createDefault();
    wb.setText({ sheet: 0, row: 0, col: 0 }, 'Region');
    wb.setText({ sheet: 0, row: 0, col: 1 }, 'Amt');
    wb.setText({ sheet: 0, row: 1, col: 0 }, 'East');
    wb.setNumber({ sheet: 0, row: 1, col: 1 }, 5);
    const created = createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
      destination: { sheet: 0, row: 4, col: 0 },
      rowField: '',
      valueField: '',
    });
    expect(created.ok).toBe(true);
    const layer = buildPivotStyleLayer(wb, 0, {});
    expect(layer.placeholders).toEqual([
      { pivotIndex: 0, range: { sheet: 0, r0: 4, c0: 0, r1: 21, c1: 2 }, name: 'PivotTable1' },
    ]);
    expect(layer.ranges).toEqual([{ sheet: 0, r0: 4, c0: 0, r1: 21, c1: 2 }]);
    expect(layer.formatAt(4, 0)).toBeUndefined();
    expect(layer.buttons.size).toBe(0);
    wb.dispose();
  });

  it('projects no caption cell, so the anchor reads blank', async () => {
    const wb = await WorkbookHandle.createDefault();
    wb.setText({ sheet: 0, row: 0, col: 0 }, 'Region');
    wb.setText({ sheet: 0, row: 0, col: 1 }, 'Amt');
    wb.setText({ sheet: 0, row: 1, col: 0 }, 'East');
    wb.setNumber({ sheet: 0, row: 1, col: 1 }, 5);
    const created = createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
      destination: { sheet: 0, row: 4, col: 0 },
      rowField: '',
      valueField: '',
    });
    expect(created.ok).toBe(true);
    expect([...wb.pivotCells(0)]).toEqual([]);
    expect([...wb.cells(0)].some((c) => c.addr.row === 4 && c.addr.col === 0)).toBe(false);
    wb.dispose();
  });

  it('counts every cell of the placeholder area as inside the report', async () => {
    const wb = await WorkbookHandle.createDefault();
    wb.setText({ sheet: 0, row: 0, col: 0 }, 'Region');
    wb.setText({ sheet: 0, row: 0, col: 1 }, 'Amt');
    wb.setText({ sheet: 0, row: 1, col: 0 }, 'East');
    wb.setNumber({ sheet: 0, row: 1, col: 1 }, 5);
    const created = createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
      destination: { sheet: 0, row: 2, col: 0 },
      rowField: '',
      valueField: '',
    });
    expect(created.ok).toBe(true);
    const at = (row: number, col: number) => findPivotTableAtCell(wb, { sheet: 0, row, col });
    expect(at(2, 0)).not.toBeNull();
    expect(at(9, 1)).not.toBeNull();
    expect(at(19, 2)).not.toBeNull();
    expect(at(2, 3)).toBeNull();
    expect(at(20, 0)).toBeNull();
    wb.dispose();
  });
});
