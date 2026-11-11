import { describe, expect, it } from 'vitest';
import {
  type CustomTableStyle,
  defaultTableOverlay,
  type TableOverlay,
} from '../../../src/commands/format-as-table.js';
import { tableCellFormat } from '../../../src/render/grid/table-format.js';

const edge = { style: 'thin' as const, color: '#000000' };
const custom = (
  stripeSizes: CustomTableStyle['stripeSizes'],
  borders: Record<string, typeof edge>,
): CustomTableStyle => ({
  id: 'custom-table:Bands',
  label: 'Bands',
  elements: { firstRowStripe: { borders }, firstColumnStripe: { borders } },
  stripeSizes,
});
const table = (patch: Partial<TableOverlay>): TableOverlay => ({
  ...defaultTableOverlay('t', { sheet: 0, r0: 0, c0: 0, r1: 8, c1: 5 }),
  styleName: 'Bands',
  ...patch,
});

describe('table stripe bands', () => {
  it('draws a multi-row band border on the band edge, not on every row', () => {
    const style = custom({ firstRowStripe: 2 }, { top: edge, bottom: edge });
    const t = table({ showTotal: true });
    const at = (row: number) => tableCellFormat(t, row, 1, undefined, [style])?.borders;
    // Header row 0; first bands are rows 1-2 and 4-5 (the second band is one row).
    expect(at(1)).toEqual({ top: edge });
    expect(at(2)).toEqual({ bottom: edge });
    expect(at(3)).toBeUndefined();
    expect(at(4)).toEqual({ top: edge });
    expect(at(5)).toEqual({ bottom: edge });
  });

  it('clips a band that runs past the body to the last body row', () => {
    const style = custom({ firstRowStripe: 3 }, { bottom: edge });
    const t = table({ showTotal: true });
    // Bands are rows 1-3 and 5-7; the last band ends at body row 7, above the total row.
    expect(tableCellFormat(t, 7, 1, undefined, [style])?.borders).toEqual({ bottom: edge });
    expect(tableCellFormat(t, 6, 1, undefined, [style])?.borders).toBeUndefined();
  });

  it('draws a multi-column band border on the band edge and clips column bands to the body', () => {
    const style = custom({ firstColumnStripe: 2 }, { left: edge, right: edge, top: edge });
    const t = table({ colStripes: true, banded: false });
    const at = (row: number, col: number) =>
      tableCellFormat(t, row, col, undefined, [style])?.borders;
    expect(at(3, 0)).toEqual({ left: edge });
    expect(at(3, 1)).toEqual({ right: edge });
    // The top edge lands on the first body row, not on the header row.
    expect(at(1, 0)).toEqual({ left: edge, top: edge });
    expect(at(0, 0)).toBeUndefined();
  });
});
