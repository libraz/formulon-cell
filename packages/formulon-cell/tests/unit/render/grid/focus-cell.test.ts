import { describe, expect, it } from 'vitest';
import { cellRect, layoutForView } from '../../../../src/render/geometry.js';
import { focusCellBand, paintFocusCellBand } from '../../../../src/render/grid/focus-cell.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';

const view = (
  active: { row: number; col: number },
  range: { r0: number; c0: number; r1: number; c1: number },
) => {
  const state = createSpreadsheetStore().getState();
  state.viewport = { rowStart: 0, rowCount: 20, colStart: 0, colCount: 10, zoom: 1, widthPx: 800 };
  state.selection = {
    ...state.selection,
    active: { sheet: 0, ...active },
    anchor: { sheet: 0, ...active },
    range: { sheet: 0, ...range },
  };
  return { ...state, layout: layoutForView(state) };
};

describe('Focus Cell band', () => {
  it("bands the active cell's whole row and column", () => {
    const v = view({ row: 2, col: 2 }, { r0: 2, c0: 2, r1: 2, c1: 2 });
    const { band, selection } = focusCellBand(v);
    const c3 = cellRect(v.layout, v.viewport, 2, 2);
    const a3 = cellRect(v.layout, v.viewport, 2, 0);
    const c1 = cellRect(v.layout, v.viewport, 0, 2);
    // One strip along row 3 from column A, one down column C from row 1.
    expect(band).toContainEqual(expect.objectContaining({ x: a3.x, y: c3.y, h: c3.h }));
    expect(band).toContainEqual(expect.objectContaining({ x: c3.x, y: c1.y, w: c3.w }));
    expect(band.every((r) => r.w > c3.w || r.h > c3.h)).toBe(true);
    expect(selection).toEqual([c3]);
  });

  it('leaves the whole selection out of the band, active cell included', () => {
    const v = view({ row: 2, col: 2 }, { r0: 2, c0: 2, r1: 7, c1: 4 });
    const fills: unknown[] = [];
    const clipped: [number, number, number, number][] = [];
    let rule: string | undefined;
    const ctx = {
      save: () => undefined,
      restore: () => undefined,
      beginPath: () => undefined,
      rect: (x: number, y: number, w: number, h: number) => clipped.push([x, y, w, h]),
      clip: (r: string) => {
        rule = r;
      },
      set fillStyle(v: string) {
        fills.push(v);
      },
      fillRect: () => undefined,
    } as unknown as CanvasRenderingContext2D;
    paintFocusCellBand({ ctx, dpr: 1, cssWidth: 800, cssHeight: 400 }, v);
    const sel = focusCellBand(v).selection[0];
    if (!sel) throw new Error('no selection rect');
    expect(rule).toBe('evenodd');
    expect(clipped).toEqual([
      [0, 0, 800, 400],
      [sel.x, sel.y, sel.w, sel.h],
    ]);
    expect(fills).toEqual(['rgba(35, 170, 60, 0.2)']);
  });
});
