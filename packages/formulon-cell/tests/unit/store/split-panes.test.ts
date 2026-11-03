import { describe, expect, it } from 'vitest';
import { splitPosition } from '../../../src/store/split-panes.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('splitPosition', () => {
  it('measures the split across the row-number strip the view draws', () => {
    // Mac, MS PGothic 11 digits (5.5 pt advance, 6 px): rows past 999 widen the strip to 5 + 4 x 6 = 29 px.
    const state = createSpreadsheetStore().getState();
    state.layout = {
      ...state.layout,
      screenScale: 0.75,
      headerColWidth: 23,
      headerRowHeight: 22,
      headerDigitAdvance: 5.5,
      defaultColWidth: 100,
      splitPanes: true,
      freezeCols: 1,
      freezeRows: 0,
    };
    const at = (rowStart: number) =>
      splitPosition({ ...state, viewport: { ...state.viewport, rowStart, rowCount: 40, zoom: 1 } });
    // Twips: 15 per canonical px; the screen strip is 1 / 0.75 canonical px per px.
    expect(at(0)?.x).toBe(Math.round((23 / 0.75 + 100) * 15));
    expect(at(990)?.x).toBe(Math.round((29 / 0.75 + 100) * 15));
  });
});
