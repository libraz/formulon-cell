import { describe, expect, it } from 'vitest';

import {
  DEFAULT_TABLE_COLOR,
  defaultTableOverlay,
  type TableOverlay,
  tableStyleSwatch,
} from '../../../../src/commands/format-as-table.js';
import {
  buildColLayout,
  buildRowLayout,
  cellRectIn,
  layoutForView,
} from '../../../../src/render/geometry.js';
import { paintCells } from '../../../../src/render/grid/cells.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

const theme = {
  bg: '#fff',
  bgRail: '#eee',
  fg: '#000',
  fgMute: '#666',
  rule: '#ccc',
  accent: '#0078d4',
  cellNumFg: '#000',
  cellFormulaFg: '#333',
  fontUi: 'sans-serif',
  fontMono: 'monospace',
  textCell: 13,
} as ResolvedTheme;

interface Fill {
  style: string;
  x: number;
  y: number;
}

/** A canvas that records the style and origin of every filled rectangle. */
function fillSpy(): { ctx: CanvasRenderingContext2D; fills: Fill[] } {
  const fills: Fill[] = [];
  const target: Record<string, unknown> = {
    fillStyle: '',
    font: '',
    globalAlpha: 1,
    measureText: () => ({ width: 7, actualBoundingBoxAscent: 9, actualBoundingBoxDescent: 2 }),
    fillRect: (x: number, y: number) => fills.push({ style: String(target.fillStyle), x, y }),
  };
  const ctx = new Proxy(target, {
    get: (object, property) => (property in object ? object[property as string] : () => undefined),
    set: (object, property, value) => {
      object[property as string] = value;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
  return { ctx, fills };
}

/** Paint a header-less 4x4 table and report the banded cells as `row,col` keys. */
function bandedCells(over: Partial<TableOverlay>): string[] {
  const store = createSpreadsheetStore();
  mutators.upsertTableOverlay(store, {
    ...defaultTableOverlay('t1', { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 3 }),
    showHeader: false,
    banded: false,
    ...over,
  });
  store.setState((state) => ({
    ...state,
    viewport: { ...state.viewport, widthPx: 640, rowCount: 8, colCount: 8 },
  }));
  const view = { ...store.getState(), layout: layoutForView(store.getState()) };
  const cols = buildColLayout(view.layout, view.viewport);
  const rows = buildRowLayout(view.layout, view.viewport);
  const { ctx, fills } = fillSpy();
  paintCells(
    {
      ctx,
      wb: null,
      locale: 'en-US',
      getDisplay: undefined,
      sheetBackgroundImage: null,
      fontEffects: false,
    },
    view,
    theme,
    cols,
    rows,
  );
  const band = tableStyleSwatch('medium', DEFAULT_TABLE_COLOR).band;
  const out: string[] = [];
  for (let row = 0; row < 4; row += 1) {
    for (let col = 0; col < 4; col += 1) {
      const rect = cellRectIn(view.layout, cols, rows, row, col);
      if (fills.some((f) => f.style === band && f.x === rect.x && f.y === rect.y)) {
        out.push(`${row},${col}`);
      }
    }
  }
  return out;
}

describe('table column stripes', () => {
  it('paints every other column from the first', () => {
    const cells = bandedCells({ colStripes: true });
    expect(cells).toHaveLength(8);
    expect(new Set(cells.map((key) => key.split(',')[1]))).toEqual(new Set(['0', '2']));
  });

  it('paints nothing without the flag', () => {
    expect(bandedCells({ colStripes: false })).toEqual([]);
  });

  it('leaves the header and total rows unstriped', () => {
    const cells = bandedCells({ colStripes: true, showHeader: true, showTotal: true });
    expect(cells.map((key) => key.split(',')[0])).not.toContain('0');
    expect(cells.map((key) => key.split(',')[0])).not.toContain('3');
    expect(cells).toHaveLength(4);
  });

  it('keeps first and last column emphasis independent of the stripes', () => {
    const cells = bandedCells({ colStripes: true, firstCol: true, lastCol: true });
    expect(cells).toHaveLength(8);
  });
});
