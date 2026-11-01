import { describe, expect, it } from 'vitest';

import { createConditionalCache } from '../../../../src/render/conditional.js';
import {
  buildColLayout,
  buildRowLayout,
  cellRectIn,
  layoutForView,
} from '../../../../src/render/geometry.js';
import { paintCells } from '../../../../src/render/grid/cells.js';
import type { CellFormat } from '../../../../src/store/store.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

const theme = {
  bg: '#fff',
  rule: '#ccc',
  accent: '#0078d4',
  accentSoft: 'rgba(0,120,212,0.1)',
  cellErrorFg: '#c00',
  cellBoolFg: '#060',
  cellNumFg: '#000',
  cellAutoFg: '#000',
  fontUi: 'sans-serif',
  fontMono: 'monospace',
  textCell: 13,
} as ResolvedTheme;

type Op =
  | { kind: 'rect'; x: number; w: number }
  | { kind: 'fillRect'; style: string; x: number }
  | { kind: 'text'; text: string; y: number };

/** Canvas stub logging clip rects, fills and text; text is 7px a character. */
function recordingCtx(): { ctx: CanvasRenderingContext2D; ops: Op[] } {
  const ops: Op[] = [];
  const target: Record<string, unknown> = {
    fillStyle: '',
    font: '',
    textAlign: 'left',
    textBaseline: 'alphabetic',
    direction: 'inherit',
    measureText: (text: string) => ({ width: text.length * 7 }) as TextMetrics,
    rect: (x: number, _y: number, w: number) => ops.push({ kind: 'rect', x, w }),
    fillRect: (x: number) => ops.push({ kind: 'fillRect', style: String(target.fillStyle), x }),
    fillText: (text: string, _x: number, y: number) => ops.push({ kind: 'text', text, y }),
  };
  const ctx = new Proxy(target, {
    get: (object, property: string) => (property in object ? object[property] : () => undefined),
    set: (object, property: string, value: unknown) => {
      object[property] = value;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
  return { ctx, ops };
}

function paintRow(
  cells: { col: number; text?: string; format?: CellFormat }[],
  merge?: { r0: number; c0: number; r1: number; c1: number },
) {
  const store = createSpreadsheetStore();
  if (merge) mutators.mergeRange(store, { sheet: 0, ...merge });
  for (const cell of cells) {
    const addr = { sheet: 0, row: 1, col: cell.col };
    if (cell.text !== undefined) mutators.setCell(store, addr, { kind: 'text', value: cell.text });
    if (cell.format) mutators.setCellFormat(store, addr, cell.format);
  }
  store.setState((state) => ({
    ...state,
    viewport: { ...state.viewport, widthPx: 1200, rowCount: 4, colCount: 10 },
  }));
  const state = store.getState();
  const view = { ...state, layout: layoutForView(state) };
  const cols = buildColLayout(view.layout, view.viewport);
  const rows = buildRowLayout(view.layout, view.viewport);
  const { ctx, ops } = recordingCtx();
  const result = paintCells(
    {
      ctx,
      wb: null,
      locale: 'en-US',
      getDisplay: undefined,
      sheetBackgroundImage: null,
      fontEffects: false,
      conditionalCache: createConditionalCache(),
    },
    view,
    theme,
    cols,
    rows,
  );
  const rect = (col: number) => cellRectIn(view.layout, cols, rows, 1, col);
  return { ops, result, rect };
}

const longText = 'x'.repeat(40);

describe('text overflow', () => {
  it('spills right-aligned text leftward until a non-empty cell', () => {
    const { ops, result, rect } = paintRow([
      { col: 1, text: 'y', format: { align: 'right' } },
      { col: 5, text: longText, format: { align: 'right' } },
    ]);
    const right = rect(5).x + rect(5).w;
    const clip = ops.find((op) => op.kind === 'rect' && op.x < rect(5).x && op.x + op.w === right);
    expect(clip).toMatchObject({ x: rect(2).x });
    expect(result.coveredRules.has('1:4')).toBe(true);
    expect(result.coveredRules.has('1:1')).toBe(false);
  });

  it('paints overflow over a neighbour fill and border, after the fill', () => {
    const { ops, result } = paintRow([
      { col: 1, text: longText },
      { col: 2, format: { fill: '#ffff00' } },
      { col: 3, format: { borders: { left: { style: 'thin' } } } },
    ]);
    const fill = ops.findIndex((op) => op.kind === 'fillRect' && op.style === '#ffff00');
    const text = ops.findIndex((op) => op.kind === 'text' && op.text === longText);
    expect(fill).toBeGreaterThanOrEqual(0);
    expect(text).toBeGreaterThan(fill);
    expect(result.coveredRules.has('1:1')).toBe(true);
    expect(result.coveredRules.has('1:2')).toBe(true);
  });

  it('keeps rules the run does not reach and clips centred text at a blocked side', () => {
    const { ops, result, rect } = paintRow([
      { col: 3, text: 'z' },
      { col: 4, text: longText, format: { align: 'center' } },
    ]);
    const clip = ops.find((op) => op.kind === 'rect' && op.w > rect(4).w);
    expect(clip).toMatchObject({ x: rect(4).x });
    expect(result.coveredRules.has('1:3')).toBe(false);
    expect(result.coveredRules.has('1:4')).toBe(true);
  });
});

describe('merged text placement', () => {
  it('centres a merge one pixel higher than a single cell', () => {
    const middle: CellFormat = { vAlign: 'middle' };
    const single = paintRow([{ col: 1, text: 'Hxg', format: middle }]);
    const merged = paintRow([{ col: 1, text: 'Hxg', format: middle }], {
      r0: 1,
      c0: 1,
      r1: 2,
      c1: 1,
    });
    const y = (ops: Op[]) => ops.find((op) => op.kind === 'text' && op.text === 'Hxg');
    const rowH = single.rect(1).h;
    expect(y(merged.ops)).toMatchObject({
      y: (y(single.ops) as { y: number }).y + rowH / 2 - 1,
    });
  });
});
