import { describe, expect, it } from 'vitest';

import type { CellValue } from '../../../../src/engine/types.js';
import { createConditionalCache } from '../../../../src/render/conditional.js';
import {
  buildColLayout,
  buildRowLayout,
  cellRectIn,
  layoutForView,
} from '../../../../src/render/geometry.js';
import { paintCells } from '../../../../src/render/grid/cells.js';
import {
  type ConditionalRule,
  createSpreadsheetStore,
  mutators,
} from '../../../../src/store/store.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

const conditionalFill = '#conditional';
const theme = {
  bg: '#fff',
  bgElev: '#fff',
  accentSoft: 'rgba(0,120,212,0.1)',
  textCell: 13,
  fontUi: 'sans-serif',
} as ResolvedTheme;

const visibleAddrs = [
  { sheet: 0, row: 6, col: 5 },
  { sheet: 0, row: 6, col: 16_001 },
  { sheet: 0, row: 900_001, col: 5 },
  { sheet: 0, row: 900_001, col: 16_001 },
] as const;

const cellValues: { name: string; rule: ConditionalRule; value: CellValue; formula?: string }[] = [
  {
    name: 'text condition',
    rule: {
      kind: 'text-contains',
      range: { sheet: 0, r0: 6, c0: 5, r1: 900_001, c1: 16_001 },
      text: 'match',
      apply: { fill: conditionalFill },
    },
    value: { kind: 'text', value: 'match' },
  },
  {
    name: 'date condition',
    rule: {
      kind: 'date-occurring',
      range: { sheet: 0, r0: 6, c0: 5, r1: 900_001, c1: 16_001 },
      period: 'today',
      apply: { fill: conditionalFill },
    },
    value: { kind: 'number', value: Math.floor(Date.now() / 86_400_000) + 25569 },
  },
  {
    name: 'blank condition',
    rule: {
      kind: 'blanks',
      range: { sheet: 0, r0: 6, c0: 5, r1: 900_001, c1: 16_001 },
      apply: { fill: conditionalFill },
    },
    value: { kind: 'blank' },
    formula: '=1',
  },
  {
    name: 'cell-value condition',
    rule: {
      kind: 'cell-value',
      range: { sheet: 0, r0: 6, c0: 5, r1: 900_001, c1: 16_001 },
      op: '>',
      a: 1,
      apply: { fill: conditionalFill },
    },
    value: { kind: 'number', value: 5 },
  },
  {
    name: 'formula condition',
    rule: {
      kind: 'formula',
      range: { sheet: 0, r0: 6, c0: 5, r1: 900_001, c1: 16_001 },
      formula: '>1',
      apply: { fill: conditionalFill },
    },
    value: { kind: 'number', value: 5 },
  },
];

function canvasSpy() {
  const fills: { style: string; rect: [number, number, number, number] }[] = [];
  const target: Record<string, unknown> = {
    fillStyle: '',
    fillRect(x: number, y: number, w: number, h: number) {
      fills.push({
        style: String(this.fillStyle),
        rect: [x, y, w, h],
      });
    },
    measureText: () => ({ width: 8, actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2 }),
  };
  const ctx = new Proxy(target, {
    get(object, property: string | symbol) {
      if (property in object) return object[property as string];
      return () => undefined;
    },
    set(object, property: string | symbol, value: unknown) {
      object[property as string] = value;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
  return { ctx, fills };
}

describe('grid conditional-format viewport includes actual visible axes', () => {
  it.each(cellValues)(
    'paints frozen and distant body cells for $name',
    ({ rule, value, formula }) => {
      const store = createSpreadsheetStore();
      mutators.setFreezePanes(store, 2, 2, { row: 5, col: 4 });
      for (const addr of visibleAddrs) mutators.setCell(store, addr, value, formula ?? null);

      store.setState((state) => ({
        ...state,
        layout: {
          ...state.layout,
          hiddenRows: new Set([900_000]),
          hiddenCols: new Set([16_000]),
        },
        viewport: {
          ...state.viewport,
          rowStart: 900_000,
          rowCount: 2,
          colStart: 16_000,
          colCount: 2,
          widthPx: 640,
          navigationRange: {
            sheet: 0,
            r0: 6,
            c0: 5,
            r1: 900_001,
            c1: 16_001,
          },
        },
        selection: {
          ...state.selection,
          active: { sheet: 0, row: 100, col: 100 },
          range: { sheet: 0, r0: 100, c0: 100, r1: 100, c1: 100 },
          anchor: { sheet: 0, row: 100, col: 100 },
        },
        conditional: { ...state.conditional, rules: [rule] },
      }));

      const state = store.getState();
      const view = { ...state, layout: layoutForView(state) };
      const rows = buildRowLayout(view.layout, view.viewport);
      const cols = buildColLayout(view.layout, view.viewport);
      expect(rows.visible).toEqual([6, 900_001]);
      expect(cols.visible).toEqual([5, 16_001]);

      const spy = canvasSpy();
      paintCells(
        {
          ctx: spy.ctx,
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

      const conditionalRects = spy.fills
        .filter(({ style }) => style === conditionalFill)
        .map(({ rect }) => rect);
      const expectedRects = visibleAddrs.map(({ row, col }) => {
        const rect = cellRectIn(view.layout, cols, rows, row, col);
        return [rect.x, rect.y, rect.w, rect.h];
      });
      expect(conditionalRects).toEqual(expectedRects);
    },
  );
});
