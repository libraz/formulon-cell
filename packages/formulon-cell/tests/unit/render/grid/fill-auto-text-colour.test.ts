import { describe, expect, it } from 'vitest';

import { defaultTableOverlay } from '../../../../src/commands/format-as-table.js';
import { createPivotTableFromRange } from '../../../../src/commands/pivot-table.js';
import { pivotStyleLayer } from '../../../../src/engine/pivot-style.js';
import { cssColorFromArgb } from '../../../../src/engine/theme-color.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createConditionalCache } from '../../../../src/render/conditional.js';
import { buildColLayout, buildRowLayout, layoutForView } from '../../../../src/render/geometry.js';
import { paintCells } from '../../../../src/render/grid/cells.js';
import { automaticTextColor } from '../../../../src/render/painters/text.js';
import {
  type CellFormat,
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import { type ResolvedTheme, WORKBOOK_TEXT_COLOR } from '../../../../src/theme/resolve.js';

const base = {
  rule: '#3b3a39',
  accent: '#4cc26b',
  accentSoft: 'rgba(76,194,107,0.18)',
  cellErrorFg: '#f1707b',
  cellBoolFg: '#6dd587',
  fontUi: 'sans-serif',
  fontMono: 'monospace',
  textCell: 13,
};

/** A dark sheet: light automatic text, fill text taken from the workbook. */
const dark = {
  ...base,
  bg: '#1b1a19',
  cellNumFg: '#ffffff',
  cellAutoFg: '#f3f2f1',
  cellFillAutoFg: WORKBOOK_TEXT_COLOR,
} as ResolvedTheme;

/** A light sheet: fill text is the automatic text colour. */
const light = {
  ...base,
  bg: '#ffffff',
  cellNumFg: '#0a0c10',
  cellAutoFg: '#0a0c10',
  cellFillAutoFg: '#0a0c10',
} as ResolvedTheme;

/** Canvas stub recording the fill style each text run is drawn with. */
function textSpy(): { ctx: CanvasRenderingContext2D; texts: Map<string, string> } {
  const texts = new Map<string, string>();
  const target: Record<string, unknown> = {
    fillStyle: '',
    font: '',
    textAlign: 'left',
    textBaseline: 'alphabetic',
    direction: 'inherit',
    measureText: (text: string) => ({ width: text.length * 7 }) as TextMetrics,
    fillText: (text: string) => texts.set(text, String(target.fillStyle)),
  };
  const ctx = new Proxy(target, {
    get: (object, property: string) =>
      property in object ? object[property] : () => ({ addColorStop: () => undefined }),
    set: (object, property: string, value: unknown) => {
      object[property] = value;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
  return { ctx, texts };
}

function paint(
  store: SpreadsheetStore,
  theme: ResolvedTheme,
  wb: WorkbookHandle | null = null,
): Map<string, string> {
  store.setState((state) => ({
    ...state,
    viewport: { ...state.viewport, widthPx: 800, rowCount: 12, colCount: 6 },
  }));
  const state = store.getState();
  const view = { ...state, layout: layoutForView(state) };
  const { ctx, texts } = textSpy();
  paintCells(
    {
      ctx,
      wb,
      locale: 'en-US',
      getDisplay: undefined,
      sheetBackgroundImage: null,
      fontEffects: false,
      conditionalCache: createConditionalCache(),
    },
    view,
    theme,
    buildColLayout(view.layout, view.viewport),
    buildRowLayout(view.layout, view.viewport),
  );
  return texts;
}

/** One text cell at B2 with an optional cell format. */
function cellStore(text: string, format?: CellFormat): SpreadsheetStore {
  const store = createSpreadsheetStore();
  const addr = { sheet: 0, row: 1, col: 1 };
  mutators.setCell(store, addr, { kind: 'text', value: text });
  if (format) mutators.setCellFormat(store, addr, format);
  return store;
}

/** B2 holding `text` under a rule that fills any cell containing it. */
function conditionalStore(text: string): SpreadsheetStore {
  const store = cellStore(text);
  store.setState((state) => ({
    ...state,
    conditional: {
      ...state.conditional,
      rules: [
        {
          kind: 'text-contains',
          range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
          text,
          apply: { fill: '#ffc7ce' },
        },
      ],
    },
  }));
  return store;
}

/** A Medium 2 table at A1:B5 with `text` in its first stripe (A2) and `plain` in the second (A3). */
function tableStore(text: string, plain: string): SpreadsheetStore {
  const store = createSpreadsheetStore();
  mutators.upsertTableOverlay(
    store,
    defaultTableOverlay('t1', { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 1 }),
  );
  mutators.setCell(store, { sheet: 0, row: 1, col: 0 }, { kind: 'text', value: text });
  mutators.setCell(store, { sheet: 0, row: 2, col: 0 }, { kind: 'text', value: plain });
  return store;
}

/** A report built from a two-row source, with `text` written into a cell its style fills without a font colour. */
async function withPivotFill(
  run: (wb: WorkbookHandle, store: SpreadsheetStore, text: string) => void,
): Promise<void> {
  const wb = await WorkbookHandle.createDefault();
  try {
    wb.setText({ sheet: 0, row: 0, col: 0 }, 'Region');
    wb.setText({ sheet: 0, row: 0, col: 1 }, 'Amt');
    wb.setText({ sheet: 0, row: 1, col: 0 }, 'East');
    wb.setNumber({ sheet: 0, row: 1, col: 1 }, 120);
    const created = createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
      destination: { sheet: 0, row: 4, col: 0 },
      rowField: 'Region',
      valueField: 'Amt',
    });
    expect(created.ok).toBe(true);
    const store = createSpreadsheetStore();
    const layer = pivotStyleLayer(wb, store.getState());
    let target: { row: number; col: number } | undefined;
    for (let row = 4; row < 10 && !target; row += 1) {
      for (let col = 0; col < 3 && !target; col += 1) {
        const fmt = layer.formatAt(row, col);
        if (fmt?.fill && !fmt.color) target = { row, col };
      }
    }
    expect(target).toBeDefined();
    const addr = { sheet: 0, ...(target as { row: number; col: number }) };
    const text = wb.getDisplayText(0, addr.row, addr.col) || 'pivot';
    mutators.setCell(store, addr, { kind: 'text', value: text });
    run(wb, store, text);
  } finally {
    wb.dispose();
  }
}

describe('automatic text on a painted fill', () => {
  describe('under a dark theme it takes the workbook Text 1', () => {
    it('on a plain cell fill', () => {
      expect(paint(cellStore('East', { fill: '#dce6f1' }), dark).get('East')).toBe('#000000');
    });

    it('on a pattern fill', () => {
      const store = cellStore('East', { fillPattern: 'gray125', fillPatternColor: '#808080' });
      expect(paint(store, dark).get('East')).toBe('#000000');
    });

    it('for a number on a cell fill', () => {
      const store = createSpreadsheetStore();
      mutators.setCell(store, { sheet: 0, row: 1, col: 1 }, { kind: 'number', value: 120 });
      mutators.setCellFormat(store, { sheet: 0, row: 1, col: 1 }, { fill: '#dce6f1' });
      expect(paint(store, dark).get('120')).toBe('#000000');
    });

    it('on a conditional-format fill', () => {
      expect(paint(conditionalStore('late'), dark).get('late')).toBe('#000000');
    });

    it('on a table style stripe', () => {
      expect(paint(tableStore('East', 'West'), dark).get('East')).toBe('#000000');
    });

    it('on a PivotTable style fill, from the workbook theme', async () => {
      await withPivotFill((wb, store, text) => {
        const dk1 = cssColorFromArgb(wb.getTheme()?.colors[0] ?? 0xff000000);
        expect(paint(store, dark, wb).get(text)).toBe(dk1);
      });
    });
  });

  describe('unchanged', () => {
    it('keeps the automatic colour without a fill', () => {
      expect(paint(cellStore('East'), dark).get('East')).toBe(dark.cellAutoFg);
      expect(paint(tableStore('East', 'West'), dark).get('West')).toBe(dark.cellAutoFg);
    });

    it('keeps a stated colour on a fill', () => {
      const store = cellStore('East', { fill: '#dce6f1', color: '#ff0000' });
      expect(paint(store, dark).get('East')).toBe('#ff0000');
    });

    it('keeps a stated black without a fill', () => {
      expect(paint(cellStore('East', { color: '#000000' }), dark).get('East')).toBe('#000000');
    });

    it('keeps every fill source on the automatic colour under a light theme', async () => {
      expect(paint(cellStore('East', { fill: '#dce6f1' }), light).get('East')).toBe(
        light.cellAutoFg,
      );
      expect(paint(conditionalStore('late'), light).get('late')).toBe(light.cellAutoFg);
      expect(paint(tableStore('East', 'West'), light).get('East')).toBe(light.cellAutoFg);
      await withPivotFill((wb, store, text) => {
        expect(paint(store, light, wb).get(text)).toBe(light.cellAutoFg);
      });
    });
  });

  describe('automaticTextColor', () => {
    it('defers to the workbook Text 1 the painter is given', () => {
      expect(automaticTextColor(dark, { fill: '#dce6f1' }, false, '#1f497d')).toBe('#1f497d');
      expect(automaticTextColor(dark, undefined, false, '#1f497d')).toBe(dark.cellAutoFg);
      expect(automaticTextColor(light, { fill: '#dce6f1' }, false, '#1f497d')).toBe(
        light.cellAutoFg,
      );
    });
  });
});
