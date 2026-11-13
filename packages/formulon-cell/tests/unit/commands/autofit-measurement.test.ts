import { describe, expect, it } from 'vitest';
import {
  autofitColWidth,
  autofitRowHeight,
  computeAutofitColWidth,
  computeAutofitRowHeight,
  FILTER_DROPDOWN_RESERVED_WIDTH,
  TABLE_HEADER_RESERVED_WIDTH,
} from '../../../src/commands/autofit-measurement.js';
import { formatAsTable } from '../../../src/commands/format-as-table.js';
import { History } from '../../../src/commands/history.js';
import {
  autofitColsWidth,
  autofitColsWidthForRanges,
  autofitRowsHeight,
  autofitRowsHeightForRanges,
} from '../../../src/commands/row-col-layout.js';
import type { CellValue } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import type { SpreadsheetInstance } from '../../../src/mount/types.js';
import { gridOriginX, gridOriginY, hitZone, layoutForView } from '../../../src/render/geometry.js';
import { filterButtonRect } from '../../../src/render/painters/filter-button.js';
import { paintTableHeaderChevron } from '../../../src/render/painters/markers.js';
import { CELL_TEXT_INDENT_PX, cellTextHorizontalMetrics } from '../../../src/render/text-wrap.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';
import type { ResolvedTheme } from '../../../src/theme/resolve.js';
import {
  applyCellFormatAction,
  type CellFormatActionDeps,
} from '../../../src/toolbar/ribbon/cell-format-action.js';

/** The layout mutator's row ceiling: 409 pt in canonical px. */
const MAX_ROW_HEIGHT_PX = (409.5 * 96) / 72;

type Seed = [row: number, col: number, value: string | number, formula?: string];

const storeWith = (cells: Seed[]): SpreadsheetStore => {
  const store = createSpreadsheetStore();
  const map = new Map<string, { value: CellValue; formula: string | null }>();
  for (const [row, col, value, formula] of cells) {
    map.set(`0:${row}:${col}`, {
      value: typeof value === 'number' ? { kind: 'number', value } : { kind: 'text', value },
      formula: formula ?? null,
    });
  }
  store.setState((s) => ({ ...s, data: { ...s.data, cells: map } }));
  return store;
};

const withFilter = (store: SpreadsheetStore): void =>
  store.setState((s) => ({
    ...s,
    ui: { ...s.ui, filterRanges: [{ sheet: 0, r0: 0, c0: 0, r1: 5, c1: 0 }] },
  }));

/** Deterministic 2d context: 7px per character, recording what was measured. */
const fakeMeasureCtx = (): { ctx: CanvasRenderingContext2D; measured: string[] } => {
  const measured: string[] = [];
  const ctx = {
    font: '',
    measureText(text: string) {
      measured.push(text);
      return { width: text.length * 7 } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
  return { ctx, measured };
};

// The headless DOM has no canvas, so these run on the 0.54em-per-char estimate:
// 'HeaderName' at 13px measures 70.2px.
describe('computeAutofitColWidth', () => {
  it('fits content plus padding', () => {
    expect(computeAutofitColWidth(storeWith([[0, 0, 'HeaderName']]).getState(), 0, null)).toBe(87);
  });

  it('measures the widest explicit line and leaves clamping to the store', () => {
    expect(computeAutofitColWidth(storeWith([[0, 0, 'x'.repeat(100)]]).getState(), 0, null)).toBe(
      718,
    );
    const store = storeWith([[0, 0, 'x'.repeat(200)]]);
    autofitColsWidth(store, null, 0, 0);
    expect(store.getState().layout.colWidths.get(0)).toBe(800);
  });

  it('reserves the filter dropdown footprint on filter header cells only', () => {
    const store = storeWith([
      [0, 0, 'HeaderName'],
      [1, 0, 'HeaderName'],
    ]);
    const plain = computeAutofitColWidth(store.getState(), 0, null);
    withFilter(store);
    expect(computeAutofitColWidth(store.getState(), 0, null)).toBe(105);
    expect(computeAutofitColWidth(store.getState(), 0, null) - plain).toBe(
      FILTER_DROPDOWN_RESERVED_WIDTH,
    );
  });

  it('reserves the table header chevron on table header cells only', () => {
    const store = storeWith([
      [0, 0, 'HeaderName'],
      [1, 0, 'HeaderName'],
    ]);
    const plain = computeAutofitColWidth(store.getState(), 0, null);
    formatAsTable(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 }, { showHeader: false });
    expect(computeAutofitColWidth(store.getState(), 0, null)).toBe(plain);
    formatAsTable(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 }, { showHeader: true });
    expect(computeAutofitColWidth(store.getState(), 0, null) - plain).toBe(
      TABLE_HEADER_RESERVED_WIDTH,
    );
    withFilter(store);
    expect(computeAutofitColWidth(store.getState(), 0, null) - plain).toBe(
      Math.max(TABLE_HEADER_RESERVED_WIDTH, FILTER_DROPDOWN_RESERVED_WIDTH),
    );
  });

  it('measures in the theme font, monospace while formulas are shown', () => {
    const store = storeWith([[0, 0, 'v', '=A2']]);
    const theme = { textCell: 15, fontUi: 'Body Font, sans-serif', fontMono: 'Code Mono' };
    const { ctx } = fakeMeasureCtx();
    computeAutofitColWidth(store.getState(), 0, ctx, { theme });
    expect(ctx.font).toBe('400 15px "Body Font", sans-serif');
    store.setState((s) => ({ ...s, ui: { ...s.ui, showFormulas: true } }));
    computeAutofitColWidth(store.getState(), 0, ctx, { theme });
    expect(ctx.font).toBe('400 15px "Code Mono"');
    expect(
      computeAutofitRowHeight(store.getState(), 0, null, { theme: { ...theme, textCell: 30 } }),
    ).toBe(38);
  });

  it('measures a missing cell font with the workbook Normal font in CSS pixels', () => {
    const store = storeWith([[0, 0, 'HeaderName']]);
    const { ctx } = fakeMeasureCtx();
    computeAutofitColWidth(store.getState(), 0, ctx, {
      workbookDisplayFont: { name: 'Aptos Narrow', size: 12 },
    });
    expect(ctx.font).toBe('400 16px "Aptos Narrow", system-ui, sans-serif');
  });

  it('measures formula text when formulas are shown', () => {
    const store = storeWith([[0, 0, 'v', '=REPT("abcdefgh",3)']]);
    expect(computeAutofitColWidth(store.getState(), 0, null)).toBe(24);
    store.setState((s) => ({ ...s, ui: { ...s.ui, showFormulas: true } }));
    expect(computeAutofitColWidth(store.getState(), 0, null)).toBe(150);
  });

  it('restricts the scan to a row span when one is given', () => {
    const state = storeWith([[5, 0, 'x'.repeat(40)]]).getState();
    expect(computeAutofitColWidth(state, 0, null)).toBe(297);
    expect(computeAutofitColWidth(state, 0, null, { span: { from: 0, to: 2 } })).toBe(
      state.layout.defaultColWidth,
    );
  });

  it('formats numbers with the requested locale and measures with the cell font', () => {
    const store = storeWith([[0, 0, 1234567.5]]);
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true, fontSize: 16 });
    const { ctx, measured } = fakeMeasureCtx();
    computeAutofitColWidth(store.getState(), 0, ctx, { locale: 'de-DE' });
    expect(measured).toEqual(['1234567,5']);
    expect(ctx.font).toBe('700 16px system-ui, sans-serif');
    measured.length = 0;
    computeAutofitColWidth(store.getState(), 0, ctx);
    expect(measured).toEqual(['1234567.5']);
  });
});

describe('computeAutofitColWidth, Excel exclusions', () => {
  // Excel: a wrapped cell does not widen its column, a column holding only a
  // merged cell fits to its other cells, and an empty column takes the default
  // width (8.0 chars) rather than a floor.
  it('skips wrapped cells', () => {
    const store = storeWith([
      [0, 0, 'x'.repeat(40)],
      [1, 0, 'HeaderName'],
    ]);
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { wrap: true });
    expect(computeAutofitColWidth(store.getState(), 0, null)).toBe(87);
  });

  it('skips merged cells, anchor included', () => {
    const store = storeWith([
      [0, 2, 'merged wide text goes here'],
      [1, 2, 'x'],
    ]);
    mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 3 });
    // 'x' at 13 px: 7.02 px of text plus the 16 px padding.
    expect(computeAutofitColWidth(store.getState(), 2, null)).toBe(24);
  });

  it('fits an empty column to the default width and has no floor', () => {
    const empty = storeWith([]);
    expect(computeAutofitColWidth(empty.getState(), 0, null)).toBe(
      empty.getState().layout.defaultColWidth,
    );
    expect(computeAutofitColWidth(storeWith([[0, 0, 'a']]).getState(), 0, null)).toBe(24);
  });
});

describe('computeAutofitRowHeight on the Mac point model', () => {
  // Excel for Mac (default 14 pt rows): one Arial 11 line fits the 14 pt
  // tabulated row, Arial 28 takes 35 pt, and a wrapped Arial 11 cell takes
  // 15 pt per line: 30 pt for two lines, 45 pt for three.
  const pt = (points: number) => (points * 4) / 3;
  const macStore = (cells: Seed[]) => {
    const store = storeWith(cells);
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, screenScale: 0.75, defaultRowHeight: pt(14) },
    }));
    return store;
  };
  const arial = (store: SpreadsheetStore, row: number, points: number, wrap = false) =>
    mutators.setCellFormat(
      store,
      { sheet: 0, row, col: 0 },
      { fontFamily: 'Arial', fontSize: pt(points), ...(wrap ? { wrap: true } : {}) },
    );

  it('sizes a single line to the tabulated row of its font', () => {
    const store = macStore([
      [0, 0, 'plain'],
      [1, 0, 'Big'],
    ]);
    arial(store, 0, 11);
    arial(store, 1, 28);
    expect(computeAutofitRowHeight(store.getState(), 0, null)).toBeCloseTo(pt(14), 6);
    expect(computeAutofitRowHeight(store.getState(), 1, null)).toBeCloseTo(pt(35), 6);
  });

  it('floors only rows of wrapped text below the default row', () => {
    const store = macStore([
      [0, 0, 'small'],
      [1, 0, 'small'],
    ]);
    arial(store, 0, 8);
    arial(store, 1, 8, true);
    expect(computeAutofitRowHeight(store.getState(), 0, null)).toBeCloseTo(pt(14), 6);
    expect(computeAutofitRowHeight(store.getState(), 1, null)).toBeCloseTo(pt(12), 6);
  });

  it('gives each wrapped Arial 11 line 15 pt', () => {
    const store = macStore([
      [0, 0, 'line1\nline2'],
      [1, 0, 'line1\nline2\nline3'],
    ]);
    arial(store, 0, 11, true);
    arial(store, 1, 11, true);
    expect(computeAutofitRowHeight(store.getState(), 0, null)).toBeCloseTo(pt(30), 6);
    expect(computeAutofitRowHeight(store.getState(), 1, null)).toBeCloseTo(pt(45), 6);
  });

  it('word-wraps Arial 11 at width 16 to the line counts the desktop picks', () => {
    // Arial advance widths per 1000 em for the letters and space in the sample text.
    const advance: Record<string, number> = {
      ' ': 278,
      a: 556,
      b: 556,
      c: 500,
      d: 556,
      e: 556,
      f: 278,
      g: 556,
      h: 556,
      i: 222,
      j: 222,
      k: 500,
      l: 222,
      m: 833,
      n: 556,
      o: 556,
      p: 556,
      q: 556,
      r: 333,
      s: 500,
      t: 278,
      u: 556,
      v: 500,
      w: 722,
      x: 500,
      y: 500,
      z: 500,
    };
    const ctx = {
      font: '',
      measureText(this: { font: string }, text: string) {
        const size = Number(this.font.match(/([\d.]+)px/)?.[1] ?? 0);
        const em = [...text].reduce((sum, ch) => sum + (advance[ch] ?? 556), 0);
        return { width: (em / 1000) * size } as TextMetrics;
      },
    } as unknown as CanvasRenderingContext2D;
    const sample = 'abcd efgh ijkl mnop qrst uvwx yzab cdef ghij klmn opqr stuv';
    const lengths = [8, 20, 33, 47];
    const store = macStore(lengths.map((n, row): Seed => [row, 0, sample.slice(0, n)]));
    // Width 16 characters of the 7 px Normal digit plus the 5 px cell margin.
    mutators.setColWidth(store, 0, 16 * 7 + 5);
    lengths.forEach((_, row) => {
      arial(store, row, 11, true);
    });
    const lines = lengths.map(
      (_, row) => computeAutofitRowHeight(store.getState(), row, ctx) / pt(15),
    );
    expect(lines.map((n) => Math.round(n * 1e6) / 1e6)).toEqual([1, 2, 3, 3]);
  });
});

describe('computeAutofitRowHeight, rotated text', () => {
  // Excel: 'rot' at 90 degrees in a 14 pt row autofits the row to 19 pt — the
  // run's length stands up in the row.
  it('grows the row by the run projected onto the vertical', () => {
    const { ctx } = fakeMeasureCtx();
    const store = storeWith([[0, 0, 'rotated text']]);
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { rotation: 90 });
    // 12 chars x 7 px = 84 px standing up, plus the row's padding.
    const padding = 20 - Math.round(13 * 1.28);
    expect(computeAutofitRowHeight(store.getState(), 0, ctx)).toBe(84 + padding);
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { rotation: 45 });
    const line = Math.round(13 * 1.28);
    expect(computeAutofitRowHeight(store.getState(), 0, ctx)).toBe(
      Math.ceil(84 * Math.SQRT1_2 + line * Math.SQRT1_2) + padding,
    );
  });

  it('skips merged cells, as Excel AutoFit does on either axis', () => {
    const store = storeWith([[0, 0, 'a\nb\nc']]);
    mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    expect(computeAutofitRowHeight(store.getState(), 0, null)).toBe(20);
  });

  it('stacks vertical text one character per line', () => {
    const { ctx } = fakeMeasureCtx();
    const store = storeWith([[0, 0, 'abcd']]);
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { rotation: 'vertical' });
    const line = Math.round(13 * 1.28);
    expect(computeAutofitRowHeight(store.getState(), 0, ctx)).toBe(4 * line + (20 - line));
  });
});

describe('computeAutofitRowHeight', () => {
  it('fits one Normal-font line to the default height and grows by a pitch per extra line', () => {
    expect(computeAutofitRowHeight(storeWith([]).getState(), 0, null)).toBe(20);
    expect(computeAutofitRowHeight(storeWith([[0, 0, 'HeaderName']]).getState(), 0, null)).toBe(20);
    expect(computeAutofitRowHeight(storeWith([[0, 0, 'a\nb\nc']]).getState(), 0, null)).toBe(54);

    const wrapped = storeWith([[0, 0, 'alpha beta gamma delta epsilon zeta']]);
    mutators.setColWidth(wrapped, 0, 70);
    mutators.setCellFormat(wrapped, { sheet: 0, row: 0, col: 0 }, { wrap: true });
    expect(computeAutofitRowHeight(wrapped.getState(), 0, null)).toBe(105);
  });

  it('restricts the scan to a column span and leaves clamping to the store', () => {
    const state = storeWith([[0, 5, 'y\ny\ny\ny']]).getState();
    expect(computeAutofitRowHeight(state, 0, null)).toBe(71);
    expect(computeAutofitRowHeight(state, 0, null, { span: { from: 0, to: 2 } })).toBe(20);

    const tall = storeWith([[0, 0, 'a\n'.repeat(40)]]);
    expect(computeAutofitRowHeight(tall.getState(), 0, null)).toBe(700);
    autofitRowsHeight(tall, null, 0, 0);
    expect(tall.getState().layout.rowHeights.get(0)).toBe(MAX_ROW_HEIGHT_PX);
  });

  it('uses the renderer padding and indent width contract for wrapped AutoFit', () => {
    const wrapped = storeWith([[0, 0, 'abcdefghi']]);
    mutators.setColWidth(wrapped, 0, 76);
    mutators.setCellFormat(wrapped, { sheet: 0, row: 0, col: 0 }, { wrap: true, indent: 1 });
    const { ctx } = fakeMeasureCtx();

    // 63px of text fits the old 64px (width - 12) calculation, but not the
    // shared renderer contract: 76 - 6px padding - one indent level.
    expect(cellTextHorizontalMetrics(76, 1).availableWidth).toBe(70 - CELL_TEXT_INDENT_PX);
    expect(computeAutofitRowHeight(wrapped.getState(), 0, ctx)).toBe(37);
  });
});

describe('autofit on the Mac point model', () => {
  /** Arial 'A'/'V' advance (1366 / 2048 em) with the pair kern a canvas run applies. */
  const arialAvCtx = (): CanvasRenderingContext2D =>
    ({
      font: '',
      measureText(this: { font: string }, text: string) {
        const size = Number(this.font.match(/([\d.]+)px/)?.[1] ?? 0);
        const kerns = Math.max(0, text.length - 1);
        return { width: ((text.length * 1366 - kerns * 152) / 2048) * size } as TextMetrics;
      },
    }) as unknown as CanvasRenderingContext2D;

  it('fits the run laid out on whole-point advances at the screen size', () => {
    const store = storeWith([[0, 0, 'AVAVAVAV']]);
    // Arial 11 pt: 8 advances of round(7.34) = 7 pt = 56 pt, stored as 74.67 px.
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { fontSize: (11 * 4) / 3 });
    const fractional = computeAutofitColWidth(store.getState(), 0, arialAvCtx());
    store.setState((s) => ({ ...s, layout: { ...s.layout, screenScale: 0.75 } }));
    expect(computeAutofitColWidth(store.getState(), 0, arialAvCtx())).toBe(75 + 16);
    expect(fractional).toBe(Math.ceil(((8 * 1366 - 7 * 152) / 2048) * ((11 * 4) / 3)) + 16);
  });
});

describe('public autofitColWidth / autofitRowHeight', () => {
  it('measure the given span through the shared implementation', () => {
    const store = storeWith([
      [0, 0, 'HeaderName'],
      [5, 0, 'x'.repeat(40)],
    ]);
    withFilter(store);
    const host = document.createElement('div');
    expect(autofitColWidth({ store, host }, 0, 0, 2, 'en')).toBe(105);
    expect(autofitColWidth({ store, host }, 0, 0, 10, 'en')).toBe(297);
    expect(autofitRowHeight({ store, host }, 0, 0, 0, 'ja')).toBe(20);
  });

  it('measure in the font the host theme sets', () => {
    const store = storeWith([[0, 0, 'HeaderName']]);
    const host = document.createElement('div');
    host.style.cssText = '--fc-text-cell: 26px;';
    document.body.appendChild(host);
    expect(autofitColWidth({ store, host }, 0, 0, 0, 'en')).toBe(157);
    host.remove();
  });

  it('uses the optional workbook display font for public helper calls', () => {
    const store = storeWith([[0, 0, 'HeaderName']]);
    const host = document.createElement('div');
    const workbook = { workbookDisplayFont: { name: 'Aptos Narrow', size: 12 } };
    expect(autofitColWidth({ store, host, workbook }, 0, 0, 0, 'en')).toBe(103);
  });
});

describe('ribbon Format > AutoFit', () => {
  const run = async (
    action: 'col-autofit' | 'row-autofit',
    store: SpreadsheetStore,
    range: { r0: number; c0: number; r1: number; c1: number },
    extraRanges: Array<{ sheet: number; r0: number; c0: number; r1: number; c1: number }> = [],
  ): Promise<{ wb: WorkbookHandle; history: History }> => {
    // The engine's profile book, whose Normal font the measurement reads; the stub states none.
    const wb = await WorkbookHandle.createDefault({ locale: 'en' });
    const history = new History();
    const host = document.createElement('div');
    store.setState((s) => ({
      ...s,
      selection: {
        ...s.selection,
        range: { sheet: 0, ...range },
        extraRanges: extraRanges.map((selected) => ({ ...selected })),
      },
    }));
    const inst = { store, history, workbook: wb, host } as unknown as SpreadsheetInstance;
    await applyCellFormatAction(action, {
      inst,
      range: { sheet: 0, ...range },
      ribbonLang: 'en',
      ribbonMenuText: {},
      sheetTabColorByAction: () => undefined,
    } as unknown as CellFormatActionDeps);
    return { wb, history };
  };

  it('fits columns to the selected rows with the filter reserve', async () => {
    const store = storeWith([
      [0, 0, 'a'],
      [0, 1, 'HeaderName'],
      [9, 1, 'x'.repeat(60)],
    ]);
    store.setState((s) => ({
      ...s,
      ui: { ...s.ui, filterRanges: [{ sheet: 0, r0: 0, c0: 0, r1: 5, c1: 1 }] },
    }));
    const { history } = await run('col-autofit', store, { r0: 0, c0: 0, r1: 2, c1: 1 });
    // 'a' (7.02 px) plus the 18 px filter reserve and the 16 px padding.
    expect(store.getState().layout.colWidths.get(0)).toBe(42);
    expect(store.getState().layout.colWidths.get(1)).toBe(114);
    expect(history.undo()).toBe(true);
    expect(store.getState().layout.colWidths.size).toBe(0);
  });

  it('fits rows to the selected columns with store-clamped heights', async () => {
    const store = storeWith([
      [0, 0, 'a\n'.repeat(40)],
      [1, 5, 'y\ny\ny\ny'],
    ]);
    await run('row-autofit', store, { r0: 0, c0: 0, r1: 1, c1: 2 });
    expect(store.getState().layout.rowHeights.get(0)).toBe(MAX_ROW_HEIGHT_PX);
    expect(store.getState().layout.rowHeights.get(1)).toBe(20);
  });

  it('fits every disjoint row and leaves the unselected row unchanged', async () => {
    const store = storeWith([
      [0, 0, 'a\n'.repeat(20)],
      [0, 1, 'this cell is outside the first selected span'],
      [1, 0, 'gap row'],
      [2, 2, 'b\n'.repeat(30)],
      [2, 0, 'this cell is outside the second selected span'],
    ]);
    const workbookDisplayFont = { name: 'Calibri', size: 11 } as const;
    // Measurement is intentionally unclamped; the layout mutator's 409 pt
    // ceiling is what the store keeps.
    const expectedFirst = Math.min(
      MAX_ROW_HEIGHT_PX,
      computeAutofitRowHeight(store.getState(), 0, null, {
        span: { from: 0, to: 0 },
        workbookDisplayFont,
      }),
    );
    const expectedSecond = Math.min(
      MAX_ROW_HEIGHT_PX,
      computeAutofitRowHeight(store.getState(), 2, null, {
        span: { from: 2, to: 2 },
        workbookDisplayFont,
      }),
    );

    await run('row-autofit', store, { r0: 0, c0: 0, r1: 0, c1: 0 }, [
      { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
    ]);

    expect(store.getState().layout.rowHeights.get(0)).toBe(expectedFirst);
    expect(store.getState().layout.rowHeights.get(2)).toBe(expectedSecond);
    expect(store.getState().layout.rowHeights.has(1)).toBe(false);
  });

  it('takes the maximum for an overlapping row without merging its spans', async () => {
    const store = storeWith([
      [0, 0, 'short'],
      [0, 1, 'unselected gap\n'.repeat(40)],
      [0, 2, 'also short'],
    ]);
    const workbookDisplayFont = { name: 'Calibri', size: 11 } as const;
    const expected = Math.max(
      computeAutofitRowHeight(store.getState(), 0, null, {
        span: { from: 0, to: 0 },
        workbookDisplayFont,
      }),
      computeAutofitRowHeight(store.getState(), 0, null, {
        span: { from: 2, to: 2 },
        workbookDisplayFont,
      }),
    );

    await run('row-autofit', store, { r0: 0, c0: 0, r1: 0, c1: 0 }, [
      { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
    ]);

    expect(store.getState().layout.rowHeights.get(0)).toBe(expected);
  });

  it('dedupes duplicate ranges and gives the same max after permutation', () => {
    const ranges = [
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 2 },
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
    ];
    const cells: Seed[] = [
      [0, 0, 'short'],
      [0, 1, 'tall\n'.repeat(8)],
      [0, 2, 'also short'],
    ];
    const first = storeWith(cells);
    const second = storeWith(cells);
    autofitRowsHeightForRanges(first, new History(), ranges);
    autofitRowsHeightForRanges(second, new History(), [...ranges].reverse());

    expect(first.getState().layout.rowHeights.get(0)).toBeGreaterThan(20);
    expect(first.getState().layout.rowHeights).toEqual(second.getState().layout.rowHeights);
  });

  it('fits disjoint columns with their own row spans and one Undo restores all', async () => {
    const store = storeWith([
      [0, 0, 'first selected column has a long value'],
      [1, 1, 'gap column'],
      [2, 2, 'second selected column has another long value'],
      [4, 0, 'outside first column span '.repeat(10)],
      [0, 2, 'outside second column span '.repeat(10)],
    ]);
    const workbookDisplayFont = { name: 'Calibri', size: 11 } as const;
    const expectedFirst = computeAutofitColWidth(store.getState(), 0, null, {
      span: { from: 0, to: 0 },
      workbookDisplayFont,
    });
    const expectedSecond = computeAutofitColWidth(store.getState(), 2, null, {
      span: { from: 2, to: 2 },
      workbookDisplayFont,
    });

    const { history } = await run('col-autofit', store, { r0: 0, c0: 0, r1: 0, c1: 0 }, [
      { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
    ]);

    expect(store.getState().layout.colWidths.get(0)).toBe(expectedFirst);
    expect(store.getState().layout.colWidths.get(2)).toBe(expectedSecond);
    expect(store.getState().layout.colWidths.has(1)).toBe(false);
    expect(history.undo()).toBe(true);
    expect(store.getState().layout.colWidths.size).toBe(0);
    expect(history.redo()).toBe(true);
    expect(store.getState().layout.colWidths.get(0)).toBe(expectedFirst);
    expect(store.getState().layout.colWidths.get(2)).toBe(expectedSecond);
  });

  it('rejects an over-budget multi-range row autofit before mutating layout', () => {
    const store = storeWith([[0, 0, 'value']]);
    const history = new History();
    autofitRowsHeightForRanges(store, history, [{ sheet: 0, r0: 0, c0: 0, r1: 100_000, c1: 0 }]);

    expect(store.getState().layout.rowHeights.size).toBe(0);
    expect(history.undo()).toBe(false);
  });

  it('rejects over-budget unique axis-span pairs before measuring or mutating', () => {
    const store = storeWith([[0, 0, 'value']]);
    const history = new History();
    const ranges: Array<{ sheet: number; r0: number; c0: number; r1: number; c1: number }> = [];
    for (let row = 0; row < 50_001; row += 1) {
      ranges.push(
        { sheet: 0, r0: row, c0: 0, r1: row, c1: 0 },
        { sheet: 0, r0: row, c0: 2, r1: row, c1: 2 },
      );
    }

    autofitRowsHeightForRanges(store, history, ranges);

    expect(store.getState().layout.rowHeights.size).toBe(0);
    expect(history.undo()).toBe(false);
  });

  it('rejects a large overlapping plan before materializing spans or measuring', () => {
    const store = storeWith([]);
    const history = new History();
    const ranges: Array<{ sheet: number; r0: number; c0: number; r1: number; c1: number }> = [];
    for (let col = 0; col <= 2_000; col += 1) {
      ranges.push({ sheet: 0, r0: 0, c0: col, r1: 50, c1: 2_000 });
    }

    // Every range is distinct, but the 51-row plans overlap almost entirely.
    // Make measurement observable: if the preflight guard is bypassed, the
    // first row scan attempts to iterate this map and fails the test.
    const cells = new Map<string, { value: CellValue; formula: string | null }>();
    Object.defineProperty(cells, Symbol.iterator, {
      value: () => {
        throw new Error('autofit measurement should not run after budget rejection');
      },
    });
    store.setState((state) => ({ ...state, data: { ...state.data, cells } }));

    expect(() => autofitRowsHeightForRanges(store, history, ranges)).not.toThrow();
    expect(store.getState().layout.rowHeights.size).toBe(0);
    expect(history.undo()).toBe(false);
  });

  it('rejects a mixed current-sheet and foreign-sheet selection atomically', () => {
    const store = storeWith([
      [0, 0, 'current sheet value'],
      [0, 2, 'foreign sheet value'],
    ]);
    const history = new History();
    autofitColsWidthForRanges(store, history, [
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      { sheet: 1, r0: 0, c0: 2, r1: 0, c1: 2 },
    ]);

    expect(store.getState().layout.colWidths.size).toBe(0);
    expect(history.undo()).toBe(false);
  });
});

describe('TABLE_HEADER_RESERVED_WIDTH', () => {
  it('matches the painted table header chevron', () => {
    const ctx = new Proxy({} as Record<string | symbol, unknown>, {
      get: () => () => {},
      set: () => true,
    }) as unknown as CanvasRenderingContext2D;
    const bounds = { x: 0, y: 0, w: 100, h: 20 };
    const chevron = paintTableHeaderChevron(ctx, bounds, {
      rule: '#000',
      fg: '#000',
    } as ResolvedTheme);
    expect(bounds.x + bounds.w - chevron.x).toBe(TABLE_HEADER_RESERVED_WIDTH);
  });
});

describe('FILTER_DROPDOWN_RESERVED_WIDTH', () => {
  const filterState = () => {
    const state = createSpreadsheetStore().getState();
    state.ui = { ...state.ui, filterRanges: [{ sheet: 0, r0: 0, c0: 0, r1: 5, c1: 0 }] };
    state.viewport = { rowStart: 0, rowCount: 3, colStart: 0, colCount: 3, zoom: 1, widthPx: 300 };
    return state;
  };

  it('matches the hit zone of the in-cell filter button', () => {
    const state = filterState();
    const layout = layoutForView(state);
    const colRight = gridOriginX(layout) + state.layout.defaultColWidth;
    const y = gridOriginY(layout) + 5;
    const fr = state.ui.filterRanges;
    const buttonLeft = colRight - FILTER_DROPDOWN_RESERVED_WIDTH;
    expect(hitZone(layout, state.viewport, buttonLeft, y, fr)).toEqual({
      kind: 'col-filter-btn',
      col: 0,
      range: fr[0],
    });
    expect(hitZone(layout, state.viewport, buttonLeft - 1, y, fr)).toEqual({
      kind: 'cell',
      row: 0,
      col: 0,
    });
  });

  it('matches the painted in-cell filter button', () => {
    const cell = { x: 10, y: 20, w: 80, h: 20 };
    expect(cell.x + cell.w - filterButtonRect(cell).x).toBe(FILTER_DROPDOWN_RESERVED_WIDTH);
  });
});
