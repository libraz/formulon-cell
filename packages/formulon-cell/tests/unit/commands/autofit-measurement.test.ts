import { describe, expect, it } from 'vitest';
import {
  autofitColWidth,
  autofitRowHeight,
  computeAutofitColWidth,
  computeAutofitRowHeight,
  FILTER_DROPDOWN_RESERVED_WIDTH,
} from '../../../src/commands/autofit-measurement.js';
import { History } from '../../../src/commands/history.js';
import { autofitColsWidth, autofitRowsHeight } from '../../../src/commands/structure.js';
import type { CellValue } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import type { SpreadsheetInstance } from '../../../src/mount/types.js';
import {
  buildColLayout,
  buildRowLayout,
  gridOriginX,
  hitZone,
  layoutForView,
} from '../../../src/render/geometry.js';
import { paintHeaders } from '../../../src/render/grid/headers.js';
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
    ui: { ...s.ui, filterRange: { sheet: 0, r0: 0, c0: 0, r1: 5, c1: 0 } },
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
  it('fits content plus padding, never below the 48px floor', () => {
    expect(computeAutofitColWidth(storeWith([[0, 0, 'HeaderName']]).getState(), 0, null)).toBe(87);
    expect(computeAutofitColWidth(storeWith([[0, 0, 'a']]).getState(), 0, null)).toBe(48);
    expect(computeAutofitColWidth(storeWith([]).getState(), 0, null)).toBe(48);
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

  it('measures formula text when formulas are shown', () => {
    const store = storeWith([[0, 0, 'v', '=REPT("abcdefgh",3)']]);
    expect(computeAutofitColWidth(store.getState(), 0, null)).toBe(48);
    store.setState((s) => ({ ...s, ui: { ...s.ui, showFormulas: true } }));
    expect(computeAutofitColWidth(store.getState(), 0, null)).toBe(150);
  });

  it('restricts the scan to a row span when one is given', () => {
    const state = storeWith([[5, 0, 'x'.repeat(40)]]).getState();
    expect(computeAutofitColWidth(state, 0, null)).toBe(297);
    expect(computeAutofitColWidth(state, 0, null, { span: { from: 0, to: 2 } })).toBe(48);
  });

  it('formats numbers with the requested locale and measures with the cell font', () => {
    const store = storeWith([[0, 0, 1234567.5]]);
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true, fontSize: 16 });
    const { ctx, measured } = fakeMeasureCtx();
    computeAutofitColWidth(store.getState(), 0, ctx, { locale: 'de-DE' });
    expect(measured).toEqual(['1.234.567,5']);
    expect(ctx.font).toBe('700 16px system-ui, sans-serif');
    measured.length = 0;
    computeAutofitColWidth(store.getState(), 0, ctx);
    expect(measured).toEqual(['1,234,567.5']);
  });
});

describe('computeAutofitRowHeight', () => {
  it('grows for explicit and wrapped lines, never below the default height', () => {
    expect(computeAutofitRowHeight(storeWith([]).getState(), 0, null)).toBe(20);
    expect(computeAutofitRowHeight(storeWith([[0, 0, 'HeaderName']]).getState(), 0, null)).toBe(25);
    expect(computeAutofitRowHeight(storeWith([[0, 0, 'a\nb\nc']]).getState(), 0, null)).toBe(59);

    const wrapped = storeWith([[0, 0, 'alpha beta gamma delta epsilon zeta']]);
    mutators.setColWidth(wrapped, 0, 70);
    mutators.setCellFormat(wrapped, { sheet: 0, row: 0, col: 0 }, { wrap: true });
    expect(computeAutofitRowHeight(wrapped.getState(), 0, null)).toBe(110);
  });

  it('restricts the scan to a column span and leaves clamping to the store', () => {
    const state = storeWith([[0, 5, 'y\ny\ny\ny']]).getState();
    expect(computeAutofitRowHeight(state, 0, null)).toBe(76);
    expect(computeAutofitRowHeight(state, 0, null, { span: { from: 0, to: 2 } })).toBe(20);

    const tall = storeWith([[0, 0, 'a\n'.repeat(40)]]);
    expect(computeAutofitRowHeight(tall.getState(), 0, null)).toBe(705);
    autofitRowsHeight(tall, null, 0, 0);
    expect(tall.getState().layout.rowHeights.get(0)).toBe(400);
  });
});

describe('public autofitColWidth / autofitRowHeight', () => {
  it('measure the given span through the shared implementation', () => {
    const store = storeWith([
      [0, 0, 'HeaderName'],
      [5, 0, 'x'.repeat(40)],
    ]);
    withFilter(store);
    expect(autofitColWidth({ store }, 0, 0, 2, 'en')).toBe(105);
    expect(autofitColWidth({ store }, 0, 0, 10, 'en')).toBe(297);
    expect(autofitRowHeight({ store }, 0, 0, 0, 'ja')).toBe(25);
  });
});

describe('ribbon Format > AutoFit', () => {
  const run = async (
    action: 'col-autofit' | 'row-autofit',
    store: SpreadsheetStore,
    range: { r0: number; c0: number; r1: number; c1: number },
  ): Promise<{ wb: WorkbookHandle; history: History }> => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const history = new History();
    const inst = { store, history, workbook: wb } as unknown as SpreadsheetInstance;
    await applyCellFormatAction(action, {
      inst,
      range: { sheet: 0, ...range },
      ribbonLang: 'en',
      ribbonMenuText: {},
      sheetTabColorByAction: () => undefined,
    } as unknown as CellFormatActionDeps);
    return { wb, history };
  };

  it('fits columns to the selected rows with the shared floor and filter reserve', async () => {
    const store = storeWith([
      [0, 0, 'a'],
      [0, 1, 'HeaderName'],
      [9, 1, 'x'.repeat(60)],
    ]);
    store.setState((s) => ({
      ...s,
      ui: { ...s.ui, filterRange: { sheet: 0, r0: 0, c0: 0, r1: 5, c1: 1 } },
    }));
    const { history } = await run('col-autofit', store, { r0: 0, c0: 0, r1: 2, c1: 1 });
    expect(store.getState().layout.colWidths.get(0)).toBe(48);
    expect(store.getState().layout.colWidths.get(1)).toBe(105);
    expect(history.undo()).toBe(true);
    expect(store.getState().layout.colWidths.size).toBe(0);
  });

  it('fits rows to the selected columns with store-clamped heights', async () => {
    const store = storeWith([
      [0, 0, 'a\n'.repeat(40)],
      [1, 5, 'y\ny\ny\ny'],
    ]);
    await run('row-autofit', store, { r0: 0, c0: 0, r1: 1, c1: 2 });
    expect(store.getState().layout.rowHeights.get(0)).toBe(400);
    expect(store.getState().layout.rowHeights.get(1)).toBe(20);
  });
});

describe('FILTER_DROPDOWN_RESERVED_WIDTH', () => {
  const filterState = () => {
    const state = createSpreadsheetStore().getState();
    state.ui = { ...state.ui, filterRange: { sheet: 0, r0: 0, c0: 0, r1: 5, c1: 0 } };
    state.viewport = { rowStart: 0, rowCount: 3, colStart: 0, colCount: 3, zoom: 1, widthPx: 300 };
    return state;
  };

  it('matches the hit zone of the header filter button', () => {
    const state = filterState();
    const layout = layoutForView(state);
    const colRight = gridOriginX(layout) + state.layout.defaultColWidth;
    const fr = state.ui.filterRange;
    const buttonLeft = colRight - FILTER_DROPDOWN_RESERVED_WIDTH;
    expect(hitZone(layout, state.viewport, buttonLeft, 10, fr)).toEqual({
      kind: 'col-filter-btn',
      col: 0,
    });
    expect(hitZone(layout, state.viewport, buttonLeft - 1, 10, fr)).toEqual({
      kind: 'col-header',
      col: 0,
    });
  });

  it('matches the painted header filter button', () => {
    const state = filterState();
    const xs: number[] = [];
    const noop = (): void => {};
    const ctx = new Proxy(
      {
        quadraticCurveTo(_cx: number, _cy: number, x: number) {
          xs.push(x);
        },
      } as Record<string | symbol, unknown>,
      {
        get: (target, prop) => (prop in target ? target[prop] : noop),
        set: () => true,
      },
    ) as unknown as CanvasRenderingContext2D;
    const layout = layoutForView(state);
    paintHeaders(
      { ctx, dpr: 1, cssWidth: 300, cssHeight: 140 },
      { ...state, layout },
      {
        textHeader: 11,
        fontUi: 'sans-serif',
      } as ResolvedTheme,
      buildColLayout(state.layout, state.viewport),
      buildRowLayout(state.layout, state.viewport),
    );
    expect(xs.length).toBeGreaterThan(0);
    const colRight = gridOriginX(layout) + state.layout.defaultColWidth;
    expect(colRight - Math.min(...xs)).toBe(FILTER_DROPDOWN_RESERVED_WIDTH);
  });
});
