import { beforeEach, describe, expect, it, vi } from 'vitest';
import * as errorIndicators from '../../../../src/commands/error-indicators.js';
import { createConditionalCache } from '../../../../src/render/conditional.js';
import { buildColLayout, buildRowLayout, layoutForView } from '../../../../src/render/geometry.js';
import { paintCells } from '../../../../src/render/grid/cells.js';
import * as renderScaleModule from '../../../../src/store/render-scale.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

vi.mock('../../../../src/commands/error-indicators.js', async (orig) => {
  const mod = await orig<typeof import('../../../../src/commands/error-indicators.js')>();
  return { ...mod, errorCheckFlagged: vi.fn(mod.errorCheckFlagged) };
});
vi.mock('../../../../src/store/render-scale.js', async (orig) => {
  const mod = await orig<typeof import('../../../../src/store/render-scale.js')>();
  return { ...mod, renderScale: vi.fn(mod.renderScale) };
});

const theme = {
  bg: '#fff',
  accent: '#0078d4',
  accentSoft: 'rgba(0,120,212,0.1)',
  fg: '#000',
  rule: '#ccc',
  cellAutoFg: '#000',
  fontUi: 'sans-serif',
  fontMono: 'monospace',
  textCell: 13,
} as unknown as ResolvedTheme;

class CountingMap<K, V> extends Map<K, V> {
  gets = 0;
  override get(key: K): V | undefined {
    this.gets += 1;
    return super.get(key);
  }
}

function makeCtx(counters: { createPattern: number }): CanvasRenderingContext2D {
  const target: Record<string, unknown> = {
    fillStyle: '',
    font: '',
    textAlign: 'left',
    measureText: () =>
      ({ width: 7, actualBoundingBoxAscent: 9, actualBoundingBoxDescent: 2 }) as TextMetrics,
    createPattern: () => {
      counters.createPattern += 1;
      return {};
    },
  };
  return new Proxy(target, {
    get: (o, p: string) => (p in o ? o[p] : () => undefined),
    set: (o, p: string, v) => {
      o[p] = v;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
}

function paint(opts: { merges?: number; background?: boolean; cells?: number }) {
  const store = createSpreadsheetStore();
  const cells = opts.cells ?? 0;
  for (let i = 0; i < cells; i += 1) {
    mutators.setCell(
      store,
      { sheet: 0, row: Math.floor(i / 10), col: i % 10 },
      { kind: 'number', value: i },
      null,
    );
  }
  for (let i = 0; i < (opts.merges ?? 0); i += 1) {
    mutators.mergeRange(store, { sheet: 0, r0: 20 + i * 2, c0: 1, r1: 21 + i * 2, c1: 2 });
  }
  store.setState((s) => ({
    ...s,
    viewport: { ...s.viewport, widthPx: 900, rowCount: 40, colCount: 12 },
  }));
  const state = store.getState();
  const byAnchor = new CountingMap(state.merges.byAnchor);
  const view = { ...state, layout: layoutForView(state), merges: { ...state.merges, byAnchor } };
  const counters = { createPattern: 0 };
  paintCells(
    {
      ctx: makeCtx(counters),
      wb: null,
      locale: 'en-US',
      getDisplay: undefined,
      sheetBackgroundImage: opts.background
        ? ({ naturalWidth: 4, naturalHeight: 4 } as HTMLImageElement)
        : null,
      fontEffects: false,
      conditionalCache: createConditionalCache(),
    },
    view,
    theme,
    buildColLayout(view.layout, view.viewport),
    buildRowLayout(view.layout, view.viewport),
  );
  return { byAnchor, counters };
}

beforeEach(() => {
  vi.mocked(errorIndicators.errorCheckFlagged).mockClear();
  vi.mocked(renderScaleModule.renderScale).mockClear();
});

describe('paintCells per-cell cost', () => {
  it('looks a merge up only for anchors, not for every painted cell', () => {
    const { byAnchor } = paint({ cells: 100, merges: 3 });
    // One lookup per anchor (the merged rects), none for the 100 plain cells.
    expect(byAnchor.gets).toBeLessThanOrEqual(3 + 3);
  });

  it('shares one options object across every error-check call', () => {
    paint({ cells: 100 });
    const calls = vi.mocked(errorIndicators.errorCheckFlagged).mock.calls;
    expect(calls.length).toBeGreaterThanOrEqual(100);
    expect(new Set(calls.map((c) => c[3])).size).toBe(1);
  });

  it('builds the background pattern once per paint', () => {
    const { counters } = paint({ cells: 0, merges: 4, background: true });
    expect(counters.createPattern).toBe(1);
  });

  it('derives the render scale once per paint, not once per cell', () => {
    paint({ cells: 100 });
    // Cell geometry still reads it once per painted cell; the painter adds none.
    expect(vi.mocked(renderScaleModule.renderScale).mock.calls.length).toBeLessThan(150);
  });
});
