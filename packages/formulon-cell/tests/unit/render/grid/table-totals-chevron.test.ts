import { afterEach, describe, expect, it } from 'vitest';

import { defaultTableOverlay } from '../../../../src/commands/format-as-table.js';
import { buildColLayout, buildRowLayout, layoutForView } from '../../../../src/render/geometry.js';
import { getTotalsChevron, setTotalsChevron } from '../../../../src/render/grid/hit-state.js';
import { GridRenderer } from '../../../../src/render/grid.js';
import { createSpreadsheetStore, mutators, type State } from '../../../../src/store/store.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

const theme = {
  accent: '#0078d4',
  accentSoft: 'rgba(0,120,212,0.1)',
  bg: '#ffffff',
  bgRail: '#eeeeee',
  fg: '#1f1f1f',
  rule: '#cccccc',
} as unknown as ResolvedTheme;

/** A canvas that accepts every call and keeps assigned properties. */
function ctxStub(): CanvasRenderingContext2D {
  const target: Record<string, unknown> = {};
  return new Proxy(target, {
    get: (object, property) =>
      property in object ? object[property as string] : () => ({ width: 0 }),
    set: (object, property, value) => {
      object[property as string] = value;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
}

function paintActive(state: State): void {
  const renderer = new GridRenderer({
    host: { getBoundingClientRect: () => ({ width: 640, height: 480 }) } as unknown as HTMLElement,
    canvas: { getContext: () => ctxStub(), style: {} } as unknown as HTMLCanvasElement,
    getState: () => state,
    getTheme: () => theme,
  });
  const view = { ...state, layout: layoutForView(state) };
  (
    renderer as unknown as {
      paintActive(state: State, theme: ResolvedTheme, cols: unknown, rows: unknown): void;
    }
  ).paintActive(
    view,
    theme,
    buildColLayout(view.layout, view.viewport),
    buildRowLayout(view.layout, view.viewport),
  );
}

function tableStore(over: Partial<ReturnType<typeof defaultTableOverlay>> = {}) {
  const store = createSpreadsheetStore();
  mutators.upsertTableOverlay(store, {
    ...defaultTableOverlay('t1', { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 2 }),
    source: 'engine',
    name: 'Table1',
    showTotal: true,
    ...over,
  });
  return store;
}

afterEach(() => setTotalsChevron(null));

describe('totals-row function chevron', () => {
  it('is painted on the active cell of an engine table total row', () => {
    const store = tableStore();
    mutators.setActive(store, { sheet: 0, row: 3, col: 1 });
    paintActive(store.getState());
    expect(getTotalsChevron()).toMatchObject({ row: 3, col: 1, tableId: 't1', column: 1 });
  });

  it('is absent on a data row, without a total row, and on a session table', () => {
    const store = tableStore();
    mutators.setActive(store, { sheet: 0, row: 2, col: 1 });
    paintActive(store.getState());
    expect(getTotalsChevron()).toBeNull();

    const noTotal = tableStore({ showTotal: false });
    mutators.setActive(noTotal, { sheet: 0, row: 3, col: 1 });
    paintActive(noTotal.getState());
    expect(getTotalsChevron()).toBeNull();

    const session = tableStore({ source: 'session' });
    mutators.setActive(session, { sheet: 0, row: 3, col: 1 });
    paintActive(session.getState());
    expect(getTotalsChevron()).toBeNull();
  });
});
