import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { attachWheel } from '../../../src/interact/wheel.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

// happy-dom's WheelEvent constructor only reads delta* from the init bag and
// drops modifier keys, so we patch them on after construction.
const fire = (
  el: HTMLElement,
  init: {
    deltaX?: number;
    deltaY?: number;
    ctrlKey?: boolean;
    metaKey?: boolean;
    shiftKey?: boolean;
    deltaMode?: number;
    clientX?: number;
    clientY?: number;
  },
): WheelEvent => {
  const e = new WheelEvent('wheel', {
    bubbles: true,
    cancelable: true,
    deltaX: init.deltaX ?? 0,
    deltaY: init.deltaY ?? 0,
  });
  Object.defineProperty(e, 'deltaMode', { value: init.deltaMode ?? 0 });
  Object.defineProperty(e, 'clientX', { value: init.clientX ?? 0 });
  Object.defineProperty(e, 'clientY', { value: init.clientY ?? 0 });
  Object.defineProperty(e, 'ctrlKey', { value: init.ctrlKey ?? false });
  Object.defineProperty(e, 'metaKey', { value: init.metaKey ?? false });
  Object.defineProperty(e, 'shiftKey', { value: init.shiftKey ?? false });
  el.dispatchEvent(e);
  return e;
};

describe('attachWheel', () => {
  let grid: HTMLElement;
  let store: SpreadsheetStore;
  let detach: () => void;

  beforeEach(() => {
    grid = document.createElement('div');
    document.body.appendChild(grid);
    store = createSpreadsheetStore();
    detach = attachWheel({ grid, store });
  });

  afterEach(() => {
    detach();
    grid.remove();
  });

  it('line-mode deltas scroll one row or column per line', () => {
    fire(grid, { deltaY: 3, deltaMode: 1 });
    expect(store.getState().viewport.rowStart).toBe(3);
    fire(grid, { deltaX: 2, deltaMode: 1 });
    expect(store.getState().viewport.colStart).toBe(2);
    fire(grid, { deltaY: 1, deltaMode: 1, shiftKey: true });
    expect(store.getState().viewport.colStart).toBe(3);
  });

  it('page-mode deltas scroll a page of visible rows or columns', () => {
    const { rowCount, colCount } = store.getState().viewport;
    fire(grid, { deltaY: 1, deltaMode: 2 });
    expect(store.getState().viewport.rowStart).toBe(rowCount);
    fire(grid, { deltaX: 1, deltaMode: 2 });
    expect(store.getState().viewport.colStart).toBe(colCount);
  });

  it('vertical wheel scrolls rows by delta / defaultRowHeight', () => {
    const rh = store.getState().layout.defaultRowHeight; // 20
    fire(grid, { deltaY: rh * 3 });
    expect(store.getState().viewport.rowStart).toBe(3);
    expect(store.getState().viewport.colStart).toBe(0);
  });

  it('horizontal wheel scrolls cols by delta / defaultColWidth', () => {
    const cw = store.getState().layout.defaultColWidth; // 64
    fire(grid, { deltaX: cw * 2 });
    expect(store.getState().viewport.colStart).toBe(2);
    expect(store.getState().viewport.rowStart).toBe(0);
  });

  it('walks columns the other way on a right-to-left sheet', () => {
    const cw = store.getState().layout.defaultColWidth;
    mutators.setRightToLeft(store, true);
    mutators.scrollBy(store, 0, 5);

    // Scrolling right now uncovers lower column indices, because they sit to
    // the right of where the view already is.
    fire(grid, { deltaX: cw * 2 });
    expect(store.getState().viewport.colStart).toBe(3);
    fire(grid, { deltaX: -cw * 2 });
    expect(store.getState().viewport.colStart).toBe(5);
  });

  it('shift+wheel routes deltaY into horizontal scroll', () => {
    const cw = store.getState().layout.defaultColWidth;
    fire(grid, { deltaY: cw * 2, shiftKey: true });
    expect(store.getState().viewport.colStart).toBe(2);
    expect(store.getState().viewport.rowStart).toBe(0);
  });

  it('accumulates sub-row trackpad deltas across events', () => {
    const rh = store.getState().layout.defaultRowHeight;
    // A single tiny delta should not move yet — under 1 row of accumulated px.
    fire(grid, { deltaY: 1 });
    expect(store.getState().viewport.rowStart).toBe(0);
    // Drip-feed exactly `rh` more pixels in tiny chunks → one row scroll.
    for (let i = 0; i < rh - 1; i += 1) fire(grid, { deltaY: 1 });
    expect(store.getState().viewport.rowStart).toBe(1);
    // Drip-feed another rh px → second row.
    for (let i = 0; i < rh; i += 1) fire(grid, { deltaY: 1 });
    expect(store.getState().viewport.rowStart).toBe(2);
  });

  it('a reversed gesture does not inherit the previous remainder', () => {
    const rh = store.getState().layout.defaultRowHeight;
    fire(grid, { deltaY: 10, deltaMode: 1 });
    expect(store.getState().viewport.rowStart).toBe(10);
    // Three quarters of a row down, then two half-row nudges back up.
    fire(grid, { deltaY: Math.floor(rh * 0.75) });
    fire(grid, { deltaY: -Math.ceil(rh / 2) });
    fire(grid, { deltaY: -Math.ceil(rh / 2) });
    expect(store.getState().viewport.rowStart).toBe(9);
  });

  it('a delta-mode change drops the remainder of the other unit', () => {
    const rh = store.getState().layout.defaultRowHeight;
    fire(grid, { deltaY: Math.floor(rh * 0.75) });
    fire(grid, { deltaY: 1, deltaMode: 1 });
    expect(store.getState().viewport.rowStart).toBe(1);
    // Had the 0.75 row survived, this quarter row would complete another row.
    fire(grid, { deltaY: Math.ceil(rh * 0.5) });
    expect(store.getState().viewport.rowStart).toBe(1);
  });

  it('clamps at the top — negative deltaY past zero stays at rowStart 0', () => {
    fire(grid, { deltaY: -1000 });
    expect(store.getState().viewport.rowStart).toBe(0);
  });

  it('respects freeze rows as the lower bound', () => {
    store.setState((s) => ({ ...s, layout: { ...s.layout, freezeRows: 3 } }));
    // Even a big negative delta cannot pull rowStart below freezeRows.
    fire(grid, { deltaY: -10_000 });
    expect(store.getState().viewport.rowStart).toBeGreaterThanOrEqual(3);
  });

  it('clamps at sheet bottom so the viewport keeps rowCount rows visible', () => {
    // 1_048_576 - rowCount(40) = 1_048_536 is the maximum rowStart.
    fire(grid, { deltaY: 1_000_000_000 });
    const { rowStart, rowCount } = store.getState().viewport;
    expect(rowStart).toBe(1_048_576 - rowCount);
  });

  it('ctrl+wheel zooms instead of scrolling', () => {
    const before = store.getState().viewport;
    fire(grid, { deltaY: -100, ctrlKey: true });
    const after = store.getState().viewport;
    expect(after.zoom).toBeGreaterThan(before.zoom);
    // No scroll happened.
    expect(after.rowStart).toBe(before.rowStart);
    expect(after.colStart).toBe(before.colStart);
  });

  it('cmd+wheel (metaKey) zooms', () => {
    const before = store.getState().viewport.zoom;
    fire(grid, { deltaY: -100, metaKey: true });
    expect(store.getState().viewport.zoom).toBeGreaterThan(before);
  });

  it('preventDefault is called when scrolling moves the viewport', () => {
    const rh = store.getState().layout.defaultRowHeight;
    const e = fire(grid, { deltaY: rh });
    expect(e.defaultPrevented).toBe(true);
  });

  it('preventDefault is NOT called when the delta rounds to zero', () => {
    const e = fire(grid, { deltaY: 1 });
    expect(e.defaultPrevented).toBe(false);
  });

  it('detach removes the listener', () => {
    detach();
    const rh = store.getState().layout.defaultRowHeight;
    fire(grid, { deltaY: rh * 5 });
    expect(store.getState().viewport.rowStart).toBe(0);
    // Re-attach so afterEach detach is a no-op-safe call.
    detach = attachWheel({ grid, store });
  });

  it('scrolls the top split pane when the wheel is over it', () => {
    mutators.setFreezePanes(store, 5, 0, { row: 0, col: 0 }, true);
    mutators.scrollBy(store, 20, 0);
    const rh = store.getState().layout.defaultRowHeight;
    // Header (20 px) plus two rows down: inside the top pane.
    fire(grid, { deltaY: rh * 3, clientX: 100, clientY: 60 });
    expect(store.getState().layout.freezeOriginRow).toBe(3);
    expect(store.getState().viewport.rowStart).toBe(25);
    // Below the split the bottom pane scrolls as usual.
    fire(grid, { deltaY: rh * 2, clientX: 100, clientY: 300 });
    expect(store.getState().viewport.rowStart).toBe(27);
    expect(store.getState().layout.freezeOriginRow).toBe(3);
  });
});

describe('attachWheel on the Mac profile with frozen panes', () => {
  let host: HTMLElement;
  let grid: HTMLElement;
  let store: SpreadsheetStore;
  let detach: () => void;

  beforeEach(() => {
    host = document.createElement('div');
    host.className = 'fc-host';
    host.dataset.fcPlatform = 'mac';
    grid = document.createElement('div');
    host.appendChild(grid);
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    store.setState((s) => ({ ...s, layout: { ...s.layout, freezeRows: 1, freezeCols: 1 } }));
    mutators.scrollBy(store, 1, 1);
    detach = attachWheel({ grid, store });
  });

  afterEach(() => {
    detach();
    host.remove();
  });

  const row = (): number => store.getState().viewport.rowStart;
  const col = (): number => store.getState().viewport.colStart;
  const body = { clientX: 400, clientY: 330 };

  it('scrolls 0.7 rows per line, rounding half down', () => {
    const start = row();
    for (const [lines, rows] of [
      [1, 1],
      [3, 2],
      [5, 3],
      [10, 7],
      [20, 14],
    ] as const) {
      detach();
      detach = attachWheel({ grid, store });
      const before = row();
      fire(grid, { deltaY: lines, deltaMode: 1, ...body });
      expect(row() - before).toBe(rows);
    }
    expect(row()).toBeGreaterThan(start);
  });

  it('accumulates single-line events: five of them move three rows', () => {
    const before = row();
    for (let i = 0; i < 5; i += 1) fire(grid, { deltaY: 1, deltaMode: 1, ...body });
    expect(row() - before).toBe(3);
  });

  it('scrolls 0.15 columns per line with a dead zone for one small event', () => {
    const before = col();
    fire(grid, { deltaX: 5, deltaMode: 1, ...body });
    expect(col()).toBe(before);
    fire(grid, { deltaX: 20, deltaMode: 1, ...body });
    expect(col() - before).toBeGreaterThanOrEqual(3);
  });

  it('scrolls both axes in one event', () => {
    const r = row();
    const c = col();
    fire(grid, { deltaY: 20, deltaX: 20, deltaMode: 1, ...body });
    expect(row() - r).toBe(14);
    expect(col() - c).toBe(3);
  });

  it('scrolls rows only over the frozen column', () => {
    const r = row();
    const c = col();
    const { headerColWidth } = store.getState().layout;
    fire(grid, { deltaY: 20, deltaX: 20, deltaMode: 1, clientX: headerColWidth + 5, clientY: 330 });
    expect(row() - r).toBe(14);
    expect(col()).toBe(c);
  });

  it('ignores the frozen row, the corner and the headers', () => {
    const r = row();
    const c = col();
    const { headerColWidth, headerRowHeight, defaultRowHeight } = store.getState().layout;
    for (const at of [
      { clientX: 400, clientY: headerRowHeight + 2 },
      { clientX: headerColWidth + 2, clientY: headerRowHeight + 2 },
      { clientX: 400, clientY: 5 },
      { clientX: 5, clientY: headerRowHeight + defaultRowHeight * 4 },
    ]) {
      fire(grid, { deltaY: 20, deltaX: 20, deltaMode: 1, ...at });
    }
    expect(row()).toBe(r);
    expect(col()).toBe(c);
  });
});
