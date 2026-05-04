import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { attachWheel } from '../../../src/interact/wheel.js';
import { type SpreadsheetStore, createSpreadsheetStore } from '../../../src/store/store.js';

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
  },
): WheelEvent => {
  const e = new WheelEvent('wheel', {
    bubbles: true,
    cancelable: true,
    deltaX: init.deltaX ?? 0,
    deltaY: init.deltaY ?? 0,
  });
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

  it('vertical wheel scrolls rows by delta / defaultRowHeight', () => {
    const rh = store.getState().layout.defaultRowHeight; // 24
    fire(grid, { deltaY: rh * 3 });
    expect(store.getState().viewport.rowStart).toBe(3);
    expect(store.getState().viewport.colStart).toBe(0);
  });

  it('horizontal wheel scrolls cols by delta / defaultColWidth', () => {
    const cw = store.getState().layout.defaultColWidth; // 80
    fire(grid, { deltaX: cw * 2 });
    expect(store.getState().viewport.colStart).toBe(2);
    expect(store.getState().viewport.rowStart).toBe(0);
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
});
