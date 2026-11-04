import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  AUTOSCROLL_BASE_RATE,
  AUTOSCROLL_EDGE_INSET_PX,
  createDragAutoscroll,
} from '../../../src/interact/drag-autoscroll.js';
import { attachPointer } from '../../../src/interact/pointer.js';
import {
  colLeadingEdge,
  layoutForView,
  rowHeight,
  rowTopEdge,
} from '../../../src/render/geometry.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const BOUNDS = { left: 100, top: 100, right: 500, bottom: 400 };
const BOTH = { row: true, col: true };

describe('createDragAutoscroll', () => {
  beforeEach(() => vi.useFakeTimers());
  afterEach(() => vi.useRealTimers());

  it('stays idle while the pointer is well inside the body', () => {
    const scrollBy = vi.fn();
    const a = createDragAutoscroll({ scrollBy, onScrolled: vi.fn() });
    a.update({ x: 300, y: 250 }, BOUNDS, BOTH);
    vi.advanceTimersByTime(1000);
    expect(scrollBy).not.toHaveBeenCalled();
  });

  it('scrolls once the pointer is inside the last row at the edge', () => {
    const scrollBy = vi.fn();
    const a = createDragAutoscroll({ scrollBy, onScrolled: vi.fn() });
    a.update({ x: 300, y: BOUNDS.bottom - AUTOSCROLL_EDGE_INSET_PX + 1 }, BOUNDS, BOTH);
    vi.advanceTimersByTime(1000);
    const rows = scrollBy.mock.calls.reduce((n, c) => n + c[0], 0);
    expect(rows).toBeGreaterThanOrEqual(AUTOSCROLL_BASE_RATE - 1);
    a.stop();
  });

  it('runs faster the further past the edge and lands in the measured band at ~18 px', () => {
    const total = (y: number): number => {
      const scrollBy = vi.fn();
      const a = createDragAutoscroll({ scrollBy, onScrolled: vi.fn() });
      a.update({ x: 300, y }, BOUNDS, BOTH);
      vi.advanceTimersByTime(1000);
      a.stop();
      return scrollBy.mock.calls.reduce((n, c) => n + c[0], 0);
    };
    const near = total(BOUNDS.bottom + 18);
    expect(near).toBeGreaterThanOrEqual(12);
    expect(near).toBeLessThanOrEqual(32);
    expect(total(BOUNDS.bottom + 100)).toBeGreaterThan(near);
  });

  it('scrolls up and left with negative deltas and honours the axis mask', () => {
    const scrollBy = vi.fn();
    const a = createDragAutoscroll({ scrollBy, onScrolled: vi.fn() });
    a.update({ x: 50, y: 50 }, BOUNDS, { row: false, col: true });
    vi.advanceTimersByTime(1000);
    expect(scrollBy.mock.calls.every((c) => c[0] === 0 && c[1] < 0)).toBe(true);
    a.stop();
  });

  it('stops when the pointer returns inside or stop() is called', () => {
    const scrollBy = vi.fn();
    const a = createDragAutoscroll({ scrollBy, onScrolled: vi.fn() });
    a.update({ x: 300, y: 450 }, BOUNDS, BOTH);
    vi.advanceTimersByTime(500);
    a.update({ x: 300, y: 250 }, BOUNDS, BOTH);
    const n = scrollBy.mock.calls.length;
    vi.advanceTimersByTime(500);
    expect(scrollBy.mock.calls.length).toBe(n);
  });
});

describe('pointer drag autoscroll', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let detach: (() => void) | undefined;

  const fire = (type: string, x: number, y: number): void => {
    host.dispatchEvent(
      new PointerEvent(type, {
        clientX: x,
        clientY: y,
        button: 0,
        bubbles: true,
        cancelable: true,
        pointerId: 1,
      }),
    );
  };

  beforeEach(async () => {
    vi.useFakeTimers();
    host = document.createElement('div');
    const captured = new Set<number>();
    Object.assign(host, {
      setPointerCapture: (id: number) => captured.add(id),
      releasePointerCapture: (id: number) => captured.delete(id),
      hasPointerCapture: (id: number) => captured.has(id),
    });
    host.getBoundingClientRect = () =>
      ({ left: 0, top: 0, right: 800, bottom: 400, width: 800, height: 400 }) as DOMRect;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    store.setState((s) => ({ ...s, viewport: { ...s.viewport, widthPx: 800 } }));
    wb = await WorkbookHandle.createDefault({ preferStub: true });
    detach = attachPointer(host, store, wb);
  });

  afterEach(() => {
    detach?.();
    host.remove();
    vi.useRealTimers();
  });

  const cellCenter = (row: number): { x: number; y: number } => {
    const s = store.getState();
    const layout = layoutForView(s);
    return {
      x: colLeadingEdge(layout, s.viewport, 1) + 5,
      y: rowTopEdge(layout, s.viewport, row) + rowHeight(layout, row, s.viewport) / 2,
    };
  };

  it('scrolls and extends the selection at the bottom edge, then stops on pointerup', () => {
    const start = cellCenter(1);
    fire('pointerdown', start.x, start.y);
    fire('pointermove', start.x, 398);
    const before = store.getState().viewport.rowStart;
    const selBefore = store.getState().selection.range.r1;
    vi.advanceTimersByTime(1000);
    const after = store.getState();
    expect(after.viewport.rowStart).toBeGreaterThan(before);
    expect(after.selection.range.r1).toBeGreaterThan(selBefore);
    expect(after.selection.range.r0).toBe(1);

    fire('pointerup', start.x, 398);
    const frozen = store.getState().viewport.rowStart;
    vi.advanceTimersByTime(1000);
    expect(store.getState().viewport.rowStart).toBe(frozen);
  });

  it('stops on pointercancel and on teardown', () => {
    const start = cellCenter(1);
    fire('pointerdown', start.x, start.y);
    fire('pointermove', start.x, 398);
    fire('pointercancel', start.x, 398);
    const a = store.getState().viewport.rowStart;
    vi.advanceTimersByTime(1000);
    expect(store.getState().viewport.rowStart).toBe(a);

    fire('pointerdown', start.x, start.y);
    fire('pointermove', start.x, 398);
    detach?.();
    detach = undefined;
    const b = store.getState().viewport.rowStart;
    vi.advanceTimersByTime(1000);
    expect(store.getState().viewport.rowStart).toBe(b);
  });
});
