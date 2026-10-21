import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachPointer } from '../../../src/interact/pointer.js';
import { attachWheel } from '../../../src/interact/wheel.js';
import {
  colLeadingEdge,
  colWidth,
  layoutForView,
  rowHeight,
  rowTopEdge,
} from '../../../src/render/geometry.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const stubPointerCapture = (host: HTMLElement): void => {
  const captured = new Set<number>();
  Object.assign(host, {
    setPointerCapture(id: number) {
      captured.add(id);
    },
    releasePointerCapture(id: number) {
      captured.delete(id);
    },
    hasPointerCapture(id: number) {
      return captured.has(id);
    },
  });
};

const firePointer = (host: HTMLElement, type: string, x: number, y: number): void => {
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

/** Drag row 0's bottom edge so the row spans `screenPx` on screen. */
const dragRow0To = (host: HTMLElement, store: SpreadsheetStore, screenPx: number): void => {
  const s = store.getState();
  const layout = layoutForView(s);
  const top = rowTopEdge(layout, s.viewport, 0);
  const bottom = top + rowHeight(layout, 0, s.viewport);
  firePointer(host, 'pointerdown', 10, bottom - 1);
  firePointer(host, 'pointermove', 10, top + screenPx);
  firePointer(host, 'pointerup', 10, top + screenPx);
};

/** Drag column 0's trailing edge so the column spans `screenPx` on screen. */
const dragCol0To = (host: HTMLElement, store: SpreadsheetStore, screenPx: number): void => {
  const s = store.getState();
  const layout = layoutForView(s);
  const leading = colLeadingEdge(layout, s.viewport, 0);
  const trailing = leading + colWidth(layout, 0, s.viewport);
  firePointer(host, 'pointerdown', trailing - 1, 10);
  firePointer(host, 'pointermove', leading + screenPx, 10);
  firePointer(host, 'pointerup', leading + screenPx, 10);
};

describe('screen-scaled resize and wheel', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let detach: (() => void) | undefined;

  const useLayout = (screenScale: number | undefined, zoom = 1): void => {
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        screenScale,
        defaultRowHeight: screenScale === 0.75 ? 80 / 3 : 20,
        defaultColWidth: screenScale === 0.75 ? 100 : 64,
        headerColWidth: 40,
        headerRowHeight: 24,
      },
      viewport: { ...s.viewport, zoom },
    }));
  };

  beforeEach(async () => {
    host = document.createElement('div');
    stubPointerCapture(host);
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await WorkbookHandle.createDefault({ preferStub: true });
  });

  afterEach(() => {
    detach?.();
    detach = undefined;
    host.remove();
  });

  it('Mac row drag to 30 screen px stores 40 canonical px (30 pt)', () => {
    useLayout(0.75);
    detach = attachPointer(host, store, wb);
    dragRow0To(host, store, 30);
    const h = store.getState().layout.rowHeights.get(0);
    expect(h).toBe(40);
    expect(((h ?? 0) * 3) / 4).toBe(30);
  });

  it('Mac column drag to 75 screen px stores 100 canonical px', () => {
    useLayout(0.75);
    detach = attachPointer(host, store, wb);
    dragCol0To(host, store, 75);
    expect(store.getState().layout.colWidths.get(0)).toBe(100);
  });

  it('default-profile row drag to 30 screen px stores 30 canonical px (22.5 pt)', () => {
    useLayout(undefined);
    detach = attachPointer(host, store, wb);
    dragRow0To(host, store, 30);
    expect(store.getState().layout.rowHeights.get(0)).toBe(30);
  });

  it('row drag at zoom 1.5 divides the screen delta by the zoom', () => {
    useLayout(undefined, 1.5);
    detach = attachPointer(host, store, wb);
    dragRow0To(host, store, 45);
    expect(store.getState().layout.rowHeights.get(0)).toBe(30);
  });

  it('a wheel delta of one Mac row scrolls one row', () => {
    useLayout(0.75);
    detach = attachWheel({ grid: host, store });
    host.dispatchEvent(new WheelEvent('wheel', { bubbles: true, cancelable: true, deltaY: 20 }));
    expect(store.getState().viewport.rowStart).toBe(1);
  });
});
