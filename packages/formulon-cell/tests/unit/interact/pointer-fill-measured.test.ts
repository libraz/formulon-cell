import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { fillRange, fillShrinkDropped, fillShrinkFor } from '../../../src/commands/fill.js';
import { autoFillDownExtent } from '../../../src/commands/fill-plan.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachPointer } from '../../../src/interact/pointer.js';
import { gridHitState } from '../../../src/render/grid.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const seed = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  cells: Array<{ row: number; col: number; value: number | string }>,
): void => {
  store.setState((s) => {
    const map = new Map(s.data.cells);
    for (const c of cells) {
      const addr = { sheet: 0, row: c.row, col: c.col };
      if (typeof c.value === 'number') {
        wb.setNumber(addr, c.value);
        map.set(addrKey(addr), { value: { kind: 'number', value: c.value }, formula: null });
      } else {
        wb.setText(addr, c.value);
        map.set(addrKey(addr), { value: { kind: 'text', value: c.value }, formula: null });
      }
    }
    return { ...s, data: { ...s.data, cells: map } };
  });
  wb.recalc();
};

const fire = (host: HTMLElement, type: string, x: number, y: number): void => {
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

describe('fill-handle shrink', () => {
  const src = { sheet: 0, r0: 1, c0: 3, r1: 5, c1: 4 };

  it('keeps rows when the vertical offset dominates', () => {
    expect(fillShrinkFor(src, { row: 3, col: 4 }, { dx: -12, dy: -34 })).toEqual({
      ...src,
      r1: 3,
    });
  });

  it('keeps columns when the horizontal offset dominates', () => {
    expect(fillShrinkFor(src, { row: 3, col: 3 }, { dx: -77, dy: -34 })).toEqual({
      ...src,
      c1: 3,
    });
  });

  it('is not a shrink outside the source or without movement', () => {
    expect(fillShrinkFor(src, { row: 7, col: 4 }, { dx: 0, dy: 40 })).toBeNull();
    expect(fillShrinkFor(src, { row: 5, col: 4 }, { dx: -3, dy: -3 })).toBeNull();
  });

  it('drops the trailing rows or columns', () => {
    expect(fillShrinkDropped(src, { ...src, r1: 3 })).toEqual({ ...src, r0: 4 });
    expect(fillShrinkDropped(src, { ...src, c1: 3 })).toEqual({ ...src, c0: 4 });
  });
});

describe('pointer: fill shrink and tooltip', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let detach: () => void;

  beforeEach(async () => {
    host = document.createElement('div');
    Object.assign(host, {
      setPointerCapture() {},
      releasePointerCapture() {},
      hasPointerCapture: () => false,
    });
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        defaultColWidth: 104,
        defaultRowHeight: 28,
        headerColWidth: 52,
        headerRowHeight: 30,
      },
    }));
    wb = await WorkbookHandle.createDefault({ preferStub: true });
  });

  afterEach(() => {
    detach?.();
    document.body.innerHTML = '';
  });

  // Column B (x 156..260), rows 1..5 (y 58..198); corner (260, 198).
  const selectColumn = (): void => {
    seed(
      store,
      wb,
      [10, 20, 30, 40, 50].map((value, i) => ({ row: i + 1, col: 1, value })),
    );
    mutators.setActive(store, { sheet: 0, row: 1, col: 1 });
    mutators.extendRangeTo(store, { sheet: 0, row: 5, col: 1 });
    gridHitState(host).fillHandle = { x: 257, y: 195, w: 6, h: 6 };
  };

  it('pulling the handle back clears the dropped contents and shrinks the selection', () => {
    selectColumn();
    detach = attachPointer(host, store, wb);
    fire(host, 'pointerdown', 259, 197);
    fire(host, 'pointermove', 220, 130);
    expect(store.getState().ui.fillPreview).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 3, c1: 1 });
    expect(document.querySelector('.fc-fill-tooltip:not([hidden])')).toBeNull();
    fire(host, 'pointerup', 220, 130);
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 3, col: 1 })).toEqual({ kind: 'number', value: 30 });
    expect(wb.getValue({ sheet: 0, row: 4, col: 1 }).kind).toBe('blank');
    expect(wb.getValue({ sheet: 0, row: 5, col: 1 }).kind).toBe('blank');
    const sel = store.getState().selection;
    expect(sel.range).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 3, c1: 1 });
    expect(sel.active).toEqual({ sheet: 0, row: 1, col: 1 });
  });

  it('anchors the value tooltip to the drag range edge and follows the pointer row', () => {
    seed(store, wb, [{ row: 0, col: 1, value: 1 }]);
    mutators.setActive(store, { sheet: 0, row: 0, col: 1 });
    gridHitState(host).fillHandle = { x: 257, y: 53, w: 6, h: 6 };
    detach = attachPointer(host, store, wb);
    fire(host, 'pointerdown', 259, 55);
    fire(host, 'pointermove', 200, 130);
    const tip = document.querySelector<HTMLElement>('.fc-fill-tooltip');
    expect(tip?.hidden).toBe(false);
    // Column B's right edge is x=260; the box sits 5 px beyond it.
    const left = Number.parseFloat(tip?.style.left ?? '');
    const top = Number.parseFloat(tip?.style.top ?? '');
    expect(left).toBeCloseTo(265);
    expect(top).toBeCloseTo(127.5);
    fire(host, 'pointermove', 240, 130);
    expect(Number.parseFloat(tip?.style.left ?? '')).toBeCloseTo(left);
    fire(host, 'pointerup', 240, 130);
  });
});

describe('double-click fill extent', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await WorkbookHandle.createDefault({ preferStub: true });
  });

  it('stops before an occupied destination cell', () => {
    seed(store, wb, [
      ...Array.from({ length: 10 }, (_, i) => ({ row: i, col: 0, value: i + 1 })),
      { row: 0, col: 1, value: 1 },
      { row: 3, col: 1, value: 'x' },
    ]);
    const dest = autoFillDownExtent(store.getState(), { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 });
    expect(dest).toEqual({ sheet: 0, r0: 0, c0: 1, r1: 2, c1: 1 });
  });

  it('does nothing when the next destination cell is occupied', () => {
    seed(store, wb, [
      { row: 1, col: 0, value: 1 },
      { row: 2, col: 0, value: 2 },
      { row: 1, col: 1, value: 'x' },
    ]);
    expect(
      autoFillDownExtent(store.getState(), { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }),
    ).toBeNull();
  });
});

describe('fill over a filtered range', () => {
  it('leaves the rows the filter hides empty', async () => {
    const store = createSpreadsheetStore();
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    seed(store, wb, [{ row: 0, col: 0, value: 7 }]);
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenRows: new Set([2]) },
      ui: { ...s.ui, filterRanges: [{ sheet: 0, r0: 0, c0: 0, r1: 5, c1: 0 }] },
    }));
    const src = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
    fillRange(store.getState(), wb, src, { ...src, r1: 4 }, { copyOnly: true });
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'number', value: 7 });
    expect(wb.getValue({ sheet: 0, row: 2, col: 0 }).kind).toBe('blank');
    expect(wb.getValue({ sheet: 0, row: 3, col: 0 })).toEqual({ kind: 'number', value: 7 });
  });
});
