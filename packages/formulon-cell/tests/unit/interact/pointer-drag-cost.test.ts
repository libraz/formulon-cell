import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as fillModule from '../../../src/commands/fill.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachPointer } from '../../../src/interact/pointer.js';
import * as geometry from '../../../src/render/geometry.js';
import { gridHitState } from '../../../src/render/grid.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

vi.mock('../../../src/commands/fill.js', async (orig) => {
  const mod = await orig<typeof import('../../../src/commands/fill.js')>();
  return { ...mod, fillLastValue: vi.fn(mod.fillLastValue) };
});
vi.mock('../../../src/render/geometry.js', async (orig) => {
  const mod = await orig<typeof import('../../../src/render/geometry.js')>();
  return { ...mod, layoutForView: vi.fn(mod.layoutForView) };
});

const MOVES = 50;

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

describe('pointer drag per-move cost', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let detach: (() => void) | undefined;
  let rectReads: number;

  beforeEach(async () => {
    host = document.createElement('div');
    Object.assign(host, {
      setPointerCapture() {},
      releasePointerCapture() {},
      hasPointerCapture: () => false,
    });
    rectReads = 0;
    host.getBoundingClientRect = () => {
      rectReads += 1;
      return { left: 0, top: 0, right: 800, bottom: 400, width: 800, height: 400 } as DOMRect;
    };
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      viewport: { ...s.viewport, widthPx: 800 },
      layout: {
        ...s.layout,
        defaultColWidth: 104,
        defaultRowHeight: 28,
        headerColWidth: 52,
        headerRowHeight: 30,
      },
    }));
    wb = await WorkbookHandle.createDefault({ preferStub: true });
    vi.mocked(fillModule.fillLastValue).mockClear();
    vi.mocked(geometry.layoutForView).mockClear();
  });

  afterEach(() => {
    detach?.();
    document.body.innerHTML = '';
  });

  const seedCell = (): void => {
    const addr = { sheet: 0, row: 0, col: 1 };
    wb.setNumber(addr, 1);
    store.setState((s) => {
      const map = new Map(s.data.cells);
      map.set(addrKey(addr), { value: { kind: 'number', value: 1 }, formula: null });
      return { ...s, data: { ...s.data, cells: map } };
    });
    wb.recalc();
  };

  const countSetState = (): { count: () => number } => {
    const original = store.setState.bind(store);
    let n = 0;
    store.setState = ((...args: Parameters<typeof original>) => {
      n += 1;
      return original(...args);
    }) as typeof store.setState;
    return { count: () => n };
  };

  it('fill drag inside one cell: one rect read, one layout, one tooltip value per move', () => {
    seedCell();
    mutators.setActive(store, { sheet: 0, row: 0, col: 1 });
    gridHitState(host).fillHandle = { x: 257, y: 53, w: 6, h: 6 };
    detach = attachPointer(host, store, wb);
    fire(host, 'pointerdown', 259, 55);
    fire(host, 'pointermove', 200, 130);
    rectReads = 0;
    vi.mocked(fillModule.fillLastValue).mockClear();
    vi.mocked(geometry.layoutForView).mockClear();
    for (let i = 0; i < MOVES; i += 1) fire(host, 'pointermove', 200 + (i % 5), 130 + (i % 3));
    const tip = document.querySelector<HTMLElement>('.fc-fill-tooltip');
    expect(tip?.hidden).toBe(false);
    expect(rectReads).toBe(MOVES);
    expect(vi.mocked(fillModule.fillLastValue).mock.calls.length).toBeLessThanOrEqual(1);
    expect(vi.mocked(geometry.layoutForView).mock.calls.length).toBeLessThanOrEqual(2 * MOVES);
    fire(host, 'pointerup', 200, 130);
  });

  it('cell drag inside one cell: one rect read per move, one selection write', () => {
    mutators.setActive(store, { sheet: 0, row: 1, col: 1 });
    detach = attachPointer(host, store, wb);
    fire(host, 'pointerdown', 200, 80);
    fire(host, 'pointermove', 300, 125);
    const spy = countSetState();
    rectReads = 0;
    for (let i = 0; i < MOVES; i += 1) fire(host, 'pointermove', 300 + (i % 5), 125 + (i % 3));
    expect(rectReads).toBe(MOVES);
    expect(spy.count()).toBe(0);
    fire(host, 'pointerup', 300, 125);
  });
});
