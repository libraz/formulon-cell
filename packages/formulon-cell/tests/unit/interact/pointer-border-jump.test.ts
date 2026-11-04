import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachPointer } from '../../../src/interact/pointer.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

// Layout pinned below: header column 52 px, header row 30 px, columns 104 px,
// rows 28 px. Column A spans x 52..156, B 156..260; row 1 y 30..58, row 2
// y 58..86, row 6 y 170..198.
const fire = (host: HTMLElement, type: string, x: number, y: number): Event => {
  const init = { clientX: x, clientY: y, button: 0, bubbles: true, cancelable: true };
  const e =
    type === 'dblclick'
      ? new MouseEvent(type, init)
      : new PointerEvent(type, { ...init, pointerId: 1 });
  host.dispatchEvent(e);
  return e;
};

const doubleClick = (host: HTMLElement, x: number, y: number): Event => {
  for (let i = 0; i < 2; i += 1) {
    fire(host, 'pointerdown', x, y);
    fire(host, 'pointerup', x, y);
  }
  return fire(host, 'dblclick', x, y);
};

describe('Mac double-click on the active cell border', () => {
  let fcHost: HTMLElement;
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let detach: () => void;

  beforeEach(async () => {
    fcHost = document.createElement('div');
    fcHost.className = 'fc-host';
    fcHost.dataset.fcPlatform = 'mac';
    host = document.createElement('div');
    Object.assign(host, {
      setPointerCapture() {},
      releasePointerCapture() {},
      hasPointerCapture: () => false,
    });
    fcHost.appendChild(host);
    document.body.appendChild(fcHost);
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
    // A1:D6 holds a header row and five data rows.
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      for (let row = 0; row < 6; row += 1) {
        for (let col = 0; col < 4; col += 1) {
          cells.set(addrKey({ sheet: 0, row, col }), {
            value: { kind: 'number', value: row * 10 + col },
            formula: null,
          });
        }
      }
      return { ...s, data: { ...s.data, cells } };
    });
    detach = attachPointer(host, store, wb);
  });

  afterEach(() => {
    detach();
    fcHost.remove();
    wb.dispose();
  });

  const active = () => store.getState().selection.active;

  it('jumps to the edge of the data region from the edge that was double-clicked', () => {
    mutators.setActive(store, { sheet: 0, row: 1, col: 0 });
    expect(doubleClick(host, 100, 86).defaultPrevented).toBe(true);
    expect(active()).toEqual({ sheet: 0, row: 5, col: 0 });

    mutators.setActive(store, { sheet: 0, row: 1, col: 0 });
    doubleClick(host, 156, 70);
    expect(active()).toEqual({ sheet: 0, row: 1, col: 3 });

    mutators.setActive(store, { sheet: 0, row: 5, col: 1 });
    doubleClick(host, 200, 170);
    expect(active()).toEqual({ sheet: 0, row: 0, col: 1 });
    expect(store.getState().selection.range).toEqual({ sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 });
  });

  it('stays put on the left edge of column A', () => {
    mutators.setActive(store, { sheet: 0, row: 1, col: 0 });
    doubleClick(host, 53, 70);
    expect(active()).toEqual({ sheet: 0, row: 1, col: 0 });
  });

  it('keeps a plain click a few pixels past the edge selecting the neighbour', () => {
    mutators.setActive(store, { sheet: 0, row: 1, col: 0 });
    fire(host, 'pointerdown', 100, 90);
    fire(host, 'pointerup', 100, 90);
    expect(active()).toEqual({ sheet: 0, row: 2, col: 0 });
  });

  it('keeps the selection when the border itself is clicked', () => {
    mutators.setActive(store, { sheet: 0, row: 1, col: 0 });
    fire(host, 'pointerdown', 100, 87);
    fire(host, 'pointerup', 100, 87);
    expect(active()).toEqual({ sheet: 0, row: 1, col: 0 });
  });
});
