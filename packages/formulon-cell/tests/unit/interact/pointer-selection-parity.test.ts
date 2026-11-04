import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachPointer } from '../../../src/interact/pointer.js';
import { gridHitState } from '../../../src/render/grid.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

// Header column 52 px, header row 30 px, columns 104 px, rows 28 px.
// Column n spans x 52+104n .. 156+104n; row n spans y 30+28n .. 58+28n.
const cx = (col: number): number => 52 + 104 * col + 52;
const cy = (row: number): number => 30 + 28 * row + 14;

const fire = (host: HTMLElement, type: string, x: number, y: number, button = 0): Event => {
  const init = { clientX: x, clientY: y, button, bubbles: true, cancelable: true };
  const e =
    type === 'dblclick'
      ? new MouseEvent(type, init)
      : new PointerEvent(type, { ...init, pointerId: 1 });
  host.dispatchEvent(e);
  return e;
};

describe('pointer selection parity', () => {
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
    detach = attachPointer(host, store, wb);
  });

  afterEach(() => {
    detach();
    fcHost.remove();
    wb.dispose();
  });

  const sel = () => store.getState().selection;
  const select = (r0: number, c0: number, r1: number, c1: number) => {
    mutators.setActive(store, { sheet: 0, row: r0, col: c0 });
    mutators.setRange(store, { sheet: 0, r0, c0, r1, c1 });
  };
  const press = (x: number, y: number, button = 0) => {
    fire(host, 'pointerdown', x, y, button);
    if (button === 0) fire(host, 'pointerup', x, y);
  };
  const drag = (from: [number, number], to: [number, number]) => {
    fire(host, 'pointerdown', from[0], from[1]);
    fire(host, 'pointermove', to[0], to[1]);
    fire(host, 'pointerup', to[0], to[1]);
  };

  describe('right-click', () => {
    it('keeps a selection that already contains the cell', () => {
      select(1, 1, 3, 3);
      press(cx(2), cy(2), 2);
      expect(sel().range).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 3, c1: 3 });
      expect(sel().active).toMatchObject({ row: 1, col: 1 });
    });

    it('moves the selection to a cell outside it before the menu opens', () => {
      select(1, 1, 3, 3);
      press(cx(5), cy(7), 2);
      expect(sel().range).toEqual({ sheet: 0, r0: 7, c0: 5, r1: 7, c1: 5 });
      expect(sel().active).toMatchObject({ row: 7, col: 5 });
    });

    it('selects the whole sheet from the corner', () => {
      select(1, 1, 3, 3);
      press(10, 10, 2);
      expect(sel().range.r1).toBeGreaterThan(1000);
      expect(sel().active).toMatchObject({ row: 0, col: 0 });
    });
  });

  describe('plain click and drag', () => {
    it('collapses a click inside the selection to the clicked cell', () => {
      select(1, 1, 3, 3);
      press(cx(2), cy(2));
      expect(sel().range).toEqual({ sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 });
      select(1, 1, 3, 3);
      press(cx(1), cy(1));
      expect(sel().range).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 });
    });

    it('starts a new selection from the press cell when dragging inside the selection', () => {
      select(1, 1, 3, 3);
      drag([cx(2), cy(2)], [cx(5), cy(7)]);
      expect(sel().range).toEqual({ sheet: 0, r0: 2, c0: 2, r1: 7, c1: 5 });
      expect(sel().active).toMatchObject({ row: 2, col: 2 });
    });

    it('extends to column A when the drag reaches the row headers', () => {
      select(2, 2, 2, 2);
      drag([cx(2), cy(2)], [20, cy(2)]);
      expect(sel().range).toEqual({ sheet: 0, r0: 2, c0: 0, r1: 2, c1: 2 });
      expect(sel().active).toMatchObject({ row: 2, col: 2 });
    });

    it('keeps the press cell active when dragging up-left', () => {
      drag([cx(3), cy(4)], [cx(1), cy(1)]);
      expect(sel().range).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 4, c1: 3 });
      expect(sel().active).toMatchObject({ row: 4, col: 3 });
    });
  });

  describe('merged cells', () => {
    beforeEach(() => {
      mutators.mergeRange(store, { sheet: 0, r0: 1, c0: 7, r1: 2, c1: 8 });
    });

    it('selects the whole merge with its anchor active from any member cell', () => {
      press(cx(8), cy(2));
      expect(sel().range).toEqual({ sheet: 0, r0: 1, c0: 7, r1: 2, c1: 8 });
      expect(sel().active).toMatchObject({ row: 1, col: 7 });
    });

    it('grows a drag that ends inside the merge to the merge edges', () => {
      drag([cx(0), cy(0)], [cx(7), cy(1)]);
      expect(sel().range).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 2, c1: 8 });
      expect(sel().active).toMatchObject({ row: 0, col: 0 });
    });

    it('keeps the merge anchor active when dragging out of the merge', () => {
      drag([cx(7), cy(1)], [cx(8), cy(4)]);
      expect(sel().range).toEqual({ sheet: 0, r0: 1, c0: 7, r1: 4, c1: 8 });
      expect(sel().active).toMatchObject({ row: 1, col: 7 });
    });
  });

  describe('corner click', () => {
    it('selects the sheet and activates the top-left visible cell', () => {
      store.setState((s) => ({ ...s, viewport: { ...s.viewport, rowStart: 19, colStart: 4 } }));
      mutators.setActive(store, { sheet: 0, row: 24, col: 6 });
      press(10, 10);
      expect(sel().range.r0).toBe(0);
      expect(sel().range.r1).toBeGreaterThan(1000);
      expect(sel().active).toMatchObject({ row: 19, col: 4 });
      expect(store.getState().viewport.rowStart).toBe(19);
      press(10, 10);
      expect(sel().active).toMatchObject({ row: 19, col: 4 });
    });

    it('activates A1 when the sheet is not scrolled', () => {
      select(1, 1, 3, 3);
      press(10, 10);
      expect(sel().active).toMatchObject({ row: 0, col: 0 });
    });
  });

  describe('double-click on the selection border', () => {
    const dbl = (x: number, y: number) => {
      for (let i = 0; i < 2; i += 1) press(x, y);
      fire(host, 'dblclick', x, y);
    };
    beforeEach(() => {
      // A1:A10 and B1:E1 hold data.
      store.setState((s) => {
        const cells = new Map(s.data.cells);
        const put = (row: number, col: number) =>
          cells.set(addrKey({ sheet: 0, row, col }), {
            value: { kind: 'number', value: 1 },
            formula: null,
          });
        for (let r = 0; r < 10; r += 1) put(r, 0);
        for (let c = 1; c < 5; c += 1) put(0, c);
        return { ...s, data: { ...s.data, cells } };
      });
    });

    it('jumps from the frame cell under the pointer, not the active cell', () => {
      select(0, 0, 2, 1);
      dbl(cx(0), 30 + 28 * 3);
      expect(sel().active).toMatchObject({ row: 9, col: 0 });
    });

    it('selects a blank frame cell instead of jumping', () => {
      select(0, 0, 2, 1);
      dbl(cx(1), 30 + 28 * 3);
      expect(sel().range).toEqual({ sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 });
      select(0, 0, 2, 1);
      dbl(52 + 104 * 2, cy(1));
      expect(sel().range).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 });
    });

    it('does nothing on a corner of the frame', () => {
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      dbl(52 + 104, 30 + 28);
      expect(sel().active).toMatchObject({ row: 0, col: 0 });
      dbl(52, 30 + 28);
      expect(sel().active).toMatchObject({ row: 0, col: 0 });
    });

    it('stays on the top and left edges of A1', () => {
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      dbl(cx(0), 30);
      expect(sel().active).toMatchObject({ row: 0, col: 0 });
      dbl(52, cy(0));
      expect(sel().active).toMatchObject({ row: 0, col: 0 });
    });
  });

  describe('fill chip anchor', () => {
    it('announces the bottom-right corner of the last filled cell', () => {
      gridHitState(host).fillHandle = { x: 152, y: 54, w: 6, h: 6 };
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      let detail: { clientX: number; clientY: number } | null = null;
      host.addEventListener('fc:autofilloptions', (e) => {
        detail = (e as CustomEvent).detail;
      });
      fire(host, 'pointerdown', 155, 57);
      fire(host, 'pointermove', 100, 100);
      fire(host, 'pointerup', 130, 110);
      // Row 2 (y 86..114) in column A (x 52..156).
      expect(detail).toMatchObject({ clientX: 156, clientY: 114 });
    });
  });

  describe('right-button fill drag', () => {
    it('opens the fill menu on release instead of filling or opening the context menu', () => {
      gridHitState(host).fillHandle = { x: 152, y: 54, w: 6, h: 6 };
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      let detail: { dest: unknown; clientX: number } | null = null;
      host.addEventListener('fc:autofillmenu', (e) => {
        detail = (e as CustomEvent).detail;
      });
      let contextMenus = 0;
      host.addEventListener('contextmenu', () => {
        contextMenus += 1;
      });
      fire(host, 'pointerdown', 155, 57, 2);
      host.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }));
      expect(store.getState().ui.fillPreview).not.toBeNull();
      fire(host, 'pointermove', 100, 100, 2);
      fire(host, 'pointerup', 130, 110, 2);
      expect(detail).toMatchObject({
        dest: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 },
        clientX: 130,
      });
      expect(contextMenus).toBe(0);
      expect(store.getState().ui.fillPreview).toBeNull();
    });

    it('opens nothing when released on the source', () => {
      gridHitState(host).fillHandle = { x: 152, y: 54, w: 6, h: 6 };
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      let opened = false;
      host.addEventListener('fc:autofillmenu', () => {
        opened = true;
      });
      fire(host, 'pointerdown', 155, 57, 2);
      fire(host, 'pointerup', 155, 57, 2);
      expect(opened).toBe(false);
    });
  });

  describe('table resize grip', () => {
    const seedTable = (): void => {
      const put = (row: number, col: number, value: number | string) => {
        const addr = { sheet: 0, row, col };
        if (typeof value === 'number') wb.setNumber(addr, value);
        else wb.setText(addr, value);
        store.setState((s) => {
          const cells = new Map(s.data.cells);
          cells.set(addrKey(addr), {
            value: typeof value === 'number' ? { kind: 'number', value } : { kind: 'text', value },
            formula: null,
          });
          return { ...s, data: { ...s.data, cells } };
        });
      };
      for (const [c, h] of ['col1', 'h2', 'h3'].entries()) put(0, c, h);
      for (let r = 1; r < 7; r += 1) for (let c = 0; c < 3; c += 1) put(r, c, r * 10 + c);
      put(7, 0, 1);
      put(7, 1, 2);
      mutators.upsertTableOverlay(store, {
        id: 't1',
        source: 'session',
        range: { sheet: 0, r0: 0, c0: 0, r1: 6, c1: 2 },
        style: 'medium',
        showHeader: true,
        showTotal: false,
        banded: true,
      });
    };
    const range = () => store.getState().tables.tables[0]?.range;

    it('grows by columns, names the new headers on, and refuses to swallow occupied rows', () => {
      seedTable();
      mutators.setActive(store, { sheet: 0, row: 0, col: 4 });
      fire(host, 'pointerdown', 361, 223);
      expect(store.getState().ui.fillPreview).not.toBeNull();
      fire(host, 'pointermove', cx(4), cy(7));
      expect(store.getState().ui.fillPreview).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 6, c1: 4 });
      // The selection stays where it was while the target previews.
      expect(sel().active).toMatchObject({ row: 0, col: 4 });
      fire(host, 'pointerup', cx(4), cy(7));
      expect(range()).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 6, c1: 4 });
      expect(wb.getValue({ sheet: 0, row: 0, col: 3 })).toEqual({ kind: 'text', value: 'h4' });
      expect(wb.getValue({ sheet: 0, row: 0, col: 4 })).toEqual({ kind: 'text', value: 'h5' });
      expect(store.getState().ui.fillPreview).toBeNull();
    });

    it('shrinks by columns, keeps rows, and keeps the cells that leave the table', () => {
      seedTable();
      fire(host, 'pointerdown', 361, 223);
      fire(host, 'pointermove', cx(1), cy(2));
      fire(host, 'pointerup', cx(1), cy(2));
      expect(range()).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 6, c1: 1 });
      expect(wb.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({ kind: 'text', value: 'h3' });
    });

    it('grows into blank rows below the table', () => {
      seedTable();
      store.setState((s) => {
        const cells = new Map(s.data.cells);
        cells.delete(addrKey({ sheet: 0, row: 7, col: 0 }));
        cells.delete(addrKey({ sheet: 0, row: 7, col: 1 }));
        return { ...s, data: { ...s.data, cells } };
      });
      fire(host, 'pointerdown', 361, 223);
      fire(host, 'pointermove', cx(2), cy(9));
      fire(host, 'pointerup', cx(2), cy(9));
      expect(range()).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 9, c1: 2 });
    });
  });
});
