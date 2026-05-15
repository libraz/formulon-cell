import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachPointer } from '../../../src/interact/pointer.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

// We mock getFillHandleRect on the rendering module so tests can place a fake
// fill handle without running the renderer. By default, every test runs with
// the real implementation (returns null) — only the dedicated fill suite
// reassigns the spy.
vi.mock('../../../src/render/grid.js', async () => {
  const actual = await vi.importActual<typeof import('../../../src/render/grid.js')>(
    '../../../src/render/grid.js',
  );
  return { ...actual, getFillHandleRect: vi.fn(() => null) };
});

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

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

const fireDown = (
  host: HTMLElement,
  x: number,
  y: number,
  init: PointerEventInit = {},
): PointerEvent => {
  const e = new PointerEvent('pointerdown', {
    clientX: x,
    clientY: y,
    button: 0,
    bubbles: true,
    cancelable: true,
    pointerId: 1,
    ...init,
  });
  host.dispatchEvent(e);
  return e;
};

const fireMove = (host: HTMLElement, x: number, y: number): PointerEvent => {
  const e = new PointerEvent('pointermove', {
    clientX: x,
    clientY: y,
    bubbles: true,
    cancelable: true,
    pointerId: 1,
  });
  host.dispatchEvent(e);
  return e;
};

const fireUp = (host: HTMLElement, x: number, y: number): PointerEvent => {
  const e = new PointerEvent('pointerup', {
    clientX: x,
    clientY: y,
    bubbles: true,
    cancelable: true,
    pointerId: 1,
  });
  host.dispatchEvent(e);
  return e;
};

const fireDblClick = (host: HTMLElement, x: number, y: number): MouseEvent => {
  const e = new MouseEvent('dblclick', {
    clientX: x,
    clientY: y,
    bubbles: true,
    cancelable: true,
  });
  host.dispatchEvent(e);
  return e;
};

// Default layout coordinates (headerCol=52, headerRow=30, defaultColW=104,
// defaultRowH=28, RESIZE_SLACK=4):
//   col 0 x∈[52,156), col 1 x∈[156,260), col 2 x∈[260,364)
//   row 0 y∈[30,58),  row 1 y∈[58,86),   row 2 y∈[86,114)
//   corner:        (10, 10)
//   col-header c0: (100, 10)   right-edge of c0 at x=156 → (151, 10) is still header
//   col-resize c0: (154, 10)
//   row-header r0: (10, 40)    bottom-edge at y=58 → (10, 53) is still header
//   row-resize r0: (10, 56)
//   cell (0,0):    (100, 40)
//   cell (1,1):    (200, 70)
//   cell (2,2):    (300, 100)

describe('attachPointer', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let detach: () => void;

  beforeEach(async () => {
    host = document.createElement('div');
    stubPointerCapture(host);
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    // Pin layout to the legacy roomy defaults so the hardcoded pixel
    // coordinates in this file stay meaningful even when production
    // defaults change.
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
    wb = await newWb();
  });

  afterEach(() => {
    detach?.();
    document.body.innerHTML = '';
  });

  describe('cell selection', () => {
    it('left-click sets the active cell', () => {
      detach = attachPointer(host, store, wb);
      fireDown(host, 200, 70); // cell (1, 1)
      fireUp(host, 200, 70);
      expect(store.getState().selection.active).toEqual({ sheet: 0, row: 1, col: 1 });
    });

    it('shift-click extends the range from the anchor', () => {
      detach = attachPointer(host, store, wb);
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      fireDown(host, 300, 100, { shiftKey: true }); // cell (2, 2)
      fireUp(host, 300, 100);
      expect(store.getState().selection.range).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 });
    });

    it('drag from one cell to another extends the selection live', () => {
      detach = attachPointer(host, store, wb);
      fireDown(host, 100, 40); // anchor (0, 0)
      fireMove(host, 300, 100); // (2, 2)
      expect(store.getState().selection.range).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 });
      fireUp(host, 300, 100);
    });

    it('non-primary mouse button is ignored on pointerdown', () => {
      detach = attachPointer(host, store, wb);
      const e = new PointerEvent('pointerdown', {
        clientX: 100,
        clientY: 50,
        button: 2,
        bubbles: true,
        cancelable: true,
        pointerId: 1,
      });
      host.dispatchEvent(e);
      // Active stays at default (0, 0).
      expect(store.getState().selection.active).toEqual({ sheet: 0, row: 0, col: 0 });
    });
  });

  describe('headers', () => {
    it('col-header click selects the entire column', () => {
      detach = attachPointer(host, store, wb);
      fireDown(host, 100, 10); // col 0 header
      const sel = store.getState().selection;
      expect(sel.range).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 1_048_575, c1: 0 });
    });

    it('drag across col headers extends to a column range', () => {
      detach = attachPointer(host, store, wb);
      fireDown(host, 100, 10); // col 0 header
      fireMove(host, 280, 10); // col 2 header (52+104+104=260, so 280 is in col 2)
      expect(store.getState().selection.range).toEqual({
        sheet: 0,
        r0: 0,
        c0: 0,
        r1: 1_048_575,
        c1: 2,
      });
      fireUp(host, 280, 10);
    });

    it('row-header click selects the entire row', () => {
      detach = attachPointer(host, store, wb);
      fireDown(host, 10, 40); // row 0 header
      expect(store.getState().selection.range).toEqual({
        sheet: 0,
        r0: 0,
        c0: 0,
        r1: 0,
        c1: 16_383,
      });
    });

    it('Ctrl/Cmd-click on row headers builds a disjoint row selection', () => {
      detach = attachPointer(host, store, wb);

      fireDown(host, 10, 460); // row 15 header
      fireDown(host, 10, 516, { ctrlKey: true }); // row 17 header
      fireDown(host, 10, 572, { ctrlKey: true }); // row 19 header

      const sel = store.getState().selection;
      expect(sel.range).toEqual({ sheet: 0, r0: 19, c0: 0, r1: 19, c1: 16383 });
      expect(sel.extraRanges).toEqual([
        { sheet: 0, r0: 15, c0: 0, r1: 15, c1: 16383 },
        { sheet: 0, r0: 17, c0: 0, r1: 17, c1: 16383 },
      ]);
    });

    it('Shift-click on row headers selects a contiguous row band', () => {
      detach = attachPointer(host, store, wb);

      fireDown(host, 10, 460); // row 15 header
      fireDown(host, 10, 572, { shiftKey: true }); // row 19 header

      const sel = store.getState().selection;
      expect(sel.range).toEqual({ sheet: 0, r0: 15, c0: 0, r1: 19, c1: 16383 });
      expect(sel.anchor).toEqual({ sheet: 0, row: 15, col: 0 });
      expect(sel.active).toEqual({ sheet: 0, row: 19, col: 0 });
    });

    it('Shift-click on column headers selects a contiguous column band', () => {
      detach = attachPointer(host, store, wb);

      fireDown(host, 100, 10); // col 0 header
      fireDown(host, 300, 10, { shiftKey: true }); // col 2 header

      const sel = store.getState().selection;
      expect(sel.range).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 1048575, c1: 2 });
      expect(sel.anchor).toEqual({ sheet: 0, row: 0, col: 0 });
      expect(sel.active).toEqual({ sheet: 0, row: 0, col: 2 });
    });

    it('corner click selects the entire sheet', () => {
      detach = attachPointer(host, store, wb);
      fireDown(host, 10, 10);
      const range = store.getState().selection.range;
      expect(range.r0).toBe(0);
      expect(range.c0).toBe(0);
      expect(range.r1).toBe(1_048_575);
      expect(range.c1).toBe(16_383);
    });
  });

  describe('resize', () => {
    it('col-resize drag updates colWidths', () => {
      detach = attachPointer(host, store, wb);
      fireDown(host, 154, 10); // col 0 right-edge
      // Drag the right edge to x=200; col 0 left edge is at 52, so width = 200-52 = 148.
      fireMove(host, 200, 10);
      expect(store.getState().layout.colWidths.get(0)).toBe(148);
      expect(host.style.cursor).toBe('col-resize');
      fireUp(host, 200, 10);
    });

    it('row-resize drag updates rowHeights', () => {
      detach = attachPointer(host, store, wb);
      fireDown(host, 10, 56); // row 0 bottom-edge (y=58 - slack 4 = 54)
      fireMove(host, 10, 100); // row 0 top edge is 30, so height = 100-30 = 70.
      expect(store.getState().layout.rowHeights.get(0)).toBe(70);
      expect(host.style.cursor).toBe('row-resize');
      fireUp(host, 10, 100);
    });

    it('col-resize push a single history entry on pointerup', () => {
      const history = new History();
      detach = attachPointer(host, store, wb, undefined, history);
      fireDown(host, 154, 10);
      fireMove(host, 200, 10);
      fireUp(host, 200, 10);
      expect(store.getState().layout.colWidths.get(0)).toBe(148);
      // Undo restores the original width (undefined → falls back to default).
      expect(history.undo()).toBe(true);
      expect(store.getState().layout.colWidths.get(0)).toBeUndefined();
    });

    it('double-click on col-resize zone autofits to content', () => {
      seed(store, wb, [{ row: 0, col: 0, value: 'a very long string that needs more width' }]);
      detach = attachPointer(host, store, wb);
      const e = fireDblClick(host, 154, 10);
      expect(e.defaultPrevented).toBe(true);
      // Width must be at least the minimum (48) and reflect the content.
      const w = store.getState().layout.colWidths.get(0) ?? 0;
      expect(w).toBeGreaterThanOrEqual(48);
    });

    it('double-click col autofit includes cells outside the viewport', () => {
      seed(store, wb, [
        {
          row: 100,
          col: 0,
          value: 'this offscreen value should still determine the autofit width',
        },
      ]);
      detach = attachPointer(host, store, wb);
      fireDblClick(host, 154, 10);
      const w = store.getState().layout.colWidths.get(0) ?? 0;
      expect(w).toBeGreaterThan(200);
    });

    it('double-click col autofit accounts for cell font size', () => {
      seed(store, wb, [{ row: 0, col: 0, value: 'large font text' }]);
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { fontSize: 30 });
      detach = attachPointer(host, store, wb);
      fireDblClick(host, 154, 10);
      const w = store.getState().layout.colWidths.get(0) ?? 0;
      expect(w).toBeGreaterThan(180);
    });

    it('double-click col autofit reserves room for filter header buttons', () => {
      seed(store, wb, [{ row: 0, col: 0, value: 'HeaderName' }]);
      detach = attachPointer(host, store, wb);
      fireDblClick(host, 154, 10);
      const plain = store.getState().layout.colWidths.get(0) ?? 0;
      detach();

      store.setState((s) => ({
        ...s,
        layout: { ...s.layout, colWidths: new Map() },
        ui: { ...s.ui, filterRange: { sheet: 0, r0: 0, c0: 0, r1: 5, c1: 0 } },
      }));
      detach = attachPointer(host, store, wb);
      fireDblClick(host, 154, 10);
      const filtered = store.getState().layout.colWidths.get(0) ?? 0;

      expect(filtered).toBeGreaterThanOrEqual(plain + 18);
    });

    it('double-click on row-resize zone resets to defaultRowHeight', () => {
      // Pre-set a custom height so the reset is observable. With height=80,
      // row 0 occupies y∈[30, 110) — the resize edge sits near y=110 (slack 4).
      mutators.setRowHeight(store, 0, 80);
      detach = attachPointer(host, store, wb);
      const e = fireDblClick(host, 10, 108);
      expect(e.defaultPrevented).toBe(true);
      expect(store.getState().layout.rowHeights.get(0)).toBe(28); // defaultRowHeight
    });

    it('double-click row autofit grows for multiline content', () => {
      seed(store, wb, [{ row: 0, col: 0, value: 'first line\nsecond line\nthird line' }]);
      detach = attachPointer(host, store, wb);
      const e = fireDblClick(host, 10, 56);
      expect(e.defaultPrevented).toBe(true);
      expect(store.getState().layout.rowHeights.get(0)).toBeGreaterThan(40);
    });

    it('double-click row autofit grows for wrapped text', () => {
      seed(store, wb, [
        { row: 0, col: 0, value: 'alpha beta gamma delta epsilon zeta eta theta iota kappa' },
      ]);
      mutators.setColWidth(store, 0, 70);
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { wrap: true });
      detach = attachPointer(host, store, wb);
      const e = fireDblClick(host, 10, 56);
      expect(e.defaultPrevented).toBe(true);
      expect(store.getState().layout.rowHeights.get(0)).toBeGreaterThan(40);
    });

    it('double-click autofit size changes are undoable', () => {
      const history = new History();
      seed(store, wb, [{ row: 0, col: 0, value: 'wide enough for undo history' }]);
      detach = attachPointer(host, store, wb, undefined, history);

      fireDblClick(host, 154, 10);
      const fitted = store.getState().layout.colWidths.get(0) ?? 0;
      expect(fitted).toBeGreaterThan(104);
      expect(history.undo()).toBe(true);
      expect(store.getState().layout.colWidths.get(0)).toBeUndefined();
      expect(history.redo()).toBe(true);
      expect(store.getState().layout.colWidths.get(0)).toBe(fitted);
    });

    it('double-click outside resize zones is a no-op', () => {
      detach = attachPointer(host, store, wb);
      const e = fireDblClick(host, 200, 70); // cell area
      expect(e.defaultPrevented).toBe(false);
    });
  });

  describe('cursor feedback', () => {
    it('hover over col-resize zone sets col-resize cursor', () => {
      detach = attachPointer(host, store, wb);
      fireMove(host, 154, 10);
      expect(host.style.cursor).toBe('col-resize');
    });

    it('hover over row-resize zone sets row-resize cursor', () => {
      detach = attachPointer(host, store, wb);
      fireMove(host, 10, 56);
      expect(host.style.cursor).toBe('row-resize');
    });

    it('hover over a normal cell clears the cursor', () => {
      detach = attachPointer(host, store, wb);
      // Prime the cursor with a resize hover, then move into a cell.
      fireMove(host, 154, 10);
      expect(host.style.cursor).toBe('col-resize');
      fireMove(host, 200, 70);
      expect(host.style.cursor).toBe('');
    });

    it('pointerleave clears the cursor when no drag is active', () => {
      detach = attachPointer(host, store, wb);
      host.style.cursor = 'col-resize';
      host.dispatchEvent(new PointerEvent('pointerleave', { pointerId: 1, bubbles: true }));
      expect(host.style.cursor).toBe('');
    });
  });

  describe('teardown', () => {
    it('detach removes listeners and clears the cursor', () => {
      detach = attachPointer(host, store, wb);
      detach();
      detach = (): void => {};
      // After detach, a pointerdown does nothing.
      fireDown(host, 200, 70);
      expect(store.getState().selection.active).toEqual({ sheet: 0, row: 0, col: 0 });
    });
  });

  describe('fill handle', () => {
    it('drag from the fill handle previews and commits a fill', async () => {
      const grid = await import('../../../src/render/grid.js');
      // Place a fake fill handle at the bottom-right of cell (0, 0).
      // Cell (0,0) is at x=52..156, y=30..58. Handle ~ (152, 54, 6, 6).
      const stub = grid.getFillHandleRect as unknown as ReturnType<typeof vi.fn>;
      stub.mockReturnValue({ x: 152, y: 54, w: 6, h: 6 });

      seed(store, wb, [{ row: 0, col: 0, value: 5 }]);
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });

      const onAfterCommit = vi.fn();
      detach = attachPointer(host, store, wb, onAfterCommit);

      fireDown(host, 155, 57); // grab handle
      expect(host.style.cursor).toBe('crosshair');
      // Drag downward into the cell column at row 2 (y∈[86,114)).
      fireMove(host, 100, 100);
      expect(store.getState().ui.fillPreview).not.toBeNull();
      fireUp(host, 100, 100);

      // After commit, the fill preview is cleared and the active range moved.
      expect(store.getState().ui.fillPreview).toBeNull();
      expect(onAfterCommit).toHaveBeenCalled();
      wb.recalc();
      // Fill replicated 5 down the column. The selection now covers (0..2, 0).
      expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'number', value: 5 });
      expect(wb.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({ kind: 'number', value: 5 });
      const sel = store.getState().selection.range;
      expect(sel).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 });

      // Reset for other tests.
      stub.mockReturnValue(null);
    });

    it('hovering over the fill handle sets the crosshair cursor', async () => {
      const grid = await import('../../../src/render/grid.js');
      const stub = grid.getFillHandleRect as unknown as ReturnType<typeof vi.fn>;
      stub.mockReturnValue({ x: 152, y: 54, w: 6, h: 6 });

      detach = attachPointer(host, store, wb);
      fireMove(host, 155, 57);
      expect(host.style.cursor).toBe('crosshair');

      stub.mockReturnValue(null);
    });

    it('double-click on the fill handle extends down to match the left neighbour', async () => {
      const grid = await import('../../../src/render/grid.js');
      const stub = grid.getFillHandleRect as unknown as ReturnType<typeof vi.fn>;
      stub.mockReturnValue({ x: 256, y: 54, w: 6, h: 6 });

      // Left column (col 0) has data in rows 0..3; source at (0, 1) carries
      // the value to be filled. spreadsheet rule: extend until the left neighbour
      // ends, i.e. through row 3.
      seed(store, wb, [
        { row: 0, col: 0, value: 'a' },
        { row: 1, col: 0, value: 'b' },
        { row: 2, col: 0, value: 'c' },
        { row: 3, col: 0, value: 'd' },
        { row: 0, col: 1, value: 10 },
      ]);
      mutators.setActive(store, { sheet: 0, row: 0, col: 1 });

      const onAfterCommit = vi.fn();
      detach = attachPointer(host, store, wb, onAfterCommit);

      const e = fireDblClick(host, 258, 56);
      expect(e.defaultPrevented).toBe(true);
      expect(onAfterCommit).toHaveBeenCalled();
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'number', value: 10 });
      expect(wb.getValue({ sheet: 0, row: 3, col: 1 })).toEqual({ kind: 'number', value: 10 });
      const sel = store.getState().selection.range;
      expect(sel).toEqual({ sheet: 0, r0: 0, c0: 1, r1: 3, c1: 1 });

      stub.mockReturnValue(null);
    });

    it('double-click on the fill handle is a no-op when neither neighbour has data', async () => {
      const grid = await import('../../../src/render/grid.js');
      const stub = grid.getFillHandleRect as unknown as ReturnType<typeof vi.fn>;
      stub.mockReturnValue({ x: 152, y: 54, w: 6, h: 6 });

      seed(store, wb, [{ row: 0, col: 0, value: 1 }]);
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });

      const onAfterCommit = vi.fn();
      detach = attachPointer(host, store, wb, onAfterCommit);
      fireDblClick(host, 155, 57);
      expect(onAfterCommit).not.toHaveBeenCalled();

      stub.mockReturnValue(null);
    });
  });

  describe('hyperlink follow', () => {
    const setHyperlink = (row: number, col: number, url: string): void => {
      store.setState((s) => {
        const formats = new Map(s.format.formats);
        formats.set(addrKey({ sheet: 0, row, col }), { hyperlink: url });
        return { ...s, format: { ...s.format, formats } };
      });
    };

    it('Cmd/Ctrl+click on a hyperlinked cell opens the URL and selects the cell', () => {
      setHyperlink(1, 1, 'https://example.com/');
      const openSpy = vi.spyOn(window, 'open').mockImplementation(() => null);
      detach = attachPointer(host, store, wb);
      const e = fireDown(host, 200, 70, { metaKey: true });
      fireUp(host, 200, 70);
      expect(e.defaultPrevented).toBe(true);
      expect(openSpy).toHaveBeenCalledWith('https://example.com/', '_blank', 'noopener,noreferrer');
      expect(store.getState().selection.active).toEqual({ sheet: 0, row: 1, col: 1 });
      openSpy.mockRestore();
    });

    it('plain click on a hyperlinked cell does not navigate', () => {
      setHyperlink(1, 1, 'https://example.com/');
      const openSpy = vi.spyOn(window, 'open').mockImplementation(() => null);
      detach = attachPointer(host, store, wb);
      fireDown(host, 200, 70);
      fireUp(host, 200, 70);
      expect(openSpy).not.toHaveBeenCalled();
      openSpy.mockRestore();
    });

    it('Cmd/Ctrl+click on a non-hyperlinked cell falls through to selection', () => {
      const openSpy = vi.spyOn(window, 'open').mockImplementation(() => null);
      detach = attachPointer(host, store, wb);
      fireDown(host, 200, 70, { metaKey: true });
      fireUp(host, 200, 70);
      expect(openSpy).not.toHaveBeenCalled();
      expect(store.getState().selection.active).toEqual({ sheet: 0, row: 1, col: 1 });
      openSpy.mockRestore();
    });

    it('rejects unsafe protocols (javascript:, data:)', () => {
      setHyperlink(1, 1, 'javascript:alert(1)');
      const openSpy = vi.spyOn(window, 'open').mockImplementation(() => null);
      detach = attachPointer(host, store, wb);
      fireDown(host, 200, 70, { ctrlKey: true });
      fireUp(host, 200, 70);
      expect(openSpy).not.toHaveBeenCalled();
      openSpy.mockRestore();
    });

    it('accepts mailto: links', () => {
      setHyperlink(1, 1, 'mailto:hi@example.com');
      const openSpy = vi.spyOn(window, 'open').mockImplementation(() => null);
      detach = attachPointer(host, store, wb);
      fireDown(host, 200, 70, { metaKey: true });
      fireUp(host, 200, 70);
      expect(openSpy).toHaveBeenCalledWith(
        'mailto:hi@example.com',
        '_blank',
        'noopener,noreferrer',
      );
      openSpy.mockRestore();
    });
  });

  describe('autofilter chevron', () => {
    const setFilterRange = (range: {
      sheet: number;
      r0: number;
      c0: number;
      r1: number;
      c1: number;
    }): void => {
      store.setState((s) => ({ ...s, ui: { ...s.ui, filterRange: range } }));
    };

    it('clicking the chevron in a filterRange column dispatches fc:openfilter', () => {
      setFilterRange({ sheet: 0, r0: 0, c0: 0, r1: 5, c1: 2 });
      detach = attachPointer(host, store, wb);
      const events: CustomEvent[] = [];
      host.addEventListener('fc:openfilter', (e) => events.push(e as CustomEvent));
      // Col 0 right edge x=156, resize slack 4 → 152..156 is col-resize.
      // Filter chevron sits at 138..152.
      const evt = fireDown(host, 145, 10);
      fireUp(host, 145, 10);
      expect(evt.defaultPrevented).toBe(true);
      expect(events).toHaveLength(1);
      const detail = events[0]?.detail as {
        col: number;
        range: { c0: number; c1: number };
      };
      expect(detail.col).toBe(0);
      expect(detail.range.c1).toBe(2);
    });

    it('clicking the chevron does not move the active cell', () => {
      setFilterRange({ sheet: 0, r0: 0, c0: 0, r1: 5, c1: 2 });
      mutators.setActive(store, { sheet: 0, row: 3, col: 3 });
      detach = attachPointer(host, store, wb);
      fireDown(host, 145, 10);
      fireUp(host, 145, 10);
      expect(store.getState().selection.active).toEqual({ sheet: 0, row: 3, col: 3 });
    });

    it('hovering over the chevron sets the pointer cursor', () => {
      setFilterRange({ sheet: 0, r0: 0, c0: 0, r1: 5, c1: 2 });
      detach = attachPointer(host, store, wb);
      fireMove(host, 145, 10);
      expect(host.style.cursor).toBe('pointer');
    });

    it('chevron is not surfaced for columns outside filterRange', () => {
      setFilterRange({ sheet: 0, r0: 0, c0: 0, r1: 5, c1: 0 }); // only col 0
      detach = attachPointer(host, store, wb);
      const events: CustomEvent[] = [];
      host.addEventListener('fc:openfilter', (e) => events.push(e as CustomEvent));
      // Col 1 mid — should select col 1 normally.
      fireDown(host, 200, 10);
      fireUp(host, 200, 10);
      expect(events).toHaveLength(0);
    });
  });

  describe('multi-range Ctrl/Cmd+click', () => {
    it('Cmd+click on a non-hyperlinked cell appends an extra range', () => {
      detach = attachPointer(host, store, wb);
      // First, plain click on (0,0).
      fireDown(host, 100, 40);
      fireUp(host, 100, 40);
      // Then Cmd+click on (2,2).
      fireDown(host, 300, 100, { metaKey: true });
      fireUp(host, 300, 100);
      const sel = store.getState().selection;
      // Active and primary range moved to (2,2).
      expect(sel.active).toEqual({ sheet: 0, row: 2, col: 2 });
      expect(sel.range).toEqual({ sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 });
      // The previous (0,0) cell got demoted to extraRanges.
      expect(sel.extraRanges).toHaveLength(1);
      expect(sel.extraRanges?.[0]).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    });

    it('Cmd+click does not start a drag-extend on the primary range', () => {
      detach = attachPointer(host, store, wb);
      fireDown(host, 100, 40);
      fireUp(host, 100, 40);
      fireDown(host, 300, 100, { metaKey: true });
      // Move without releasing — must not extend primary, since drag is none.
      fireMove(host, 300, 70); // move to (1,2)
      const sel = store.getState().selection;
      expect(sel.range).toEqual({ sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 });
      fireUp(host, 300, 70);
    });

    it('plain click after a multi-range clears extras', () => {
      detach = attachPointer(host, store, wb);
      fireDown(host, 100, 40);
      fireUp(host, 100, 40);
      fireDown(host, 300, 100, { metaKey: true });
      fireUp(host, 300, 100);
      // Plain click on (1,1).
      fireDown(host, 200, 70);
      fireUp(host, 200, 70);
      const sel = store.getState().selection;
      expect(sel.active).toEqual({ sheet: 0, row: 1, col: 1 });
      expect(sel.extraRanges).toEqual([]);
    });
  });
});
