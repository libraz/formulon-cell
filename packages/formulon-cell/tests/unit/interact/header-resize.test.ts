import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { ja } from '../../../src/i18n/strings.js';
import { attachPointer } from '../../../src/interact/pointer.js';
import {
  colLeadingEdge,
  colWidth,
  layoutForView,
  rowHeight,
  rowTopEdge,
} from '../../../src/render/geometry.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const stubPointerCapture = (host: HTMLElement): void => {
  const captured = new Set<number>();
  Object.assign(host, {
    setPointerCapture: (id: number) => captured.add(id),
    releasePointerCapture: (id: number) => captured.delete(id),
    hasPointerCapture: (id: number) => captured.has(id),
  });
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

const dblclick = (host: HTMLElement, x: number, y: number): void => {
  host.dispatchEvent(
    new MouseEvent('dblclick', { clientX: x, clientY: y, bubbles: true, cancelable: true }),
  );
};

// Mac profile: 1 screen px == 1 pt, so pt offsets in the facts are px offsets here.
describe('header border resize', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let history: History;
  let detach: (() => void) | undefined;

  const geo = () => {
    const s = store.getState();
    return { s, layout: layoutForView(s) };
  };
  /** Screen x of column `c`'s trailing edge. */
  const colRight = (c: number): number => {
    const { s, layout } = geo();
    return colLeadingEdge(layout, s.viewport, c) + colWidth(layout, c, s.viewport);
  };
  const colLeft = (c: number): number => {
    const { s, layout } = geo();
    return colLeadingEdge(layout, s.viewport, c);
  };
  const rowBottom = (r: number): number => {
    const { s, layout } = geo();
    return rowTopEdge(layout, s.viewport, r) + rowHeight(layout, r, s.viewport);
  };
  const rowTop = (r: number): number => {
    const { s, layout } = geo();
    return rowTopEdge(layout, s.viewport, r);
  };
  const widths = (...cols: number[]) => cols.map((c) => store.getState().layout.colWidths.get(c));
  const heights = (...rows: number[]) => rows.map((r) => store.getState().layout.rowHeights.get(r));
  const tip = (): HTMLElement | null => document.querySelector('.fc-resize-tooltip');

  beforeEach(async () => {
    host = document.createElement('div');
    stubPointerCapture(host);
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        screenScale: 0.75,
        defaultRowHeight: 80 / 3,
        defaultColWidth: 100,
        headerColWidth: 40,
        headerRowHeight: 24,
      },
    }));
    wb = await WorkbookHandle.createDefault({ preferStub: true });
    history = new History();
    detach = attachPointer(
      host,
      store,
      wb,
      undefined,
      history,
      undefined,
      undefined,
      undefined,
      undefined,
      () => ja.resizeTip,
    );
  });

  afterEach(() => {
    detach?.();
    detach = undefined;
    document.body.innerHTML = '';
  });

  const dragColBorder = (c: number, dx: number, offset = 0): void => {
    const x = colRight(c) + offset;
    fire(host, 'pointerdown', x, 10);
    fire(host, 'pointermove', x + dx, 10);
    fire(host, 'pointerup', x + dx, 10);
  };
  const dragRowBorder = (r: number, dy: number, offset = 0): void => {
    const y = rowBottom(r) + offset;
    fire(host, 'pointerdown', 10, y);
    fire(host, 'pointermove', 10, y + dy);
    fire(host, 'pointerup', 10, y + dy);
  };

  describe('multi-select drag', () => {
    it('resizes every selected whole column from any selected border', () => {
      mutators.selectCols(store, 1, 3);
      dragColBorder(3, 40);
      const w = (75 + 40) / 0.75;
      expect(widths(0, 1, 2, 3, 4)).toEqual([undefined, w, w, w, undefined]);
      expect(store.getState().selection.range).toMatchObject({ c0: 1, c1: 3 });
      expect(store.getState().selection.active).toMatchObject({ row: 0, col: 1 });

      dragColBorder(1, 20);
      const w2 = (75 + 40 + 20) / 0.75;
      expect(widths(1, 2, 3)).toEqual([w2, w2, w2]);
    });

    it('resizes only the hit column outside the selection and keeps the selection', () => {
      mutators.selectCols(store, 1, 3);
      dragColBorder(5, 40);
      expect(widths(1, 2, 3, 5)).toEqual([undefined, undefined, undefined, (75 + 40) / 0.75]);
      expect(store.getState().selection.range).toMatchObject({ c0: 1, c1: 3 });
    });

    it('resizes only the dragged column for a cell-range selection', () => {
      mutators.setRange(store, { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 3 });
      dragColBorder(3, 40);
      expect(widths(1, 2, 3)).toEqual([undefined, undefined, (75 + 40) / 0.75]);
    });

    it('resizes every selected whole row', () => {
      mutators.selectRows(store, 1, 3);
      dragRowBorder(3, 20);
      const h = (20 + 20) / 0.75;
      expect(heights(0, 1, 2, 3, 4)).toEqual([undefined, h, h, h, undefined]);
      dragRowBorder(1, -4);
      const h2 = (40 - 4) / 0.75;
      expect(heights(1, 2, 3)).toEqual([h2, h2, h2]);
    });

    it('resizes every area of a multi-area whole-column selection', () => {
      mutators.selectCols(store, 1, 1);
      store.setState((s) => ({
        ...s,
        selection: {
          ...s.selection,
          extraRanges: [{ sheet: 0, r0: 0, c0: 4, r1: 1048575, c1: 5 }],
        },
      }));
      dragColBorder(5, 25);
      const w = (75 + 25) / 0.75;
      expect(widths(0, 1, 4, 5, 6)).toEqual([undefined, w, w, w, undefined]);
    });

    it('undoes a multi-track drag in one step', () => {
      mutators.selectCols(store, 1, 3);
      dragColBorder(3, 40);
      expect(history.undo()).toBe(true);
      expect(widths(1, 2, 3)).toEqual([undefined, undefined, undefined]);
      expect(history.canUndo()).toBe(false);
    });
  });

  describe('hit zone', () => {
    it('grabs a column border within 6 pt either side and resizes 1:1 from the grab point', () => {
      for (const o of [-6, 0, 6]) {
        store.setState((s) => ({ ...s, layout: { ...s.layout, colWidths: new Map() } }));
        dragColBorder(3, 20, o);
        expect(widths(3)).toEqual([(75 + 20) / 0.75]);
      }
    });

    it('selects columns rather than resizing at +/-7 pt', () => {
      fire(host, 'pointerdown', colRight(3) - 7, 10);
      expect(store.getState().selection.range).toMatchObject({ c0: 3, c1: 3 });
      fire(host, 'pointerup', colRight(3) - 7, 10);
      fire(host, 'pointerdown', colRight(3) + 7, 10);
      expect(store.getState().selection.range).toMatchObject({ c0: 4, c1: 4 });
    });

    it('grabs a row border within 3 pt and selects beyond', () => {
      for (const o of [-3, 3]) {
        store.setState((s) => ({ ...s, layout: { ...s.layout, rowHeights: new Map() } }));
        dragRowBorder(3, 20, o);
        expect(heights(3)).toEqual([(20 + 20) / 0.75]);
      }
      fire(host, 'pointerdown', 10, rowBottom(3) - 4);
      expect(store.getState().selection.range).toMatchObject({ r0: 3, r1: 3 });
    });
  });

  describe('size readout', () => {
    it('shows the dragged row in points with the pixel count', () => {
      mutators.selectRows(store, 1, 3);
      const y = rowBottom(3);
      fire(host, 'pointerdown', 10, y);
      fire(host, 'pointermove', 10, y + 20);
      expect(tip()?.textContent).toBe('高さ: 40.00 (40 ピクセル)');
      expect(tip()?.hidden).toBe(false);
      fire(host, 'pointerup', 10, y + 20);
      expect(tip()?.hidden).toBe(true);
    });

    it('shows the column size and follows the pointer x', () => {
      const x = colRight(2);
      fire(host, 'pointerdown', x, 10);
      fire(host, 'pointermove', x + 40, 10);
      expect(tip()?.textContent).toMatch(/^幅: \d+\.\d{2} \(115 ピクセル\)$/);
      expect(tip()?.style.left).toBe(`${x + 40}px`);
      fire(host, 'pointerup', x + 40, 10);
    });

    it('omits the pixel count at zero', () => {
      const x = colRight(2);
      fire(host, 'pointerdown', x, 10);
      fire(host, 'pointermove', x - 75, 10);
      expect(tip()?.textContent).toBe('幅: 0.00');
      fire(host, 'pointerup', x - 75, 10);
    });
  });

  describe('drag to zero', () => {
    it('hides the column at its leading edge and keeps the previous one', () => {
      const x = colRight(3);
      fire(host, 'pointerdown', x, 10);
      fire(host, 'pointermove', colLeft(3), 10);
      expect(store.getState().layout.hiddenCols.has(3)).toBe(true);
      fire(host, 'pointerup', colLeft(3), 10);
      expect(widths(2)).toEqual([undefined]);
    });

    it('shrinks the previous column by twice the overshoot', () => {
      const x = colRight(3);
      const left = colLeft(3);
      fire(host, 'pointerdown', x, 10);
      fire(host, 'pointermove', left - 4, 10);
      const s = store.getState();
      expect(s.layout.hiddenCols.has(3)).toBe(true);
      expect(s.layout.colWidths.get(2)).toBe((75 - 8) / 0.75);
      fire(host, 'pointerup', left - 4, 10);
    });

    it('hides exhausted columns and carries the remainder back', () => {
      const x = colRight(3);
      const left = colLeft(3);
      fire(host, 'pointerdown', x, 10);
      // 2 * 43 = 86 = column C (75) + 11 off column B.
      fire(host, 'pointermove', left - 43, 10);
      const s = store.getState();
      expect([...s.layout.hiddenCols].sort()).toEqual([2, 3]);
      expect(s.layout.colWidths.get(1)).toBe((75 - 11) / 0.75);
      fire(host, 'pointerup', left - 43, 10);
    });

    it('hides a row dragged to its top edge', () => {
      const y = rowBottom(3);
      fire(host, 'pointerdown', 10, y);
      fire(host, 'pointermove', 10, rowTop(3));
      expect(store.getState().layout.hiddenRows.has(3)).toBe(true);
      fire(host, 'pointerup', 10, rowTop(3));
    });

    it('restores the layout on pointercancel', () => {
      const x = colRight(3);
      fire(host, 'pointerdown', x, 10);
      fire(host, 'pointermove', colLeft(3) - 20, 10);
      host.dispatchEvent(new PointerEvent('pointercancel', { pointerId: 1, bubbles: true }));
      const s = store.getState();
      expect(s.layout.hiddenCols.size).toBe(0);
      expect(s.layout.colWidths.size).toBe(0);
      expect(history.canUndo()).toBe(false);
      expect(tip()?.hidden).toBe(true);
    });
  });

  describe('double-click autofit', () => {
    const seedText = (): void => {
      const cells = [
        { row: 1, col: 1, text: 'short' },
        { row: 1, col: 2, text: 'a considerably longer text value' },
        { row: 2, col: 3, text: 'zzzzzzzzzzzzzzzzzzzzzzzz' },
      ];
      store.setState((s) => {
        const map = new Map(s.data.cells);
        for (const c of cells) {
          const addr = { sheet: 0, row: c.row, col: c.col };
          wb.setText(addr, c.text);
          map.set(addrKey(addr), { value: { kind: 'text', value: c.text }, formula: null });
        }
        return { ...s, data: { ...s.data, cells: map } };
      });
    };

    it('fits each selected column to its own content from any selected border', () => {
      seedText();
      mutators.selectCols(store, 1, 3);
      dblclick(host, colRight(2), 10);
      const [b, c, d, a, e] = widths(1, 2, 3, 0, 4);
      expect(b).toBeDefined();
      expect(c).toBeGreaterThan(b ?? 0);
      expect(c).toBeGreaterThan(d ?? 0);
      expect(d).toBeGreaterThan(b ?? 0);
      expect([a, e]).toEqual([undefined, undefined]);
      expect(store.getState().selection.range).toMatchObject({ c0: 1, c1: 3 });
      expect(history.undo()).toBe(true);
      expect(widths(1, 2, 3)).toEqual([undefined, undefined, undefined]);
    });

    it('fits only the hit column outside the selection', () => {
      seedText();
      mutators.selectCols(store, 1, 3);
      dblclick(host, colRight(4), 10);
      expect(widths(1, 2, 3)).toEqual([undefined, undefined, undefined]);
      expect(widths(4)[0]).toBeDefined();
    });

    it('fits each selected row to its own content', () => {
      mutators.selectRows(store, 1, 3);
      store.setState((s) => ({
        ...s,
        layout: {
          ...s.layout,
          rowHeights: new Map([
            [1, 90],
            [2, 90],
            [3, 90],
          ]),
        },
      }));
      dblclick(host, 10, rowBottom(3));
      const [r1, r2, r3] = heights(1, 2, 3);
      expect(r1).toBeLessThan(90);
      expect(r2).toBeLessThan(90);
      expect(r3).toBeLessThan(90);
    });
  });

  describe('hidden boundary', () => {
    beforeEach(() => {
      store.setState((s) => ({
        ...s,
        layout: { ...s.layout, hiddenCols: new Set([3]), hiddenRows: new Set([3]) },
      }));
    });

    it('double-click on the right of the boundary unhides at the autofit width', () => {
      const x = colRight(2);
      dblclick(host, x + 2, 10);
      const s = store.getState();
      expect(s.layout.hiddenCols.size).toBe(0);
      expect(s.layout.colWidths.get(3)).toBeDefined();
      expect(history.undo()).toBe(true);
      expect(store.getState().layout.hiddenCols.has(3)).toBe(true);
    });

    it('double-click on the left of the boundary fits the visible neighbour only', () => {
      dblclick(host, colRight(2) - 2, 10);
      const s = store.getState();
      expect(s.layout.hiddenCols.has(3)).toBe(true);
      expect(s.layout.colWidths.get(2)).toBeDefined();
    });

    it('dragging right from the boundary unhides at the drag distance', () => {
      const x = colRight(2);
      fire(host, 'pointerdown', x + 2, 10);
      fire(host, 'pointermove', x + 20, 10);
      fire(host, 'pointerup', x + 20, 10);
      const s = store.getState();
      expect(s.layout.hiddenCols.size).toBe(0);
      expect(s.layout.colWidths.get(3)).toBe(18 / 0.75);
      expect(s.layout.colWidths.get(2)).toBeUndefined();
    });

    it('dragging from the left of the boundary resizes the visible neighbour', () => {
      const x = colRight(2);
      fire(host, 'pointerdown', x - 3, 10);
      fire(host, 'pointermove', x - 3 + 20, 10);
      fire(host, 'pointerup', x - 3 + 20, 10);
      const s = store.getState();
      expect(s.layout.hiddenCols.has(3)).toBe(true);
      expect(s.layout.colWidths.get(2)).toBe((75 + 20) / 0.75);
    });

    it('unhides a hidden row by dragging down from its boundary', () => {
      const y = rowBottom(2);
      fire(host, 'pointerdown', 10, y + 1);
      fire(host, 'pointermove', 10, y + 21);
      fire(host, 'pointerup', 10, y + 21);
      const s = store.getState();
      expect(s.layout.hiddenRows.size).toBe(0);
      expect(s.layout.rowHeights.get(3)).toBe(20 / 0.75);
    });
  });

  describe('header selection drag', () => {
    it('keeps the press header as the active cell on a reverse drag', () => {
      fire(host, 'pointerdown', (colLeft(3) + colRight(3)) / 2, 10);
      fire(host, 'pointermove', (colLeft(1) + colRight(1)) / 2, 10);
      const sel = store.getState().selection;
      expect(sel.range).toMatchObject({ c0: 1, c1: 3 });
      expect(sel.active).toMatchObject({ row: 0, col: 3 });
      fire(host, 'pointerup', 0, 10);
    });

    it('extends to the last visible column when dragged past the window end', () => {
      const s0 = store.getState();
      const last = s0.viewport.colStart + s0.viewport.colCount - 1;
      fire(host, 'pointerdown', (colLeft(1) + colRight(1)) / 2, 10);
      fire(host, 'pointermove', 100000, 10);
      expect(store.getState().selection.range).toMatchObject({ c0: 1, c1: last });
      fire(host, 'pointerup', 100000, 10);
    });
  });
});
