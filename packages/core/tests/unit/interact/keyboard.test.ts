import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { WorkbookHandle, addrKey } from '../../../src/engine/workbook-handle.js';
import { attachKeyboard } from '../../../src/interact/keyboard.js';
import {
  type SpreadsheetStore,
  createSpreadsheetStore,
  mutators,
} from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const seed = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  cells: Array<{ row: number; col: number; value: number | string | boolean; formula?: string }>,
): void => {
  store.setState((s) => {
    const map = new Map(s.data.cells);
    for (const c of cells) {
      const addr = { sheet: 0, row: c.row, col: c.col };
      if (c.formula) {
        wb.setFormula(addr, c.formula);
        map.set(addrKey(addr), {
          value:
            typeof c.value === 'number'
              ? { kind: 'number', value: c.value }
              : typeof c.value === 'boolean'
                ? { kind: 'bool', value: c.value }
                : { kind: 'text', value: c.value },
          formula: c.formula,
        });
      } else if (typeof c.value === 'number') {
        wb.setNumber(addr, c.value);
        map.set(addrKey(addr), { value: { kind: 'number', value: c.value }, formula: null });
      } else if (typeof c.value === 'boolean') {
        wb.setBool(addr, c.value);
        map.set(addrKey(addr), { value: { kind: 'bool', value: c.value }, formula: null });
      } else {
        wb.setText(addr, c.value);
        map.set(addrKey(addr), { value: { kind: 'text', value: c.value }, formula: null });
      }
    }
    return { ...s, data: { ...s.data, cells: map } };
  });
  wb.recalc();
};

const fire = (
  host: HTMLElement,
  key: string,
  init: Partial<KeyboardEventInit> = {},
): KeyboardEvent => {
  const e = new KeyboardEvent('keydown', {
    key,
    bubbles: true,
    cancelable: true,
    ...init,
  });
  host.dispatchEvent(e);
  return e;
};

describe('attachKeyboard', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let onBeginEdit: ReturnType<typeof vi.fn>;
  let onClearActive: ReturnType<typeof vi.fn>;
  let onAfterHistory: ReturnType<typeof vi.fn>;
  let onGoTo: ReturnType<typeof vi.fn>;
  let detach: () => void;

  const setup = (history: History | null = null): void => {
    detach = attachKeyboard({
      host,
      store,
      wb,
      history,
      onBeginEdit,
      onClearActive,
      onAfterHistory,
      onGoTo,
    });
  };

  beforeEach(async () => {
    host = document.createElement('div');
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await newWb();
    onBeginEdit = vi.fn();
    onClearActive = vi.fn();
    onAfterHistory = vi.fn();
    onGoTo = vi.fn();
  });

  afterEach(() => {
    detach?.();
    document.body.innerHTML = '';
  });

  describe('navigation', () => {
    it('ArrowDown moves active down by one', () => {
      setup();
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      const e = fire(host, 'ArrowDown');
      expect(e.defaultPrevented).toBe(true);
      expect(store.getState().selection.active).toEqual({ sheet: 0, row: 1, col: 0 });
    });

    it('ArrowUp at row 0 clamps to row 0', () => {
      setup();
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      fire(host, 'ArrowUp');
      expect(store.getState().selection.active.row).toBe(0);
    });

    it('Shift+Arrow extends the selection range from the anchor; active follows the cursor', () => {
      setup();
      mutators.setActive(store, { sheet: 0, row: 5, col: 5 });
      fire(host, 'ArrowRight', { shiftKey: true });
      const sel = store.getState().selection;
      // anchor stays at (5, 5); active is the new edge.
      expect(sel.range).toEqual({ sheet: 0, r0: 5, c0: 5, r1: 5, c1: 6 });
      expect(sel.active).toEqual({ sheet: 0, row: 5, col: 6 });
      expect(sel.anchor).toEqual({ sheet: 0, row: 5, col: 5 });
    });

    it('Tab moves right; Shift+Tab moves left', () => {
      setup();
      mutators.setActive(store, { sheet: 0, row: 2, col: 5 });
      fire(host, 'Tab');
      expect(store.getState().selection.active).toEqual({ sheet: 0, row: 2, col: 6 });
      fire(host, 'Tab', { shiftKey: true });
      expect(store.getState().selection.active).toEqual({ sheet: 0, row: 2, col: 5 });
    });

    it('Shift+Tab does not extend the range — Tab is excluded from extend-range', () => {
      setup();
      mutators.setActive(store, { sheet: 0, row: 5, col: 5 });
      fire(host, 'Tab', { shiftKey: true });
      // active moves, range follows active (single-cell), not extended.
      const sel = store.getState().selection;
      expect(sel.active).toEqual({ sheet: 0, row: 5, col: 4 });
      expect(sel.range).toEqual({ sheet: 0, r0: 5, c0: 4, r1: 5, c1: 4 });
    });

    it('Home jumps to col 0; Ctrl+Home jumps to (0, 0)', () => {
      setup();
      mutators.setActive(store, { sheet: 0, row: 5, col: 7 });
      fire(host, 'Home');
      expect(store.getState().selection.active).toEqual({ sheet: 0, row: 5, col: 0 });
      mutators.setActive(store, { sheet: 0, row: 5, col: 7 });
      fire(host, 'Home', { ctrlKey: true });
      expect(store.getState().selection.active).toEqual({ sheet: 0, row: 0, col: 0 });
    });

    it('Ctrl+End jumps to last populated cell on the active sheet', () => {
      setup();
      seed(store, wb, [
        { row: 1, col: 1, value: 1 },
        { row: 5, col: 7, value: 2 },
        { row: 3, col: 9, value: 3 },
      ]);
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      fire(host, 'End', { ctrlKey: true });
      expect(store.getState().selection.active).toEqual({ sheet: 0, row: 5, col: 9 });
    });

    it('PageDown/PageUp jump by viewport.rowCount-1', () => {
      setup();
      // viewport defaults to rowCount: 40.
      mutators.setActive(store, { sheet: 0, row: 50, col: 0 });
      fire(host, 'PageDown');
      expect(store.getState().selection.active.row).toBe(50 + 39);
      fire(host, 'PageUp');
      expect(store.getState().selection.active.row).toBe(50);
    });
  });

  describe('Ctrl+Arrow jumpEdge', () => {
    it('empty origin + populated cells skips blanks to the next populated', () => {
      setup();
      seed(store, wb, [
        { row: 5, col: 0, value: 1 },
        { row: 10, col: 0, value: 2 },
      ]);
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      fire(host, 'ArrowDown', { ctrlKey: true });
      expect(store.getState().selection.active.row).toBe(5);
    });

    it('populated origin + populated neighbor walks to the end of the run', () => {
      setup();
      seed(store, wb, [
        { row: 0, col: 0, value: 1 },
        { row: 1, col: 0, value: 2 },
        { row: 2, col: 0, value: 3 },
        { row: 5, col: 0, value: 4 }, // gap at rows 3,4
      ]);
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      fire(host, 'ArrowDown', { ctrlKey: true });
      expect(store.getState().selection.active.row).toBe(2);
    });

    it('populated origin + empty neighbor jumps over the gap to the next populated', () => {
      setup();
      seed(store, wb, [
        { row: 0, col: 0, value: 1 },
        { row: 5, col: 0, value: 2 },
      ]);
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      fire(host, 'ArrowDown', { ctrlKey: true });
      expect(store.getState().selection.active.row).toBe(5);
    });

    it('no populated cells in the direction → goes to sheet edge', () => {
      setup();
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      fire(host, 'ArrowDown', { ctrlKey: true });
      expect(store.getState().selection.active.row).toBe(1_048_575);
    });

    it('already at sheet edge — Ctrl+Arrow is a no-op', () => {
      setup();
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      fire(host, 'ArrowUp', { ctrlKey: true });
      expect(store.getState().selection.active.row).toBe(0);
    });
  });

  describe('edit triggers', () => {
    it('Enter calls onBeginEdit("")', () => {
      setup();
      const e = fire(host, 'Enter');
      expect(onBeginEdit).toHaveBeenCalledWith('');
      expect(e.defaultPrevented).toBe(true);
    });

    it('printable character calls onBeginEdit with that key', () => {
      setup();
      fire(host, 'a');
      expect(onBeginEdit).toHaveBeenCalledWith('a');
    });

    it('printable + Alt or meta does not begin edit', () => {
      setup();
      fire(host, 'a', { altKey: true });
      fire(host, 'a', { metaKey: true });
      expect(onBeginEdit).not.toHaveBeenCalled();
    });

    it('F2 seeds the editor with the cell formula when present', () => {
      setup();
      seed(store, wb, [{ row: 0, col: 0, value: 5, formula: '=2+3' }]);
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      fire(host, 'F2');
      expect(onBeginEdit).toHaveBeenCalledWith('=2+3');
    });

    it('F2 seeds with formatted number when no formula', () => {
      setup();
      seed(store, wb, [{ row: 0, col: 0, value: 42 }]);
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      fire(host, 'F2');
      expect(onBeginEdit).toHaveBeenCalledWith('42');
    });

    it('F2 seeds with TRUE/FALSE for booleans', () => {
      setup();
      seed(store, wb, [{ row: 0, col: 0, value: true }]);
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      fire(host, 'F2');
      expect(onBeginEdit).toHaveBeenCalledWith('TRUE');
    });

    it('F2 seeds with raw text', () => {
      setup();
      seed(store, wb, [{ row: 0, col: 0, value: 'hello' }]);
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      fire(host, 'F2');
      expect(onBeginEdit).toHaveBeenCalledWith('hello');
    });

    it('F2 seeds with empty string when the cell is blank', () => {
      setup();
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      fire(host, 'F2');
      expect(onBeginEdit).toHaveBeenCalledWith('');
    });
  });

  describe('Backspace / Delete', () => {
    it('clears all populated cells in the selection range and notifies onClearActive', () => {
      setup();
      seed(store, wb, [
        { row: 0, col: 0, value: 1 },
        { row: 0, col: 1, value: 2 },
        { row: 5, col: 5, value: 99 }, // outside range — preserved
      ]);
      store.setState((s) => ({
        ...s,
        selection: {
          ...s.selection,
          active: { sheet: 0, row: 0, col: 0 },
          anchor: { sheet: 0, row: 0, col: 0 },
          range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
        },
      }));
      const e = fire(host, 'Delete');
      expect(e.defaultPrevented).toBe(true);
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');
      expect(wb.getValue({ sheet: 0, row: 0, col: 1 }).kind).toBe('blank');
      expect(wb.getValue({ sheet: 0, row: 5, col: 5 })).toEqual({ kind: 'number', value: 99 });
      expect(onClearActive).toHaveBeenCalled();
    });

    it('Backspace also clears the range', () => {
      setup();
      seed(store, wb, [{ row: 0, col: 0, value: 1 }]);
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      fire(host, 'Backspace');
      wb.recalc();
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');
    });
  });

  describe('Go To', () => {
    it('F5 calls onGoTo', () => {
      setup();
      const e = fire(host, 'F5');
      expect(onGoTo).toHaveBeenCalled();
      expect(e.defaultPrevented).toBe(true);
    });

    it('Ctrl+G calls onGoTo (case-insensitive)', () => {
      setup();
      fire(host, 'g', { ctrlKey: true });
      fire(host, 'G', { metaKey: true });
      expect(onGoTo).toHaveBeenCalledTimes(2);
    });
  });

  describe('undo / redo', () => {
    it('Cmd+Z routes through the workbook when no history is provided', () => {
      const undo = vi.spyOn(wb, 'undo').mockReturnValue(true);
      setup();
      const e = fire(host, 'z', { metaKey: true });
      expect(undo).toHaveBeenCalled();
      expect(onAfterHistory).toHaveBeenCalled();
      expect(e.defaultPrevented).toBe(true);
      undo.mockRestore();
    });

    it('Cmd+Shift+Z routes through redo', () => {
      const redo = vi.spyOn(wb, 'redo').mockReturnValue(true);
      setup();
      fire(host, 'z', { metaKey: true, shiftKey: true });
      expect(redo).toHaveBeenCalled();
      redo.mockRestore();
    });

    it('Cmd+Y routes through redo', () => {
      const redo = vi.spyOn(wb, 'redo').mockReturnValue(true);
      setup();
      fire(host, 'y', { metaKey: true });
      expect(redo).toHaveBeenCalled();
      redo.mockRestore();
    });

    it('uses the shared History instance when provided', () => {
      const history = new History();
      const undo = vi.spyOn(history, 'undo').mockReturnValue(true);
      const redo = vi.spyOn(history, 'redo').mockReturnValue(true);
      setup(history);
      fire(host, 'z', { metaKey: true });
      fire(host, 'z', { metaKey: true, shiftKey: true });
      fire(host, 'y', { metaKey: true });
      expect(undo).toHaveBeenCalledTimes(1);
      expect(redo).toHaveBeenCalledTimes(2);
    });

    it('does not call onAfterHistory when undo/redo returns false', () => {
      const undo = vi.spyOn(wb, 'undo').mockReturnValue(false);
      setup();
      fire(host, 'z', { metaKey: true });
      expect(onAfterHistory).not.toHaveBeenCalled();
      undo.mockRestore();
    });
  });

  describe('inactive', () => {
    it('does nothing while the editor is in non-idle mode', () => {
      setup();
      store.setState((s) => ({ ...s, ui: { ...s.ui, editor: { kind: 'enter', raw: '' } } }));
      mutators.setActive(store, { sheet: 0, row: 5, col: 5 });
      fire(host, 'ArrowDown');
      expect(store.getState().selection.active).toEqual({ sheet: 0, row: 5, col: 5 });
      expect(onBeginEdit).not.toHaveBeenCalled();
    });

    it('Escape with editor idle is a no-op', () => {
      setup();
      const e = fire(host, 'Escape');
      expect(e.defaultPrevented).toBe(false);
    });

    it('unrecognized keys do not preventDefault', () => {
      setup();
      const e = fire(host, 'Insert');
      expect(e.defaultPrevented).toBe(false);
    });

    it('detach removes the listener so subsequent keys are ignored', () => {
      setup();
      detach();
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      fire(host, 'ArrowDown');
      expect(store.getState().selection.active).toEqual({ sheet: 0, row: 0, col: 0 });
    });
  });

  describe('outline group shortcuts', () => {
    it('Alt+Shift+Right groups rows when selection spans multiple rows', () => {
      setup();
      // Select rows 1..3.
      mutators.setRange(store, { sheet: 0, r0: 1, c0: 0, r1: 3, c1: 0 });
      const e = fire(host, 'ArrowRight', { altKey: true, shiftKey: true });
      expect(e.defaultPrevented).toBe(true);
      const layout = store.getState().layout;
      expect(layout.outlineRows.get(2)).toBe(1);
      expect(layout.outlineRows.size).toBe(3);
    });

    it('Alt+Shift+Left ungroups the same selection', () => {
      setup();
      mutators.setRange(store, { sheet: 0, r0: 1, c0: 0, r1: 3, c1: 0 });
      fire(host, 'ArrowRight', { altKey: true, shiftKey: true });
      fire(host, 'ArrowLeft', { altKey: true, shiftKey: true });
      expect(store.getState().layout.outlineRows.size).toBe(0);
    });

    it('groups columns when the selection spans more cols than rows', () => {
      setup();
      mutators.setRange(store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 4 });
      fire(host, 'ArrowRight', { altKey: true, shiftKey: true });
      const layout = store.getState().layout;
      expect(layout.outlineCols.get(2)).toBe(1);
      expect(layout.outlineRows.size).toBe(0);
    });
  });
});
