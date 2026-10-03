import { afterEach, beforeEach, describe, expect, it, type Mock, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import type { InteractionPolicy } from '../../../src/commands/interaction-policy.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { InlineEditor } from '../../../src/interact/editor.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const flushRaf = async (): Promise<void> => {
  // The editor focuses on requestAnimationFrame; flush a microtask + a frame.
  await new Promise<void>((r) => requestAnimationFrame(() => r()));
};

describe('InlineEditor', () => {
  let host: HTMLElement;
  let grid: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let onAfterCommit: Mock<() => void>;
  let editor: InlineEditor;

  beforeEach(async () => {
    host = document.createElement('div');
    grid = document.createElement('div');
    host.appendChild(grid);
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    // Pin layout so hardcoded pixel coordinates below stay meaningful even
    // when the production defaults shift.
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
    onAfterCommit = vi.fn<() => void>();
    editor = new InlineEditor({ host, grid, store, wb, onAfterCommit });
  });

  afterEach(() => {
    if (editor.isActive()) editor.cancel();
    document.body.innerHTML = '';
  });

  it('begin appends an input to the grid and switches editor mode to enter', () => {
    mutators.setActive(store, { sheet: 0, row: 2, col: 3 });
    editor.begin('hi');
    expect(editor.isActive()).toBe(true);
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement | null;
    expect(input).not.toBeNull();
    expect(input?.value).toBe('hi');
    const mode = store.getState().ui.editor;
    expect(mode.kind).toBe('enter');
    expect(mode.kind === 'enter' && mode.raw).toBe('hi');
  });

  it('begin cancels copy mode so the marquee does not survive a cell edit', () => {
    mutators.setActive(store, { sheet: 0, row: 2, col: 3 });
    mutators.setCopyRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1048575, c1: 0 });
    editor.begin('x');
    expect(store.getState().ui.copyRange).toBeNull();
    expect(store.getState().ui.copyRanges).toBeNull();
  });

  it('begin positions the input over the active cell using cellRect', () => {
    mutators.setActive(store, { sheet: 0, row: 1, col: 1 });
    editor.begin('');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    // Test fixture pins layout to (headerCol=52, headerRow=30,
    // defaultColW=104, defaultRowH=28). Cell (1, 1): x=52+104=156,
    // y=30+28=58, w=104, h=28.
    expect(input.style.left).toBe('156px');
    expect(input.style.top).toBe('58px');
    expect(input.style.width).toBe('104px');
    expect(input.style.height).toBe('28px');
  });

  it('follows its cell when the sheet scrolls mid-edit', () => {
    mutators.setActive(store, { sheet: 0, row: 5, col: 1 });
    editor.begin('hi');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    expect(input.style.top).toBe(`${30 + 5 * 28}px`);

    mutators.scrollBy(store, 2, 0);
    // Two rows scrolled away → the cell (and the editor) climb two rows.
    expect(input.style.top).toBe(`${30 + 3 * 28}px`);
    expect(input.style.left).toBe('156px');

    mutators.scrollBy(store, 0, 1);
    expect(input.style.left).toBe(`${52 + 0 * 104}px`);
  });

  it('parks the editor outside the grid once its cell scrolls off the leading edge', () => {
    mutators.setActive(store, { sheet: 0, row: 1, col: 1 });
    editor.begin('hi');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;

    mutators.scrollBy(store, 4, 0);
    // Negative offsets: the grid clips the editor instead of letting it hang
    // over the column headers, and it keeps its size so it can come back.
    expect(Number.parseFloat(input.style.top)).toBeLessThan(0);
    expect(Number.parseFloat(input.style.left)).toBeLessThan(0);
    expect(input.style.height).toBe('28px');
    // The edit is still live — value and editor mode survive the trip.
    expect(input.value).toBe('hi');
    expect(editor.isActive()).toBe(true);

    mutators.scrollBy(store, -4, 0);
    expect(input.style.top).toBe('58px');
    expect(input.style.left).toBe('156px');
  });

  it('keeps the editor on a frozen cell in place while the body scrolls', () => {
    mutators.setFreezePanes(store, 2, 2);
    mutators.setActive(store, { sheet: 0, row: 1, col: 1 });
    editor.begin('hi');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    const top = input.style.top;
    const left = input.style.left;

    mutators.scrollBy(store, 10, 10);
    expect(input.style.top).toBe(top);
    expect(input.style.left).toBe(left);
  });

  it('takes the cell fill and text color so the edit reads as in-cell', () => {
    const addr = { sheet: 0, row: 1, col: 1 };
    mutators.setCellFormat(store, addr, { fill: '#ffeeaa', color: '#003366' });
    mutators.setActive(store, addr);
    editor.begin('hi');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    expect(input.style.background).toBe('#ffeeaa');
    expect(input.style.color).toBe('#003366');
  });

  it('repaints when a format is applied mid-edit', () => {
    const addr = { sheet: 0, row: 1, col: 1 };
    mutators.setActive(store, addr);
    editor.begin('hi');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    expect(input.style.background).toBe('');

    mutators.setCellFormat(store, addr, { fill: '#112233' });
    expect(input.style.background).toBe('#112233');
  });

  it('stops tracking the store once the edit ends', () => {
    mutators.setActive(store, { sheet: 0, row: 5, col: 1 });
    editor.begin('hi');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    editor.cancel();
    const top = input.style.top;
    mutators.scrollBy(store, 2, 0);
    expect(input.style.top).toBe(top);
  });

  it('mirrors spreadsheet default edit alignment for typed values', () => {
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });

    editor.begin('42');
    let input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    expect(input.style.textAlign).toBe('right');

    input.value = 'TRUE';
    input.dispatchEvent(new Event('input'));
    expect(input.style.textAlign).toBe('center');

    input.value = '=A1';
    input.dispatchEvent(new Event('input'));
    expect(input.style.textAlign).toBe('left');

    editor.cancel();
    editor.begin('text');
    input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    expect(input.style.textAlign).toBe('left');
  });

  it('honors explicit cell alignment while editing', () => {
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { align: 'center' });

    editor.begin('42');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    expect(input.style.textAlign).toBe('center');

    input.value = '=A1';
    input.dispatchEvent(new Event('input'));
    expect(input.style.textAlign).toBe('center');
  });

  it('cancel removes the input and resets editor mode', () => {
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    editor.begin('x');
    editor.cancel();
    expect(editor.isActive()).toBe(false);
    expect(grid.querySelector('textarea.fc-host__editor')).toBeNull();
    expect(store.getState().ui.editor.kind).toBe('idle');
  });

  it('cancel without an active begin is a no-op', () => {
    editor.cancel();
    expect(editor.isActive()).toBe(false);
    expect(store.getState().ui.editor.kind).toBe('idle');
  });

  it('commit writes the value via writeInput and advances the active cell down by default', () => {
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    editor.begin('');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    input.value = '42';
    editor.commit();
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 42 });
    expect(onAfterCommit).toHaveBeenCalledTimes(1);
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 1, col: 0 });
    expect(editor.isActive()).toBe(false);
  });

  it('commit recomputes formulas that depend on the edited cell', () => {
    const a1 = { sheet: 0, row: 0, col: 0 };
    const b1 = { sheet: 0, row: 0, col: 1 };
    wb.setNumber(a1, 20);
    wb.setFormula(b1, '=A1*2');
    expect(wb.getValue(b1)).toEqual({ kind: 'number', value: 40 });

    mutators.setActive(store, a1);
    editor.begin('');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    input.value = '25';
    editor.commit();

    // No explicit recalc — committing the edit must be enough.
    expect(wb.getValue(b1)).toEqual({ kind: 'number', value: 50 });
  });

  it('commit applies pending empty-cell format to the typed cell and then clears it', () => {
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(store, addr);
    mutators.setPendingFormat(store, {
      addr,
      format: { align: 'center', bold: true, numFmt: { kind: 'text' } },
    });
    editor.begin('');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    input.value = '00123';

    editor.commit('none');
    wb.recalc();

    expect(wb.getValue(addr)).toEqual({ kind: 'text', value: '00123' });
    expect(store.getState().format.formats.get('0:0:0')).toMatchObject({
      align: 'center',
      bold: true,
      numFmt: { kind: 'text' },
    });
    expect(store.getState().ui.pendingFormat).toBeNull();
  });

  it('cancel clears pending empty-cell format without writing it to the cell', () => {
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(store, addr);
    mutators.setPendingFormat(store, { addr, format: { bold: true } });
    editor.begin('');

    editor.cancel();

    expect(store.getState().format.formats.get('0:0:0')).toBeUndefined();
    expect(store.getState().ui.pendingFormat).toBeNull();
  });

  it('commit preserves numeric-looking input as text for cells formatted as Text', () => {
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(store, addr);
    mutators.setCellFormat(store, addr, { numFmt: { kind: 'text' } });
    editor.begin('');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    input.value = '00123';

    editor.commit('none');
    wb.recalc();

    expect(wb.getValue(addr)).toEqual({ kind: 'text', value: '00123' });
    expect(store.getState().selection.active).toEqual(addr);
  });

  it('commit("right") advances column instead of row', () => {
    mutators.setActive(store, { sheet: 0, row: 4, col: 4 });
    editor.begin('');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    input.value = 'foo';
    editor.commit('right');
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 4, col: 5 });
  });

  it('commit("none") does not move the active cell', () => {
    mutators.setActive(store, { sheet: 0, row: 4, col: 4 });
    editor.begin('');
    editor.commit('none');
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 4, col: 4 });
  });

  it('commit without begin is a no-op', () => {
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    editor.commit();
    expect(onAfterCommit).not.toHaveBeenCalled();
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 0, col: 0 });
  });

  it('suspends and restores the displayed raw edit without a workbook write or focus steal', async () => {
    const realWorkbook = await WorkbookHandle.createDefault();
    const localHost = document.createElement('div');
    const localGrid = document.createElement('div');
    localHost.appendChild(localGrid);
    document.body.appendChild(localHost);
    const localStore = createSpreadsheetStore();
    const localEditor = new InlineEditor({
      host: localHost,
      grid: localGrid,
      store: localStore,
      wb: realWorkbook,
      onAfterCommit: () => {},
    });
    try {
      const anchor = { sheet: 0, row: 2, col: 3 };
      mutators.setActive(localStore, anchor);
      mutators.setRange(localStore, { sheet: 0, r0: 2, c0: 3, r1: 4, c1: 5 });
      mutators.setPendingFormat(localStore, { addr: anchor, format: { bold: true } });
      mutators.setR1C1(localStore, true);
      localEditor.begin('=SUM(A1:B2)');
      mutators.setCopyRanges(
        localStore,
        [
          { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 2 },
          { sheet: 0, r0: 7, c0: 0, r1: 7, c1: 0 },
        ],
        'cut',
      );
      const input = localGrid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
      input.value = '=SUM(A1:B2)';
      input.setSelectionRange(6, 10, 'backward');
      input.dispatchEvent(new Event('input'));
      const before = localStore.getState();
      const lease = localEditor.suspendForFormulaPalette({
        getLocale: () => 'ja',
        contextCurrent: () => true,
      });
      expect(lease).not.toBeNull();
      expect(localEditor.isActive()).toBe(true);
      expect(localEditor.isFormulaEdit()).toBe(false);
      expect(localGrid.querySelector('textarea.fc-host__editor')).toBeNull();
      expect(localStore.getState().ui.editor.kind).toBe('idle');
      expect(localStore.getState().ui.pendingFormat).toEqual(before.ui.pendingFormat);
      expect(localStore.getState().ui.editorRefs).toEqual([]);
      expect(realWorkbook.getValue(anchor)).toEqual({ kind: 'blank' });

      // Palette navigation may alter the visible selection while it is open;
      // the lease restores the captured grid selection rather than treating
      // that navigation as an edit to the workbook.
      localStore.setState((state) => ({
        ...state,
        selection: {
          ...state.selection,
          active: { sheet: 0, row: 9, col: 9 },
          anchor: { sheet: 0, row: 9, col: 9 },
          range: { sheet: 0, r0: 9, c0: 9, r1: 9, c1: 9 },
        },
      }));
      const restored = lease?.userCancel() as HTMLTextAreaElement | null | undefined;
      expect(restored).not.toBeNull();
      expect(restored?.value).toBe('=SUM(A1:B2)');
      expect(restored?.selectionStart).toBe(6);
      expect(restored?.selectionEnd).toBe(10);
      expect(restored?.selectionDirection).toBe('backward');
      expect(localStore.getState().selection).toEqual(before.selection);
      expect(localStore.getState().ui.pendingFormat).toEqual(before.ui.pendingFormat);
      expect(localStore.getState().ui.copyRanges).toEqual(before.ui.copyRanges);
      expect(localStore.getState().ui.copyMode).toBe(before.ui.copyMode);
      expect(localStore.getState().ui.editor).toEqual(before.ui.editor);
      expect(localStore.getState().ui.editorRefs).toEqual(before.ui.editorRefs);
      expect(localStore.getState().ui.r1c1).toBe(true);
      expect(realWorkbook.getValue(anchor)).toEqual({ kind: 'blank' });
      expect(document.activeElement).not.toBe(restored);
    } finally {
      localEditor.cancel();
      realWorkbook.dispose();
      localHost.remove();
    }
  });

  it('refuses suspension during IME composition and invalidates stale leases', () => {
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(store, addr);
    editor.begin('=A1');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    input.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
    expect(
      editor.suspendForFormulaPalette({ getLocale: () => 'en', contextCurrent: () => true }),
    ).toBeNull();
    expect(editor.isActive()).toBe(true);
    input.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true }));
    const lease = editor.suspendForFormulaPalette({
      getLocale: () => 'en',
      contextCurrent: () => true,
    });
    expect(lease).not.toBeNull();
    mutators.setR1C1(store, true);
    expect(lease?.userCancel()).toBeNull();
    expect(editor.isActive()).toBe(false);
    expect(grid.querySelector('textarea.fc-host__editor')).toBeNull();
    editor.cancel();
  });

  it('reports a suspended edit as active and cancel releases its lease', () => {
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    editor.begin('=A1');
    const lease = editor.suspendForFormulaPalette({
      getLocale: () => 'en',
      contextCurrent: () => true,
    });
    expect(lease?.valid()).toBe(true);
    expect(editor.isActive()).toBe(true);
    editor.cancel();
    expect(editor.isActive()).toBe(false);
    expect(lease?.valid()).toBe(false);
    expect(lease?.userCancel()).toBeNull();
    expect(grid.querySelector('textarea.fc-host__editor')).toBeNull();
  });

  it('Enter key commits and advances down', () => {
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    editor.begin('');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    input.value = '7';
    const e = new KeyboardEvent('keydown', { key: 'Enter', cancelable: true, bubbles: true });
    input.dispatchEvent(e);
    expect(e.defaultPrevented).toBe(true);
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 7 });
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 1, col: 0 });
  });

  it('does not commit Enter while IME composition is active', () => {
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    editor.begin('');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    input.value = 'に';
    input.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));

    const composingEnter = new KeyboardEvent('keydown', {
      key: 'Enter',
      cancelable: true,
      bubbles: true,
    });
    input.dispatchEvent(composingEnter);

    expect(composingEnter.defaultPrevented).toBe(false);
    expect(onAfterCommit).not.toHaveBeenCalled();
    expect(editor.isActive()).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');

    input.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true }));
    const enter = new KeyboardEvent('keydown', { key: 'Enter', cancelable: true, bubbles: true });
    input.dispatchEvent(enter);

    expect(enter.defaultPrevented).toBe(true);
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'に' });
    expect(editor.isActive()).toBe(false);
  });

  it('Escape cancels without writing', () => {
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    editor.begin('');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    input.value = '99';
    const e = new KeyboardEvent('keydown', { key: 'Escape', cancelable: true, bubbles: true });
    input.dispatchEvent(e);
    expect(e.defaultPrevented).toBe(true);
    expect(onAfterCommit).not.toHaveBeenCalled();
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');
    expect(editor.isActive()).toBe(false);
  });

  it('does not cancel Escape while IME composition is active', () => {
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    editor.begin('');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    input.value = '変換中';
    input.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));

    const composingEscape = new KeyboardEvent('keydown', {
      key: 'Escape',
      cancelable: true,
      bubbles: true,
    });
    input.dispatchEvent(composingEscape);

    expect(composingEscape.defaultPrevented).toBe(false);
    expect(onAfterCommit).not.toHaveBeenCalled();
    expect(editor.isActive()).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');

    input.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true }));
    const escapeKey = new KeyboardEvent('keydown', {
      key: 'Escape',
      cancelable: true,
      bubbles: true,
    });
    input.dispatchEvent(escapeKey);

    expect(escapeKey.defaultPrevented).toBe(true);
    expect(editor.isActive()).toBe(false);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');
  });

  it('Tab advances right; default Shift+Tab commits in place', () => {
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    editor.begin('');
    let input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    input.value = 'a';
    input.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Tab', cancelable: true, bubbles: true }),
    );
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 0, col: 1 });

    editor.begin('');
    input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    input.value = 'b';
    input.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Tab', shiftKey: true, cancelable: true, bubbles: true }),
    );
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 0, col: 1 });
  });

  it('Mac Tab commits and advances through the selected rectangle', () => {
    host.classList.add('fc-host');
    host.dataset.fcPlatform = 'mac';
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    editor.begin('');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    input.value = 'entered';
    input.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Tab', cancelable: true, bubbles: true }),
    );

    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'entered' });
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 0, col: 1 });
    expect(store.getState().selection.anchor).toEqual({ sheet: 0, row: 0, col: 0 });
    expect(store.getState().selection.range).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
  });

  it('Mac Cmd+Enter fills the selection in one undoable controller batch', () => {
    host.classList.add('fc-host');
    host.dataset.fcPlatform = 'mac';
    const history = new History();
    const controller = new InteractionController({
      store,
      getWb: () => wb,
      history,
    });
    const unregister = registerInteractionController(store, controller);
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    editor.begin('');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    input.value = '9';
    const e = new KeyboardEvent('keydown', {
      key: 'Enter',
      metaKey: true,
      cancelable: true,
      bubbles: true,
    });
    input.dispatchEvent(e);

    expect(e.defaultPrevented).toBe(true);
    expect(editor.isActive()).toBe(false);
    expect(history.canUndo()).toBe(true);
    for (let row = 0; row <= 1; row += 1) {
      for (let col = 0; col <= 1; col += 1) {
        expect(wb.getValue({ sheet: 0, row, col })).toEqual({ kind: 'number', value: 9 });
      }
    }
    expect(history.undo()).toBe(true);
    for (let row = 0; row <= 1; row += 1) {
      for (let col = 0; col <= 1; col += 1) {
        expect(wb.getValue({ sheet: 0, row, col })).toEqual({ kind: 'blank' });
      }
    }
    unregister();
    controller.dispose();
  });

  it('Mac Shift+Enter commits upward without inserting a newline', () => {
    host.classList.add('fc-host');
    host.dataset.fcPlatform = 'mac';
    mutators.setActive(store, { sheet: 0, row: 2, col: 0 });
    editor.begin('');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    input.value = 'up';
    const e = new KeyboardEvent('keydown', {
      key: 'Enter',
      shiftKey: true,
      cancelable: true,
      bubbles: true,
    });
    input.dispatchEvent(e);

    expect(e.defaultPrevented).toBe(true);
    expect(input.isConnected).toBe(false);
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 1, col: 0 });
    expect(wb.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({ kind: 'text', value: 'up' });
  });

  it('default Shift+Enter inserts a newline without committing', () => {
    mutators.setActive(store, { sheet: 0, row: 2, col: 0 });
    editor.begin('before');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    const e = new KeyboardEvent('keydown', {
      key: 'Enter',
      shiftKey: true,
      cancelable: true,
      bubbles: true,
    });
    input.dispatchEvent(e);

    expect(e.defaultPrevented).toBe(true);
    expect(editor.isActive()).toBe(true);
    expect(input.value).toBe('before\n');
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 2, col: 0 });
    expect(onAfterCommit).not.toHaveBeenCalled();
  });

  it('Cmd+Enter keeps newline behavior on the default platform', () => {
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    editor.begin('a');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    const e = new KeyboardEvent('keydown', {
      key: 'Enter',
      metaKey: true,
      cancelable: true,
      bubbles: true,
    });
    input.dispatchEvent(e);

    expect(e.defaultPrevented).toBe(true);
    expect(editor.isActive()).toBe(true);
    expect(input.value).toBe('a\n');
    expect(onAfterCommit).not.toHaveBeenCalled();
  });

  it('Alt+Enter inserts a newline while editing', () => {
    host.classList.add('fc-host');
    host.dataset.fcPlatform = 'mac';
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    editor.begin('a');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    input.setSelectionRange(1, 1);
    const e = new KeyboardEvent('keydown', {
      key: 'Enter',
      altKey: true,
      cancelable: true,
      bubbles: true,
    });
    input.dispatchEvent(e);

    expect(e.defaultPrevented).toBe(true);
    expect(editor.isActive()).toBe(true);
    expect(input.value).toBe('a\n');
  });

  it('Cmd+T rotates a formula reference on Mac without committing', () => {
    host.classList.add('fc-host');
    host.dataset.fcPlatform = 'mac';
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    editor.begin('=A1');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    const e = new KeyboardEvent('keydown', {
      key: 't',
      metaKey: true,
      cancelable: true,
      bubbles: true,
    });
    input.dispatchEvent(e);

    expect(e.defaultPrevented).toBe(true);
    expect(input.value).toBe('=$A$1');
    expect(editor.isActive()).toBe(true);
    expect(onAfterCommit).not.toHaveBeenCalled();
  });

  it('blur on the input commits as "none"', () => {
    mutators.setActive(store, { sheet: 0, row: 3, col: 3 });
    editor.begin('');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    input.value = 'blur';
    input.dispatchEvent(new Event('blur'));
    expect(onAfterCommit).toHaveBeenCalledTimes(1);
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 3, col: 3 });
    expect(editor.isActive()).toBe(false);
  });

  it('writeInput failures are swallowed (warning logged) and the editor still tears down', () => {
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    editor.begin('');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    input.value = 'x';

    // Force writeInput to throw by stubbing the workbook setter.
    const setText = wb.setText.bind(wb);
    const setNumber = wb.setNumber.bind(wb);
    const setBool = wb.setBool.bind(wb);
    wb.setText = () => {
      throw new Error('boom');
    };
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});

    editor.commit('none');
    expect(warn).toHaveBeenCalled();
    expect(editor.isActive()).toBe(false);

    warn.mockRestore();
    wb.setText = setText;
    wb.setNumber = setNumber;
    wb.setBool = setBool;
  });

  describe('single-cell commit write path', () => {
    const withUnrestrictedController = () => {
      const history = new History();
      const onChanged = vi.fn();
      const controller = new InteractionController({ store, getWb: () => wb, history, onChanged });
      const unregister = registerInteractionController(store, controller);
      const onValidation = vi.fn();
      const surface = new InlineEditor({ host, grid, store, wb, onAfterCommit, onValidation });
      const typeAndCommit = (raw: string): void => {
        surface.begin('');
        const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
        input.value = raw;
        surface.commit('none');
      };
      return {
        history,
        onChanged,
        onValidation,
        surface,
        typeAndCommit,
        dispose: () => {
          if (surface.isActive()) surface.cancel();
          unregister();
          controller.dispose();
        },
      };
    };

    it('routes through a registered controller even without a policy', () => {
      const { history, onChanged, typeAndCommit, dispose } = withUnrestrictedController();
      mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
      typeAndCommit('5');

      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 5 });
      expect(onChanged).toHaveBeenCalledWith(expect.objectContaining({ status: 'applied' }));
      expect(history.canUndo()).toBe(true);
      dispose();
    });

    it('keeps validation titles and warning notices on the controller path', () => {
      const { onValidation, surface, typeAndCommit, dispose } = withUnrestrictedController();
      const stopAddr = { sheet: 0, row: 0, col: 0 };
      const warnAddr = { sheet: 0, row: 1, col: 0 };
      mutators.setCellFormat(store, stopAddr, {
        validation: {
          kind: 'whole',
          op: '=',
          a: 5,
          errorStyle: 'stop',
          errorTitle: 'Five only',
          errorMessage: 'Value must be five.',
        },
      });
      mutators.setCellFormat(store, warnAddr, {
        validation: {
          kind: 'whole',
          op: '=',
          a: 5,
          errorStyle: 'warning',
          errorMessage: 'Value is not five.',
        },
      });

      mutators.setActive(store, stopAddr);
      typeAndCommit('4');
      expect(surface.isActive()).toBe(true);
      expect(wb.getValue(stopAddr)).toEqual({ kind: 'blank' });
      expect(onValidation).toHaveBeenLastCalledWith({
        severity: 'stop',
        title: 'Five only',
        message: 'Value must be five.',
      });
      surface.cancel();

      mutators.setActive(store, warnAddr);
      typeAndCommit('4');
      expect(surface.isActive()).toBe(false);
      expect(wb.getValue(warnAddr)).toEqual({ kind: 'number', value: 4 });
      expect(onValidation).toHaveBeenLastCalledWith(
        expect.objectContaining({ severity: 'warning', message: 'Value is not five.' }),
      );
      dispose();
    });
  });

  it('isActive reflects whether an input is mounted', async () => {
    expect(editor.isActive()).toBe(false);
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    editor.begin('hi');
    expect(editor.isActive()).toBe(true);
    await flushRaf();
    editor.cancel();
    expect(editor.isActive()).toBe(false);
  });

  describe('commitMulti (Ctrl+Enter selection fill)', () => {
    const selectRange = (r0: number, c0: number, r1: number, c1: number): void => {
      store.setState((s) => ({
        ...s,
        selection: {
          active: { sheet: 0, row: r0, col: c0 },
          anchor: { sheet: 0, row: r0, col: c0 },
          range: { sheet: 0, r0, c0, r1, c1 },
          extraRanges: [],
        },
      }));
    };

    it('fills every selected cell with the typed value', () => {
      selectRange(0, 0, 0, 2);
      editor.begin('7');
      editor.commitMulti();
      wb.recalc();
      for (let col = 0; col <= 2; col += 1) {
        expect(wb.getValue({ sheet: 0, row: 0, col })).toEqual({ kind: 'number', value: 7 });
      }
    });

    const withController = (
      policy?: InteractionPolicy,
    ): { history: History; onValidation: Mock; fillEditor: InlineEditor; dispose: () => void } => {
      const history = new History();
      const controller = new InteractionController({ store, getWb: () => wb, history });
      if (policy) controller.setPolicy(policy);
      const unregister = registerInteractionController(store, controller);
      const onValidation = vi.fn();
      const fillEditor = new InlineEditor({ host, grid, store, wb, onAfterCommit, onValidation });
      return {
        history,
        onValidation,
        fillEditor,
        dispose: () => {
          if (fillEditor.isActive()) fillEditor.cancel();
          unregister();
          controller.dispose();
        },
      };
    };

    it('Ctrl+Enter off Mac fills through the controller, once per cell and never into merge bodies', () => {
      const { history, fillEditor, dispose } = withController();
      mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 2 });
      selectRange(0, 0, 0, 3);
      store.setState((s) => ({
        ...s,
        selection: { ...s.selection, extraRanges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 }] },
      }));
      fillEditor.begin('5');
      const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
      const e = new KeyboardEvent('keydown', {
        key: 'Enter',
        ctrlKey: true,
        cancelable: true,
        bubbles: true,
      });
      input.dispatchEvent(e);

      expect(fillEditor.isActive()).toBe(false);
      expect(history.canUndo()).toBe(true);
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 5 });
      expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 5 });
      expect(wb.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({ kind: 'blank' });
      expect(wb.getValue({ sheet: 0, row: 0, col: 3 })).toEqual({ kind: 'number', value: 5 });
      dispose();
    });

    it('reports the controller rejection reason and keeps the editor open', () => {
      const { onValidation, fillEditor, dispose } = withController({
        defaultOperation: 'allow',
        editable: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }],
      });
      selectRange(0, 0, 0, 1);
      fillEditor.begin('5');
      fillEditor.commitMulti();

      expect(fillEditor.isActive()).toBe(true);
      expect(onValidation).toHaveBeenCalledWith(
        expect.objectContaining({ severity: 'stop', message: 'cell is outside editable cells' }),
      );
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });
      dispose();
    });

    it('reports an over-limit controller fill instead of silently ignoring it', () => {
      const { onValidation, fillEditor, dispose } = withController();
      selectRange(0, 0, 100_000, 0);
      fillEditor.begin('7');
      fillEditor.commitMulti();

      expect(fillEditor.isActive()).toBe(true);
      expect(onValidation).toHaveBeenCalledWith(
        expect.objectContaining({ severity: 'stop', message: expect.stringContaining('100,000') }),
      );
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');
      dispose();
    });

    it('does not materialize huge selections', () => {
      selectRange(0, 0, 100_000, 0);
      editor.begin('7');
      editor.commitMulti();
      wb.recalc();

      expect(wb.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');
      expect(onAfterCommit).not.toHaveBeenCalled();
      expect(editor.isActive()).toBe(true);
    });

    it('validates NON-anchor cells too, aborting on a stop rule (M-16)', () => {
      const onValidation = vi.fn();
      const validatedEditor = new InlineEditor({
        host,
        grid,
        store,
        wb,
        onAfterCommit,
        onValidation,
      });
      // B1 restricts entries to a list; the anchor A1 is unrestricted.
      mutators.setCellFormat(
        store,
        { sheet: 0, row: 0, col: 1 },
        {
          validation: { kind: 'list', source: ['X'] },
        },
      );
      selectRange(0, 0, 0, 1);
      validatedEditor.begin('Y'); // not in B1's list
      validatedEditor.commitMulti();
      wb.recalc();
      // The stop rule fired for B1: the fill aborted and B1 kept no value.
      expect(onValidation).toHaveBeenCalledTimes(1);
      expect(wb.getValue({ sheet: 0, row: 0, col: 1 }).kind).toBe('blank');
      validatedEditor.cancel();
    });

    it('fills all cells when every value satisfies its rule', () => {
      const onValidation = vi.fn();
      const validatedEditor = new InlineEditor({
        host,
        grid,
        store,
        wb,
        onAfterCommit,
        onValidation,
      });
      mutators.setCellFormat(
        store,
        { sheet: 0, row: 0, col: 1 },
        {
          validation: { kind: 'list', source: ['X'] },
        },
      );
      selectRange(0, 0, 0, 1);
      validatedEditor.begin('X'); // valid for B1
      validatedEditor.commitMulti();
      wb.recalc();
      expect(onValidation).not.toHaveBeenCalled();
      expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'text', value: 'X' });
      validatedEditor.cancel();
    });
  });
});
