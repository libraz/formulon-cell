import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { InlineEditor } from '../../../src/interact/editor.js';
import {
  type SpreadsheetStore,
  createSpreadsheetStore,
  mutators,
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
  let onAfterCommit: ReturnType<typeof vi.fn>;
  let editor: InlineEditor;

  beforeEach(async () => {
    host = document.createElement('div');
    grid = document.createElement('div');
    host.appendChild(grid);
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await newWb();
    onAfterCommit = vi.fn();
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

  it('begin positions the input over the active cell using cellRect', () => {
    mutators.setActive(store, { sheet: 0, row: 1, col: 1 });
    editor.begin('');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    // headerColWidth=52, headerRowHeight=30, defaultColWidth=104, defaultRowHeight=28.
    // Cell (1, 1): x=52+104=156, y=30+28=58, w=104, h=28.
    expect(input.style.left).toBe('156px');
    expect(input.style.top).toBe('58px');
    expect(input.style.width).toBe('104px');
    expect(input.style.height).toBe('28px');
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

  it('commit("right") advances column instead of row', () => {
    mutators.setActive(store, { sheet: 0, row: 4, col: 4 });
    editor.begin('');
    const input = grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;
    input.value = 'foo';
    editor.commit('right');
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 4, col: 5 });
  });

  it('commit("none") does not move the active cell (Shift+Tab semantics)', () => {
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

  it('Tab commits and advances right; Shift+Tab commits in place', () => {
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
    // Active stays where it was after the first commit (still at col 1).
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 0, col: 1 });
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

  it('isActive reflects whether an input is mounted', async () => {
    expect(editor.isActive()).toBe(false);
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    editor.begin('hi');
    expect(editor.isActive()).toBe(true);
    await flushRaf();
    editor.cancel();
    expect(editor.isActive()).toBe(false);
  });
});
