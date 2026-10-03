import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { InlineEditor } from '../../../../src/interact/editor.js';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/index.js';
import { type PaletteSetupArgs, paletteRoot, setupPalette } from './fixtures.js';

describe('interact/mac-formula-palette suspended edit lease', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet({ workbook: await WorkbookHandle.createDefault() });
    expect(sheet.workbook.isStub).toBe(false);
  });

  afterEach(() => sheet.dispose());

  const setup = (...args: PaletteSetupArgs) => setupPalette(sheet, ...args);

  const attachInlineEditor = (): { editor: InlineEditor; grid: HTMLElement } => {
    const grid = document.createElement('div');
    sheet.host.appendChild(grid);
    const editor = new InlineEditor({
      host: sheet.host,
      grid,
      store: sheet.instance.store,
      wb: sheet.workbook,
      onAfterCommit: () => {},
    });
    return { editor, grid };
  };

  it('suspends a mid-edit inline editor and Cancel restores its text and focus', () => {
    const { editor, grid } = attachInlineEditor();
    const { formulaBar, palette } = setup(
      undefined,
      undefined,
      undefined,
      undefined,
      (_, context) => editor.suspendForFormulaPalette(context),
    );
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    editor.begin('=1+');

    palette.open();
    expect(grid.querySelector('textarea.fc-host__editor')).toBeNull();
    expect(editor.isActive()).toBe(true);
    expect(formulaBar.input.value).toBe('=1+');
    const trigger = document.createElement('button');
    sheet.host.appendChild(trigger);
    palette.setReturnFocusTarget(trigger);

    palette.close();
    const restored = grid.querySelector<HTMLTextAreaElement>('textarea.fc-host__editor');
    expect(restored?.value).toBe('=1+');
    expect(document.activeElement).toBe(restored);
    expect(sheet.instance.store.getState().ui.editor.kind).not.toBe('idle');
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBeNull();
    editor.cancel();
    expect(editor.isActive()).toBe(false);
  });

  it('Done commits the palette formula and discards the suspended inline edit', () => {
    const { editor, grid } = attachInlineEditor();
    const { palette } = setup(undefined, undefined, undefined, undefined, (_, context) =>
      editor.suspendForFormulaPalette(context),
    );
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    editor.begin('=');

    palette.open('ACOS');
    const arg = paletteRoot(palette).querySelector<HTMLInputElement>('[data-argument-index="0"]');
    expect(arg).not.toBeNull();
    if (arg) {
      arg.value = '1';
      arg.dispatchEvent(new Event('input', { bubbles: true }));
    }
    paletteRoot(palette).querySelector<HTMLButtonElement>('[data-action="done"]')?.click();

    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=ACOS(1)');
    expect(editor.isActive()).toBe(false);
    expect(grid.querySelector('textarea.fc-host__editor')).toBeNull();
    palette.close();
    expect(grid.querySelector('textarea.fc-host__editor')).toBeNull();
  });

  it('projects an existing compound inline formula and commits one history entry', () => {
    const anchor = { sheet: 0, row: 0, col: 0 };
    sheet.workbook.setNumber(anchor, 9);
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, anchor);
    sheet.instance.history.clear();

    const { editor, grid } = attachInlineEditor();
    const { formulaBar, palette } = setup(
      undefined,
      undefined,
      undefined,
      undefined,
      (_, context) => editor.suspendForFormulaPalette(context),
    );
    editor.begin('=1+SUM(2,3)*4');
    const inline = grid.querySelector<HTMLTextAreaElement>('textarea.fc-host__editor');
    inline?.setSelectionRange(7, 10, 'backward');

    palette.open();
    const root = paletteRoot(palette);
    expect(root.dataset.state).toBe('arguments-editing');
    expect(formulaBar.input.value).toBe('=1+SUM(2,3)*4');
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="0"]')?.value).toBe('2');
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="1"]')?.value).toBe('3');

    const first = root.querySelector<HTMLInputElement>('[data-argument-index="0"]');
    expect(first).not.toBeNull();
    if (first) {
      first.value = '5';
      first.dispatchEvent(new Event('input', { bubbles: true }));
    }
    expect(formulaBar.input.value).toBe('=1+SUM(5,3)*4');
    expect(sheet.workbook.getValue(anchor)).toEqual({ kind: 'number', value: 9 });
    expect(sheet.instance.history.canUndo()).toBe(false);

    root.querySelector<HTMLButtonElement>('[data-action="done"]')?.click();
    expect(sheet.workbook.cellFormula(anchor)).toBe('=1+SUM(5,3)*4');
    expect(sheet.workbook.getValue(anchor)).toEqual({ kind: 'number', value: 33 });
    expect(sheet.instance.history.canUndo()).toBe(true);
    expect(sheet.instance.undo()).toBe(true);
    expect(sheet.workbook.cellFormula(anchor)).toBeNull();
    expect(sheet.workbook.getValue(anchor)).toEqual({ kind: 'number', value: 9 });
    expect(sheet.instance.redo()).toBe(true);
    expect(sheet.workbook.cellFormula(anchor)).toBe('=1+SUM(5,3)*4');
    editor.cancel();
  });

  it('restores the exact compound caret selection when the leased palette draft is cancelled', () => {
    const { editor, grid } = attachInlineEditor();
    const { formulaBar, palette } = setup(
      undefined,
      undefined,
      undefined,
      undefined,
      (_, context) => editor.suspendForFormulaPalette(context),
    );
    editor.begin('=1+SUM(2,3)*4');
    const inline = grid.querySelector<HTMLTextAreaElement>('textarea.fc-host__editor');
    inline?.setSelectionRange(7, 10, 'backward');

    palette.open();
    expect(formulaBar.input.value).toBe('=1+SUM(2,3)*4');
    palette.close();

    const restored = grid.querySelector<HTMLTextAreaElement>('textarea.fc-host__editor');
    expect(restored?.value).toBe('=1+SUM(2,3)*4');
    expect(restored?.selectionStart).toBe(7);
    expect(restored?.selectionEnd).toBe(10);
    expect(restored?.selectionDirection).toBe('backward');
    expect(document.activeElement).toBe(restored);
    editor.cancel();
  });

  it('suspends a formula-bar edit and Cancel returns focus to the restored formula bar', () => {
    const { formulaBar, palette } = setup(
      undefined,
      undefined,
      undefined,
      undefined,
      (bar, context) => bar.suspendForFormulaPalette(context),
    );
    formulaBar.input.focus();
    formulaBar.input.value = '=2*';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    expect(formulaBar.controller.isEditing()).toBe(true);

    palette.open();
    palette.close();
    expect(formulaBar.input.value).toBe('=2*');
    expect(formulaBar.controller.isEditing()).toBe(true);
    expect(document.activeElement).toBe(formulaBar.input);
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBeNull();
    formulaBar.controller.cancelFx();
  });

  it('opens a compound formula-bar call at its backward caret and restores it on Cancel', () => {
    const { formulaBar, palette } = setup(
      undefined,
      undefined,
      undefined,
      undefined,
      (bar, context) => bar.suspendForFormulaPalette(context),
    );
    const raw = '=1+SUM(2,3)*4';
    formulaBar.input.focus();
    formulaBar.input.value = raw;
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    formulaBar.input.setSelectionRange(7, 10, 'backward');
    expect(formulaBar.controller.isEditing()).toBe(true);

    palette.open();
    const root = paletteRoot(palette);
    expect(root.dataset.state).toBe('arguments-editing');
    expect(formulaBar.input.value).toBe(raw);
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="0"]')?.value).toBe('2');
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="1"]')?.value).toBe('3');

    palette.close();
    expect(formulaBar.input.value).toBe(raw);
    expect(formulaBar.input.selectionStart).toBe(7);
    expect(formulaBar.input.selectionEnd).toBe(10);
    expect(formulaBar.input.selectionDirection).toBe('backward');
    expect(document.activeElement).toBe(formulaBar.input);
    formulaBar.controller.cancelFx();
  });

  it('discards a suspended owner on a system boundary without restoring focus or selection', () => {
    const { formulaBar, palette } = setup(
      undefined,
      undefined,
      undefined,
      undefined,
      (bar, context) => bar.suspendForFormulaPalette(context),
    );
    const oldAnchor = { sheet: 0, row: 0, col: 0 };
    const nextSelection = { sheet: 0, row: 3, col: 3 };
    mutators.setActive(sheet.instance.store, oldAnchor);
    formulaBar.input.focus();
    formulaBar.input.value = '=2*';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));

    palette.open();
    expect(formulaBar.input.value).toBe('=2*');
    mutators.setActive(sheet.instance.store, nextSelection);
    palette.discard();

    expect(sheet.instance.store.getState().selection.active).toEqual(nextSelection);
    expect(formulaBar.controller.isEditing()).toBe(false);
    expect(document.activeElement).not.toBe(formulaBar.input);
    expect(sheet.workbook.cellFormula(oldAnchor)).toBeNull();
    expect(sheet.instance.history.canUndo()).toBe(false);
  });
});
