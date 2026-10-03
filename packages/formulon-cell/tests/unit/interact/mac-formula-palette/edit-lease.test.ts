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
});
