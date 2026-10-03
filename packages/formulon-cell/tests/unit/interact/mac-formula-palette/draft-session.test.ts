import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { defaultStrings } from '../../../../src/i18n/strings.js';
import type { FormulaBarController } from '../../../../src/mount/formula-bar.js';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/index.js';
import { type PaletteSetupArgs, paletteRoot, setupPalette } from './fixtures.js';

describe('interact/mac-formula-palette draft session', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet({ workbook: await WorkbookHandle.createDefault() });
    expect(sheet.workbook.isStub).toBe(false);
  });

  afterEach(() => sheet.dispose());

  const setup = (...args: PaletteSetupArgs) => setupPalette(sheet, ...args);

  it('keeps unsynchronized raw authoritative across a visible category request and picker insert', () => {
    const { formulaBar, palette } = setup();
    palette.open('ACOS');
    formulaBar.input.value = '=MISSING_FUNCTION(';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));

    palette.open(undefined, { category: 'logical' });
    const root = paletteRoot(palette);
    expect(root.dataset.state).toBe('picker');
    expect(root.dataset.rawSynchronized).toBe('false');
    expect(formulaBar.input.value).toBe('=MISSING_FUNCTION(');
    root.querySelector<HTMLElement>('[data-function-name="IF"]')?.click();
    root.querySelector<HTMLButtonElement>('[data-action="insert-function"]')?.click();
    expect(formulaBar.input.value).toBe('=MISSING_FUNCTION(');
    expect(root.dataset.rawSynchronized).toBe('false');
    expect(root.querySelector('.fc-mac-formula-palette__guard')?.textContent).toBe(
      defaultStrings.fxDialog.macPalette?.draftConflict,
    );
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBeNull();
  });

  it('does not steal inline editor ownership during repeated palette requests', () => {
    const { formulaBar, palette } = setup();
    palette.open('ACOS');
    const rawBefore = formulaBar.input.value;
    const outside = document.createElement('textarea');
    sheet.host.appendChild(outside);
    mutators.setEditor(sheet.instance.store, { kind: 'edit', raw: 'inline', caret: 6 });
    outside.focus();

    palette.open();
    expect(document.activeElement).toBe(outside);
    expect(formulaBar.input.value).toBe(rawBefore);
    expect(paletteRoot(palette).querySelector('.fc-mac-formula-palette__guard')?.textContent).toBe(
      defaultStrings.fxDialog.macPalette?.draftConflict,
    );

    palette.open('COUNTIF');
    expect(document.activeElement).toBe(outside);
    expect(formulaBar.input.value).toBe(rawBefore);
    palette.open(undefined, { category: 'logical' });
    expect(document.activeElement).toBe(outside);
    expect(formulaBar.input.value).toBe(rawBefore);
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBeNull();
    expect(sheet.instance.history.canUndo()).toBe(false);
    outside.remove();
    mutators.setEditor(sheet.instance.store, { kind: 'idle' });
    palette.close();
  });

  it('keeps one history entry across Done, category insertion, cancel, and the next Done', () => {
    const { formulaBar, palette } = setup();
    palette.open('ACOS');
    const acosArg = paletteRoot(palette).querySelector<HTMLInputElement>(
      '[data-argument-index="0"]',
    );
    expect(acosArg).not.toBeNull();
    if (acosArg) {
      acosArg.value = '1';
      acosArg.dispatchEvent(new Event('input', { bubbles: true }));
    }
    paletteRoot(palette).querySelector<HTMLButtonElement>('[data-action="done"]')?.click();
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=ACOS(1)');

    palette.open(undefined, { category: 'logical' });
    let root = paletteRoot(palette);
    root.querySelector<HTMLElement>('[data-function-name="IF"]')?.click();
    root.querySelector<HTMLButtonElement>('[data-action="insert-function"]')?.click();
    expect(root.dataset.state).toBe('arguments-editing');
    expect(formulaBar.input.value).toBe('=IF()');
    palette.close();
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=ACOS(1)');

    palette.open(undefined, { category: 'logical' });
    root = paletteRoot(palette);
    root.querySelector<HTMLElement>('[data-function-name="IF"]')?.click();
    root.querySelector<HTMLButtonElement>('[data-action="insert-function"]')?.click();
    const values = ['TRUE', '1', '2'];
    values.forEach((value, index) => {
      const field = root.querySelector<HTMLInputElement>(`[data-argument-index="${index}"]`);
      expect(field).not.toBeNull();
      if (field) {
        field.value = value;
        field.dispatchEvent(new Event('input', { bubbles: true }));
      }
    });
    root.querySelector<HTMLButtonElement>('[data-action="done"]')?.click();
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=IF(TRUE,1,2)');
    expect(sheet.instance.undo()).toBe(true);
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=ACOS(1)');
    expect(sheet.instance.redo()).toBe(true);
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=IF(TRUE,1,2)');
  });

  it('keeps the picker open after a refused fresh draft and allows retrying Insert', () => {
    const { beginDraft, formulaBar, palette } = setup();
    palette.open('ACOS');
    const acosArg = paletteRoot(palette).querySelector<HTMLInputElement>(
      '[data-argument-index="0"]',
    );
    expect(acosArg).not.toBeNull();
    if (acosArg) {
      acosArg.value = '1';
      acosArg.dispatchEvent(new Event('input', { bubbles: true }));
    }
    paletteRoot(palette).querySelector<HTMLButtonElement>('[data-action="done"]')?.click();
    expect(formulaBar.input.value).toBe('=ACOS(1)');

    palette.open(undefined, { category: 'logical' });
    beginDraft.mockImplementation(() => null);
    const root = paletteRoot(palette);
    root.querySelector<HTMLElement>('[data-function-name="IF"]')?.click();
    root.querySelector<HTMLButtonElement>('[data-action="insert-function"]')?.click();
    expect(root.dataset.state).toBe('picker');
    expect(formulaBar.input.value).toBe('=ACOS(1)');
    expect(root.querySelector('[data-section="logical"]')).not.toBeNull();
    expect(root.querySelector('.fc-mac-formula-palette__guard')?.textContent).toBe(
      defaultStrings.fxDialog.macPalette?.draftConflict,
    );

    beginDraft.mockImplementation(
      (...args: Parameters<FormulaBarController['beginExternalDraft']>) =>
        formulaBar.controller.beginExternalDraft(...args),
    );
    root.querySelector<HTMLElement>('[data-function-name="IF"]')?.click();
    root.querySelector<HTMLButtonElement>('[data-action="insert-function"]')?.click();
    expect(paletteRoot(palette).dataset.state).toBe('arguments-editing');
    expect(formulaBar.input.value).toBe('=IF()');
    palette.close();
  });

  it('does not let picker insertion overwrite compound or malformed initial raw drafts', () => {
    const { formulaBar, palette } = setup();
    for (const raw of ['=SUM(1)+2', '=SUM({1,2]']) {
      palette.open();
      formulaBar.input.value = raw;
      formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
      const root = paletteRoot(palette);
      root.querySelector<HTMLElement>('[data-function-name="IF"]')?.click();
      root.querySelector<HTMLButtonElement>('[data-action="insert-function"]')?.click();
      expect(formulaBar.input.value).toBe(raw);
      expect(root.dataset.rawSynchronized).toBe('false');
      expect(root.querySelector('.fc-mac-formula-palette__guard')?.textContent).toBe(
        defaultStrings.fxDialog.macPalette?.draftConflict,
      );

      palette.open('IF');
      expect(formulaBar.input.value).toBe(raw);
      expect(paletteRoot(palette).dataset.rawSynchronized).toBe('false');
      expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBeNull();
      expect(sheet.instance.history.canUndo()).toBe(false);
      palette.close();
    }
  });

  it('keeps the compound source when Show All switches to a different function', () => {
    const { formulaBar, palette } = setup();
    palette.open('SUM');
    const raw = '=1+SUM(2,3)*4';
    formulaBar.input.value = raw;
    formulaBar.input.setSelectionRange(raw.indexOf('2') + 1, raw.indexOf('2') + 1);
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    expect(paletteRoot(palette).dataset.state).toBe('arguments-editing');

    paletteRoot(palette).querySelector<HTMLButtonElement>('[data-action="show-all"]')?.click();
    const root = paletteRoot(palette);
    expect(root.dataset.state).toBe('picker');
    expect(formulaBar.input.value).toBe(raw);
    root.querySelector<HTMLElement>('[data-function-name="AVERAGE"]')?.click();
    root.querySelector<HTMLButtonElement>('[data-action="insert-function"]')?.click();

    expect(formulaBar.input.value).toBe('=1+AVERAGE()*4');
    expect(paletteRoot(palette).dataset.state).toBe('arguments-editing');
  });

  it('rebuilds a known incomplete call only after an explicit function request', () => {
    const { formulaBar, palette } = setup();
    palette.open();
    formulaBar.input.value = '=SUM(';
    formulaBar.input.setSelectionRange(
      formulaBar.input.value.length,
      formulaBar.input.value.length,
    );
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));

    const root = paletteRoot(palette);
    expect(root.dataset.rawSynchronized).toBe('true');
    expect(root.dataset.state).toBe('picker');
    root.querySelector<HTMLElement>('[data-function-name="IF"]')?.click();
    root.querySelector<HTMLButtonElement>('[data-action="insert-function"]')?.click();

    expect(formulaBar.input.value).toBe('=IF()');
    expect(paletteRoot(palette).dataset.state).toBe('arguments-editing');
    expect(
      paletteRoot(palette).querySelector<HTMLButtonElement>('[data-action="done"]'),
    ).not.toBeNull();
  });

  it('reopens the existing picker without reanchoring or starting another draft', () => {
    const { anchor, beginDraft, formulaBar, opener, palette } = setup();
    const initialAnchor = { ...anchor };
    palette.open();
    expect(beginDraft).toHaveBeenCalledTimes(1);

    anchor.row = 7;
    palette.open();
    expect(beginDraft).toHaveBeenCalledTimes(1);
    expect(formulaBar.input.value).toBe('=');
    expect(paletteRoot(palette).dataset.state).toBe('picker');
    expect(document.activeElement).toBe(paletteRoot(palette).querySelector('input[type="search"]'));
    expect(beginDraft.mock.calls[0]?.[0]).toEqual(initialAnchor);

    palette.close();
    expect(document.activeElement).toBe(opener);
  });

  it('replaces a synchronized editing draft on a repeat seeded open without writing the cell', () => {
    const { formulaBar, palette } = setup();
    palette.open('ACOS');
    const arg = paletteRoot(palette).querySelector<HTMLInputElement>('[data-argument-index="0"]');
    expect(arg).not.toBeNull();
    if (arg) {
      arg.value = '1';
      arg.dispatchEvent(new Event('input', { bubbles: true }));
    }
    palette.open('COUNTIF');

    expect(paletteRoot(palette).dataset.state).toBe('arguments-editing');
    expect(formulaBar.input.value).toBe('=COUNTIF()');
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBeNull();
    palette.close();
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBeNull();
  });

  it('keeps unsynchronized raw input authoritative and guards missing repeat seeds', () => {
    const { formulaBar, palette } = setup();
    palette.open('ACOS');
    formulaBar.input.value = '=MISSING_FUNCTION(';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    palette.open('COUNTIF');
    expect(formulaBar.input.value).toBe('=MISSING_FUNCTION(');
    expect(paletteRoot(palette).dataset.rawSynchronized).toBe('false');

    palette.open('MISSING_FUNCTION');
    expect(formulaBar.input.value).toBe('=MISSING_FUNCTION(');
    expect(paletteRoot(palette).dataset.rawSynchronized).toBe('false');
    palette.close();
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBeNull();
  });

  it('reopens a committed formula as a fresh draft and keeps workbook history write-free until Done', () => {
    const { formulaBar, palette } = setup();
    palette.open('ACOS');
    const acosArg = paletteRoot(palette).querySelector<HTMLInputElement>(
      '[data-argument-index="0"]',
    );
    expect(acosArg).not.toBeNull();
    if (acosArg) {
      acosArg.value = '1';
      acosArg.dispatchEvent(new Event('input', { bubbles: true }));
    }
    paletteRoot(palette).querySelector<HTMLButtonElement>('[data-action="done"]')?.click();
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=ACOS(1)');

    palette.open('COUNTIF');
    expect(formulaBar.input.value).toBe('=COUNTIF()');
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=ACOS(1)');
    palette.close();
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=ACOS(1)');

    palette.open('COUNTIF');
    const range = paletteRoot(palette).querySelector<HTMLInputElement>('[data-argument-index="0"]');
    const criterion = paletteRoot(palette).querySelector<HTMLInputElement>(
      '[data-argument-index="1"]',
    );
    expect(range).not.toBeNull();
    expect(criterion).not.toBeNull();
    if (range) {
      range.value = 'A1:A2';
      range.dispatchEvent(new Event('input', { bubbles: true }));
    }
    if (criterion) {
      criterion.value = '1';
      criterion.dispatchEvent(new Event('input', { bubbles: true }));
    }
    paletteRoot(palette).querySelector<HTMLButtonElement>('[data-action="done"]')?.click();
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=COUNTIF(A1:A2,1)');

    expect(sheet.instance.undo()).toBe(true);
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=ACOS(1)');
    expect(sheet.instance.redo()).toBe(true);
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=COUNTIF(A1:A2,1)');
  });

  it('guards a repeat open when the live workbook loses the selected function', () => {
    const { formulaBar, palette } = setup();
    palette.open('ACOS');
    expect(paletteRoot(palette).dataset.state).toBe('arguments-editing');
    const names = sheet.workbook.functionNames();
    const namesSpy = vi
      .spyOn(sheet.workbook, 'functionNames')
      .mockReturnValue(names?.filter((name) => name !== 'ACOS') ?? []);

    palette.open();
    expect(paletteRoot(palette).dataset.state).toBe('arguments-editing');
    expect(formulaBar.input.value).toBe('=ACOS()');
    expect(paletteRoot(palette).dataset.rawSynchronized).toBe('false');
    expect(paletteRoot(palette).querySelector('.fc-mac-formula-palette__guard')?.textContent).toBe(
      defaultStrings.fxDialog.macPalette?.unavailable,
    );
    namesSpy.mockRestore();
  });

  it('retains the latest raw draft and blocks writes when the same-workbook catalog withdraws it', () => {
    const { beginDraft, formulaBar, palette } = setup();
    palette.open('SUM');
    const first = paletteRoot(palette).querySelector<HTMLInputElement>('[data-argument-index="0"]');
    expect(first).not.toBeNull();
    if (first) {
      first.value = '9';
      first.dispatchEvent(new Event('input', { bubbles: true }));
    }
    const names = sheet.workbook.functionNames();
    const namesSpy = vi
      .spyOn(sheet.workbook, 'functionNames')
      .mockReturnValue(names?.filter((name) => name !== 'SUM') ?? []);

    palette.open();
    const root = paletteRoot(palette);
    expect(root.dataset.state).toBe('arguments-editing');
    expect(root.dataset.rawSynchronized).toBe('false');
    expect(formulaBar.input.value).toBe('=SUM(9)');
    expect(beginDraft).toHaveBeenCalledTimes(1);
    expect(root.querySelector<HTMLButtonElement>('[data-action="done"]')?.disabled).toBe(true);
    palette.rangeInsertTarget()?.insertRefAtCaret('A1');
    expect(formulaBar.input.value).toBe('=SUM(9)');
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBeNull();

    namesSpy.mockRestore();
  });

  it('keeps ACOS transient until Done, previews without native writes, and round-trips one history entry', () => {
    const { anchor, formulaBar, mirror, palette } = setup();
    sheet.workbook.setNumber(anchor, 1);
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, anchor);
    sheet.instance.history.clear();

    palette.open();
    expect(formulaBar.input.value).toBe('=');
    expect(mirror).toHaveBeenLastCalledWith(anchor, '=');
    expect(sheet.workbook.getValue(anchor)).toEqual({ kind: 'number', value: 1 });
    expect(sheet.workbook.cellFormula(anchor)).toBeNull();
    expect(sheet.instance.history.canUndo()).toBe(false);

    const acos = paletteRoot(palette).querySelector<HTMLButtonElement>(
      '[data-function-name="ACOS"]',
    );
    expect(acos).not.toBeNull();
    acos?.click();
    paletteRoot(palette)
      .querySelector<HTMLButtonElement>('[data-action="insert-function"]')
      ?.click();

    expect(paletteRoot(palette).dataset.state).toBe('arguments-editing');
    const arg = paletteRoot(palette).querySelector<HTMLInputElement>('[data-argument-index="0"]');
    expect(arg).not.toBeNull();
    expect(paletteRoot(palette).querySelector('[data-action="range-picker"] svg')).not.toBeNull();
    expect(formulaBar.input.value).toBe('=ACOS()');
    arg?.focus();
    if (arg) {
      arg.value = '1';
      arg.dispatchEvent(new Event('input', { bubbles: true }));
    }
    expect(formulaBar.input.value).toBe('=ACOS(1)');
    expect(paletteRoot(palette).querySelector('[data-role="preview-value"]')?.textContent).toBe(
      '0',
    );
    expect(sheet.workbook.getValue(anchor)).toEqual({ kind: 'number', value: 1 });
    expect(sheet.instance.history.canUndo()).toBe(false);

    paletteRoot(palette).querySelector<HTMLButtonElement>('[data-action="done"]')?.click();
    expect(sheet.workbook.cellFormula(anchor)).toBe('=ACOS(1)');
    expect(sheet.workbook.getValue(anchor)).toEqual({ kind: 'number', value: 0 });
    expect(sheet.instance.history.canUndo()).toBe(true);
    expect(palette.isOpen()).toBe(true);
    expect(paletteRoot(palette).dataset.state).toBe('arguments-committed');
    expect(
      paletteRoot(palette).querySelector<HTMLInputElement>('[data-argument-index="0"]')?.value,
    ).toBe('1');
    expect(paletteRoot(palette).querySelector('[data-role="preview-value"]')?.textContent).toBe(
      '0',
    );

    palette.close();
    expect(sheet.workbook.cellFormula(anchor)).toBe('=ACOS(1)');
    expect(sheet.instance.undo()).toBe(true);
    expect(sheet.workbook.cellFormula(anchor)).toBeNull();
    expect(sheet.workbook.getValue(anchor)).toEqual({ kind: 'number', value: 1 });
    expect(sheet.instance.history.canUndo()).toBe(false);
    expect(sheet.instance.redo()).toBe(true);
    expect(sheet.workbook.cellFormula(anchor)).toBe('=ACOS(1)');
  });

  it('records a synchronized explicit trailing-blank commit in Recent', () => {
    const { formulaBar, palette } = setup();
    palette.open('IF');
    formulaBar.input.value = '=IF(FALSE,1,)';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    paletteRoot(palette).querySelector<HTMLButtonElement>('[data-action="done"]')?.click();
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=IF(FALSE,1,)');
    palette.close();
    palette.open();
    expect(
      paletteRoot(palette).querySelectorAll('[data-section="recent"] [data-function-name="IF"]'),
    ).toHaveLength(1);
  });

  it('records the authoritative function when formula-bar input changes the selected call', () => {
    const { formulaBar, palette } = setup();
    palette.open('ACOS');
    formulaBar.input.value = '=SUM(1)';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    paletteRoot(palette).querySelector<HTMLButtonElement>('[data-action="done"]')?.click();
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=SUM(1)');
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'number',
      value: 1,
    });
    palette.close();
    palette.open();
    expect(
      paletteRoot(palette).querySelectorAll('[data-section="recent"] [data-function-name="SUM"]'),
    ).toHaveLength(1);
    palette.close();
    expect(sheet.instance.undo()).toBe(true);
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBeNull();
    expect(sheet.instance.redo()).toBe(true);
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=SUM(1)');
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'number',
      value: 1,
    });
  });

  it('discards a stale draft inertly when the live workbook changes', async () => {
    const anchor = { sheet: 0, row: 0, col: 0 };
    sheet.workbook.setNumber(anchor, 1);
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, anchor);
    sheet.instance.history.clear();
    let liveWorkbook = sheet.workbook;
    const { formulaBar, palette } = setup(() => liveWorkbook);
    palette.open('ACOS');
    const arg = paletteRoot(palette).querySelector<HTMLInputElement>('[data-argument-index="0"]');
    expect(arg).not.toBeNull();
    if (arg) {
      arg.value = '1';
      arg.dispatchEvent(new Event('input', { bubbles: true }));
    }
    const nextWorkbook = await WorkbookHandle.createDefault();
    expect(nextWorkbook.isStub).toBe(false);
    liveWorkbook = nextWorkbook;
    paletteRoot(palette).querySelector<HTMLButtonElement>('[data-action="done"]')?.click();
    expect(palette.isOpen()).toBe(false);
    expect(formulaBar.input.value).toBe('=ACOS(1)');
    expect(sheet.workbook.cellFormula(anchor)).toBeNull();
    expect(nextWorkbook.cellFormula(anchor)).toBeNull();
    expect(sheet.instance.history.canUndo()).toBe(false);
    palette.close();
    nextWorkbook.dispose();
  });

  it('starts a fresh external draft for the first post-Done field edit', () => {
    const { beginDraft, formulaBar, palette } = setup();
    palette.open('ACOS');
    const root = paletteRoot(palette);
    const arg = root.querySelector<HTMLInputElement>('[data-argument-index="0"]');
    expect(arg).not.toBeNull();
    if (arg) {
      arg.value = '1';
      arg.dispatchEvent(new Event('input', { bubbles: true }));
    }
    root.querySelector<HTMLButtonElement>('[data-action="done"]')?.click();
    expect(beginDraft).toHaveBeenCalledTimes(1);
    const committedArg = paletteRoot(palette).querySelector<HTMLInputElement>(
      '[data-argument-index="0"]',
    );
    expect(committedArg).not.toBeNull();
    if (committedArg) {
      committedArg.value = '2';
      committedArg.dispatchEvent(new Event('input', { bubbles: true }));
    }
    expect(beginDraft).toHaveBeenCalledTimes(2);
    expect(formulaBar.input.value).toBe('=ACOS(2)');
    palette.close();
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=ACOS(1)');
  });

  it('restores the committed display and explains a refused post-Done draft', () => {
    const { beginDraft, formulaBar, palette } = setup();
    palette.open('ACOS');
    const root = paletteRoot(palette);
    const arg = root.querySelector<HTMLInputElement>('[data-argument-index="0"]');
    expect(arg).not.toBeNull();
    if (arg) {
      arg.value = '1';
      arg.dispatchEvent(new Event('input', { bubbles: true }));
    }
    root.querySelector<HTMLButtonElement>('[data-action="done"]')?.click();
    beginDraft.mockImplementation(() => null);
    const committedArg = paletteRoot(palette).querySelector<HTMLInputElement>(
      '[data-argument-index="0"]',
    );
    expect(committedArg).not.toBeNull();
    if (committedArg) {
      committedArg.value = '2';
      committedArg.dispatchEvent(new Event('input', { bubbles: true }));
    }
    expect(formulaBar.input.value).toBe('=ACOS(1)');
    expect(paletteRoot(palette).querySelector('[data-role="draft-conflict"]')?.textContent).toBe(
      defaultStrings.fxDialog.macPalette?.draftConflict,
    );
    expect(
      paletteRoot(palette).querySelector<HTMLInputElement>('[data-argument-index="0"]')?.value,
    ).toBe('1');
  });

  it('promotes a direct formula-bar accept to the committed palette state', () => {
    const { formulaBar, palette } = setup();
    palette.open('ACOS');
    const arg = paletteRoot(palette).querySelector<HTMLInputElement>('[data-argument-index="0"]');
    expect(arg).not.toBeNull();
    if (arg) {
      arg.value = '1';
      arg.dispatchEvent(new Event('input', { bubbles: true }));
    }
    formulaBar.controller.acceptFx();
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=ACOS(1)');
    expect(paletteRoot(palette).dataset.state).toBe('arguments-committed');
    expect(paletteRoot(palette).querySelector('[data-role="preview-value"]')?.textContent).toBe(
      '0',
    );
  });

  it('does not resurrect the committed pane when mirror cleanup closes it reentrantly', () => {
    const { mirror, palette } = setup();
    let closed = false;
    mirror.mockImplementation((_anchor, raw) => {
      if (raw === null && !closed) {
        closed = true;
        palette.close();
      }
    });
    palette.open('ACOS');
    const arg = paletteRoot(palette).querySelector<HTMLInputElement>('[data-argument-index="0"]');
    expect(arg).not.toBeNull();
    if (arg) {
      arg.value = '1';
      arg.dispatchEvent(new Event('input', { bubbles: true }));
    }
    paletteRoot(palette).querySelector<HTMLButtonElement>('[data-action="done"]')?.click();
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=ACOS(1)');
    expect(palette.isOpen()).toBe(false);
  });
});
