import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings/en.js';
import { ja } from '../../../src/i18n/strings/ja.js';
import { defaultStrings } from '../../../src/i18n/strings.js';
import { InlineEditor } from '../../../src/interact/editor.js';
import type {
  FormulaEditLease,
  FormulaEditLeaseContext,
} from '../../../src/interact/formula-edit-lease.js';
import {
  attachMacFormulaPalette,
  type MacFormulaArgumentHelp,
} from '../../../src/interact/mac-formula-palette.js';
import {
  attachFormulaBarController,
  type FormulaBarController,
} from '../../../src/mount/formula-bar.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/index.js';

const attachFormulaBarHarness = (
  sheet: MountedStubSheet,
): {
  controller: FormulaBarController;
  input: HTMLTextAreaElement;
  detach: () => void;
} => {
  const formulabar = document.createElement('div');
  const input = document.createElement('textarea');
  const cancel = document.createElement('button');
  const accept = document.createElement('button');
  formulabar.append(cancel, accept, input);
  sheet.host.appendChild(formulabar);
  const autocomplete = {
    isOpen: () => false,
    move: () => {},
    acceptHighlighted: () => false,
    close: () => {},
    refresh: () => {},
  };
  const argHelper = { close: () => {}, refresh: () => {} };
  const controller = attachFormulaBarController({
    formulabar,
    fxAccept: accept,
    fxCancel: cancel,
    fxInput: input,
    getArgHelper: () => argHelper,
    getAutocomplete: () => autocomplete,
    getStrings: () => defaultStrings,
    cancelBindingEditor: () => {},
    host: sheet.host,
    onValidation: vi.fn(),
    store: sheet.instance.store,
    updateChrome: () => {},
    wb: () => sheet.workbook,
  });
  return { controller, input, detach: controller.detach };
};

describe('interact/mac-formula-palette', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet({ workbook: await WorkbookHandle.createDefault() });
    expect(sheet.workbook.isStub).toBe(false);
  });

  afterEach(() => sheet.dispose());

  const setup = (
    getWorkbook: () => WorkbookHandle = () => sheet.workbook,
    getPaletteStrings: () => typeof defaultStrings = () => defaultStrings,
    getPaletteLocale: () => string = () => 'en-US',
    getArgumentHelp?: (
      name: string,
      index: number,
      locale: string,
    ) => MacFormulaArgumentHelp | null,
    suspend?: (
      formulaBar: FormulaBarController,
      context: FormulaEditLeaseContext,
    ) => FormulaEditLease | null,
  ) => {
    const opener = document.createElement('button');
    opener.textContent = 'open';
    sheet.host.appendChild(opener);
    opener.focus();
    const dock = document.createElement('div');
    dock.className = 'fc-host__taskpane-dock';
    sheet.host.appendChild(dock);
    const formulaBar = attachFormulaBarHarness(sheet);
    const beginDraft = vi.fn((...args: Parameters<FormulaBarController['beginExternalDraft']>) =>
      formulaBar.controller.beginExternalDraft(...args),
    );
    const mirror = vi.fn();
    const anchor = { sheet: 0, row: 0, col: 0 };
    const palette = attachMacFormulaPalette({
      host: sheet.host,
      dock,
      store: sheet.instance.store,
      getWb: getWorkbook,
      getLocale: getPaletteLocale,
      getStrings: getPaletteStrings,
      getAnchor: () => anchor,
      beginDraft,
      projectMirror: mirror,
      ...(getArgumentHelp ? { getArgumentHelp } : {}),
      ...(suspend
        ? { suspendActiveEdit: (context) => suspend(formulaBar.controller, context) }
        : {}),
    });
    return { anchor, beginDraft, dock, formulaBar, mirror, opener, palette };
  };

  it('opens a nonmodal complementary 300px pane with consecutive Recent and All sections', () => {
    const { dock, mirror, opener, palette } = setup();

    expect(dock.hidden).toBe(true);
    palette.open();

    const root = dock.querySelector<HTMLElement>('.fc-mac-formula-palette');
    expect(root).not.toBeNull();
    expect(root?.getAttribute('role')).toBe('complementary');
    expect(root?.hasAttribute('aria-modal')).toBe(false);
    expect(root?.dataset.state).toBe('picker');
    expect(root?.querySelector('[data-section="recent"]')).not.toBeNull();
    expect(root?.querySelector('[data-section="all"]')).not.toBeNull();
    expect(root?.querySelector('[data-action="insert-function"]')).not.toBeNull();
    expect(root?.querySelector('[data-action="close"] svg')).not.toBeNull();
    expect(root?.querySelector<HTMLButtonElement>('[data-action="close"]')?.title).toBe(
      defaultStrings.fxDialog.macPalette?.close,
    );
    expect(root?.style.inlineSize).toBe('300px');
    expect(palette.isOpen()).toBe(true);
    expect(dock.hidden).toBe(false);

    palette.close();
    expect(palette.isOpen()).toBe(false);
    expect(dock.hidden).toBe(true);
    expect(document.activeElement).toBe(opener);
    expect(mirror).toHaveBeenLastCalledWith(expect.any(Object), null);
    palette.open();
    palette.detach();
    expect(dock.hidden).toBe(true);
  });

  it('does not steal focus when closed, while an actual close restores the opener', () => {
    const { opener, palette } = setup();
    const outside = document.createElement('button');
    outside.textContent = 'outside';
    sheet.host.appendChild(outside);
    outside.focus();

    palette.close();
    expect(document.activeElement).toBe(outside);

    opener.focus();
    palette.open();
    palette.close();
    expect(document.activeElement).toBe(opener);
    outside.focus();
    palette.close();
    expect(document.activeElement).toBe(outside);
  });

  it('projects default, recent, and family picker categories without a selector', () => {
    const { palette } = setup();

    palette.open(undefined, { category: 'recent' });
    let root = paletteRoot(palette);
    expect(root.querySelectorAll('[data-section]')).toHaveLength(1);
    expect(root.querySelector('[data-section="recent"] h3')?.textContent).toBe(
      defaultStrings.fxDialog.categoryRecent,
    );
    expect(root.querySelector('[data-section="all"]')).toBeNull();
    expect(root.querySelector('select')).toBeNull();

    palette.close();
    palette.open(undefined, { category: 'logical' });
    root = paletteRoot(palette);
    expect(root.querySelectorAll('[data-section]')).toHaveLength(1);
    expect(root.querySelector('[data-section="logical"] h3')?.textContent).toBe(
      defaultStrings.fxDialog.categoryLogical,
    );
    expect(root.querySelector('[data-function-name="IF"]')).not.toBeNull();
    expect(root.querySelector('[data-function-name="SUM"]')).toBeNull();
    expect(root.querySelector('[data-section="all"]')).toBeNull();
    expect(root.querySelector('[data-section="recent"]')).toBeNull();
    expect(root.querySelector('select')).toBeNull();

    palette.close();
    palette.open(undefined, { category: 'all' });
    root = paletteRoot(palette);
    expect(root.querySelector('[data-section="recent"]')).not.toBeNull();
    expect(root.querySelector('[data-section="all"]')).not.toBeNull();

    palette.close();
    palette.open('IF', { category: 'recent' });
    expect(paletteRoot(palette).dataset.state).toBe('arguments-editing');
    expect(paletteRoot(palette).querySelector('[data-section]')).toBeNull();
  });

  it.each([
    {
      locale: 'en-US',
      strings: en,
      functionDescription: 'arccosine',
      syntaxLabel: 'Syntax',
    },
    {
      locale: 'ja-JP',
      strings: ja,
      functionDescription: 'アークコサイン',
      syntaxLabel: '構文',
    },
  ])(
    'shows localized catalog help without argument help for a real workbook in $locale',
    ({ locale, strings, functionDescription, syntaxLabel }) => {
      const { palette } = setup(
        () => sheet.workbook,
        () => strings,
        () => locale,
      );
      palette.open('ACOS');

      const root = paletteRoot(palette);
      const argument = root.querySelector('.fc-mac-formula-palette__argument');
      expect(argument?.querySelector('span')?.textContent).toBe('number');
      expect(argument?.querySelector('small')).toBeNull();

      const help = root.querySelector('.fc-mac-formula-palette__help');
      const helpText = help?.querySelector('p')?.textContent ?? '';
      const syntaxText = help?.querySelectorAll('p')[1]?.textContent ?? '';
      expect(helpText).toContain(functionDescription);
      expect(syntaxText).toBe(`${syntaxLabel}: ACOS(number)`);
      expect(help?.querySelector('a')).toBeNull();
      palette.close();
    },
  );

  it('prefers workbook metadata and argument-help provider values', () => {
    sheet.workbook.setFunctionMetadataProvider({
      ACOS: { signature: 'ACOS(workbookNumber)', description: 'Workbook arccosine help.' },
    });
    const getArgumentHelp = vi.fn((name: string, index: number) =>
      name === 'ACOS' && index === 0
        ? {
            label: 'Provided number',
            description: 'Provider argument details.',
            url: 'https://example.com/acos-help',
          }
        : null,
    );
    const { palette } = setup(
      () => sheet.workbook,
      () => en,
      () => 'en-US',
      getArgumentHelp,
    );
    palette.open('ACOS');

    const root = paletteRoot(palette);
    const argument = root.querySelector('.fc-mac-formula-palette__argument');
    expect(argument?.querySelector('span')?.textContent).toBe('Provided number');
    expect(argument?.querySelector('small')?.textContent).toBe('Provider argument details.');
    const help = root.querySelector('.fc-mac-formula-palette__help');
    expect(help?.textContent).toContain('Workbook arccosine help.');
    expect(help?.textContent).toContain('Syntax: ACOS(workbookNumber)');
    expect(help?.querySelector('a')?.getAttribute('href')).toBe('https://example.com/acos-help');
    expect(getArgumentHelp).toHaveBeenCalledWith('ACOS', 0, 'en-US');
    palette.close();
  });

  it('keeps unsynchronized raw authoritative across a visible category request and picker insert', () => {
    const { formulaBar, palette } = setup();
    palette.open('ACOS');
    formulaBar.input.value = '=ACOS(';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));

    palette.open(undefined, { category: 'logical' });
    const root = paletteRoot(palette);
    expect(root.dataset.state).toBe('picker');
    expect(root.dataset.rawSynchronized).toBe('false');
    expect(formulaBar.input.value).toBe('=ACOS(');
    root.querySelector<HTMLElement>('[data-function-name="IF"]')?.click();
    root.querySelector<HTMLButtonElement>('[data-action="insert-function"]')?.click();
    expect(formulaBar.input.value).toBe('=ACOS(');
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

  it('ignores detached picker rows and reflects live metadata changes before Insert', () => {
    let metadata = {
      name: 'TEMP',
      localizedName: 'Old TEMP',
      description: 'old description',
      minArity: 1,
      maxArity: 1,
      availability: 0,
    };
    const workbook = {
      functionNames: () => ['TEMP'],
      functionMetadata: () => metadata,
    } as unknown as WorkbookHandle;
    const { palette } = setup(() => workbook);
    palette.open();
    const firstRoot = paletteRoot(palette);
    const staleRow = firstRoot.querySelector<HTMLElement>('[data-function-name="TEMP"]');
    expect(staleRow).not.toBeNull();
    staleRow?.click();
    expect(paletteRoot(palette).textContent).toContain('Old TEMP');

    metadata = { ...metadata, localizedName: 'New TEMP', description: 'new description' };
    palette.refresh();
    const summary = paletteRoot(palette).querySelector('.fc-mac-formula-palette__summary');
    expect(summary?.textContent).toContain('New TEMP');
    expect(summary?.textContent).toContain('new description');
    expect(summary?.textContent).not.toContain('Old TEMP');
    expect(summary?.textContent).not.toContain('old description');

    metadata = { ...metadata, availability: 3 };
    palette.refresh();
    const refreshedRoot = paletteRoot(palette);
    expect(
      refreshedRoot.querySelector<HTMLButtonElement>('[data-action="insert-function"]')?.disabled,
    ).toBe(true);
    staleRow?.click();
    expect(refreshedRoot.querySelector('.fc-mac-formula-palette__guard')).toBeNull();
    expect(
      refreshedRoot.querySelector<HTMLButtonElement>('[data-action="insert-function"]')?.disabled,
    ).toBe(true);
    palette.detach();
  });

  it('uses localized date-time and statistical family titles', () => {
    const { palette } = setup(
      () => sheet.workbook,
      () => ja,
      () => 'ja-JP',
    );
    palette.open(undefined, { category: 'datetime' });
    const root = paletteRoot(palette);
    expect(root.querySelector('[data-section="datetime"] h3')?.textContent).toBe(
      ja.fxDialog.categoryDateTime,
    );
    expect(root.querySelector('[data-section="datetime"] h3')?.textContent).not.toBe('Date & Time');
    palette.open(undefined, { category: 'statistical' });
    expect(paletteRoot(palette).querySelector('[data-section="statistical"] h3')?.textContent).toBe(
      '統計',
    );
  });

  it('does not let picker insertion overwrite compound or malformed initial raw drafts', () => {
    const { formulaBar, palette } = setup();
    for (const raw of ['=SUM(1)+2', '=SUM(']) {
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
    formulaBar.input.value = '=ACOS(';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    palette.open('COUNTIF');
    expect(formulaBar.input.value).toBe('=ACOS(');
    expect(paletteRoot(palette).dataset.rawSynchronized).toBe('false');

    palette.open('MISSING_FUNCTION');
    expect(formulaBar.input.value).toBe('=ACOS(');
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
    expect(paletteRoot(palette).dataset.state).toBe('picker');
    expect(formulaBar.input.value).toBe('');
    expect(paletteRoot(palette).querySelector('.fc-mac-formula-palette__guard')?.textContent).toBe(
      defaultStrings.fxDialog.macPalette?.unavailable,
    );
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

  it('round-trips nested and quoted arguments, preserves unsynchronized raw input, and exposes the focused range target', () => {
    const { anchor, formulaBar, palette } = setup();
    palette.open('IF');
    const root = paletteRoot(palette);
    expect(root.dataset.state).toBe('arguments-editing');
    const first = root.querySelector<HTMLInputElement>('[data-argument-index="0"]');
    expect(first).not.toBeNull();
    const liveTarget = palette.rangeInsertTarget();
    expect(liveTarget?.isFormulaEdit()).toBe(true);
    liveTarget?.insertRefAtCaret('B3');
    expect(formulaBar.input.value).toBe('=IF(B3)');
    formulaBar.input.value = '=IF(SUM(A1,A2)>1,"a,b",ACOS(1))';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="0"]')?.value).toBe(
      'SUM(A1,A2)>1',
    );
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="1"]')?.value).toBe('"a,b"');
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="2"]')?.value).toBe(
      'ACOS(1)',
    );

    formulaBar.input.value = '=ACOS(';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    expect(formulaBar.input.value).toBe('=ACOS(');
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="0"]')?.value).toBe(
      'SUM(A1,A2)>1',
    );

    const target = palette.rangeInsertTarget();
    expect(target?.isFormulaEdit()).toBe(false);
    expect(target).not.toBeNull();
    target?.insertRefAtCaret('B3');
    expect(formulaBar.input.value).toBe('=ACOS(');
    expect(anchor).toEqual({ sheet: 0, row: 0, col: 0 });
  });

  it('preserves explicit trailing blank argument slots when reverse-projected raw is edited', () => {
    const { formulaBar, palette } = setup();
    palette.open('IF');
    const root = paletteRoot(palette);
    formulaBar.input.value = '=IF(FALSE,1,)';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="2"]')?.value).toBe('');

    const valueIfTrue = root.querySelector<HTMLInputElement>('[data-argument-index="1"]');
    expect(valueIfTrue).not.toBeNull();
    if (valueIfTrue) {
      valueIfTrue.value = '2';
      valueIfTrue.dispatchEvent(new Event('input', { bubbles: true }));
    }
    expect(formulaBar.input.value).toBe('=IF(FALSE,2,)');

    formulaBar.input.value = '=SUM({1,2},3)';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    expect(root.dataset.rawSynchronized).toBe('false');
    expect(formulaBar.input.value).toBe('=SUM({1,2},3)');
    palette.close();
  });

  it('keeps palette-owned multiargument edits and range insertion in the assembled formula', () => {
    const { formulaBar, palette } = setup();
    palette.open('IF');
    const root = paletteRoot(palette);
    const valueIfTrue = root.querySelector<HTMLInputElement>('[data-argument-index="1"]');
    const valueIfFalse = root.querySelector<HTMLInputElement>('[data-argument-index="2"]');
    expect(valueIfTrue).not.toBeNull();
    expect(valueIfFalse).not.toBeNull();
    if (valueIfTrue) {
      valueIfTrue.value = '1';
      valueIfTrue.dispatchEvent(new Event('input', { bubbles: true }));
    }
    if (valueIfFalse) {
      valueIfFalse.value = '2';
      valueIfFalse.dispatchEvent(new Event('input', { bubbles: true }));
    }
    expect(formulaBar.input.value).toBe('=IF(,1,2)');
    palette.rangeInsertTarget()?.insertRefAtCaret('B3');
    expect(formulaBar.input.value).toBe('=IF(,1,B3)');
  });

  it('grows externally projected IF arguments through fields and range insertion', () => {
    const { formulaBar, palette } = setup();
    palette.open('IF');
    const root = paletteRoot(palette);
    formulaBar.input.value = '=IF(TRUE)';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    const valueIfTrue = root.querySelector<HTMLInputElement>('[data-argument-index="1"]');
    const valueIfFalse = root.querySelector<HTMLInputElement>('[data-argument-index="2"]');
    expect(valueIfTrue).not.toBeNull();
    expect(valueIfFalse).not.toBeNull();
    if (valueIfTrue) {
      valueIfTrue.focus();
      valueIfTrue.value = '1';
      valueIfTrue.dispatchEvent(new Event('input', { bubbles: true }));
      expect(root.querySelector('[data-argument-index="1"]')).toBe(valueIfTrue);
    }
    expect(formulaBar.input.value).toBe('=IF(TRUE,1)');
    if (valueIfFalse) {
      valueIfFalse.focus();
      valueIfFalse.value = '';
      valueIfFalse.dispatchEvent(new Event('input', { bubbles: true }));
    }
    expect(formulaBar.input.value).toBe('=IF(TRUE,1,)');
    palette.rangeInsertTarget()?.insertRefAtCaret('B3');
    expect(formulaBar.input.value).toBe('=IF(TRUE,1,B3)');
  });

  it('reverse-projects array and structured-reference fields while preserving edits and blanks', () => {
    const { formulaBar, palette } = setup();
    palette.open('SUM');
    const root = paletteRoot(palette);
    formulaBar.input.value = '=SUM({1,2;3,4},Table1[[Last, First],[Amount]],)';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    expect(root.dataset.rawSynchronized).toBe('true');
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="0"]')?.value).toBe(
      '{1,2;3,4}',
    );
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="1"]')?.value).toBe(
      'Table1[[Last, First],[Amount]]',
    );
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="2"]')?.value).toBe('');

    const array = root.querySelector<HTMLInputElement>('[data-argument-index="0"]');
    expect(array).not.toBeNull();
    if (array) {
      array.value = '{5,6;7,8}';
      array.dispatchEvent(new Event('input', { bubbles: true }));
    }
    expect(formulaBar.input.value).toBe('=SUM({5,6;7,8},Table1[[Last, First],[Amount]],)');

    const structuredReference = root.querySelector<HTMLInputElement>('[data-argument-index="1"]');
    expect(structuredReference).not.toBeNull();
    structuredReference?.focus();
    palette.rangeInsertTarget()?.insertRefAtCaret('B3');
    expect(formulaBar.input.value).toBe('=SUM({5,6;7,8},B3,)');
  });

  it('fails closed for crossed, unclosed, or different-name raw formulas', () => {
    const { formulaBar, palette } = setup();
    palette.open('SUM');
    const root = paletteRoot(palette);

    formulaBar.input.value = '=SUM({1,2],3)';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    expect(root.dataset.rawSynchronized).toBe('false');
    expect(formulaBar.input.value).toBe('=SUM({1,2],3)');
    root.querySelector<HTMLButtonElement>('[data-action="add-argument"]')?.click();
    expect(root.dataset.rawSynchronized).toBe('false');
    expect(formulaBar.input.value).toBe('=SUM({1,2],3)');

    formulaBar.input.value = '=SUM(Table1[[Last, First],3)';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    expect(root.dataset.rawSynchronized).toBe('false');
    expect(formulaBar.input.value).toBe('=SUM(Table1[[Last, First],3)');

    formulaBar.input.value = '=AVERAGE({1,2},3)';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    expect(root.dataset.rawSynchronized).toBe('false');
    expect(formulaBar.input.value).toBe('=AVERAGE({1,2},3)');
  });

  it('grows variadic arguments without dropping an existing trailing blank slot', () => {
    const { formulaBar, palette } = setup();
    palette.open('SUM');
    const root = paletteRoot(palette);
    formulaBar.input.value = '=SUM(1,)';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    root.querySelector<HTMLButtonElement>('[data-action="add-argument"]')?.click();
    expect(formulaBar.input.value).toBe('=SUM(1,,)');
    const second = paletteRoot(palette).querySelector<HTMLInputElement>(
      '[data-argument-index="1"]',
    );
    expect(second).not.toBeNull();
    if (second) {
      second.value = '2';
      second.dispatchEvent(new Event('input', { bubbles: true }));
    }
    expect(formulaBar.input.value).toBe('=SUM(1,2,)');
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

  it('does not record the selected function when an authoritative raw commit changes function', () => {
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
      paletteRoot(palette).querySelectorAll('[data-section="recent"] [data-function-name]'),
    ).toHaveLength(0);
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

  it('cancels the stale draft and returns to a guarded picker when the live workbook changes', async () => {
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
    expect(paletteRoot(palette).dataset.state).toBe('picker');
    expect(formulaBar.input.value).toBe('');
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

  it('keeps unavailable class 3 entries readable but blocks Insert and Done', () => {
    const store = createSpreadsheetStore();
    const dock = document.createElement('div');
    const host = document.createElement('div');
    host.tabIndex = -1;
    document.body.append(host);
    document.body.append(dock);
    const wb = {
      functionNames: () => ['FUTURE'],
      functionMetadata: () => ({ name: 'FUTURE', minArity: 1, maxArity: 1, availability: 3 }),
    } as unknown as WorkbookHandle;
    const beginDraft = vi.fn(() => null);
    const palette = attachMacFormulaPalette({
      host,
      dock,
      store,
      getWb: () => wb,
      getLocale: () => 'en-US',
      getStrings: () => defaultStrings,
      getAnchor: () => ({ sheet: 0, row: 0, col: 0 }),
      beginDraft,
      projectMirror: vi.fn(),
    });

    palette.open();
    const root = paletteRoot(palette);
    beginDraft.mockClear();
    const future = root.querySelector<HTMLButtonElement>('[data-function-name="FUTURE"]');
    expect(future?.getAttribute('aria-disabled')).toBe('true');
    expect(future?.getAttribute('aria-description')).toBe(
      defaultStrings.fxDialog.macPalette?.unavailable,
    );
    future?.click();
    expect(root.querySelector<HTMLButtonElement>('[data-action="insert-function"]')?.disabled).toBe(
      true,
    );
    root.querySelector<HTMLButtonElement>('[data-action="insert-function"]')?.click();
    expect(paletteRoot(palette).dataset.state).toBe('picker');
    expect(beginDraft).not.toHaveBeenCalled();
    palette.detach();
    host.remove();
    dock.remove();
  });

  it('disables Insert after the selected catalog entry disappears', () => {
    const store = createSpreadsheetStore();
    const dock = document.createElement('div');
    const host = document.createElement('div');
    document.body.append(host, dock);
    let names = ['TEMP'];
    const wb = {
      functionNames: () => names,
      functionMetadata: () => ({ name: 'TEMP', minArity: 1, maxArity: 1, availability: 1 }),
    } as unknown as WorkbookHandle;
    const palette = attachMacFormulaPalette({
      host,
      dock,
      store,
      getWb: () => wb,
      getLocale: () => 'en-US',
      getStrings: () => defaultStrings,
      getAnchor: () => ({ sheet: 0, row: 0, col: 0 }),
      beginDraft: vi.fn(() => null),
      projectMirror: vi.fn(),
    });
    palette.open();
    paletteRoot(palette).querySelector<HTMLButtonElement>('[data-function-name="TEMP"]')?.click();
    names = [];
    palette.refresh();
    expect(
      paletteRoot(palette).querySelector<HTMLButtonElement>('[data-action="insert-function"]')
        ?.disabled,
    ).toBe(true);
    palette.detach();
    host.remove();
    dock.remove();
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

  describe('suspended edit lease', () => {
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
});

const paletteRoot = (palette: { isOpen(): boolean }): HTMLElement => {
  if (!palette.isOpen()) throw new Error('palette is not open');
  const root = document.querySelector<HTMLElement>('.fc-mac-formula-palette');
  if (!root) throw new Error('palette root is missing');
  return root;
};
