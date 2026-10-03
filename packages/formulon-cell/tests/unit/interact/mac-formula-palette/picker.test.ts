import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { en } from '../../../../src/i18n/strings/en.js';
import { ja } from '../../../../src/i18n/strings/ja.js';
import { defaultStrings } from '../../../../src/i18n/strings.js';
import { attachMacFormulaPalette } from '../../../../src/interact/mac-formula-palette.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/index.js';
import { type PaletteSetupArgs, paletteRoot, setupPalette } from './fixtures.js';

describe('interact/mac-formula-palette picker', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet({ workbook: await WorkbookHandle.createDefault() });
    expect(sheet.workbook.isStub).toBe(false);
  });

  afterEach(() => sheet.dispose());

  const setup = (...args: PaletteSetupArgs) => setupPalette(sheet, ...args);

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
});
