import { afterEach, describe, expect, it, vi } from 'vitest';

import { getRecentFunctions } from '../../../src/commands/function-history.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { defaultStrings } from '../../../src/i18n/strings.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/index.js';

/**
 * Per-feature opt-out tests. `presets.*` covers bundle behavior; this spec
 * locks down that EACH individual flag (set to `false` while the others
 * remain default) drops just its own attach.
 */
describe('mount/host-features — individual feature flags', () => {
  let sheet: MountedStubSheet | undefined;

  afterEach(() => {
    sheet?.dispose();
    sheet = undefined;
  });

  it('formatDialog: false omits the format dialog handle', async () => {
    sheet = await mountStubSheet({ features: { formatDialog: false } });
    expect(sheet.instance.features.formatDialog).toBeFalsy();
  });

  it('findReplace: false omits the find/replace handle', async () => {
    sheet = await mountStubSheet({ features: { findReplace: false } });
    expect(sheet.instance.features.findReplace).toBeFalsy();
  });

  it('fxDialog: false disables the formula-bar fx button with a reason', async () => {
    sheet = await mountStubSheet({ features: { fxDialog: false } });
    expect(sheet.instance.features.fxDialog).toBeFalsy();
    const fx = sheet.host.querySelector<HTMLButtonElement>('.fc-host__formulabar-fx');
    expect(fx?.disabled).toBe(true);
    expect(fx?.dataset.disabledReason).toBe(defaultStrings.fxDialog.fxButtonUnavailable);
    expect(fx?.getAttribute('aria-description')).toBe(defaultStrings.fxDialog.fxButtonUnavailable);
  });

  it('keeps function arguments open and unrecorded when formula-bar commit is rejected', async () => {
    sheet = await mountStubSheet({ features: { fxDialog: true } });
    expect(sheet.instance.features.fxDialog).toBeTruthy();
    mutators.setSheetProtected(sheet.instance.store, 0, true);
    sheet.instance.openFunctionArguments('SUM');
    const overlay = document.querySelector<HTMLElement>('.fc-fxdialog');
    const insertBtn = document.querySelector<HTMLButtonElement>(
      '.fc-fxdialog .fc-fmtdlg__btn--primary',
    );

    insertBtn?.click();

    expect(overlay?.hidden).toBe(false);
    expect(document.activeElement).toBe(insertBtn);
    expect(getRecentFunctions(sheet.instance.store)).toEqual([]);
    expect(sheet.host.querySelector<HTMLTextAreaElement>('.fc-host__formulabar-input')?.value).toBe(
      '=SUM()',
    );
  });

  it('closes and records a function after formula-bar commit is accepted', async () => {
    sheet = await mountStubSheet({ features: { fxDialog: true } });
    expect(sheet.instance.features.fxDialog).toBeTruthy();
    sheet.instance.openFunctionArguments('SUM');
    const argInput = document.querySelector<HTMLInputElement>('.fc-fxdialog__arg-input');
    if (!argInput) throw new Error('expected SUM argument input');
    argInput.value = '1';
    argInput.dispatchEvent(new Event('input'));

    document.querySelector<HTMLButtonElement>('.fc-fxdialog .fc-fmtdlg__btn--primary')?.click();

    expect(document.querySelector<HTMLElement>('.fc-fxdialog')?.hidden).toBe(true);
    expect(getRecentFunctions(sheet.instance.store)).toEqual(['SUM']);
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=SUM(1)');
  });

  it('uses the Mac complementary pane, preserves selection seeding, and refreshes in place', async () => {
    const workbook = await WorkbookHandle.createDefault();
    expect(workbook.isStub).toBe(false);
    sheet = await mountStubSheet({
      workbook,
      ui: { platform: 'mac' },
      features: { fxDialog: true },
      locale: 'en',
    });
    const fx = sheet.host.querySelector<HTMLButtonElement>('.fc-host__formulabar-fx');
    const handle = sheet.instance.features.fxDialog;
    expect(handle).toBeTruthy();
    expect(sheet.host.querySelector('.fc-fxdialog')).toBeNull();
    expect(sheet.host.querySelector('.fc-mac-formula-palette')).toBeTruthy();

    mutators.setRange(sheet.instance.store, {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 0,
      c1: 1,
    });
    fx?.focus();
    sheet.instance.openFunctionArguments(undefined, { category: 'all' });

    const pickerRoot = sheet.host.querySelector<HTMLElement>('.fc-mac-formula-palette');
    expect(pickerRoot?.querySelector('[data-section="recent"]')).toBeTruthy();
    expect(pickerRoot?.querySelector('[data-section="all"]')).toBeTruthy();
    pickerRoot?.querySelector<HTMLButtonElement>('[data-action="close"]')?.click();
    fx?.focus();
    sheet.instance.openFunctionArguments('SUM', { category: 'all' });

    const root = sheet.host.querySelector<HTMLElement>('.fc-mac-formula-palette');
    const dock = sheet.host.querySelector<HTMLElement>('.fc-host__taskpane-dock');
    const firstArgument = root?.querySelector<HTMLInputElement>('[data-argument-index="0"]');
    expect(root?.dataset.state).toBe('arguments-editing');
    expect(firstArgument?.value).toBe('A1:B1');
    expect(dock?.hidden).toBe(false);
    const englishTitle = root?.getAttribute('aria-label');
    expect(englishTitle).toBeTruthy();

    sheet.instance.i18n.setLocale('ja');

    expect(sheet.host.querySelector('.fc-mac-formula-palette')).toBe(root);
    expect(sheet.instance.features.fxDialog).toBe(handle);
    expect(root?.getAttribute('aria-label')).not.toBe(englishTitle);
    root?.querySelector<HTMLButtonElement>('[data-action="close"]')?.click();
    expect(root?.hidden).toBe(true);
    expect(dock?.hidden).toBe(true);
    expect(document.activeElement).toBe(fx);
  });

  it('suspends a mid-edit formula bar while the Mac palette is open and restores it on close', async () => {
    const workbook = await WorkbookHandle.createDefault();
    sheet = await mountStubSheet({
      workbook,
      ui: { platform: 'mac' },
      features: { fxDialog: true },
    });
    const anchor = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(sheet.instance.store, anchor);
    const fx = sheet.host.querySelector<HTMLButtonElement>('.fc-host__formulabar-fx');
    const fxInput = sheet.host.querySelector<HTMLTextAreaElement>('.fc-host__formulabar-input');
    expect(fx).toBeTruthy();
    expect(fxInput).toBeTruthy();
    fxInput?.focus();
    if (fxInput) fxInput.value = '=2*';
    fxInput?.dispatchEvent(new Event('input', { bubbles: true }));

    const press = new MouseEvent('mousedown', { bubbles: true, cancelable: true });
    fx?.dispatchEvent(press);
    expect(press.defaultPrevented).toBe(true);
    fx?.click();
    const root = sheet.host.querySelector<HTMLElement>('.fc-mac-formula-palette');
    expect(root?.hidden).toBe(false);
    expect(workbook.cellFormula(anchor)).toBeNull();

    root?.querySelector<HTMLButtonElement>('[data-action="close"]')?.click();
    expect(root?.hidden).toBe(true);
    expect(fxInput?.value).toBe('=2*');
    expect(document.activeElement).toBe(fxInput);
    expect(workbook.cellFormula(anchor)).toBeNull();
  });

  it('cancels an anchored Mac draft on selection, policy, and feature teardown without a write', async () => {
    const workbook = await WorkbookHandle.createDefault();
    expect(workbook.isStub).toBe(false);
    sheet = await mountStubSheet({
      workbook,
      ui: { platform: 'mac' },
      features: { fxDialog: true },
    });
    const anchor = { sheet: 0, row: 0, col: 0 };
    workbook.setNumber(anchor, 7);
    mutators.replaceCells(sheet.instance.store, workbook.cells(0));
    mutators.setActive(sheet.instance.store, anchor);
    sheet.instance.history.clear();

    const root = sheet.host.querySelector<HTMLElement>('.fc-mac-formula-palette');
    sheet.instance.openFunctionArguments('SUM');
    expect(root?.hidden).toBe(false);
    expect(workbook.cellFormula(anchor)).toBeNull();
    expect(workbook.getValue(anchor)).toEqual({ kind: 'number', value: 7 });
    expect(sheet.instance.history.canUndo()).toBe(false);

    mutators.setActive(sheet.instance.store, { sheet: 0, row: 1, col: 0 });
    expect(root?.hidden).toBe(true);
    expect(workbook.cellFormula(anchor)).toBeNull();
    expect(sheet.instance.history.canUndo()).toBe(false);

    mutators.setActive(sheet.instance.store, anchor);
    sheet.instance.openFunctionArguments('SUM');
    sheet.instance.setPolicy({ readOnly: true });
    expect(root?.hidden).toBe(true);
    expect(workbook.cellFormula(anchor)).toBeNull();
    expect(sheet.instance.history.canUndo()).toBe(false);

    sheet.instance.setPolicy(undefined);
    sheet.instance.setFeatures({ fxDialog: true });
    sheet.instance.openFunctionArguments('SUM');
    sheet.instance.setFeatures({ fxDialog: false });
    expect(sheet.instance.features.fxDialog).toBeFalsy();
    expect(root?.isConnected).toBe(false);
    expect(workbook.cellFormula(anchor)).toBeNull();
    expect(sheet.instance.history.canUndo()).toBe(false);
  });

  it('threads getFunctionArgumentHelp into the Mac palette with the active locale', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const getFunctionArgumentHelp = vi.fn((name: string, index: number, locale: string) =>
      name === 'ACOS' && index === 0
        ? {
            description: locale.startsWith('ja') ? '-1 から 1 の数値' : 'Number from -1 to 1.',
            url: 'https://example.com/acos',
          }
        : null,
    );
    sheet = await mountStubSheet({
      workbook,
      ui: { platform: 'mac' },
      features: { fxDialog: true },
      locale: 'en',
      getFunctionArgumentHelp,
    });

    sheet.instance.openFunctionArguments('ACOS');
    const root = sheet.host.querySelector<HTMLElement>('.fc-mac-formula-palette');
    const argument = root?.querySelector('.fc-mac-formula-palette__argument');
    expect(argument?.querySelector('small')?.textContent).toBe('Number from -1 to 1.');
    expect(root?.querySelector('.fc-mac-formula-palette__help a')?.getAttribute('href')).toBe(
      'https://example.com/acos',
    );
    expect(getFunctionArgumentHelp).toHaveBeenCalledWith('ACOS', 0, 'en');
    root?.querySelector<HTMLButtonElement>('[data-action="close"]')?.click();

    sheet.instance.i18n.setLocale('ja');
    sheet.instance.openFunctionArguments('ACOS');
    expect(getFunctionArgumentHelp).toHaveBeenCalledWith('ACOS', 0, 'ja');
    expect(root?.querySelector('.fc-mac-formula-palette__argument small')?.textContent).toBe(
      '-1 から 1 の数値',
    );
  });

  it('shows no argument hint or help link when getFunctionArgumentHelp is absent', async () => {
    const workbook = await WorkbookHandle.createDefault();
    sheet = await mountStubSheet({
      workbook,
      ui: { platform: 'mac' },
      features: { fxDialog: true },
      locale: 'en',
    });

    sheet.instance.openFunctionArguments('ACOS');
    const root = sheet.host.querySelector<HTMLElement>('.fc-mac-formula-palette');
    expect(root?.dataset.state).toBe('arguments-editing');
    expect(root?.querySelector('.fc-mac-formula-palette__argument small')).toBeNull();
    expect(root?.querySelector('.fc-mac-formula-palette__help a')).toBeNull();
  });

  it('passes the mounted live workbook catalog to Function Arguments', async () => {
    const workbook = await WorkbookHandle.createDefault();
    expect(workbook.isStub).toBe(false);
    sheet = await mountStubSheet({ workbook, features: { fxDialog: true } });

    sheet.instance.openFunctionArguments('ACOS');

    expect(document.querySelector<HTMLElement>('[data-fx-name="ACOS"]')).toBeTruthy();
    expect(document.querySelector<HTMLElement>('.fc-fxdialog__picker')?.hidden).toBe(true);
    expect(document.querySelector<HTMLElement>('.fc-fxdialog__args')?.hidden).toBe(false);
    expect(document.querySelectorAll<HTMLInputElement>('.fc-fxdialog__arg-input')).toHaveLength(1);
  });

  it('viewToolbar: false omits the .fc-viewbar chrome', async () => {
    sheet = await mountStubSheet({ features: { viewToolbar: false } });
    expect(sheet.host.querySelector('.fc-viewbar')).toBeNull();
    expect(sheet.instance.features.viewToolbar).toBeFalsy();
  });

  it('sheetTabs: false omits the sheet-tab chrome', async () => {
    sheet = await mountStubSheet({ features: { sheetTabs: false } });
    expect(sheet.host.querySelector('.fc-host__sheetbar-tabs')).toBeNull();
  });

  it('statusBar: false omits the statusbar chrome', async () => {
    sheet = await mountStubSheet({ features: { statusBar: false } });
    expect(sheet.host.querySelector('.fc-host__statusbar')).toBeNull();
  });

  it('formulaBar: false omits the formula bar chrome', async () => {
    sheet = await mountStubSheet({ features: { formulaBar: false } });
    expect(sheet.host.querySelector('.fc-host__formulabar')).toBeNull();
  });

  it('conditional + iterative + namedRanges: false drops only those features', async () => {
    sheet = await mountStubSheet({
      features: { conditional: false, iterative: false, namedRanges: false },
    });
    const f = sheet.instance.features;
    expect(f.conditional).toBeFalsy();
    expect(f.iterative).toBeFalsy();
    expect(f.namedRanges).toBeFalsy();
    // Other defaults still on.
    expect(f.findReplace).toBeTruthy();
    expect(f.statusBar).toBeTruthy();
  });

  it('default mount enables charts + pivot dialog + workbook objects', async () => {
    sheet = await mountStubSheet();
    const f = sheet.instance.features;
    expect(f.charts).toBeTruthy();
    expect(f.pivotTableDialog).toBeTruthy();
    expect(f.workbookObjects).toBeTruthy();
  });
});
