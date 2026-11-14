import { expect, type Locator, type Page } from '@playwright/test';

import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type Platform = 'default' | 'mac';
type ProbeWindow = { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };

/** Where each profile puts the Formulas menus the scenario drives. */
const PROFILE = {
  default: {
    errorMenu: '#menu-error-checking',
    circularRow: '[data-format-submenu="circular"]',
    circularPanel: '#menu-error-checking-circular',
    calcCommand: 'calcOptions',
    calcMenu: '#menu-calc-options',
    auto: '[data-calc-option="auto"]',
    manual: '[data-calc-option="manual"]',
    formatStale: '[data-calc-option="format-stale"]',
    compatRow: '[data-format-submenu="compat-version"]',
    compatPanel: '#menu-calc-options-compat-version',
    compat: (id: string) => `[data-calc-option="compat-${id}"]`,
  },
  mac: {
    errorMenu: '#menu-mac-formulas-error-check',
    circularRow: '[data-format-submenu="mac.formulas.errorCheck.circular"]',
    circularPanel: '#menu-mac-formulas-error-check-circular',
    calcCommand: 'mac.formulas.calcOptions',
    calcMenu: '#menu-mac-formulas-calc-options',
    auto: '[data-ribbon-command="mac.formulas.calc.auto"]',
    manual: '[data-ribbon-command="mac.formulas.calc.manual"]',
    formatStale: '[data-ribbon-command="mac.formulas.calc.formatStale"]',
    compatRow: '[data-format-submenu="mac.formulas.calc.compatVersion"]',
    compatPanel: '#menu-mac-formulas-calc-options-compat-version',
    compat: (id: string) =>
      `[data-ribbon-command="mac.formulas.calc.compatVersion.${id === 'learn-more' ? 'learnMore' : id}"]`,
  },
} as const;

/** Error Checking › Circular References lists the cycle and selects a member; Calculation Options
 *  carries Format Stale Values and the Compatibility Version, which undo and an XLSX round trip
 *  keep. */
export async function runCalcOptionsMenusScenario(page: Page, platform: Platform): Promise<void> {
  const p = PROFILE[platform];
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform });
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.expectNoStub();
  await page.evaluate(() => {
    const inst = (window as unknown as ProbeWindow).__fcInst;
    const sheet = inst.workbook.sheet(0);
    inst.workbook.batch(() => {
      sheet.setFormula(0, 1, '=C1');
      sheet.setFormula(0, 2, '=B1');
      sheet.setNumber(1, 0, 1);
      sheet.setFormula(1, 1, '=A2*2');
      // U+1D49C is one code point and two UTF-16 units.
      sheet.setFormula(2, 0, '=LEN("\u{1D49C}")');
    });
    inst.history.clear();
  });
  await selectRibbonTab(page, 'formulas');

  const errorMenu = page.locator(p.errorMenu);
  const openErrorMenu = async (): Promise<void> => {
    if (platform === 'mac') await sp.clickRibbon('mac.formulas.errorCheck');
    // The default profile's primary face runs Error Checking; the chevron opens the menu.
    else
      await page
        .locator('[data-ribbon-command="errorChecking"] .fc-tb__rb-split-chevron')
        .first()
        .click();
    await expect(errorMenu).toBeVisible();
  };

  // Circular References: the cycle's cells, the active one checked; picking one selects it.
  await sp.goTo('C1');
  await openErrorMenu();
  await errorMenu.locator(p.circularRow).click();
  const circular = page.locator(p.circularPanel);
  await expect(circular).toBeVisible();
  const entries = circular.locator('[data-circular-reference]');
  await expect(entries).toHaveText(['$B$1', '$C$1']);
  await expect(entries.nth(0)).toHaveAttribute('role', 'menuitemradio');
  await expect(entries.nth(0)).toHaveAttribute('data-circular-reference', '0:0:1');
  await expect(entries.nth(1)).toHaveAttribute('aria-checked', 'true');
  await expect(entries.nth(0)).toHaveAttribute('aria-checked', 'false');
  await entries.nth(0).click();
  await expect(page.locator('.fc-host__formulabar-tag')).toHaveValue('B1');

  // Breaking the cycle greys out the row.
  await sp.enter('C1', '5');
  await openErrorMenu();
  await expect(errorMenu.locator(p.circularRow)).toHaveAttribute('aria-disabled', 'true');
  await page.keyboard.press('Escape');

  // Format Stale Values is a checked preference; under manual calculation an edit leaves its
  // dependants stale.
  const calcMenu = page.locator(p.calcMenu);
  const openCalcMenu = async (): Promise<void> => {
    await sp.clickRibbon(p.calcCommand);
    await expect(calcMenu).toBeVisible();
  };
  const formatStale = calcMenu.locator(p.formatStale);
  const stalePreference = () =>
    page.evaluate(() => {
      const w = window as unknown as ProbeWindow;
      return w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst).store.getState().ui
        .formatStaleValues;
    });
  const staleB2 = () =>
    page.evaluate(() => {
      const w = window as unknown as ProbeWindow;
      return w.__fcProbes['api/workbook']
        .workbookHandleOf(w.__fcInst.workbook)
        .isCellStale({ sheet: 0, row: 1, col: 1 });
    });
  await openCalcMenu();
  await expect(formatStale).toHaveAttribute('role', 'menuitemcheckbox');
  await expect(formatStale).toHaveAttribute('aria-checked', 'true');
  await calcMenu.locator(p.manual).click();
  await sp.enter('A2', '3');
  await expect.poll(staleB2).toBe(true);
  await openCalcMenu();
  await formatStale.click();
  await expect.poll(stalePreference).toBe(false);
  await openCalcMenu();
  await expect(formatStale).toHaveAttribute('aria-checked', 'false');
  await formatStale.click();
  await expect.poll(stalePreference).toBe(true);
  await openCalcMenu();
  await calcMenu.locator(p.auto).click();
  await expect.poll(staleB2).toBe(false);
  await expect.poll(() => sp.readValue('B2')).toEqual({ kind: 'number', value: 6 });

  // Compatibility Version: a new book is version 2; version 1 counts UTF-16 units.
  const version = () =>
    page.evaluate(() => {
      const w = window as unknown as ProbeWindow;
      return w.__fcProbes['api/workbook']
        .workbookHandleOf(w.__fcInst.workbook)
        .compatibilityVersion();
    });
  const compat = page.locator(p.compatPanel);
  const openCompat = async (): Promise<void> => {
    await openCalcMenu();
    await calcMenu.locator(p.compatRow).click();
    await expect(compat).toBeVisible();
  };
  const item = (id: string): Locator => compat.locator(p.compat(id));
  expect(await version()).toBe(2);
  await expect.poll(() => sp.readValue('A3')).toEqual({ kind: 'number', value: 1 });
  await openCompat();
  await expect(item('1')).toHaveAttribute('role', 'menuitemradio');
  await expect(item('1')).toHaveAttribute('aria-checked', 'false');
  await expect(item('2')).toHaveAttribute('aria-checked', 'true');
  // Version 3 is listed only on a book stamped 3.
  await expect(item('3')).toBeHidden();
  await expect(item('learn-more')).toHaveAttribute('aria-disabled', 'true');
  await item('1').click();
  await expect.poll(version).toBe(1);
  await expect.poll(() => sp.readValue('A3')).toEqual({ kind: 'number', value: 2 });
  await page.locator('.fc-host').focus();
  await sp.shortcut('z');
  await expect.poll(version).toBe(2);
  await expect.poll(() => sp.readValue('A3')).toEqual({ kind: 'number', value: 1 });
  await sp.shortcut('Shift+z');
  await expect.poll(version).toBe(1);

  await page.evaluate(async () => {
    const w = window as unknown as ProbeWindow;
    const { Workbook } = w.__fcProbes['api/workbook'];
    const next = await Workbook.load(w.__fcInst.workbook.save().bytes);
    if (next.engine !== 'wasm') throw new Error('Reload must use the real WASM engine');
    w.__fcInst.update({ workbook: next });
  });
  await expect.poll(version).toBe(1);
  await expect.poll(() => sp.readValue('A3')).toEqual({ kind: 'number', value: 2 });
  await openCompat();
  await expect(item('1')).toHaveAttribute('aria-checked', 'true');
  await page.keyboard.press('Escape');

  // A book stamped 3 lists version 3 and checks it.
  await page.evaluate(() => {
    const w = window as unknown as ProbeWindow;
    w.__fcProbes['api/workbook'].workbookHandleOf(w.__fcInst.workbook).setCompatibilityVersion(3);
  });
  await openCompat();
  await expect(item('3')).toHaveAttribute('aria-checked', 'true');
  await expect(item('2')).toHaveAttribute('aria-checked', 'false');
  await page.keyboard.press('Escape');
  await sp.expectNoConsoleErrors();
}
