import { expect, type Page } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

export async function runExcelChromeBackstageSearchScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  const fileTab = page.locator('[data-ribbon-tab="file"]').first();
  await expect(fileTab).toBeVisible();
  await fileTab.click();

  const backstage = page.locator('.demo__backstage[role="dialog"]').first();
  await expect(backstage).toBeVisible();
  await expect(backstage.getByRole('button', { name: 'Info', exact: true })).toBeVisible();
  await expect(backstage.getByRole('button', { name: 'New', exact: true })).toBeVisible();
  await expect(backstage.getByRole('button', { name: 'Open', exact: true })).toBeVisible();
  await expect(backstage.getByRole('button', { name: 'Save', exact: true })).toBeVisible();
  await expect(backstage.getByRole('button', { name: 'Save As', exact: true })).toBeVisible();
  await expect(backstage.getByRole('button', { name: 'Print', exact: true })).toBeVisible();
  await expect(backstage.getByRole('button', { name: 'Share', exact: true })).toBeVisible();
  await expect(backstage.getByRole('button', { name: 'Export', exact: true })).toBeVisible();
  await expect(backstage.getByRole('button', { name: 'Options', exact: true })).toBeVisible();

  await page.evaluate(() => {
    const win = window as unknown as {
      __fcInst?: {
        print?: (mode?: 'print' | 'pdf') => void;
        workbook?: {
          setText?: (addr: { sheet: number; row: number; col: number }, value: string) => void;
        };
      };
      __fcPrintModes?: string[];
    };
    win.__fcPrintModes = [];
    if (win.__fcInst) {
      win.__fcInst.workbook?.setText?.({ sheet: 0, row: 0, col: 0 }, 'Preview Cell');
      win.__fcInst.print = (mode = 'print') => {
        win.__fcPrintModes?.push(mode);
      };
    }
  });

  await backstage.getByRole('button', { name: 'Print', exact: true }).first().click();
  await expect(backstage.locator('[data-demo-print-preview]')).toBeVisible();
  await expect(backstage.frameLocator('.demo__print-frame').locator('body')).toContainText(
    'Preview Cell',
  );
  await expect(backstage.getByRole('button', { name: 'Export to PDF', exact: true })).toBeVisible();
  await backstage
    .locator('[data-demo-print-preview]')
    .getByRole('button', { name: 'Print' })
    .click();
  await expect
    .poll(() =>
      page.evaluate(
        () => (window as unknown as { __fcPrintModes?: string[] }).__fcPrintModes ?? [],
      ),
    )
    .toContain('print');

  await backstage.getByRole('button', { name: 'Export', exact: true }).first().click();
  await expect
    .poll(() =>
      page.evaluate(
        () => (window as unknown as { __fcPrintModes?: string[] }).__fcPrintModes ?? [],
      ),
    )
    .toEqual(['print', 'pdf']);

  await backstage.getByRole('button', { name: 'Options', exact: true }).click();
  await expect(page.getByRole('complementary', { name: 'Options panel' })).toBeVisible();

  await backstage.getByRole('button', { name: /^Page Setup\b/ }).click();
  const pageSetup = page.getByRole('dialog', { name: 'Page Setup' });
  await expect(pageSetup).toBeVisible();
  await pageSetup.getByRole('button', { name: 'Close', exact: true }).click();
  await expect(pageSetup).toBeHidden();

  await backstage.getByRole('button', { name: 'Info', exact: true }).click();
  await backstage.getByRole('button', { name: /^Edit Links\b/ }).click();
  const externalLinks = page.getByRole('dialog', { name: 'External Links' });
  await expect(externalLinks).toBeVisible();
  await externalLinks.getByRole('button', { name: 'Close', exact: true }).click();
  await expect(externalLinks).toBeHidden();

  await backstage.getByRole('button', { name: 'Close', exact: true }).click();
  await expect(backstage).toBeHidden();
  await expect(page.locator('[data-ribbon-tab="home"][aria-selected="true"]')).toHaveCount(1);

  await page.keyboard.press('F6');
  await expect(page.locator('.demo__quick button[aria-label="Save"]')).toBeFocused();
  await page.keyboard.press('F6');
  await expect(page.locator('[data-ribbon-tab="home"]')).toBeFocused();
  await page.keyboard.press('F6');
  await expect(page.locator('.fc-host__formulabar-tag')).toBeFocused();
  await page.keyboard.press('F6');
  await expect(page.locator('.fc-host')).toBeFocused();
  await page.keyboard.press('F6');
  await expect(page.locator('.fc-host__statusbar')).toBeFocused();
  await page.keyboard.press('Shift+F6');
  await expect(page.locator('.fc-host')).toBeFocused();

  const search = page.getByRole('combobox', { name: 'Search commands' });
  await page.keyboard.press('Alt+Q');
  await expect(search).toBeFocused();
  await search.fill('training');

  const helpResult = page
    .locator('.demo__command-item')
    .filter({ hasText: 'Help and training' })
    .first();
  await expect(helpResult).toBeVisible();
  await search.press('ArrowDown');
  await expect(search).toHaveAttribute('aria-activedescendant', 'demo-search-option-0');
  await expect(page.locator('#demo-search-option-0')).toHaveAttribute('aria-selected', 'true');
  await search.press('Enter');

  await expect(page.locator('[data-ribbon-tab="help"][aria-selected="true"]')).toHaveCount(1);
}
