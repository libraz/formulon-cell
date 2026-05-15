import type { Page } from '@playwright/test';
import { expect, test } from '@playwright/test';

import { mountVisualPage } from './helpers.js';

type DialogCase = {
  readonly name: string;
  readonly snapshot: string;
  readonly locator: string;
  readonly open: (page: Page) => Promise<void>;
};

const clickRibbonCommand = async (page: Page, tab: string, command: string): Promise<void> => {
  await page.getByRole('tab', { name: tab, exact: true }).click();
  await page.locator(`[data-ribbon-command="${command}"]`).click();
};

const openInstanceDialog = async (
  page: Page,
  method: 'openCellStylesGallery' | 'openExternalLinksDialog' | 'openIterativeDialog',
): Promise<void> => {
  await page.evaluate((name) => {
    const inst = (
      window as unknown as {
        __fcInst?: Record<string, unknown>;
      }
    ).__fcInst;
    const open = inst?.[name];
    if (typeof open !== 'function') throw new Error(`Missing visual dialog opener: ${name}`);
    open.call(inst);
  }, method);
};

const dialogCases: readonly DialogCase[] = [
  {
    name: 'format-cells',
    snapshot: 'dialog-format-cells.png',
    locator: '[class="fc-fmtdlg"]',
    open: (page) => clickRibbonCommand(page, 'File', 'formatCells'),
  },
  {
    name: 'page-setup',
    snapshot: 'dialog-page-setup.png',
    locator: '.fc-pgsetup',
    open: (page) => clickRibbonCommand(page, 'File', 'pageSetup'),
  },
  {
    name: 'go-to-special',
    snapshot: 'dialog-go-to-special.png',
    locator: '.fc-goto',
    open: (page) => clickRibbonCommand(page, 'File', 'gotoSpecial'),
  },
  {
    name: 'conditional-rules',
    snapshot: 'dialog-conditional-rules.png',
    locator: '.fc-cfrulesdlg',
    open: (page) => clickRibbonCommand(page, 'Home', 'rules'),
  },
  {
    name: 'external-links',
    snapshot: 'dialog-external-links.png',
    locator: '.fc-extlinkdlg',
    open: (page) => clickRibbonCommand(page, 'File', 'links'),
  },
  {
    name: 'cell-styles',
    snapshot: 'dialog-cell-styles.png',
    locator: '.fc-stylegallery',
    open: (page) => clickRibbonCommand(page, 'Home', 'cellStyles'),
  },
  {
    name: 'iterative-calculation',
    snapshot: 'dialog-iterative-calculation.png',
    locator: '.fc-iterdlg',
    open: (page) => openInstanceDialog(page, 'openIterativeDialog'),
  },
  {
    name: 'function-arguments',
    snapshot: 'dialog-function-arguments.png',
    locator: '.fc-fxdialog',
    open: (page) => clickRibbonCommand(page, 'Insert', 'fxInsert'),
  },
];

for (const c of dialogCases) {
  test(`@visual dialog baseline — ${c.name}`, async ({ page }) => {
    await mountVisualPage(page, '/?theme=light&locale=en');
    await c.open(page);

    const dialog = page.locator(c.locator).first();
    await expect(dialog).toBeVisible();
    await page.waitForTimeout(150);

    await expect(dialog).toHaveScreenshot(c.snapshot, {
      maxDiffPixels: 100,
      animations: 'disabled',
    });
  });
}
