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
  method:
    | 'openCellStylesGallery'
    | 'openExternalLinksDialog'
    | 'openFindReplace'
    | 'openFormatDialog'
    | 'openGoToSpecial'
    | 'openIterativeDialog'
    | 'openNamedRangeDialog'
    | 'openPageSetup',
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
    open: (page) => openInstanceDialog(page, 'openFormatDialog'),
  },
  {
    name: 'page-setup',
    snapshot: 'dialog-page-setup.png',
    locator: '.fc-pgsetup',
    open: (page) => openInstanceDialog(page, 'openPageSetup'),
  },
  {
    name: 'go-to-special',
    snapshot: 'dialog-go-to-special.png',
    locator: '.fc-goto',
    open: (page) => openInstanceDialog(page, 'openGoToSpecial'),
  },
  {
    name: 'find-replace',
    snapshot: 'dialog-find-replace.png',
    locator: '.fc-find',
    open: (page) => openInstanceDialog(page, 'openFindReplace'),
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
    open: (page) => openInstanceDialog(page, 'openExternalLinksDialog'),
  },
  {
    name: 'name-manager',
    snapshot: 'dialog-name-manager.png',
    locator: '.fc-namedlg',
    open: async (page) => {
      await page.locator('.fc-host__formulabar-tag').fill('TaxRate');
      await page.locator('.fc-host__formulabar-tag').press('Enter');
      await openInstanceDialog(page, 'openNamedRangeDialog');
    },
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

test('@visual dialog name-manager filter menu', async ({ page }) => {
  await mountVisualPage(page, '/?theme=light&locale=en');
  await page.locator('.fc-host__formulabar-tag').fill('TaxRate');
  await page.locator('.fc-host__formulabar-tag').press('Enter');
  await openInstanceDialog(page, 'openNamedRangeDialog');

  const dialog = page.locator('.fc-namedlg').first();
  await expect(dialog).toBeVisible();
  await dialog.getByRole('button', { name: 'Filter', exact: true }).click();
  const menu = page.locator('.fc-namedlg__filter-menu');
  await expect(menu).toBeVisible();
  await page.waitForTimeout(150);

  await expect(menu).toHaveScreenshot('dialog-name-manager-filter-menu.png', {
    maxDiffPixels: 100,
    animations: 'disabled',
  });
});

test('@visual dialog find-replace options', async ({ page }) => {
  await mountVisualPage(page, '/?theme=light&locale=en');
  await openInstanceDialog(page, 'openFindReplace');

  const dialog = page.locator('.fc-find').first();
  await expect(dialog).toBeVisible();
  await dialog.getByRole('button', { name: 'Options >>', exact: true }).click();
  await page.waitForTimeout(150);

  await expect(dialog).toHaveScreenshot('dialog-find-replace-options.png', {
    maxDiffPixels: 100,
    animations: 'disabled',
  });
});

test('@visual dialog name-manager new-name dialog', async ({ page }) => {
  await mountVisualPage(page, '/?theme=light&locale=en');
  await openInstanceDialog(page, 'openNamedRangeDialog');

  const dialog = page.locator('.fc-namedlg').first();
  await expect(dialog).toBeVisible();
  await dialog.getByRole('button', { name: 'New...', exact: true }).click();
  const editor = page.locator('.fc-namedlg-editor');
  await expect(editor).toBeVisible();
  await page.waitForTimeout(150);

  await expect(editor).toHaveScreenshot('dialog-name-manager-new-name.png', {
    maxDiffPixels: 100,
    animations: 'disabled',
  });
});

test('@visual dialog name-manager delete confirmation', async ({ page }) => {
  await mountVisualPage(page, '/?theme=light&locale=en');
  await page.locator('.fc-host__formulabar-tag').fill('TaxRate');
  await page.locator('.fc-host__formulabar-tag').press('Enter');
  await openInstanceDialog(page, 'openNamedRangeDialog');

  const dialog = page.locator('.fc-namedlg').first();
  await expect(dialog).toBeVisible();
  await dialog.getByRole('button', { name: 'Delete', exact: true }).click();
  const confirm = page.locator('.fc-namedlg-confirm');
  await expect(confirm).toBeVisible();
  await page.waitForTimeout(150);

  await expect(confirm).toHaveScreenshot('dialog-name-manager-delete-confirmation.png', {
    maxDiffPixels: 100,
    animations: 'disabled',
  });
});

const formatTabCases = [
  { tab: 'Alignment', snapshot: 'dialog-format-cells-alignment.png' },
  { tab: 'Font', snapshot: 'dialog-format-cells-font.png' },
  { tab: 'Border', snapshot: 'dialog-format-cells-border.png' },
  { tab: 'Fill', snapshot: 'dialog-format-cells-fill.png' },
  { tab: 'Protection', snapshot: 'dialog-format-cells-protection.png' },
] as const;

const pageSetupTabCases = [
  { tab: 'Margins', snapshot: 'dialog-page-setup-margins.png' },
  { tab: 'Header/Footer', snapshot: 'dialog-page-setup-header-footer.png' },
  { tab: 'Sheet', snapshot: 'dialog-page-setup-sheet.png' },
] as const;

for (const c of formatTabCases) {
  test(`@visual format cells tab — ${c.tab}`, async ({ page }) => {
    await mountVisualPage(page, '/?theme=light&locale=en');
    await openInstanceDialog(page, 'openFormatDialog');

    const dialog = page.locator('[class="fc-fmtdlg"]').first();
    await expect(dialog).toBeVisible();
    await dialog.getByRole('tab', { name: c.tab, exact: true }).click();
    await page.waitForTimeout(150);

    await expect(dialog).toHaveScreenshot(c.snapshot, {
      maxDiffPixels: 100,
      animations: 'disabled',
    });
  });
}

test('@visual format cells number category — Number', async ({ page }) => {
  await mountVisualPage(page, '/?theme=light&locale=en');
  await openInstanceDialog(page, 'openFormatDialog');

  const dialog = page.locator('[class="fc-fmtdlg"]').first();
  await expect(dialog).toBeVisible();
  await dialog.getByRole('option', { name: 'Number', exact: true }).click();
  await expect(dialog.getByRole('option', { name: 'Number', exact: true })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await page.waitForTimeout(150);

  await expect(dialog).toHaveScreenshot('dialog-format-cells-number-category.png', {
    maxDiffPixels: 100,
    animations: 'disabled',
  });
});

for (const c of pageSetupTabCases) {
  test(`@visual page setup tab — ${c.tab}`, async ({ page }) => {
    await mountVisualPage(page, '/?theme=light&locale=en');
    await openInstanceDialog(page, 'openPageSetup');

    const dialog = page.locator('.fc-pgsetup').first();
    await expect(dialog).toBeVisible();
    await dialog.getByRole('tab', { name: c.tab, exact: true }).click();
    await expect(dialog.getByRole('tab', { name: c.tab, exact: true })).toHaveAttribute(
      'aria-selected',
      'true',
    );
    await page.waitForTimeout(150);

    await expect(dialog).toHaveScreenshot(c.snapshot, {
      maxDiffPixels: 100,
      animations: 'disabled',
    });
  });
}
