import { expect, type Page } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

export async function runExcelFormatCategoryVisibilityScenario(
  page: Page,
  locale: 'en' | 'ja',
): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount({ locale });
  await sp.expectNoStub();
  await sp.focusHost();
  await sp.shortcut('1');
  const dialog = page.getByRole('dialog', {
    name: locale === 'ja' ? 'セルの書式設定' : 'Format Cells',
    exact: true,
  });
  await expect(dialog).toBeVisible();
  const panel = dialog.locator('[role="tabpanel"][data-fc-tab="number"]');
  const thousands = panel.locator('input[data-fc-check="thousands"]');
  const negatives = panel.locator('.fc-fmtdlg__negative');
  const rows = panel.locator('.fc-fmtdlg__cat-controls > .fc-fmtdlg__row');
  const decimals = rows.nth(0);
  const symbol = rows.nth(1);
  const bounds = await dialog.boundingBox();

  // The reference screenshots show separators only for Number, and negative
  // choices only for Number/Currency. Check layout, not just the hidden flag:
  // author display:flex rules can make hidden controls visible in a browser.
  for (const category of [
    'general',
    'fixed',
    'currency',
    'accounting',
    'date',
    'time',
    'percent',
    'fraction',
    'scientific',
    'text',
    'special',
    'custom',
    'general',
  ] as const) {
    await panel.locator(`[data-fc-cat="${category}"]`).click();
    const hasDecimals = ['fixed', 'currency', 'accounting', 'percent', 'scientific'].includes(
      category,
    );
    await expect(decimals).toBeVisible({ visible: hasDecimals });
    await expect(symbol).toBeVisible({ visible: ['currency', 'accounting'].includes(category) });
    await expect(thousands).toBeVisible({ visible: category === 'fixed' });
    await expect(negatives).toBeVisible({ visible: ['fixed', 'currency'].includes(category) });
    await expect(panel.locator('.fc-fmtdlg__pattern-list-wrap')).toBeVisible({
      visible: ['date', 'time', 'fraction', 'special'].includes(category),
    });
    await expect(rows.nth(2)).toBeVisible({ visible: category === 'custom' });
    await expect(rows.nth(3)).toBeVisible({ visible: category === 'custom' });
    await expect(rows.nth(4)).toBeVisible({
      visible: ['date', 'time', 'special'].includes(category),
    });
    await expect(rows.nth(5)).toBeVisible({ visible: category === 'date' });
    await expect
      .poll(() =>
        panel
          .locator('[hidden]')
          .evaluateAll((nodes) =>
            nodes.filter((node) => node.getClientRects().length > 0).map((node) => node.className),
          ),
      )
      .toEqual([]);
    expect(await dialog.boundingBox()).toEqual(bounds);
  }
}

export async function runExcelGeneralNotationScenario(
  page: Page,
  locale: 'en' | 'ja',
): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount({ locale });
  await sp.expectNoStub();
  for (const [input, expected] of [
    ['123456789012', '1.23457E+11'],
    ['-0.0000000001234', '-1.23400E-10'],
    ['1234.5', '1234.5'],
  ]) {
    await sp.typeIntoActiveCell(input);
    await page.keyboard.press('ArrowUp');
    await sp.shortcut('1');
    const dialog = page.getByRole('dialog', {
      name: locale === 'ja' ? 'セルの書式設定' : 'Format Cells',
      exact: true,
    });
    await expect(dialog).toBeVisible();
    await expect(dialog.locator('.fc-fmtdlg__preview-cell')).toHaveText(expected);
    await dialog
      .locator('.fc-fmtdlg__footer')
      .getByRole('button', {
        name: locale === 'ja' ? 'キャンセル' : 'Cancel',
        exact: true,
      })
      .click();
    await expect(dialog).toBeHidden();
  }
}

export async function runExcelFormatPreviewScenario(
  page: Page,
  locale: 'en' | 'ja',
): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount({ locale });
  await sp.expectNoStub();
  await sp.typeIntoActiveCell('-1234.5');
  await page.keyboard.press('ArrowUp');
  await sp.shortcut('1');
  const dialog = page.getByRole('dialog', {
    name: locale === 'ja' ? 'セルの書式設定' : 'Format Cells',
    exact: true,
  });
  await expect(dialog).toBeVisible();
  const sample = dialog.locator('.fc-fmtdlg__preview-cell');
  await expect(sample).toHaveText('-1234.5');
  await dialog.locator('[data-fc-cat="fixed"]').click();
  await expect(sample).toHaveText('-1234.50');
  const negativeMinus = dialog.locator('[data-fc-negative-style="minus"]');
  await expect(negativeMinus).toHaveText('-1234.00');
  await dialog.locator('[data-fc-check="thousands"]').check();
  await expect(sample).toHaveText('-1,234.50');
  await expect(negativeMinus).toHaveText('-1,234.00');
  await dialog.locator('[data-fc-negative-style="red-parens"]').click();
  await expect(sample).toHaveText('(1,234.50)');
  await expect(sample).toHaveCSS('color', 'rgb(192, 0, 0)');
  await dialog.locator('[data-fc-negative-style="minus"]').click();
  await expect(sample).toHaveText('-1,234.50');
  await expect(sample).not.toHaveCSS('color', 'rgb(192, 0, 0)');
  await dialog
    .locator('.fc-fmtdlg__footer')
    .getByRole('button', { name: locale === 'ja' ? 'キャンセル' : 'Cancel', exact: true })
    .click();
  await expect(dialog).toBeHidden();
  await expect(page.locator('.fc-host__formulabar-input')).toHaveValue('-1234.5');
}
