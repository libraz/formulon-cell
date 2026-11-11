import { expect, type Locator, type Page } from '@playwright/test';
import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { gridCellCenter, SpreadsheetPage } from '../pages/SpreadsheetPage.js';

const host = (page: Page): Locator => page.locator('.fc-host').first();

async function openCellMenu(page: Page): Promise<Locator> {
  await page.keyboard.press('Shift+F10');
  const menu = page.locator('.fc-ctxmenu:not(.fc-ctxmenu__sub)');
  await expect(menu).toBeVisible();
  return menu;
}

export async function runContextDialogReturnScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount({ locale: 'en' });
  await sp.expectNoStub();
  await sp.focusHost();
  for (const selector of [
    '.fc-ctxmenu__item[data-fc-action="formatCells"]',
    '.fc-ctxmenu__mini [data-fc-action="formatCells"]',
  ]) {
    const menu = await openCellMenu(page);
    const action = menu.locator(selector);
    await action.focus();
    await page.keyboard.press('Enter');
    const dialog = page.locator('.fc-fmtdlg[aria-label="Format Cells"]');
    await expect(dialog).toBeVisible();
    await expect(menu).toBeHidden();
    await page.keyboard.press('Escape');
    await expect(dialog).toBeHidden();
    await expect(host(page)).toBeFocused();
  }

  const cell = await gridCellCenter(page, 0, 0);
  const grid = await page.locator('.fc-host__grid').first().boundingBox();
  if (!grid) throw new Error('Missing grid');
  await page.mouse.click(grid.x + 5, cell.y, { button: 'right' });
  const menu = page.locator('.fc-ctxmenu:not(.fc-ctxmenu__sub)');
  await menu.locator('[data-fc-action="rowHeight"]').click();
  const prompt = page.locator('.fc-tb__dlg');
  await expect(prompt).toBeVisible();
  const before = await rowHeight(page);
  await prompt.locator('input').fill(String(before + 12));
  await prompt.getByRole('button', { name: 'Cancel', exact: true }).focus();
  await page.keyboard.press('Enter');
  await expect(prompt).toBeHidden();
  await expect(host(page)).toBeFocused();
  expect(await rowHeight(page)).toBe(before);
}

async function rowHeight(page: Page): Promise<number> {
  return page.evaluate(() => {
    const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
    const layout = w.__fcProbes['mount/internals']
      .spreadsheetInternals(w.__fcInst)
      .store.getState().layout;
    return layout.rowHeights.get(0) ?? layout.defaultRowHeight;
  });
}

async function openPasteSpecial(page: Page): Promise<Locator> {
  const menu = await openCellMenu(page);
  const trigger = menu.locator('[data-fc-submenu="pasteSpecialMenu"]');
  await trigger.focus();
  await expect(trigger).toBeFocused();
  await page.keyboard.press('ArrowRight');
  const submenu = page.locator('.fc-ctxmenu__sub');
  await expect(submenu).toBeVisible();
  const item = submenu.locator('[data-fc-action="pasteSpecial"]');
  await item.focus();
  await expect(item).toBeFocused();
  await page.keyboard.press('Enter');
  const dialog = page.locator('.fc-pastesp');
  await expect(dialog).toBeVisible();
  return dialog;
}

async function readValue(page: Page, col: number) {
  return page.evaluate((column) => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    return instance.workbook.sheet(0).getValue(0, column);
  }, col);
}

export async function runPasteSpecialCancelScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount({ locale: 'en' });
  await sp.expectNoStub();
  await sp.typeIntoActiveCell('17');
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('ArrowRight');
  await page.keyboard.type('99');
  await page.keyboard.press('Enter');
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('ArrowLeft');
  await sp.shortcut('c');
  await page.keyboard.press('ArrowRight');

  const dialog = await openPasteSpecial(page);
  await dialog.getByRole('button', { name: 'Cancel', exact: true }).focus();
  await page.keyboard.press('Enter');
  await expect(dialog).toBeHidden();
  await expect.poll(() => readValue(page, 1)).toEqual(99);
  await expect(host(page)).toBeFocused();

  await openPasteSpecial(page);
  await dialog.getByRole('button', { name: 'OK', exact: true }).focus();
  await page.keyboard.press('Enter');
  await expect(dialog).toBeHidden();
  await expect.poll(() => readValue(page, 1)).toEqual(17);
  await expect(host(page)).toBeFocused();
  await sp.shortcut('z');
  await expect.poll(() => readValue(page, 1)).toEqual(99);
}

async function openHyperlink(page: Page): Promise<Locator> {
  await page.evaluate(() => {
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.dialogs.open('hyperlink');
  });
  const dialog = page.locator('.fc-hldlg');
  await expect(dialog).toBeVisible();
  return dialog;
}

async function readLinks(page: Page): Promise<string[]> {
  return page.evaluate(() => {
    const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
    const { store } = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst);
    return [...store.getState().format.formats.values()].flatMap((format) =>
      format.hyperlink ? [format.hyperlink] : [],
    );
  });
}

export async function runHyperlinkCancelScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount({ locale: 'en' });
  await sp.expectNoStub();
  await sp.typeIntoActiveCell('original');
  await page.keyboard.press('ArrowUp');
  const before = await readValue(page, 0);
  const url = 'https://example.com/target';
  const dialog = await openHyperlink(page);
  const input = dialog.locator('input[type="url"]');
  const ok = dialog.getByRole('button', { name: 'OK', exact: true });
  await input.fill(url);
  await dialog.getByRole('button', { name: 'Cancel', exact: true }).focus();
  await page.keyboard.press('Enter');
  await expect(dialog).toBeHidden();
  expect(await readLinks(page)).toEqual([]);
  expect(await readValue(page, 0)).toEqual(before);
  await expect(host(page)).toBeFocused();

  await openHyperlink(page);
  await input.fill(url);
  await page.keyboard.press('Shift+Tab');
  await expect(ok).toBeFocused();
  await page.keyboard.press('Tab');
  await expect(input).toBeFocused();
  await ok.focus();
  await page.keyboard.press('Enter');
  await expect(dialog).toBeHidden();
  expect(await readLinks(page)).toEqual([url]);
  await expect(host(page)).toBeFocused();

  await openHyperlink(page);
  await dialog.getByRole('button', { name: 'Remove link', exact: true }).focus();
  await page.keyboard.press('Enter');
  await expect(dialog).toBeHidden();
  expect(await readLinks(page)).toEqual([]);
  await expect(host(page)).toBeFocused();
  await sp.shortcut('z');
  expect(await readLinks(page)).toEqual([url]);
}
