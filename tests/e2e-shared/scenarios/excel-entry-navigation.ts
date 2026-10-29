import { expect, type Page } from '@playwright/test';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/mount/types.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

const selection = (page: Page) =>
  page.evaluate(
    () =>
      (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.store.getState().selection,
  );

const active = async (page: Page, row: number, col: number): Promise<void> => {
  await expect.poll(() => selection(page)).toMatchObject({ active: { sheet: 0, row, col } });
};

async function mount(page: Page): Promise<UserJourneyPage> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  return sp;
}

export async function runTabReturnNavigation(page: Page): Promise<void> {
  const sp = await mount(page);
  await sp.goTo('A1');
  await page.keyboard.press('Tab');
  await active(page, 0, 1);
  await page.keyboard.press('Tab');
  await active(page, 0, 2);
  await page.keyboard.press('Enter');
  await active(page, 1, 0);
  await page.keyboard.press('Tab');
  await active(page, 1, 1);
  await page.keyboard.press('Enter');
  await active(page, 2, 0);
  await sp.goTo('C5');
  await page.keyboard.press('Tab');
  await page.keyboard.press('Tab');
  await page.keyboard.press('Enter');
  await active(page, 5, 2);
  await page.keyboard.press('Enter');
  await active(page, 6, 2);
  await sp.expectNoConsoleErrors();
}

export async function runTabReturnCommits(page: Page): Promise<void> {
  const sp = await mount(page);
  await sp.goTo('A1');
  for (const [value, key] of [
    ['1', 'Tab'],
    ['2', 'Tab'],
    ['3', 'Enter'],
  ] as const) {
    await page.keyboard.type(value);
    await page.keyboard.press(key);
  }
  await active(page, 1, 0);
  for (const [ref, value] of [
    ['A1', 1],
    ['B1', 2],
    ['C1', 3],
  ] as const)
    await expect.poll(() => sp.readValue(ref)).toEqual({ kind: 'number', value });
  await sp.goTo('A5');
  const bar = page.locator('.fc-host__formulabar-input');
  for (const [formula, key] of [
    ['=1+1', 'Tab'],
    ['=2+2', 'Tab'],
    ['=3+3', 'Enter'],
  ] as const) {
    await bar.fill(formula);
    await bar.press(key);
  }
  await active(page, 5, 0);
  for (const [ref, value] of [
    ['A5', 2],
    ['B5', 4],
    ['C5', 6],
  ] as const)
    await expect.poll(() => sp.readValue(ref)).toEqual({ kind: 'number', value });
  await sp.expectNoConsoleErrors();
}

export async function runTabReturnMergesAndHidden(page: Page): Promise<void> {
  const sp = await mount(page);
  for (const ref of ['A1:A2', 'C4:C5']) {
    await sp.goTo(ref);
    const menu = await sp.openRibbonSplitMenu('merge', 'menu-merge');
    await menu.locator('[data-merge-action="mergeCells"]').click();
  }
  await sp.goTo('A1');
  await page.keyboard.press('Tab');
  await page.keyboard.press('Tab');
  await page.keyboard.press('Enter');
  await active(page, 2, 0);
  await sp.goTo('B5');
  await page.keyboard.press('Tab');
  await active(page, 3, 2);
  await page.keyboard.press('Tab');
  await active(page, 4, 3);
  await page.keyboard.press('Enter');
  await active(page, 5, 1);
  await sp.goTo('H8:H9');
  await page.keyboard.press('Control+9');
  await sp.goTo('B12:C12');
  await page.keyboard.press('Control+0');
  await sp.goTo('A7');
  await page.keyboard.press('Tab');
  await active(page, 6, 3);
  await page.keyboard.press('Tab');
  await active(page, 6, 4);
  await page.keyboard.press('Enter');
  await active(page, 9, 0);
  await sp.expectNoConsoleErrors();
}

export async function runTabReturnResets(page: Page): Promise<void> {
  const sp = await mount(page);
  for (const [start, keys, row, col] of [
    ['A1', ['Tab', 'Tab', 'Shift+Tab', 'Enter'], 1, 1],
    ['D5', ['Shift+Tab', 'Shift+Tab', 'Enter'], 5, 1],
    ['A3', ['Tab', 'Tab', 'Shift+Enter'], 1, 2],
    ['A1', ['Tab', 'Tab', 'ArrowLeft', 'Enter'], 1, 1],
    ['A1', ['Tab', 'Tab', 'Escape', 'Enter'], 1, 0],
    ['D5', ['Shift+Tab', 'Tab', 'Tab', 'Enter'], 5, 2],
  ] as const) {
    await sp.goTo(start);
    for (const key of keys) await page.keyboard.press(key);
    await active(page, row, col);
  }
  await sp.goTo('A1');
  await page.keyboard.press('Tab');
  await page.keyboard.press('Tab');
  await sp.goTo('E4');
  await page.keyboard.press('Enter');
  await active(page, 4, 4);
  await sp.goTo('A1:C2');
  await page.keyboard.press('Tab');
  await page.keyboard.press('Tab');
  await page.keyboard.press('Enter');
  await expect
    .poll(() => selection(page))
    .toMatchObject({ active: { row: 1, col: 2 }, range: { r0: 0, c0: 0, r1: 1, c1: 2 } });
  await sp.expectNoConsoleErrors();
}
