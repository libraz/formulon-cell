import { expect, type Page } from '@playwright/test';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { gridCellCenter } from '../pages/SpreadsheetPage.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

const selection = (page: Page) =>
  page.evaluate(() =>
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.selection.get(),
  );

export async function runMacSelectionTraversal(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  const range = { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 };
  for (const [key, path] of [
    [
      'Enter',
      [
        [1, 0],
        [0, 1],
        [1, 1],
        [0, 0],
      ],
    ],
    [
      'Shift+Enter',
      [
        [1, 1],
        [0, 1],
        [1, 0],
        [0, 0],
      ],
    ],
    [
      'Tab',
      [
        [0, 1],
        [1, 0],
        [1, 1],
        [0, 0],
      ],
    ],
    [
      'Shift+Tab',
      [
        [1, 1],
        [1, 0],
        [0, 1],
        [0, 0],
      ],
    ],
  ] as const) {
    await sp.goTo('A1:B2');
    await page.locator('.fc-host').focus();
    for (const [row, col] of path) {
      await page.keyboard.press(key);
      await expect.poll(() => selection(page)).toMatchObject({ active: { row, col }, range });
    }
  }
  await sp.expectNoConsoleErrors();
}

export async function runMacSelectionEntry(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.goTo('A1:B2');
  await page.locator('.fc-host').focus();
  await page.keyboard.type('10');
  await page.keyboard.press('Enter');
  await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'number', value: 10 });
  await expect
    .poll(() => selection(page))
    .toMatchObject({
      active: { row: 1, col: 0 },
      range: { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 },
    });
  await page.keyboard.type('20');
  await page.keyboard.press('Tab');
  await expect.poll(() => sp.readValue('A2')).toEqual({ kind: 'number', value: 20 });
  await expect
    .poll(() => selection(page))
    .toMatchObject({
      active: { row: 1, col: 1 },
      range: { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 },
    });
  await sp.expectNoConsoleErrors();
}

export async function runMacCommandReturnFill(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.goTo('A1:B2');
  await page.locator('.fc-host').focus();
  await page.keyboard.type('42');
  await page.keyboard.press('Meta+Enter');
  await expect(page.locator('.fc-host__editor')).toHaveCount(0);
  for (const ref of ['A1', 'A2', 'B1', 'B2'])
    await expect.poll(() => sp.readValue(ref)).toEqual({ kind: 'number', value: 42 });
  await expect
    .poll(() => selection(page))
    .toMatchObject({
      active: { row: 0, col: 0 },
      range: { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 },
    });
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Meta+z');
  for (const ref of ['A1', 'A2', 'B1', 'B2'])
    await expect.poll(() => sp.readValue(ref)).toEqual({ kind: 'blank' });
  await sp.goTo('C1');
  await page.keyboard.type('first');
  await page.keyboard.press('Alt+Enter');
  await page.keyboard.type('second');
  await page.keyboard.press('Enter');
  await expect.poll(() => sp.readValue('C1')).toEqual({ kind: 'text', value: 'first\nsecond' });
  await sp.expectNoConsoleErrors();
}

export async function runMacCollapseSelection(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.enter('A1', 'keep');
  await sp.goTo('A1:B2');
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Shift+Backspace');
  await expect(page.locator('.fc-host__editor')).toHaveCount(0);
  await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'text', value: 'keep' });
  await expect
    .poll(() => selection(page))
    .toMatchObject({
      active: { row: 0, col: 0 },
      range: { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 },
    });
  await sp.expectNoConsoleErrors();
}

export async function runMacFormulaBarSelection(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.goTo('A1:B2');
  const fx = page.locator('.fc-host__formulabar-input');
  await fx.fill('7');
  await fx.press('Enter');
  await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'number', value: 7 });
  await expect
    .poll(() => selection(page))
    .toMatchObject({
      active: { row: 1, col: 0 },
      range: { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 },
    });
  await fx.fill('8');
  await fx.press('Tab');
  await expect.poll(() => sp.readValue('A2')).toEqual({ kind: 'number', value: 8 });
  await expect
    .poll(() => selection(page))
    .toMatchObject({
      active: { row: 1, col: 1 },
      range: { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 },
    });
  await sp.goTo('C1:D2');
  await fx.fill('=A1');
  await fx.press('Meta+Enter');
  for (const [ref, formula] of [
    ['C1', '=A1'],
    ['D1', '=B1'],
    ['C2', '=A2'],
    ['D2', '=B2'],
  ])
    await expect.poll(() => sp.readFormula(ref as string)).toBe(formula);
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Meta+z');
  for (const ref of ['C1', 'C2', 'D1', 'D2'])
    await expect.poll(() => sp.readValue(ref)).toEqual({ kind: 'blank' });
  await sp.expectNoConsoleErrors();
}

export async function runMacFormulaBarPointerReferences(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  for (const [ref, value] of [
    ['A1', '2'],
    ['B1', '4'],
    ['A2', '3'],
    ['B2', '5'],
  ])
    await sp.enter(ref as string, value as string);
  await sp.goTo('C1');
  const fx = page.locator('.fc-host__formulabar-input');
  await fx.fill('=SUM(');
  const a1 = await gridCellCenter(page, 0, 0);
  const b2 = await gridCellCenter(page, 1, 1);
  await page.mouse.move(a1.x, a1.y);
  await page.mouse.down();
  await page.mouse.move(b2.x, b2.y, { steps: 8 });
  await page.mouse.up();
  await expect(fx).toBeFocused();
  await expect(fx).toHaveValue('=SUM(A1:B2');
  await page.keyboard.type(')');
  await page.keyboard.press('Enter');
  await expect.poll(() => sp.readFormula('C1')).toBe('=SUM(A1:B2)');
  await expect.poll(() => sp.readValue('C1')).toEqual({ kind: 'number', value: 14 });
  await sp.expectNoConsoleErrors();
}

export async function runMacOptionSheetSwitch(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  const tabs = page.locator('.fc-host__sheetbar-tab');
  await page.locator('.fc-host__sheetbar-add').click();
  await expect(tabs).toHaveCount(2);
  await tabs.first().click();
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Alt+ArrowRight');
  await expect(tabs.nth(1)).toHaveAttribute('aria-selected', 'true');
  await page.keyboard.press('Alt+ArrowLeft');
  await expect(tabs.first()).toHaveAttribute('aria-selected', 'true');
  await sp.expectNoConsoleErrors();
}
