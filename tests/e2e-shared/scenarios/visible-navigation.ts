import { expect, type Page } from '@playwright/test';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/mount/types.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

const selection = (page: Page) =>
  page.evaluate(
    () =>
      (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.store.getState().selection,
  );

async function hideBands(sp: UserJourneyPage): Promise<void> {
  await sp.goTo('H2:H3');
  await sp.page.keyboard.press('Control+9');
  await sp.goTo('B8:C8');
  await sp.page.keyboard.press('Control+0');
  await expect
    .poll(() =>
      sp.page.evaluate(() => {
        const { layout } = (
          window as unknown as { __fcInst: SpreadsheetInstance }
        ).__fcInst.store.getState();
        return { rows: [...layout.hiddenRows], cols: [...layout.hiddenCols] };
      }),
    )
    .toEqual({ rows: [1, 2], cols: [1, 2] });
}

export async function runHiddenBandNavigation(
  page: Page,
  platform: 'mac' | 'default',
): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform });
  await sp.expectNoStub();
  await hideBands(sp);
  for (const [start, key, row, col] of [
    ['A1', 'ArrowDown', 3, 0],
    ['A1', 'ArrowRight', 0, 3],
    ['A1', 'Enter', 3, 0],
    ['A1', 'Tab', 0, 3],
    ['A4', 'ArrowUp', 0, 0],
    ['D1', 'ArrowLeft', 0, 0],
    ['D1', 'Shift+Tab', 0, 0],
  ] as const) {
    await sp.goTo(start);
    await page.keyboard.press(key);
    await expect.poll(() => selection(page)).toMatchObject({ active: { row, col } });
    const box = page.locator('.fc-host__formulabar-tag');
    await expect(box).toHaveValue(`${String.fromCharCode(65 + col)}${row + 1}`);
  }
  // Explicit Go To may still target a hidden cell, unlike directional movement.
  await sp.goTo('B2');
  await expect.poll(() => selection(page)).toMatchObject({ active: { row: 1, col: 1 } });
  await sp.expectNoConsoleErrors();
}

export async function runHiddenBandCommit(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await hideBands(sp);
  await sp.enter('A1', '17');
  await expect.poll(() => selection(page)).toMatchObject({ active: { row: 3, col: 0 } });
  await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'number', value: 17 });
  await sp.goTo('A4');
  await page.keyboard.type('23');
  await page.keyboard.press('Tab');
  await expect.poll(() => selection(page)).toMatchObject({ active: { row: 3, col: 3 } });
  await expect.poll(() => sp.readValue('A4')).toEqual({ kind: 'number', value: 23 });
  await sp.goTo('A1');
  const formulaBar = page.locator('.fc-host__formulabar-input');
  await formulaBar.fill('=SUM(A4,5)');
  await formulaBar.press('Enter');
  await expect.poll(() => selection(page)).toMatchObject({ active: { row: 3, col: 0 } });
  await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'number', value: 28 });
  await sp.goTo('A1');
  await formulaBar.fill('31');
  await formulaBar.press('Tab');
  await expect.poll(() => selection(page)).toMatchObject({ active: { row: 0, col: 3 } });
  await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'number', value: 31 });
  await sp.expectNoConsoleErrors();
}

export async function runHiddenBandTraversal(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await hideBands(sp);
  const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 3 };
  for (const [key, path] of [
    [
      'Enter',
      [
        [3, 0],
        [0, 3],
        [3, 3],
        [0, 0],
      ],
    ],
    [
      'Tab',
      [
        [0, 3],
        [3, 0],
        [3, 3],
        [0, 0],
      ],
    ],
    [
      'Shift+Enter',
      [
        [3, 3],
        [0, 3],
        [3, 0],
        [0, 0],
      ],
    ],
    [
      'Shift+Tab',
      [
        [3, 3],
        [3, 0],
        [0, 3],
        [0, 0],
      ],
    ],
  ] as const) {
    await sp.goTo('A1:D4');
    for (const [row, col] of path) {
      await page.keyboard.press(key);
      await expect.poll(() => selection(page)).toMatchObject({ active: { row, col }, range });
    }
  }
  await sp.expectNoConsoleErrors();
}

export async function runHiddenMergeNavigation(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  for (const ref of ['A2:A4', 'B1:D1']) {
    await sp.goTo(ref);
    const menu = await sp.openRibbonSplitMenu('merge', 'menu-merge');
    await menu.locator('[data-merge-action="mergeCells"]').click();
  }
  await hideBands(sp);
  // Excel 16.114.1 skips a merge entered at a hidden leading edge, but
  // entering its visible trailing edge selects even a hidden merge anchor.
  for (const [start, key, row, col] of [
    ['A1', 'ArrowDown', 4, 0],
    ['A5', 'ArrowUp', 1, 0],
    ['A1', 'ArrowRight', 0, 4],
    ['E1', 'ArrowLeft', 0, 1],
  ] as const) {
    await sp.goTo(start);
    await page.keyboard.press(key);
    await expect.poll(() => selection(page)).toMatchObject({ active: { row, col } });
  }
  await sp.goTo('A1');
  await page.keyboard.type('7');
  await page.keyboard.press('Enter');
  await expect.poll(() => selection(page)).toMatchObject({ active: { row: 4, col: 0 } });
  await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'number', value: 7 });
  await sp.expectNoConsoleErrors();
}

export async function runMergedSelectionCycles(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.goTo('B1:C2');
  const menu = await sp.openRibbonSplitMenu('merge', 'menu-merge');
  await menu.locator('[data-merge-action="mergeCells"]').click();
  const range = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 };
  for (const [key, path] of [
    [
      'Tab',
      [
        [0, 1],
        [1, 0],
        [2, 0],
        [2, 1],
        [2, 2],
        [0, 0],
      ],
    ],
    [
      'Shift+Tab',
      [
        [2, 2],
        [2, 1],
        [2, 0],
        [1, 0],
        [0, 1],
        [0, 0],
      ],
    ],
    [
      'Enter',
      [
        [1, 0],
        [2, 0],
        [0, 1],
        [2, 1],
        [2, 2],
        [0, 0],
      ],
    ],
    [
      'Shift+Enter',
      [
        [2, 2],
        [2, 1],
        [0, 1],
        [2, 0],
        [1, 0],
        [0, 0],
      ],
    ],
  ] as const) {
    await sp.goTo('A1:C3');
    for (const [row, col] of path) {
      await page.keyboard.press(key);
      await expect.poll(() => selection(page)).toMatchObject({ active: { row, col }, range });
    }
  }
  await sp.expectNoConsoleErrors();
}

export async function runHiddenGapMergeNavigation(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  for (const ref of ['A3:A5', 'C1:E1']) {
    await sp.goTo(ref);
    const menu = await sp.openRibbonSplitMenu('merge', 'menu-merge');
    await menu.locator('[data-merge-action="mergeCells"]').click();
  }
  await sp.goTo('H2:H4');
  await page.keyboard.press('Control+9');
  await sp.goTo('B8:D8');
  await page.keyboard.press('Control+0');
  for (const [key, row, col] of [
    ['ArrowDown', 5, 0],
    ['ArrowRight', 0, 5],
  ] as const) {
    await sp.goTo('A1');
    await page.keyboard.press(key);
    await expect.poll(() => selection(page)).toMatchObject({ active: { row, col } });
  }
  await sp.goTo('H6');
  await page.keyboard.press('Control+9');
  await sp.goTo('F8');
  await page.keyboard.press('Control+0');
  for (const [start, key, row, col] of [
    ['A7', 'ArrowUp', 2, 0],
    ['G1', 'ArrowLeft', 0, 2],
  ] as const) {
    await sp.goTo(start);
    await page.keyboard.press(key);
    await expect.poll(() => selection(page)).toMatchObject({ active: { row, col } });
  }
  await sp.expectNoConsoleErrors();
}
