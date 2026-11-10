import { expect, type Locator, type Page } from '@playwright/test';
import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

type Kind = 'insert' | 'delete';

async function values(page: Page) {
  return page.evaluate(() => {
    const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
    // The tagged cell value has no facade equivalent; read it from the engine handle.
    const handle = w.__fcProbes['api/workbook'].workbookHandleOf(w.__fcInst.workbook);
    return [0, 1, 2, 3].map((row) => handle.getValue({ sheet: 0, row, col: 0 }));
  });
}

async function openShift(page: Page, kind: Kind): Promise<Locator> {
  await page.keyboard.press('Shift+F10');
  const menu = page.locator('.fc-ctxmenu:not(.fc-ctxmenu__sub)');
  await expect(menu).toBeVisible();
  await menu.locator(`[data-fc-action="${kind}Cells"]`).click();
  const dialog = page.locator('.fc-cellshift');
  await expect(dialog).toBeVisible();
  await expect(menu).toBeHidden();
  return dialog;
}

export async function runCellShiftModalScenario(page: Page, kind: Kind): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount({ locale: 'en' });
  await sp.expectNoStub();
  await sp.typeIntoActiveCell('11');
  await page.keyboard.type('22');
  await page.keyboard.press('Enter');
  await page.keyboard.type('33');
  await page.keyboard.press('Enter');
  await sp.focusHost();
  await page.keyboard.press('ArrowRight');
  await page.keyboard.type('44');
  await page.keyboard.press('Enter');
  await sp.focusHost();

  const initial = [
    { kind: 'number', value: 11 },
    { kind: 'number', value: 22 },
    { kind: 'number', value: 33 },
    { kind: 'blank' },
  ];
  await expect.poll(() => values(page)).toEqual(initial);
  const host = page.locator('.fc-host').first();
  const primary = kind === 'insert' ? 'down' : 'up';
  const alternate = kind === 'insert' ? 'right' : 'left';
  const dialog = await openShift(page, kind);
  const checked = dialog.locator(`input[value="${primary}"]`);
  const ok = dialog.getByRole('button', { name: 'OK', exact: true });
  const cancel = dialog.getByRole('button', { name: 'Cancel', exact: true });
  await expect(checked).toBeChecked();
  await expect(ok).toBeFocused();
  await page.keyboard.press('Tab');
  await expect(checked).toBeFocused();
  await page.keyboard.press('Shift+Tab');
  await expect(ok).toBeFocused();
  await page.keyboard.press('Tab');
  await page.keyboard.press('ArrowUp');
  await expect(dialog.locator(`input[value="${alternate}"]`)).toBeChecked();
  await page.keyboard.press('Tab');
  await expect(cancel).toBeFocused();
  await page.keyboard.press('Enter');
  await expect(dialog).toHaveCount(0);
  await expect(host).toBeFocused();
  expect(await values(page)).toEqual(initial);

  await openShift(page, kind);
  await expect(checked).toBeChecked();
  await page.keyboard.press('Escape');
  await expect(dialog).toHaveCount(0);
  await expect(host).toBeFocused();
  expect(await values(page)).toEqual(initial);

  await openShift(page, kind);
  await checked.focus();
  await page.keyboard.press('Enter');
  await expect(dialog).toHaveCount(0);
  await expect(host).toBeFocused();
  const shifted =
    kind === 'insert'
      ? [{ kind: 'blank' }, ...initial.slice(0, 3)]
      : [initial[1], initial[2], { kind: 'blank' }, { kind: 'blank' }];
  await expect.poll(() => values(page)).toEqual(shifted);
  await sp.shortcut('z');
  await expect.poll(() => values(page)).toEqual(initial);
  const originalRow = [
    { kind: 'number', value: 11 },
    { kind: 'number', value: 44 },
    { kind: 'blank' },
  ];
  await expect.poll(() => rowValues(page)).toEqual(originalRow);

  await openShift(page, kind);
  await checked.focus();
  await page.keyboard.press('ArrowUp');
  await expect(dialog.locator(`input[value="${alternate}"]`)).toBeChecked();
  await ok.focus();
  await page.keyboard.press('Enter');
  await expect(dialog).toHaveCount(0);
  await expect(host).toBeFocused();
  const horizontal =
    kind === 'insert'
      ? [{ kind: 'blank' }, { kind: 'number', value: 11 }, { kind: 'number', value: 44 }]
      : [{ kind: 'number', value: 44 }, { kind: 'blank' }, { kind: 'blank' }];
  await expect.poll(() => rowValues(page)).toEqual(horizontal);
  await expect
    .poll(() => values(page))
    .toEqual([horizontal[0], initial[1], initial[2], initial[3]]);
  await sp.shortcut('z');
  await expect.poll(() => values(page)).toEqual(initial);
  await expect.poll(() => rowValues(page)).toEqual(originalRow);
}

async function rowValues(page: Page) {
  return page.evaluate(() => {
    const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
    const handle = w.__fcProbes['api/workbook'].workbookHandleOf(w.__fcInst.workbook);
    return [0, 1, 2].map((col) => handle.getValue({ sheet: 0, row: 0, col }));
  });
}
