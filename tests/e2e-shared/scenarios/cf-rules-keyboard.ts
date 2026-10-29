import { expect, type Page } from '@playwright/test';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/mount/types.js';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

export async function runCfRulesKeyboardScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount();
  await sp.expectNoStub();

  await page.evaluate(() => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    instance.openConditionalDialog({ mode: 'new', kind: 'cell-value' });
  });
  const authoring = page.locator('.fc-conddlg');
  await expect(authoring).toBeVisible();
  await authoring.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(authoring).toBeHidden();
  await page.evaluate(() => {
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.openCfRulesDialog();
  });

  const manager = page.locator('.fc-cfrulesdlg');
  const rows = manager.locator('.fc-cfrulesdlg__table tbody tr');
  await expect(manager).toBeVisible();
  await expect(rows).toHaveCount(1);
  await expect.poll(() => readRuleCounts(page)).toEqual({ session: 1, engine: 1 });

  const stopIfTrue = rows.first().locator('.fc-cfrulesdlg__cell-stop input[type="checkbox"]');
  await stopIfTrue.focus();
  await page.keyboard.press('Space');
  await expect(stopIfTrue).toBeChecked();
  await expect(rows).toHaveCount(1);
  expect(
    await page.evaluate(
      () =>
        (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.store.getState()
          .conditional.rules[0]?.stopIfTrue,
    ),
  ).toBe(true);

  const duplicate = rows.first().locator('.fc-cfrulesdlg__row-duplicate');
  await expect(duplicate).toBeEnabled();
  await duplicate.focus();
  await page.keyboard.press('Enter');
  await expect(rows).toHaveCount(2);
  await expect.poll(() => readRuleCounts(page)).toEqual({ session: 2, engine: 2 });
  await expect
    .poll(() => readStopFlags(page))
    .toEqual({
      session: [true, true],
      engine: [true, true],
    });

  const removeDuplicate = rows.nth(1).locator('.fc-cfrulesdlg__remove');
  await removeDuplicate.focus();
  await page.keyboard.press('Enter');
  await expect(rows).toHaveCount(1);
  await expect.poll(() => readRuleCounts(page)).toEqual({ session: 1, engine: 1 });

  await page.keyboard.press('Escape');
  await expect(manager).toBeHidden();
  await page.evaluate(() => {
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.openCfRulesDialog();
  });
  await expect(rows).toHaveCount(1);
  await expect(stopIfTrue).toBeChecked();
  await page.keyboard.press('Escape');
  await expect(manager).toBeHidden();
  expect(consoleErrors.read()).toEqual([]);
}

async function readRuleCounts(page: Page): Promise<{ session: number; engine: number }> {
  return page.evaluate(() => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    return {
      session: instance.store.getState().conditional.rules.length,
      engine: instance.workbook.getConditionalFormats(0).length,
    };
  });
}

async function readStopFlags(page: Page): Promise<{ session: boolean[]; engine: boolean[] }> {
  return page.evaluate(() => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    return {
      session: instance.store.getState().conditional.rules.map((rule) => rule.stopIfTrue === true),
      engine: instance.workbook.getConditionalFormats(0).map((rule) => rule.stopIfTrue),
    };
  });
}
