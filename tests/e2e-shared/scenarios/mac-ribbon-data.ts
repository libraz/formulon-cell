import { expect, type Page } from '@playwright/test';

import { UserJourneyPage } from '../pages/UserJourneyPage.js';

export async function runMacOutlineDetailScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  const hidden = () =>
    page.evaluate(() => {
      const inst = (
        window as unknown as {
          __fcInst: {
            store: { getState(): { layout: { hiddenRows: Set<number>; hiddenCols: Set<number> } } };
          };
        }
      ).__fcInst;
      const layout = inst.store.getState().layout;
      return { rows: Array.from(layout.hiddenRows), cols: Array.from(layout.hiddenCols) };
    });
  await sp.goTo('A1:A3');
  await page.locator('[data-ribbon-tab="data"]').click();
  await sp.clickRibbon('mac.data.group');
  await sp.clickRibbon('mac.data.hideDetail');
  await expect.poll(hidden).toEqual({ rows: [0, 1, 2], cols: [] });
  await sp.clickRibbon('mac.data.showDetail');
  await expect.poll(hidden).toEqual({ rows: [], cols: [] });
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Meta+z');
  await expect.poll(hidden).toEqual({ rows: [0, 1, 2], cols: [] });
  await page.keyboard.press('Meta+Shift+z');
  await expect.poll(hidden).toEqual({ rows: [], cols: [] });
  await sp.goTo('C1:E1');
  await sp.clickRibbon('mac.data.group');
  await sp.clickRibbon('mac.data.hideDetail');
  await expect.poll(hidden).toEqual({ rows: [], cols: [2, 3, 4] });
  await sp.clickRibbon('mac.data.showDetail');
  await expect.poll(hidden).toEqual({ rows: [], cols: [] });
  await sp.expectNoConsoleErrors();
}

export async function runMacGoalSeekScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.enter('A1', '3');
  await sp.enter('B1', '=A1*2');
  await sp.goTo('B1');
  await page.locator('[data-ribbon-tab="data"]').click();
  await sp.clickRibbon('mac.data.whatIf');
  await page.locator('#menu-mac-data-what-if [data-ribbon-command="mac.data.goalSeek"]').click();
  const dialog = page.locator('.fc-mac-goalseek');
  await expect(dialog).toBeVisible();
  await dialog.locator('#fc-mac-goalseek-formula-cell').fill('B1');
  await dialog.locator('#fc-mac-goalseek-target-value').fill('10');
  await dialog.locator('#fc-mac-goalseek-changing-cell').fill('A1');
  const confirm = dialog.locator('[data-fc-mac-action="goal-seek-ok"]');
  await confirm.click();
  await expect(dialog.locator('.fc-mac-goalseek__status')).toHaveAttribute('data-state', 'success');
  expect(await sp.readValue('A1')).toEqual({ kind: 'number', value: 3 });
  await confirm.click();
  await expect(dialog).toBeHidden();
  await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'number', value: 5 });
  await expect.poll(() => sp.readValue('B1')).toEqual({ kind: 'number', value: 10 });
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Meta+z');
  await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'number', value: 3 });
  await page.keyboard.press('Meta+Shift+z');
  await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'number', value: 5 });
  await sp.clickRibbon('mac.data.whatIf');
  await page.locator('#menu-mac-data-what-if [data-ribbon-command="mac.data.goalSeek"]').click();
  await dialog.locator('#fc-mac-goalseek-formula-cell').fill('B1');
  await dialog.locator('#fc-mac-goalseek-target-value').fill('24');
  await dialog.locator('#fc-mac-goalseek-changing-cell').fill('A1');
  await confirm.click();
  await expect(dialog.locator('.fc-mac-goalseek__status')).toHaveAttribute('data-state', 'success');
  await dialog.locator('[data-fc-mac-action="goal-seek-cancel"]').click();
  await expect(dialog).toBeHidden();
  expect(await sp.readValue('A1')).toEqual({ kind: 'number', value: 5 });
  await sp.expectNoConsoleErrors();
}

export async function runMacConsolidateScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.enter('A1', '1');
  await sp.enter('A2', '2');
  await sp.enter('C1', '3');
  await sp.enter('C2', '4');
  await page.locator('[data-ribbon-tab="data"]').click();
  await sp.clickRibbon('mac.data.consolidate');
  const dialog = page.locator('.fc-mac-consolidate');
  await expect(dialog).toBeVisible();
  await dialog.locator('#fc-mac-consolidate-sources').fill('A1:A2\nC1:C2');
  await dialog.locator('#fc-mac-consolidate-destination').fill('E1');
  await dialog.locator('[data-fc-mac-action="consolidate-ok"]').click();
  await expect(dialog).toBeHidden();
  await expect.poll(() => sp.readValue('E1')).toEqual({ kind: 'number', value: 4 });
  await expect.poll(() => sp.readValue('E2')).toEqual({ kind: 'number', value: 6 });
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Meta+z');
  await expect.poll(() => sp.readValue('E1')).toEqual({ kind: 'blank' });
  await expect.poll(() => sp.readValue('E2')).toEqual({ kind: 'blank' });
  await sp.expectNoConsoleErrors();
}

export async function runMacSubtotalScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  for (const [addr, value] of [
    ['A1', 'Group'],
    ['B1', 'Amount'],
    ['A2', 'a'],
    ['B2', '2'],
    ['A3', 'a'],
    ['B3', '3'],
    ['A4', 'b'],
    ['B4', '4'],
  ])
    await sp.enter(addr as string, value as string);
  await sp.goTo('A1:B4');
  await page.locator('[data-ribbon-tab="data"]').click();
  await sp.clickRibbon('mac.data.subtotal');
  const dialog = page.locator('.fc-mac-subtotal[role=dialog]');
  await expect(dialog).toBeVisible();
  await dialog.locator('#fc-mac-subtotal-range').fill('A1:B4');
  await dialog.locator('#fc-mac-subtotal-group-by').fill('A');
  await dialog.locator('#fc-mac-subtotal-columns').fill('B');
  await dialog.locator('[data-fc-mac-action="subtotal-ok"]').click();
  await expect(dialog).toBeHidden();
  await expect.poll(() => sp.readValue('B4')).toEqual({ kind: 'number', value: 5 });
  await expect.poll(() => sp.readValue('B6')).toEqual({ kind: 'number', value: 4 });
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Meta+z');
  await expect.poll(() => sp.readValue('B4')).toEqual({ kind: 'number', value: 4 });
  await expect.poll(() => sp.readValue('B6')).toEqual({ kind: 'blank' });
  await sp.expectNoConsoleErrors();
}

export async function runMacDirectDataControlsScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.enter('A1', '3');
  await sp.enter('A2', '1');
  await sp.enter('A3', '2');
  await sp.goTo('A1:A3');
  await page.locator('[data-ribbon-tab="data"]').click();
  const panel = page.locator('[data-ribbon-panel="data"]');
  for (const action of [
    'sortAsc',
    'sortDesc',
    'sortCustom',
    'filter',
    'clear',
    'reapply',
    'advancedFilter',
  ]) {
    await expect(
      panel.locator(`.fc-tb__ribbon-tools > [data-ribbon-command="mac.data.${action}"]`),
    ).toHaveCount(1);
  }
  await sp.clickRibbon('mac.data.sortAsc');
  await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'number', value: 1 });
  await expect.poll(() => sp.readValue('A3')).toEqual({ kind: 'number', value: 3 });
  await sp.clickRibbon('mac.data.sortDesc');
  await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'number', value: 3 });
  await expect.poll(() => sp.readValue('A3')).toEqual({ kind: 'number', value: 1 });
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Meta+z');
  await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'number', value: 1 });
  await sp.clickRibbon('mac.data.filter');
  const filter = panel.locator('.fc-tb__ribbon-tools > [data-ribbon-command="mac.data.filter"]');
  await expect(filter).toHaveAttribute('aria-pressed', 'true');
  await page.locator('[data-ribbon-tab="view"]').click();
  await page.locator('[data-ribbon-tab="data"]').click();
  await expect(filter).toHaveAttribute('aria-pressed', 'true');
  await sp.clickRibbon('mac.data.filter');
  await expect(filter).toHaveAttribute('aria-pressed', 'false');
  await sp.clickRibbon('mac.data.workbookLinks');
  await expect(page.getByRole('dialog')).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(page.getByRole('dialog')).toHaveCount(0);
  await sp.clickRibbon('mac.data.sortCustom');
  await expect(page.locator('.fc-sortdlg__levels')).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(page.locator('.fc-sortdlg__levels')).toHaveCount(0);
  await sp.clickRibbon('mac.data.advancedFilter');
  await expect(page.getByRole('dialog')).toBeVisible();
  await page.keyboard.press('Escape');
  await sp.expectNoConsoleErrors();
}

export async function runMacValidationMenuScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.enter('A1', '9');
  await sp.goTo('A1');
  await page.locator('[data-ribbon-tab="data"]').click();
  await sp.clickRibbon('mac.data.validation');
  const menu = page.locator('#menu-mac-data-validation');
  await expect(menu).toBeVisible();
  await expect(
    menu.locator('[data-ribbon-command="mac.data.validation.circleInvalid"]'),
  ).toBeDisabled();
  await menu.locator('[data-ribbon-command="mac.data.validation.settings"]').click();
  const dialog = page.locator('.fc-fmtdlg--data-validation');
  await expect(dialog).toBeVisible();
  await dialog.locator('label:has(select[aria-label="Kind"]) .fc-select__button').click();
  await page
    .locator('.fc-select__list:not([hidden]) [role="option"]', { hasText: 'Whole number' })
    .click();
  await dialog.getByLabel('Value', { exact: true }).fill('1');
  await dialog.getByLabel('Upper value', { exact: true }).fill('2');
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(dialog).toBeHidden();
  await sp.clickRibbon('mac.data.validation');
  await menu.locator('[data-ribbon-command="mac.data.validation.circleInvalid"]').click();
  const circles = () =>
    page.evaluate(() =>
      Array.from(
        (
          window as unknown as {
            __fcInst: {
              store: { getState(): { errorIndicators: { validationCircles: Set<string> } } };
            };
          }
        ).__fcInst.store.getState().errorIndicators.validationCircles,
      ),
    );
  await expect.poll(circles).toEqual(['0:0:0']);
  await sp.clickRibbon('mac.data.validation');
  await menu.locator('[data-ribbon-command="mac.data.validation.clearCircles"]').click();
  await expect.poll(circles).toEqual([]);
  await sp.clickRibbon('mac.data.validation');
  await menu.locator('[data-ribbon-command="mac.data.validation.clearRules"]').click();
  const validation = () =>
    page.evaluate(
      () =>
        (
          window as unknown as {
            __fcInst: {
              store: { getState(): { format: { formats: Map<string, { validation?: unknown }> } } };
            };
          }
        ).__fcInst.store
          .getState()
          .format.formats.get('0:0:0')?.validation ?? null,
    );
  await expect.poll(validation).toBeNull();
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Meta+z');
  await expect.poll(validation).not.toBeNull();
  await sp.expectNoConsoleErrors();
}
