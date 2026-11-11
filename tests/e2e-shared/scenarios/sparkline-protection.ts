import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { revealRibbonCommand, selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };
type Platform = 'default' | 'mac';

const readGroups = (page: Page) =>
  page.evaluate(() =>
    (window as unknown as InstanceWindow).__fcInst.workbook
      .sheet(0)
      .listSparklineGroups()
      .map((group) => ({
        type: group.type,
        sparklines: group.sparklines.map((s) => ({ location: s.location, source: s.source })),
      })),
  );

const canUndo = (page: Page) =>
  page.evaluate(() => (window as unknown as InstanceWindow).__fcInst.history.canUndo);

const sparklineCommand = (page: Page, id: string) =>
  page.locator(`[data-ribbon-panel="sparklineDesign"] [data-ribbon-command="${id}"]`).first();

/** The Sparkline tab on a sparkline cell of a protected sheet: every control is unavailable except
 *  Axis, which opens a menu of unavailable entries, so the group cannot change. */
const expectTabUnavailable = async (page: Page) => {
  await selectRibbonTab(page, 'sparklineDesign');
  const commands = page.locator('[data-ribbon-panel="sparklineDesign"] [data-ribbon-command]');
  const ids = await commands.evaluateAll((els) =>
    els.map(
      (el) =>
        [
          (el as HTMLElement).dataset.ribbonCommand ?? '',
          (el as HTMLButtonElement).disabled,
        ] as const,
    ),
  );
  const sparklineIds = ids.filter(([id]) => id.startsWith('sparkline'));
  expect(sparklineIds.map(([id]) => id)).toEqual(
    expect.arrayContaining(['sparklineEditData', 'sparklineTypeColumn', 'sparklineClear']),
  );
  expect(sparklineIds.filter(([id, off]) => id !== 'sparklineAxis' && !off)).toEqual([]);

  const axis = sparklineCommand(page, 'sparklineAxis');
  await revealRibbonCommand(axis);
  await expect(axis).toBeEnabled();
  await axis.click();
  const menu = page.locator('#menu-sparkline-axis');
  await expect(menu).toBeVisible();
  const entries = menu.locator('[data-sparkline-axis]');
  expect(await entries.count()).toBeGreaterThan(0);
  for (const entry of await entries.all()) await expect(entry).toBeDisabled();
  await entries.first().click({ force: true });
  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
};

/** On a protected sheet a sparkline cell still shows the Sparkline tab, but nothing on it edits the
 *  group, whether the sparkline cells are locked or not; on the Mac chrome Insert > Sparkline is
 *  unavailable too. Before protection the same tab changes the group type. */
export async function runSparklineProtectionScenario(
  page: Page,
  platform: Platform,
  sparklineCells: 'locked' | 'unlocked',
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ fixture: 'empty', locale: 'ja', platform });
  await sheet.expectNoStub();
  await sheet.seed({ A1: 1, B1: 3, C1: 2, A2: 4, B2: 1, C2: 5, A3: 2, B3: 2, C3: 6 });
  await page.evaluate(() => {
    const instance = (window as unknown as InstanceWindow).__fcInst;
    instance.workbook.sheet(0).addSparklineGroup({
      type: 0,
      sparklines: [1, 2, 3].map((row) => ({
        source: `Sheet1!A${row}:C${row}`,
        location: `E${row}`,
      })),
    });
    instance.history.clear();
  });
  const line = await readGroups(page);
  expect(line).toHaveLength(1);
  expect(line[0]?.type).toBe(0);

  // Control: unprotected, the tab edits the group.
  await sheet.goTo('E1');
  await selectRibbonTab(page, 'sparklineDesign');
  const column = sparklineCommand(page, 'sparklineTypeColumn');
  await revealRibbonCommand(column);
  await expect(column).toBeEnabled();
  await column.click();
  await expect.poll(async () => (await readGroups(page))[0]?.type).toBe(1);
  await sheet.shortcut('z');
  await expect.poll(async () => (await readGroups(page))[0]?.type).toBe(0);

  await page.evaluate((unlock) => {
    const instance = (window as unknown as InstanceWindow).__fcInst;
    const target = instance.workbook.sheet(0);
    if (unlock) target.updateCellFormat('E1:E3', { protection: { locked: false } });
    target.protect();
    instance.history.clear();
  }, sparklineCells === 'unlocked');
  const before = await readGroups(page);

  await sheet.goTo('E2');
  await expect(page.locator('[data-ribbon-tab="sparklineDesign"]')).toBeVisible();
  await expectTabUnavailable(page);

  if (platform === 'mac') {
    await selectRibbonTab(page, 'insert');
    await expect(
      page.locator('[data-ribbon-command="mac.insert.sparkline"]').first(),
    ).toBeDisabled();
  }

  expect(await readGroups(page)).toEqual(before);
  expect(await canUndo(page)).toBe(false);
  await sheet.expectNoConsoleErrors();
}
