import { expect, type Page } from '@playwright/test';

import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { revealRibbonCommand, selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
type Platform = 'default' | 'mac';

const LINE = 0 as const;
const COLUMN = 1 as const;

/** Adds one sparkline group per `[type, location, source]`, then clears the undo history. */
const addGroups = (
  page: Page,
  groups: readonly (readonly [typeof LINE | typeof COLUMN, string, string])[],
) =>
  page.evaluate((list) => {
    const instance = (window as unknown as InstanceWindow).__fcInst;
    for (const [type, location, source] of list)
      instance.workbook.sheet(0).addSparklineGroup({ type, sparklines: [{ source, location }] });
    instance.history.clear();
  }, groups);

const sparklineCommand = (page: Page, id: string) =>
  page.locator(`[data-ribbon-panel="sparklineDesign"] [data-ribbon-command="${id}"]`).first();

/** The first group's empty-cell display. */
const emptyCellsAs = (page: Page) =>
  page.evaluate(() => {
    const group = (window as unknown as InstanceWindow).__fcInst.workbook
      .sheet(0)
      .listSparklineGroups()[0] as unknown as { emptyCellsAs?: number } | undefined;
    return group?.emptyCellsAs ?? null;
  });

/** Edit Data > Hidden & Empty Cells: on the Mac a submenu of the empty-cell displays that a click
 *  opens with no dialog, a pick being one undo step; elsewhere an entry that opens the Hidden and
 *  Empty Cell Settings dialog. */
export async function runSparklineHiddenEmptyScenario(
  page: Page,
  platform: Platform,
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'ja', platform });
  await sheet.expectNoStub();
  await sheet.seed({ A1: 1, C1: 3 });
  await addGroups(page, [[LINE, 'E1', 'Sheet1!A1:C1']]);
  await sheet.goTo('E1');
  await selectRibbonTab(page, 'sparklineDesign');
  const editData = sparklineCommand(page, 'sparklineEditData');
  await revealRibbonCommand(editData);
  await editData.locator('.fc-tb__rb-split-chevron').click();
  const menu = page.locator('#menu-sparkline-edit-data');
  await expect(menu).toBeVisible();

  if (platform === 'default') {
    await expect(menu.locator('[data-format-submenu="hidden-empty"]')).toHaveCount(0);
    await menu.locator('[data-sparkline-edit="hidden-empty"]').click();
    const dialog = page.getByRole('dialog', { name: '非表示および空白のセルの設定' });
    await expect(dialog).toBeVisible();
    await dialog.getByRole('button', { name: 'キャンセル', exact: true }).last().click();
    await expect(dialog).toBeHidden();
    await sheet.expectNoConsoleErrors();
    return;
  }

  await expect(menu.locator('[data-sparkline-edit="hidden-empty"]')).toHaveCount(0);
  const trigger = menu.locator('[data-format-submenu="hidden-empty"]');
  await expect(trigger).toHaveText('非表示および空白のセル');
  await trigger.click();
  const panel = page.locator('#menu-sparkline-edit-data-hidden-empty');
  await expect(panel).toBeVisible();
  await expect(page.getByRole('dialog')).toHaveCount(0);
  const before = await emptyCellsAs(page);
  const values = await panel
    .locator('[data-sparkline-empty]:not([data-sparkline-empty="hidden"])')
    .evaluateAll((els) => els.map((el) => Number((el as HTMLElement).dataset.sparklineEmpty)));
  expect(values).toHaveLength(3);
  const next = values.find((value) => value !== before);
  await panel.locator(`[data-sparkline-empty="${next}"]`).click();
  await expect.poll(() => emptyCellsAs(page)).toBe(next);
  await expect(page.getByRole('dialog')).toHaveCount(0);
  await sheet.shortcut('z');
  await expect.poll(() => emptyCellsAs(page)).toBe(before);
  await sheet.expectNoConsoleErrors();
}

/** Presets the desktop strip leaves to the full gallery behind its chevron. */
const FULL_GALLERY_ONLY = [
  'accentLighter40-4',
  'accentLighter40-5',
  'accentLighter40-6',
  'dark-1',
  'dark-2',
  'dark-3',
  'dark-4',
];

/** The Sparkline Style strip pages through 29 of the 36 presets, seven at a time, in 80 x 56
 *  chips. */
export async function runSparklineStyleStripScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'ja', platform: 'default' });
  await sheet.expectNoStub();
  await sheet.seed({ A1: 1, B1: 3, C1: 2 });
  await addGroups(page, [[LINE, 'E1', 'Sheet1!A1:C1']]);
  await sheet.goTo('E1');
  await selectRibbonTab(page, 'sparklineDesign');
  const panel = page.locator('[data-ribbon-panel="sparklineDesign"]');
  const chips = panel.locator('[data-sparkline-quick-style]');
  await revealRibbonCommand(chips.first());
  const box = await chips.first().boundingBox();
  expect(box && [Math.round(box.width), Math.round(box.height)]).toEqual([80, 56]);

  const prev = panel.locator('.fc-tb__sparkline-quick-prev');
  for (let turn = 0; turn < 10 && (await prev.isEnabled()); turn++) await prev.click();
  const next = panel.locator('.fc-tb__sparkline-quick-next');
  const read = () =>
    chips.evaluateAll((els) => els.map((el) => (el as HTMLElement).dataset.sparklineQuickStyle));
  const seen: (string | undefined)[] = [...(await read())];
  expect(seen).toHaveLength(7);
  for (let turn = 0; turn < 10 && (await next.isEnabled()); turn++) {
    await next.click();
    seen.push(...(await read()));
  }
  expect(seen).toHaveLength(29);
  expect(new Set(seen).size).toBe(29);
  for (const id of FULL_GALLERY_ONLY) expect(seen).not.toContain(id);
  await sheet.expectNoConsoleErrors();
}

/** Ink in the top and bottom thirds of a cell on the grid canvas, the top third also split into
 *  its left and right halves: pixels darker than the gridlines, inside the cell edges. */
const cellInk = (page: Page, ref: string) =>
  page.evaluate((a1) => {
    const w = window as unknown as InstanceWindow;
    const match = a1.match(/^([A-Z]+)(\d+)$/);
    if (!match) throw new Error(`bad ref ${a1}`);
    const col = [...(match[1] ?? '')].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
    const row = Number(match[2]) - 1;
    const state = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst).store.getState();
    const { cellRectUnclamped, layoutForView } = w.__fcProbes['render/geometry'];
    const rect = cellRectUnclamped(layoutForView(state), state.viewport, row, col);
    const canvas = document.querySelector<HTMLCanvasElement>('.fc-host__canvas');
    const grid = document.querySelector<HTMLElement>('.fc-host__grid');
    const ctx = canvas?.getContext('2d');
    if (!canvas || !grid || !ctx) throw new Error('no grid canvas');
    const cbox = canvas.getBoundingClientRect();
    const gbox = grid.getBoundingClientRect();
    const scale = canvas.width / cbox.width;
    // The plot starts 2px inside the cell; skip the gridlines on every side.
    const insetX = 3;
    const insetY = 1;
    const x = Math.round((gbox.left - cbox.left + rect.x + insetX) * scale);
    const y = Math.round((gbox.top - cbox.top + rect.y + insetY) * scale);
    const width = Math.round((rect.w - insetX * 2) * scale);
    const height = Math.round((rect.h - insetY * 2) * scale);
    const { data } = ctx.getImageData(x, y, width, height);
    const third = Math.floor(height / 3);
    let top = 0;
    let bottom = 0;
    let topLeft = 0;
    let topRight = 0;
    for (let py = 0; py < height; py++) {
      for (let px = 0; px < width; px++) {
        const i = (py * width + px) * 4;
        const dark = Math.min(data[i] ?? 255, data[i + 1] ?? 255, data[i + 2] ?? 255) < 190;
        if (!dark) continue;
        if (py < third) {
          top++;
          if (px < width / 2) topLeft++;
          else topRight++;
        } else if (py >= height - third) bottom++;
      }
    }
    return { top, bottom, topLeft, topRight };
  }, ref);

/** Flat and single-value series draw as the desktop does: a line over one value draws nothing and a
 *  column over one value fills the plot; over all-equal values the line runs along the top and the
 *  columns stand at full height. */
export async function runSparklineFlatRenderScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'ja', platform: 'default' });
  await sheet.expectNoStub();
  await sheet.seed({ A1: 5, A3: 3, B3: 3, C3: 3, A4: 3, B4: 3, C4: 3 });
  await addGroups(page, [
    [LINE, 'E1', 'Sheet1!A1:A1'],
    [COLUMN, 'E2', 'Sheet1!A1:A1'],
    [LINE, 'E3', 'Sheet1!A3:C3'],
    [COLUMN, 'E4', 'Sheet1!A4:C4'],
  ]);
  // The active cell sits away from the sparklines so no selection paint lands on them.
  await sheet.goTo('H10');

  await expect.poll(async () => (await cellInk(page, 'E2')).bottom).toBeGreaterThan(0);
  const singleLine = await cellInk(page, 'E1');
  const singleColumn = await cellInk(page, 'E2');
  const flatLine = await cellInk(page, 'E3');
  const flatColumn = await cellInk(page, 'E4');
  expect([singleLine.top, singleLine.bottom], 'line over one value').toEqual([0, 0]);
  expect(singleColumn.top, 'column over one value, top').toBeGreaterThan(0);
  expect(flatLine.top, 'flat line, top').toBeGreaterThan(0);
  expect(flatLine.bottom, 'flat line, bottom').toBe(0);
  expect(flatColumn.top, 'flat columns, top').toBeGreaterThan(0);
  expect(flatColumn.bottom, 'flat columns, bottom').toBeGreaterThan(0);
  await sheet.expectNoConsoleErrors();
}

/** The first group's Plot Data Right-to-Left setting. */
const rightToLeft = (page: Page) =>
  page.evaluate(() => {
    const group = (window as unknown as InstanceWindow).__fcInst.workbook
      .sheet(0)
      .listSparklineGroups()[0] as unknown as { rightToLeft?: boolean } | undefined;
    return group?.rightToLeft ?? null;
  });

/** Plot Data Right-to-Left, from the Axis dialog on the Mac and the Axis menu elsewhere, sets the
 *  group attribute as one undo step and mirrors the plot: the tall first column moves from the left
 *  of the cell to the right. */
export async function runSparklineRightToLeftScenario(
  page: Page,
  platform: Platform,
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'ja', platform });
  await sheet.expectNoStub();
  await sheet.seed({ A1: 9, B1: 1, C1: 1, D1: 1 });
  await addGroups(page, [[COLUMN, 'F1', 'Sheet1!A1:D1']]);
  await sheet.goTo('F1');
  expect(await rightToLeft(page)).toBe(false);
  // Read the plot with the active cell away from the sparkline.
  const plot = async () => {
    await sheet.goTo('H10');
    return cellInk(page, 'F1');
  };
  await expect.poll(async () => (await plot()).topLeft).toBeGreaterThan(0);
  expect((await plot()).topRight, 'left to right').toBe(0);

  await sheet.goTo('F1');
  await selectRibbonTab(page, 'sparklineDesign');
  const label = 'データを右から左にプロット';
  if (platform === 'mac') {
    const openAxis = async () => {
      const axis = sparklineCommand(page, 'sparklineAxisDialog');
      await revealRibbonCommand(axis);
      await axis.click();
      const dialog = page.locator('.fc-macsparkaxis');
      await expect(dialog).toBeVisible();
      return { dialog, box: dialog.getByRole('checkbox', { name: label }) };
    };
    let open = await openAxis();
    await expect(open.box).toBeEnabled();
    await expect(open.box).not.toBeChecked();
    await open.box.check();
    await open.dialog.locator('.fc-fmtdlg__btn--primary').click();
    await expect(open.dialog).toBeHidden();
    await expect.poll(() => rightToLeft(page)).toBe(true);
    open = await openAxis();
    await expect(open.box).toBeChecked();
    await open.dialog.locator('[data-dialog-action="cancel"]').click();
    await expect(open.dialog).toBeHidden();
  } else {
    const openAxis = async () => {
      const axis = sparklineCommand(page, 'sparklineAxis');
      await revealRibbonCommand(axis);
      await axis.click();
      const menu = page.locator('#menu-sparkline-axis');
      await expect(menu).toBeVisible();
      return { menu, entry: menu.locator('[data-sparkline-axis="rtl"]') };
    };
    let open = await openAxis();
    await expect(open.entry).toHaveText(label);
    await expect(open.entry).toHaveAttribute('aria-checked', 'false');
    await open.entry.click();
    await expect(open.menu).toBeHidden();
    await expect.poll(() => rightToLeft(page)).toBe(true);
    open = await openAxis();
    await expect(open.entry).toHaveAttribute('aria-checked', 'true');
    await page.keyboard.press('Escape');
    await expect(open.menu).toBeHidden();
  }

  await expect.poll(async () => (await plot()).topRight).toBeGreaterThan(0);
  expect((await plot()).topLeft, 'right to left').toBe(0);

  await sheet.shortcut('z');
  await expect.poll(() => rightToLeft(page)).toBe(false);
  await expect.poll(async () => (await plot()).topRight).toBe(0);
  expect((await plot()).topLeft, 'undone').toBeGreaterThan(0);
  await sheet.expectNoConsoleErrors();
}
