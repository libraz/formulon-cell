import { expect, type Page } from '@playwright/test';

import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

export async function runMacSparklineScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.expectNoStub();
  await sp.enter('A1', '1');
  await sp.enter('B1', '4');
  await sp.enter('C1', '2');
  await sp.goTo('A1:C1');
  await selectRibbonTab(page, 'insert');
  // Sparklines is a menu; Line opens the dialog with that type chosen.
  await sp.clickRibbon('mac.insert.sparkline');
  await page
    .locator('[role=menu]:not([hidden]) [data-ribbon-command="mac.insert.sparkline.line"]')
    .click();
  const dialog = page.locator('.fc-macsparkdlg[role=dialog]');
  await expect(dialog).toBeVisible();
  await dialog.locator('.fc-macsparkdlg__source').fill('A1:C1');
  await dialog.locator('.fc-macsparkdlg__destination').fill('E1');
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(dialog).toBeHidden();
  const read = () =>
    page.evaluate(() => {
      const w = window as unknown as { __fcInst: SpreadsheetInstance };
      for (const group of w.__fcInst.workbook.sheet(0).listSparklineGroups()) {
        const at = group.sparklines.find((s) => s.location.row === 0 && s.location.col === 4);
        if (at) return { type: group.type, source: at.source?.rect ?? null };
      }
      return null;
    });
  await expect
    .poll(read)
    .toEqual({ type: 0, source: { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 2 } });
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Meta+z');
  await expect.poll(read).toBeNull();
  await page.keyboard.press('Meta+Shift+z');
  await expect
    .poll(read)
    .toEqual({ type: 0, source: { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 2 } });
  await sp.expectNoConsoleErrors();
}

export async function runMacInkScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.expectNoStub();
  await selectRibbonTab(page, 'draw');
  await sp.clickRibbon('mac.draw.penRed');
  await expect(
    page.locator(
      '[data-ribbon-panel="draw"] .fc-tb__ribbon-tools > [data-ribbon-command="mac.draw.penRed"]',
    ),
  ).toHaveAttribute('aria-pressed', 'true');
  await expect(
    page.locator(
      '[data-ribbon-panel="draw"] .fc-tb__ribbon-tools > [data-ribbon-command="mac.draw.penBlack"]',
    ),
  ).toHaveAttribute('aria-pressed', 'false');
  await expect(
    page.locator(
      '[data-ribbon-panel="draw"] .fc-tb__ribbon-tools > [data-ribbon-command="mac.draw.toggle"]',
    ),
  ).toHaveAttribute('aria-pressed', 'true');
  const input = page.locator('.fc-mac-ink__input');
  await expect(input).toBeVisible();
  const grid = await page.locator('.fc-host__grid').boundingBox();
  if (!grid) throw new Error('Spreadsheet grid has no bounds');
  const start = { x: grid.x + 150, y: grid.y + 100 };
  await page.mouse.move(start.x, start.y);
  await page.mouse.down();
  await page.mouse.move(start.x + 80, start.y + 40, { steps: 12 });
  await page.mouse.up();
  const stroke = page.locator('[data-fc-mac-ink-id]');
  await expect(stroke).toHaveCount(1);
  await expect(stroke).toHaveAttribute('stroke', '#d13438');
  await page.keyboard.press('Escape');
  await expect(input).toHaveCount(0);
  await expect(
    page.locator(
      '[data-ribbon-panel="draw"] .fc-tb__ribbon-tools > [data-ribbon-command="mac.draw.toggle"]',
    ),
  ).toHaveAttribute('aria-pressed', 'false');
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Meta+z');
  await expect(stroke).toHaveCount(0);
  await page.keyboard.press('Meta+Shift+z');
  await expect(stroke).toHaveCount(1);
  await sp.clickRibbon('mac.draw.eraser');
  await page.locator('#menu-mac-draw-eraser [data-ribbon-command="mac.draw.eraser"]').click();
  await page.mouse.click(start.x + 40, start.y + 20);
  await expect(stroke).toHaveCount(0);
  await page.keyboard.press('Escape');
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Meta+z');
  await expect(stroke).toHaveCount(1);
  // Escape exits trackpad drawing; the next click must reactivate it immediately.
  for (let attempt = 0; attempt < 2; attempt += 1) {
    await sp.clickRibbon('mac.draw.trackpad');
    await expect(input).toBeVisible();
    await expect(
      page.locator(
        '[data-ribbon-panel="draw"] .fc-tb__ribbon-tools > [data-ribbon-command="mac.draw.trackpad"]',
      ),
    ).toHaveAttribute('aria-pressed', 'true');
    await page.keyboard.press('Escape');
    await expect(input).toHaveCount(0);
    await expect(
      page.locator(
        '[data-ribbon-panel="draw"] .fc-tb__ribbon-tools > [data-ribbon-command="mac.draw.trackpad"]',
      ),
    ).toHaveAttribute('aria-pressed', 'false');
  }
  // Leaving ink mode restores ordinary spreadsheet editing.
  await sp.enter('A1', 'normal input');
  await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'text', value: 'normal input' });
  await sp.expectNoConsoleErrors();
}

export async function runMacChartKindScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.expectNoStub();
  await sp.enter('A1', '1');
  await sp.enter('A2', '3');
  await sp.goTo('A1:A2');
  await selectRibbonTab(page, 'insert');
  await page.locator('.fc-tb__rb[data-ribbon-command="mac.insert.chartLine"]').click();
  await page
    .locator('[role=menu]:not([hidden]) [data-ribbon-command="mac.insert.chartArea"]')
    .click();
  const kinds = () =>
    page.evaluate(() => {
      const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
      const { store } = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst);
      return store.getState().charts.charts.map((chart) => chart.kind);
    });
  await expect.poll(kinds).toEqual(['area']);
  // The inserted chart is selected and opens Chart Design, as in the spreadsheet.
  await expect(page.locator('[data-ribbon-tab="chartDesign"]')).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await selectRibbonTab(page, 'insert');
  await page.locator('.fc-tb__rb[data-ribbon-command="mac.insert.chartPie"]').click();
  await page
    .locator('[role=menu]:not([hidden]) [data-ribbon-command="mac.insert.chartPie"]')
    .click();
  await expect.poll(kinds).toEqual(['area', 'pie']);
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Meta+z');
  await expect.poll(kinds).toEqual(['area']);
  await sp.expectNoConsoleErrors();
}

export async function runMacInsertGalleryScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.expectNoStub();
  const menu = page.locator('#menu-mac-insert-shapes');
  await selectRibbonTab(page, 'insert');
  await sp.clickRibbon('mac.insert.shapes');
  await expect(menu.locator('.fc-tb__visual-tile')).toHaveCount(7);
  await expect(menu.locator('.fc-tb__visual-tile svg')).toHaveCount(7);
  // The saved file has no line end, so the engine-backed Arrow tile stays disabled.
  await expect(menu.locator('[data-ribbon-command="mac.insert.shapeArrow"]')).toBeDisabled();
  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
  const kinds = ['Line', 'Rectangle', 'RoundedRectangle', 'Oval', 'Triangle', 'Diamond'];
  for (const kind of kinds) {
    // Inserting a shape selects it and switches to the Shape Format tab.
    await selectRibbonTab(page, 'insert');
    await sp.clickRibbon('mac.insert.shapes');
    await menu.locator(`[data-ribbon-command="mac.insert.shape${kind}"]`).click();
    await expect(menu).toBeHidden();
  }
  const shapes = () =>
    page.evaluate(() => {
      const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
      const { store } = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst);
      return store.getState().illustrations.illustrations.map((item) => item.shape);
    });
  await expect
    .poll(shapes)
    .toEqual(['line', 'rectangle', 'rounded-rectangle', 'oval', 'triangle', 'diamond']);
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Meta+z');
  await expect.poll(shapes).toEqual(['line', 'rectangle', 'rounded-rectangle', 'oval', 'triangle']);
  await sp.expectNoConsoleErrors();
}
