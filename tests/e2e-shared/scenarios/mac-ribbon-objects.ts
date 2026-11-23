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
  await page.locator('.fc-host__grid').focus();
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
  await page.locator('.fc-host__grid').focus();
  await page.keyboard.press('Meta+z');
  await expect(stroke).toHaveCount(0);
  await page.keyboard.press('Meta+Shift+z');
  await expect(stroke).toHaveCount(1);
  await sp.clickRibbon('mac.draw.eraser');
  await page
    .locator('#menu-mac-draw-eraser [data-ribbon-command="mac.draw.eraser.stroke"]')
    .click();
  await page.mouse.click(start.x + 40, start.y + 20);
  await expect(stroke).toHaveCount(0);
  await page.keyboard.press('Escape');
  await page.locator('.fc-host__grid').focus();
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
  await page.locator('.fc-host__grid').focus();
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
  await expect(menu.locator('.fc-tb__visual-tile:disabled')).toHaveCount(0);
  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
  const kinds = ['Line', 'Arrow', 'Rectangle', 'RoundedRectangle', 'Oval', 'Triangle', 'Diamond'];
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
  const inserted = ['line', 'arrow', 'rectangle', 'rounded-rectangle', 'oval', 'triangle'];
  await expect.poll(shapes).toEqual([...inserted, 'diamond']);
  // The arrow saves as a straight connector with an end only at its tail.
  const saved = await page.evaluate(async () => {
    const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
    const { WorkbookHandle } = w.__fcProbes['engine/workbook-handle'];
    const workbook = await WorkbookHandle.loadBytes(w.__fcInst.workbook.save().bytes);
    try {
      return workbook.listShapes(0).map((shape) => ({
        preset: shape.presetGeometry,
        head: shape.line?.headEnd != null,
        tail: shape.line?.tailEnd != null,
      }));
    } finally {
      workbook.dispose();
    }
  });
  expect(saved[1]).toEqual({ preset: 'straightConnector1', head: false, tail: true });
  expect(saved.filter((shape) => shape.tail)).toHaveLength(1);
  await page.locator('.fc-host__grid').focus();
  await page.keyboard.press('Meta+z');
  await expect.poll(shapes).toEqual(inserted);
  await sp.expectNoConsoleErrors();
}

/** The Charts group draws two columns of 36 pt icon cells three high, then Waterfall over Combo as
 *  a two-row stack of 38x26 pt cells; a width-fitting stage never widens a cell past that. */
export async function runMacChartCellGeometryScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await page.setViewportSize({ width: 2560, height: 900 });
  await sp.expectNoStub();
  await selectRibbonTab(page, 'insert');
  const charts = page.locator('[data-ribbon-panel="insert"] .fc-tb__ribbon-group--charts');
  const icons = [
    'mac.insert.chartColumn',
    'mac.insert.chartLine',
    'mac.insert.chartPie',
    'mac.insert.chartHierarchy',
    'mac.insert.chartStatistical',
    'mac.insert.chartScatter',
  ];
  const stack = ['mac.insert.chartWaterfall', 'mac.insert.chartCombo'];
  type Box = { x: number; y: number; width: number; height: number };
  const measure = async (): Promise<(id: string) => Box> => {
    const boxes = await charts.evaluate(
      (group, ids) =>
        ids.map((id) => {
          const cell = group.querySelector(`.fc-tb__rb[data-ribbon-command="${id}"]`);
          if (!cell) throw new Error(`${id} is missing from the Charts group`);
          const { x, y, width, height } = cell.getBoundingClientRect();
          return [id, { x, y, width, height }] as const;
        }),
      [...icons, ...stack],
    );
    const byId = new Map<string, Box>(boxes);
    return (id) => {
      const box = byId.get(id);
      if (!box) throw new Error(`${id} was not measured`);
      return box;
    };
  };
  await expect(charts).not.toHaveClass(/fc-tb__ribbon-group--(narrow|condensed|icon-only)/);
  const drawn = await measure();
  const top = drawn('mac.insert.chartColumn');
  const secondColumn = drawn('mac.insert.chartHierarchy');
  icons.forEach((id, index) => {
    const cell = drawn(id);
    expect(cell.width, id).toBeCloseTo(36, 0);
    expect(cell.height, id).toBeCloseTo(24, 0);
    expect(cell.y - top.y, id).toBeCloseTo((index % 3) * 24, 0);
    expect(cell.x, id).toBeCloseTo((index < 3 ? top : secondColumn).x, 0);
  });
  expect(secondColumn.x).toBeGreaterThan(top.x + top.width - 1);
  const waterfall = drawn('mac.insert.chartWaterfall');
  const combo = drawn('mac.insert.chartCombo');
  for (const cell of [waterfall, combo]) {
    expect(cell.width).toBeCloseTo(38, 0);
    expect(cell.height).toBeCloseTo(26, 0);
  }
  expect(waterfall.y - top.y).toBeCloseTo(3, 0);
  expect(combo.y - waterfall.y).toBeCloseTo(28, 0);
  expect(combo.x).toBeCloseTo(waterfall.x, 0);
  const scatter = drawn('mac.insert.chartScatter');
  expect(waterfall.x).toBeGreaterThan(scatter.x + scatter.width - 1);

  // At 1280 the panel fits by stages on its neighbours; Charts keeps a stage only if it got
  // narrower, and no stage may widen a chart cell.
  await page.setViewportSize({ width: 1280, height: 900 });
  const insert = page.locator('[data-ribbon-panel="insert"]');
  await expect
    .poll(() =>
      insert.evaluate(
        (p) =>
          p.scrollWidth <= p.clientWidth &&
          p.querySelector(':scope > .fc-tb__ribbon-group--narrow') !== null,
      ),
    )
    .toBe(true);
  await expect(charts).not.toHaveClass(/fc-tb__ribbon-group--collapsed/);
  const fitted = await measure();
  for (const id of [...icons, ...stack])
    expect(fitted(id).width, id).toBeLessThanOrEqual(drawn(id).width + 0.5);
  await sp.expectNoConsoleErrors();
}

/** With a cell selected, Insert greys Slicer, Timeline and Symbol; Symbol opens up while a cell
 *  is being edited, and greys again once the edit is committed. */
export async function runMacInsertSymbolStateScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac', locale: 'en', fixture: 'empty' });
  await page.setViewportSize({ width: 1920, height: 900 });
  await sp.expectNoStub();
  await selectRibbonTab(page, 'insert');
  const command = (id: string) =>
    page.locator(`[data-ribbon-panel="insert"] [data-ribbon-command="${id}"]`).first();
  const symbol = command('mac.insert.symbol');
  await sp.goTo('B2');
  for (const id of ['mac.insert.slicer', 'mac.insert.timeline', 'mac.insert.symbol'])
    await expect(command(id)).toBeDisabled();
  await sp.typeIntoActiveCell('ab', { commit: false });
  await expect(page.locator('.fc-host__editor')).toBeVisible();
  await expect(symbol).toBeEnabled();
  await page.keyboard.press('Enter');
  await expect(symbol).toBeDisabled();
  await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'text', value: 'ab' });
  await sp.expectNoConsoleErrors();
}
