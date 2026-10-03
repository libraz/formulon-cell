import { expect, type Page } from '@playwright/test';

import { UserJourneyPage } from '../pages/UserJourneyPage.js';

export async function runMacSparklineScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.enter('A1', '1');
  await sp.enter('B1', '4');
  await sp.enter('C1', '2');
  await sp.goTo('A1:C1');
  await page.locator('[data-ribbon-tab="insert"]').click();
  await sp.clickRibbon('mac.insert.sparkline');
  const dialog = page.locator('.fc-macsparkdlg[role=dialog]');
  await expect(dialog).toBeVisible();
  await dialog.locator('.fc-macsparkdlg__source').fill('A1:C1');
  await dialog.locator('.fc-macsparkdlg__destination').fill('E1');
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(dialog).toBeHidden();
  const read = () =>
    page.evaluate(() => {
      const inst = (
        window as unknown as {
          __fcInst: { store: { getState(): { sparkline: { sparklines: Map<string, unknown> } } } };
        }
      ).__fcInst;
      return inst.store.getState().sparkline.sparklines.get('0:0:4') ?? null;
    });
  await expect.poll(read).toMatchObject({ kind: 'line' });
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Meta+z');
  await expect.poll(read).toBeNull();
  await page.keyboard.press('Meta+Shift+z');
  await expect.poll(read).toMatchObject({ kind: 'line' });
  await sp.expectNoConsoleErrors();
}

export async function runMacInkScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await page.locator('[data-ribbon-tab="draw"]').click();
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
  await sp.expectNoStub();
  await sp.enter('A1', '1');
  await sp.enter('A2', '3');
  await sp.goTo('A1:A2');
  await page.locator('[data-ribbon-tab="insert"]').click();
  await page.locator('.fc-tb__rb[data-ribbon-command="mac.insert.chartLine"]').click();
  await page
    .locator('[role=menu]:not([hidden]) [data-ribbon-command="mac.insert.chartArea"]')
    .click();
  const kinds = () =>
    page.evaluate(() =>
      (
        window as unknown as {
          __fcInst: { store: { getState(): { charts: { charts: { kind: string }[] } } } };
        }
      ).__fcInst.store
        .getState()
        .charts.charts.map((chart) => chart.kind),
    );
  await expect.poll(kinds).toEqual(['area']);
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
  await sp.expectNoStub();
  await page.locator('[data-ribbon-tab="insert"]').click();
  const kinds = ['Line', 'Arrow', 'Rectangle', 'RoundedRectangle', 'Oval', 'Triangle', 'Diamond'];
  for (const kind of kinds) {
    await sp.clickRibbon('mac.insert.shapes');
    const menu = page.locator('#menu-mac-insert-shapes');
    await expect(menu.locator('.fc-tb__visual-tile')).toHaveCount(7);
    await expect(menu.locator('.fc-tb__visual-tile svg')).toHaveCount(7);
    await menu.locator(`[data-ribbon-command="mac.insert.shape${kind}"]`).click();
    await expect(menu).toBeHidden();
  }
  const shapes = () =>
    page.evaluate(() =>
      (
        window as unknown as {
          __fcInst: {
            store: { getState(): { illustrations: { illustrations: { shape?: string }[] } } };
          };
        }
      ).__fcInst.store
        .getState()
        .illustrations.illustrations.map((item) => item.shape),
    );
  await expect
    .poll(shapes)
    .toEqual(['line', 'arrow', 'rectangle', 'rounded-rectangle', 'oval', 'triangle', 'diamond']);
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Meta+z');
  await expect
    .poll(shapes)
    .toEqual(['line', 'arrow', 'rectangle', 'rounded-rectangle', 'oval', 'triangle']);
  await sp.expectNoConsoleErrors();
}
