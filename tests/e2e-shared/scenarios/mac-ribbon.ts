import { expect, type Page } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

export async function runMacCollapsedRibbonScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await page.locator('[data-ribbon-tab="home"]').dblclick();
  await expect(page.locator('[data-ribbon-panel="home"]')).toBeHidden();
  const collapsedGrid = await page.locator('.fc-host__grid').boundingBox();
  await page.locator('[data-ribbon-tab="home"]').click();
  await expect(page.locator('[data-ribbon-panel="home"]')).toBeVisible();
  expect(await page.locator('.fc-host__grid').boundingBox()).toEqual(collapsedGrid);
  await page.locator('[data-ribbon-tab="insert"]').click();
  await expect(page.locator('[data-ribbon-panel="insert"]')).toBeVisible();
  await sp.clickRibbon('mac.insert.shapes');
  await expect(page.locator('#menu-mac-insert-shapes')).toBeVisible();
  await page.locator('[data-ribbon-tab="formulas"]').click();
  await expect(page.locator('#menu-mac-insert-shapes')).toBeHidden();
  await expect(page.locator('[data-ribbon-panel="formulas"]')).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(page.locator('[data-ribbon-panel="formulas"]')).toBeHidden();
  await page.locator('[data-ribbon-tab="data"]').click();
  await expect(page.locator('[data-ribbon-panel="data"]')).toBeVisible();
  await page.locator('.fc-host__grid').click({ position: { x: 200, y: 160 } });
  await expect(page.locator('[data-ribbon-panel="data"]')).toBeHidden();
  // A double click pins the ribbon without losing the clicked tab.
  await page.locator('[data-ribbon-tab="view"]').dblclick();
  await expect(page.locator('[data-ribbon-panel="view"]')).toBeVisible();
  await page.locator('.fc-host__grid').click({ position: { x: 200, y: 160 } });
  await expect(page.locator('[data-ribbon-panel="view"]')).toBeVisible();
  await sp.expectNoConsoleErrors();
}

export async function runMacFormulaPaletteScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.enter('A1', '1');
  await sp.goTo('A1');
  const grid = page.locator('.fc-host__grid');
  const originalGrid = await grid.boundingBox();
  if (!originalGrid) throw new Error('Missing initial grid geometry.');
  await page.locator('[data-ribbon-tab="formulas"]').click();
  const originalOpener = await page
    .locator('[data-ribbon-command="mac.formulas.insertFunction"]')
    .elementHandle();
  if (!originalOpener) throw new Error('Missing palette opener.');
  await page.locator('[data-ribbon-command="mac.formulas.insertFunction"]').focus();
  await sp.clickRibbon('mac.formulas.insertFunction');
  const palette = page.locator('.fc-mac-formula-palette[role=complementary]');
  const mirror = page.locator('.fc-host__formula-draft-mirror');
  await expect(palette).toBeVisible();
  await expect(palette).not.toHaveAttribute('aria-modal');
  const paletteBox = await palette.boundingBox();
  expect(paletteBox?.width).toBeCloseTo(300, 0);
  await expect
    .poll(async () => (await grid.boundingBox())?.width)
    .toBeCloseTo(originalGrid.width - 300, 0);
  await expect.poll(() => sp.formulaBarValue()).toBe('=');
  await expect(mirror).toBeVisible();
  await expect(mirror).toHaveText('=');
  await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'number', value: 1 });
  await expect.poll(() => sp.readFormula('A1')).toBeNull();
  const search = palette.locator('.fc-mac-formula-palette__search');
  await search.fill('ACOS');
  await palette.locator('[data-section="all"] [data-function-name="ACOS"]').click();
  await palette.getByRole('button', { name: 'Insert function', exact: true }).click();
  const argument = palette.locator('.fc-mac-formula-palette__argument').first();
  await expect(argument.locator('span')).toHaveText('number');
  await expect(argument.locator('small')).toHaveText('Must be a number from -1 to 1.');
  const help = palette.locator('.fc-mac-formula-palette__help');
  await expect(help).toContainText('arccosine');
  await expect(help).toContainText('Syntax: ACOS(number)');
  await expect(help.locator('a')).toHaveAttribute(
    'href',
    'https://support.microsoft.com/en-us/excel/functions/acos-function',
  );
  const field = palette.locator('.fc-mac-formula-palette__argument input').first();
  await field.fill('1');
  await expect.poll(() => sp.formulaBarValue()).toBe('=ACOS(1)');
  await expect(mirror).toHaveText('=ACOS(1)');
  await expect(palette.locator('[data-role="preview-value"]')).toHaveText('0');
  await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'number', value: 1 });
  await palette.getByRole('button', { name: 'Select range', exact: true }).click();
  const editingGrid = await grid.boundingBox();
  if (!editingGrid) throw new Error('Missing editing grid geometry.');
  await page.mouse.click(editingGrid.x + 26 + 75 * 1.5, editingGrid.y + 20 + 20 * 1.5);
  await expect(field).toHaveValue('B2');
  await expect.poll(() => sp.formulaBarValue()).toBe('=ACOS(B2)');
  await expect(page.locator('.fc-host__formulabar-tag')).toHaveValue('A1');
  await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'number', value: 1 });
  await expect.poll(() => sp.readFormula('B2')).toBeNull();
  await field.fill('1');
  await expect(palette.locator('[data-role="preview-value"]')).toHaveText('0');
  const argumentLayout = await palette
    .locator('.fc-mac-formula-palette__argument')
    .first()
    .evaluate((row) => {
      const label = row.querySelector('span')?.getBoundingClientRect();
      const input = row.querySelector('input')?.getBoundingClientRect();
      const range = row.querySelector('button')?.getBoundingClientRect();
      const hint = row.querySelector('small')?.getBoundingClientRect();
      if (!label || !input || !range || !hint) throw new Error('Missing argument controls.');
      return {
        labelBottom: label.bottom,
        inputTop: input.top,
        inputRight: input.right,
        rangeRight: range.right,
        rangeLeft: range.left,
        inputLeft: input.left,
        inputBottom: input.bottom,
        hintTop: hint.top,
      };
    });
  expect(argumentLayout.labelBottom).toBeLessThanOrEqual(argumentLayout.inputTop);
  expect(argumentLayout.rangeRight).toBeLessThanOrEqual(argumentLayout.inputRight);
  expect(argumentLayout.rangeLeft).toBeGreaterThan(argumentLayout.inputLeft);
  expect(argumentLayout.inputBottom).toBeLessThanOrEqual(argumentLayout.hintTop);
  await page.screenshot({
    path: `/tmp/formulon-mac-palette-${new URL(page.url()).port}-${page.context().browser()?.browserType().name()}.png`,
  });
  await palette.getByRole('button', { name: 'Done', exact: true }).click();
  await expect(palette).toBeVisible();
  await expect(palette).toHaveAttribute('data-state', 'arguments-committed');
  await expect.poll(() => sp.readFormula('A1')).toBe('=ACOS(1)');
  await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'number', value: 0 });
  await expect(mirror).toBeHidden();
  // Done retains the pane; another ribbon function must start a fresh draft
  // without changing the formula that was just committed.
  await sp.clickRibbon('mac.formulas.more');
  const menu = page.locator('#menu-mac-formulas-more');
  await menu.locator('[data-function-category-submenu="statistical"]').click();
  await menu.locator('[data-ribbon-command="mac.function.COUNTIF"]').click();
  await expect(palette.locator('.fc-mac-formula-palette__args-name')).toHaveText('COUNTIF');
  await expect(palette.locator('.fc-mac-formula-palette__argument input')).toHaveCount(2);
  await expect.poll(() => sp.readFormula('A1')).toBe('=ACOS(1)');
  await palette.getByRole('button', { name: 'Close', exact: true }).click();
  await expect(palette).toBeHidden();
  await expect
    .poll(async () => (await grid.boundingBox())?.width)
    .toBeCloseTo(originalGrid.width, 0);
  // Committing can redraw the ribbon and disconnect the original button.
  // Close returns to that exact opener when present, otherwise to the sheet.
  if (await originalOpener.evaluate((element) => element.isConnected)) {
    await expect
      .poll(() => originalOpener.evaluate((element) => element === document.activeElement))
      .toBe(true);
  } else {
    await expect(page.locator('.fc-host')).toBeFocused();
  }
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Meta+z');
  await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'number', value: 1 });
  await expect.poll(() => sp.readFormula('A1')).toBeNull();
  await page.keyboard.press('Meta+Shift+z');
  await expect.poll(() => sp.readFormula('A1')).toBe('=ACOS(1)');
  await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'number', value: 0 });
  await sp.expectNoConsoleErrors();
}

export async function runMacFunctionCatalogScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.goTo('C1');
  await page.locator('[data-ribbon-tab="formulas"]').click();
  await sp.clickRibbon('mac.formulas.math');
  const menu = page.locator('#menu-mac-formulas-math');
  await expect(menu.locator('[data-ribbon-command="mac.function.ACOS"]')).toBeVisible();
  await expect(menu.locator('[data-ribbon-command="mac.function.COUNTIF"]')).toHaveCount(0);
  await menu.locator('[data-ribbon-command="mac.function.ACOS"]').click();
  const dialog = page.locator('.fc-mac-formula-palette[role=complementary]');
  const fields = dialog.locator('.fc-mac-formula-palette__argument input');
  await expect(fields).toHaveCount(1);
  await fields.nth(0).fill('1');
  await dialog.getByRole('button', { name: 'Done', exact: true }).click();
  await expect.poll(() => sp.readFormula('C1')).toBe('=ACOS(1)');
  await expect.poll(() => sp.readValue('C1')).toEqual({ kind: 'number', value: 0 });
  await expect(dialog).toBeVisible();
  await dialog.getByRole('button', { name: 'Close', exact: true }).click();
  await sp.clickRibbon('mac.formulas.recent');
  const recentMenu = page.locator('#menu-mac-formulas-recent');
  await expect(recentMenu.locator('[data-ribbon-command="mac.function.ACOS"]')).toBeVisible();
  await page.keyboard.press('Escape');
  await sp.clickRibbon('mac.formulas.more');
  const moreMenu = page.locator('#menu-mac-formulas-more');
  const categoryTriggers = moreMenu.locator('[data-function-category-submenu]');
  await expect(categoryTriggers).toHaveCount(6);
  expect(
    await categoryTriggers.evaluateAll((buttons) =>
      buttons.map((button) => button.textContent?.trim()),
    ),
  ).toEqual(['Statistical', 'Engineering', 'Cube', 'Information', 'Compatibility', 'Web']);
  const statisticalTrigger = moreMenu.locator('[data-function-category-submenu="statistical"]');
  await statisticalTrigger.click();
  const statisticalPanel = moreMenu.locator('[data-function-category-panel="statistical"]');
  await expect(statisticalPanel).toBeVisible();
  await expect(
    statisticalPanel.locator('[data-ribbon-command="mac.function.COUNTIF"]'),
  ).toBeVisible();
  await expect(statisticalPanel.locator('[data-ribbon-command="mac.function.ACOS"]')).toHaveCount(
    0,
  );
  await statisticalPanel.locator('[data-ribbon-command="mac.function.COUNTIF"]').click();
  await expect(dialog).toBeVisible();
  await expect(dialog.locator('.fc-mac-formula-palette__sections')).toBeHidden();
  await expect(dialog.locator('.fc-mac-formula-palette__fields')).toBeVisible();
  await expect(dialog.locator('.fc-mac-formula-palette__args-name')).toHaveText('COUNTIF');
  await expect(dialog.locator('.fc-mac-formula-palette__argument input')).toHaveCount(2);
  await dialog.getByRole('button', { name: 'Close', exact: true }).click();
  await expect(dialog).toBeHidden();
  await expect(page.locator('[data-ribbon-command="mac.formulas.more"]')).toBeFocused();
  await sp.clickRibbon('mac.formulas.more');
  await moreMenu.locator('[data-function-category-submenu="statistical"]').focus();
  await page.keyboard.press('ArrowRight');
  await expect(statisticalPanel).toBeVisible();
  await expect(statisticalPanel.locator('button').first()).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(statisticalPanel).toBeHidden();
  await expect(statisticalTrigger).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(moreMenu).toBeHidden();
  await expect(page.locator('[data-ribbon-command="mac.formulas.more"]')).toBeFocused();

  // Native availability is a leaf property: recognized functions stay in
  // their Excel family, while class-3 engine stubs remain visible but cannot
  // launch the picker. INFO/CELL and the web functions with availability 0/2
  // remain ordinary enabled menu items.
  await sp.clickRibbon('mac.formulas.more');
  for (const [category, names] of [
    [
      'cube',
      [
        ['CUBESET', true],
        ['CUBEVALUE', true],
      ],
    ],
    [
      'information',
      [
        ['INFO', false],
        ['CELL', false],
      ],
    ],
    [
      'web',
      [
        ['ENCODEURL', false],
        ['FILTERXML', false],
        ['WEBSERVICE', true],
      ],
    ],
  ] as const) {
    const trigger = moreMenu.locator(`[data-function-category-submenu="${category}"]`);
    await trigger.click();
    const panel = moreMenu.locator(`[data-function-category-panel="${category}"]`);
    await expect(panel).toBeVisible();
    for (const [name, unavailable] of names) {
      const leaf = panel.locator(`[data-ribbon-command="mac.function.${name}"]`);
      await expect(leaf).toBeVisible();
      if (unavailable) {
        await expect(leaf).toBeDisabled();
        await expect(leaf).toHaveAttribute('aria-disabled', 'true');
        await expect(leaf).toHaveAttribute('data-function-unavailable', 'true');
        await expect(leaf).toHaveAttribute(
          'aria-description',
          /unavailable in the current calculation engine/i,
        );
      } else {
        await expect(leaf).not.toBeDisabled();
        await expect(leaf).toHaveAttribute('aria-disabled', 'false');
      }
    }
    await page.keyboard.press('Escape');
    await expect(panel).toBeHidden();
    await expect(moreMenu).toBeVisible();
    await expect(trigger).toBeFocused();
  }
  await page.keyboard.press('Escape');
  await expect(moreMenu).toBeHidden();

  // PY is recognized by the live catalog but has no family membership, so it
  // is checked through the All Functions picker rather than More.
  await sp.clickRibbon('mac.formulas.insertFunction');
  const fxDialog = page.locator('.fc-mac-formula-palette[role=complementary]');
  await expect(fxDialog).toBeVisible();
  const fxSearch = fxDialog.locator('.fc-mac-formula-palette__search');
  await fxSearch.fill('PY');
  const py = fxDialog.locator('[data-function-name="PY"]');
  await expect(py).toBeVisible();
  await expect(py).toHaveAttribute('aria-disabled', 'true');
  await expect(py).toHaveClass(/is-unavailable/);
  await expect(
    fxDialog.getByRole('button', { name: 'Insert function', exact: true }),
  ).toBeDisabled();
  await expect(py).toHaveAttribute(
    'aria-description',
    /unavailable in the current calculation engine/i,
  );
  await fxDialog.getByRole('button', { name: 'Close', exact: true }).click();
  await expect(fxDialog).toBeHidden();

  await page.setViewportSize({ width: 390, height: 900 });
  await sp.clickRibbon('mac.formulas.more');
  await moreMenu.scrollIntoViewIfNeeded();
  await expect(moreMenu).toBeVisible();
  await statisticalTrigger.hover();
  await expect(statisticalPanel).toBeVisible();
  const narrowParentBox = await moreMenu.boundingBox();
  const hoveredPanelBox = await statisticalPanel.boundingBox();
  expect(narrowParentBox).not.toBeNull();
  expect(hoveredPanelBox).not.toBeNull();
  if (!narrowParentBox || !hoveredPanelBox)
    throw new Error('Missing hovered More submenu geometry.');
  const hoveredRectsDisjoint =
    hoveredPanelBox.x >= narrowParentBox.x + narrowParentBox.width ||
    narrowParentBox.x >= hoveredPanelBox.x + hoveredPanelBox.width ||
    hoveredPanelBox.y >= narrowParentBox.y + narrowParentBox.height ||
    narrowParentBox.y >= hoveredPanelBox.y + hoveredPanelBox.height;
  expect(hoveredRectsDisjoint).toBe(true);
  await statisticalTrigger.click();
  await expect(statisticalPanel).toBeVisible();
  const narrowPanelBox = await statisticalPanel.boundingBox();
  const narrowViewport = page.viewportSize();
  expect(narrowPanelBox).not.toBeNull();
  expect(narrowViewport).not.toBeNull();
  if (!narrowPanelBox || !narrowViewport) throw new Error('Missing narrow More panel geometry.');
  expect(narrowPanelBox.x).toBeGreaterThanOrEqual(-1);
  expect(narrowPanelBox.x + narrowPanelBox.width).toBeLessThanOrEqual(narrowViewport.width + 1);
  expect(narrowPanelBox.y).toBeGreaterThanOrEqual(-1);
  expect(narrowPanelBox.y + narrowPanelBox.height).toBeLessThanOrEqual(narrowViewport.height + 1);
  const narrowPanelMetrics = await statisticalPanel.evaluate((element) => ({
    clientHeight: element.clientHeight,
    overflowY: getComputedStyle(element).overflowY,
    scrollHeight: element.scrollHeight,
  }));
  expect(narrowPanelMetrics.overflowY).toBe('auto');
  expect(narrowPanelMetrics.scrollHeight).toBeGreaterThan(narrowPanelMetrics.clientHeight);

  const lastStatisticalLeaf = statisticalPanel
    .locator('[data-ribbon-command^="mac.function."]')
    .last();
  await lastStatisticalLeaf.scrollIntoViewIfNeeded();
  await expect(lastStatisticalLeaf).toBeVisible();
  const lastLeafBox = await lastStatisticalLeaf.boundingBox();
  expect(lastLeafBox).not.toBeNull();
  if (!lastLeafBox) throw new Error('Missing last Statistical leaf geometry.');
  expect(lastLeafBox.y).toBeGreaterThanOrEqual(-1);
  expect(lastLeafBox.y + lastLeafBox.height).toBeLessThanOrEqual(narrowViewport.height + 1);

  const narrowCountIf = statisticalPanel.locator('[data-ribbon-command="mac.function.COUNTIF"]');
  await narrowCountIf.click();
  await expect(dialog).toBeVisible();
  await expect(dialog.locator('.fc-mac-formula-palette__args-name')).toHaveText('COUNTIF');
  await dialog.getByRole('button', { name: 'Close', exact: true }).click();
  await expect(dialog).toBeHidden();
  await expect(page.locator('[data-ribbon-command="mac.formulas.more"]')).toBeFocused();

  await sp.clickRibbon('mac.formulas.more');
  await statisticalTrigger.focus();
  await page.keyboard.press('ArrowRight');
  await expect(statisticalPanel).toBeVisible();
  await narrowCountIf.focus();
  await page.keyboard.press('Enter');
  await expect(dialog).toBeVisible();
  await expect(dialog.locator('.fc-mac-formula-palette__args-name')).toHaveText('COUNTIF');
  await dialog.getByRole('button', { name: 'Close', exact: true }).click();
  await expect(dialog).toBeHidden();
  await expect(page.locator('[data-ribbon-command="mac.formulas.more"]')).toBeFocused();

  await sp.clickRibbon('mac.formulas.more');
  await statisticalTrigger.focus();
  await page.keyboard.press('ArrowRight');
  await expect(statisticalPanel).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(statisticalPanel).toBeHidden();
  await expect(statisticalTrigger).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(moreMenu).toBeHidden();
  await expect(page.locator('[data-ribbon-command="mac.formulas.more"]')).toBeFocused();
  await sp.expectNoConsoleErrors();
}

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

export async function runMacScaleAndSheetViewScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  const read = () =>
    page.evaluate(() => {
      const inst = (
        window as unknown as {
          __fcInst: {
            store: {
              getState(): import('../../../packages/formulon-cell/src/store/types.js').State;
            };
          };
        }
      ).__fcInst;
      const state = inst.store.getState();
      return {
        setup: state.pageSetup.setupBySheet.get(0),
        views: state.sheetViews.views,
        freezeRows: state.layout.freezeRows,
      };
    });
  await page.locator('[data-ribbon-tab="pageLayout"]').click();
  for (const [id, value] of [
    ['scaleWidth', '1'],
    ['scaleHeight', '2'],
  ]) {
    const select = page.locator(`[data-ribbon-select="${id}"]`);
    await select.locator('.fc-tb__rb-dd__btn').click();
    await select.locator(`[role="option"][data-value="${value}"]`).click();
  }
  await expect.poll(async () => (await read()).setup).toMatchObject({ fitWidth: 1, fitHeight: 2 });
  await page.locator('[data-ribbon-tab="view"]').click();
  await sp.clickRibbon('mac.view.freeze');
  await page
    .locator('#menu-mac-view-freeze [data-ribbon-command="mac.view.freeze.firstRow"]')
    .click();
  await sp.clickRibbon('mac.view.sheetViewSave');
  await expect.poll(async () => (await read()).views.length).toBe(1);
  const id = (await read()).views[0].id;
  await sp.clickRibbon('mac.view.freeze');
  await page.locator('#menu-mac-view-freeze [data-ribbon-command="mac.view.freeze.off"]').click();
  await expect.poll(async () => (await read()).freezeRows).toBe(0);
  const select = page.locator('[data-ribbon-select="sheetViewSelect"]');
  await select.locator('.fc-tb__rb-dd__btn').click();
  await select.locator(`[role="option"][data-value="${id}"]`).click();
  await expect.poll(async () => (await read()).freezeRows).toBe(1);
  await sp.clickRibbon('mac.view.sheetViewDelete');
  await expect.poll(async () => (await read()).views.length).toBe(0);
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Meta+z');
  await expect.poll(async () => (await read()).views.length).toBe(1);
  await sp.expectNoConsoleErrors();
}

export async function runMacRibbonSwitchingScenario(
  page: Page,
  locale: 'en' | 'ja' = 'en',
): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount({ platform: 'mac', locale });
  await sp.expectNoStub();
  const tabs = [
    'home',
    'insert',
    'draw',
    'pageLayout',
    'formulas',
    'data',
    'review',
    'view',
    'automate',
  ];
  for (const id of [...tabs, ...tabs.toReversed()]) {
    await page.locator(`[data-ribbon-tab="${id}"]`).click();
    await expect(page.locator('[data-ribbon-tab][aria-selected="true"]')).toHaveCount(1);
    await expect(page.locator(`[data-ribbon-tab="${id}"]`)).toHaveAttribute(
      'aria-selected',
      'true',
    );
    await expect(page.locator('.fc-tb__ribbon:not([hidden])')).toHaveCount(1);
    const panel = page.locator(`[data-ribbon-panel="${id}"]`);
    await expect(panel).toBeVisible();
    expect(
      await panel.locator('[data-ribbon-command], [data-ribbon-control]').count(),
    ).toBeGreaterThan(0);
  }
  // The tab list is a keyboard surface too; its active panel must follow focus.
  await page.locator('[data-ribbon-tab="home"]').focus();
  await page.keyboard.press('ArrowRight');
  await expect(page.locator('[data-ribbon-tab="insert"]')).toBeFocused();
  await expect(page.locator('[data-ribbon-panel="insert"]')).toBeVisible();
  await page.keyboard.press('End');
  await expect(page.locator('[data-ribbon-tab="automate"]')).toBeFocused();
  await expect(page.locator('[data-ribbon-panel="automate"]')).toBeVisible();
  await page.keyboard.press('Home');
  await expect(page.locator('[data-ribbon-tab="home"]')).toBeFocused();
  await expect(page.locator('[data-ribbon-panel="home"]')).toBeVisible();
  await page.setViewportSize({ width: 1440, height: 900 });
  await page.locator('[data-ribbon-tab="formulas"]').click();
  const functionLibrary = page.locator(
    '[data-ribbon-panel="formulas"] .fc-tb__ribbon-group--function-library',
  );
  await expect(functionLibrary).toBeVisible();
  const libraryGeometry = await functionLibrary.evaluate((group) => {
    const panel = group.closest<HTMLElement>('.fc-tb__ribbon');
    const buttons = Array.from(
      group.querySelectorAll<HTMLButtonElement>(
        '.fc-tb__ribbon-tools > button[data-ribbon-command]',
      ),
    );
    const panelRect = panel?.getBoundingClientRect();
    const groupRect = group.getBoundingClientRect();
    return {
      groupWidth: groupRect.width,
      panelHeight: panelRect?.height ?? 0,
      panelLeft: panelRect?.left ?? 0,
      panelRight: panelRect?.right ?? 0,
      panelTop: panelRect?.top ?? 0,
      panelBottom: panelRect?.bottom ?? 0,
      buttons: buttons.map((button) => {
        const rect = button.getBoundingClientRect();
        const icon = button.querySelector<SVGElement>('.fc-tb__rb-icon')?.getBoundingClientRect();
        const labelElement = Array.from(button.children)
          .filter((child): child is HTMLElement => child instanceof HTMLElement)
          .find(
            (child) =>
              !child.classList.contains('fc-tb__rb-icon') &&
              !child.classList.contains('fc-tb__rb-split-chevron'),
          );
        const label = labelElement?.getBoundingClientRect();
        const labelFragments = (() => {
          if (!labelElement) return [];
          const range = document.createRange();
          range.selectNodeContents(labelElement);
          return Array.from(range.getClientRects())
            .filter((fragment) => fragment.width > 0 && fragment.height > 0)
            .map((fragment) => ({
              left: fragment.left,
              right: fragment.right,
              top: fragment.top,
              bottom: fragment.bottom,
            }));
        })();
        return {
          id: button.dataset.ribbonCommand,
          left: rect.left,
          right: rect.right,
          top: rect.top,
          bottom: rect.bottom,
          width: rect.width,
          iconBottom: icon?.bottom ?? 0,
          labelTop: label?.top ?? 0,
          labelFragments,
          hasMenu: button.getAttribute('aria-haspopup') === 'menu',
        };
      }),
    };
  });
  if (locale === 'ja') {
    expect(libraryGeometry.groupWidth).toBeGreaterThanOrEqual(502);
    expect(libraryGeometry.groupWidth).toBeLessThanOrEqual(506);
  } else {
    expect(libraryGeometry.groupWidth).toBeGreaterThanOrEqual(502);
  }
  expect(libraryGeometry.panelHeight).toBe(75);
  expect(libraryGeometry.buttons).toHaveLength(10);
  const libraryTop = libraryGeometry.buttons[0]?.top ?? 0;
  for (const [index, button] of libraryGeometry.buttons.entries()) {
    expect(Math.abs(button.top - libraryTop), button.id).toBeLessThanOrEqual(1);
    expect(button.bottom - button.top, button.id).toBeGreaterThanOrEqual(65);
    expect(button.bottom - button.top, button.id).toBeLessThanOrEqual(67);
    expect(button.iconBottom, button.id).toBeLessThanOrEqual(button.labelTop + 1);
    expect(button.top, button.id).toBeGreaterThanOrEqual(libraryGeometry.panelTop - 1);
    expect(button.bottom, button.id).toBeLessThanOrEqual(libraryGeometry.panelBottom + 1);
    if (locale === 'ja') {
      if (index === 0) {
        expect(button.width, button.id).toBeGreaterThanOrEqual(37);
        expect(button.width, button.id).toBeLessThanOrEqual(39);
      } else {
        expect(button.width, button.id).toBeGreaterThanOrEqual(49);
        expect(button.width, button.id).toBeLessThanOrEqual(51);
      }
    } else {
      expect(button.width, button.id).toBeGreaterThanOrEqual(index === 0 ? 38 : 50);
      if (index > 0) {
        const previous = libraryGeometry.buttons[index - 1];
        expect(button.left, button.id).toBeGreaterThanOrEqual((previous?.right ?? 0) - 1);
      }
    }
    expect(button.hasMenu, button.id).toBe(index > 0);
  }
  const labelsRequiringContainment = new Set([
    'mac.formulas.insertFunction',
    'mac.formulas.lookup',
    'mac.formulas.more',
  ]);
  const containmentButtons = libraryGeometry.buttons.filter((button) =>
    labelsRequiringContainment.has(button.id ?? ''),
  );
  expect(containmentButtons).toHaveLength(3);
  expect(containmentButtons.map((button) => button.id).toSorted()).toEqual(
    [...labelsRequiringContainment].toSorted(),
  );
  for (const button of containmentButtons) {
    expect(button.labelFragments.length, `${button.id} text fragments`).toBeGreaterThan(0);
    for (const fragment of button.labelFragments) {
      expect(fragment.left, `${button.id} label left`).toBeGreaterThanOrEqual(button.left - 1);
      expect(fragment.right, `${button.id} label right`).toBeLessThanOrEqual(button.right + 1);
      expect(fragment.top, `${button.id} label top`).toBeGreaterThanOrEqual(button.top - 1);
      expect(fragment.bottom, `${button.id} label bottom`).toBeLessThanOrEqual(button.bottom + 1);
      expect(fragment.left, `${button.id} panel left`).toBeGreaterThanOrEqual(
        libraryGeometry.panelLeft - 1,
      );
      expect(fragment.right, `${button.id} panel right`).toBeLessThanOrEqual(
        libraryGeometry.panelRight + 1,
      );
      expect(fragment.top, `${button.id} panel top`).toBeGreaterThanOrEqual(
        libraryGeometry.panelTop - 1,
      );
      expect(fragment.bottom, `${button.id} panel bottom`).toBeLessThanOrEqual(
        libraryGeometry.panelBottom + 1,
      );
    }
  }
  await page.locator('[data-ribbon-toggle]').click();
  await page.locator('[data-ribbon-display-option="singleLine"]').click();
  await page.setViewportSize({ width: 1440, height: 900 });
  await page.locator('[data-ribbon-tab="formulas"]').click();
  const singleLineGeometry = await functionLibrary.evaluate((group) => {
    const panel = group.closest<HTMLElement>('.fc-tb__ribbon');
    const buttons = Array.from(
      group.querySelectorAll<HTMLButtonElement>(
        '.fc-tb__ribbon-tools > button[data-ribbon-command]',
      ),
    );
    const panelRect = panel?.getBoundingClientRect();
    return {
      panelHeight: panelRect?.height ?? 0,
      buttons: buttons.map((button) => {
        const rect = button.getBoundingClientRect();
        const icon = button.querySelector<SVGElement>('.fc-tb__rb-icon')?.getBoundingClientRect();
        const label = Array.from(button.children)
          .filter((child): child is HTMLElement => child instanceof HTMLElement)
          .find(
            (child) =>
              !child.classList.contains('fc-tb__rb-icon') &&
              !child.classList.contains('fc-tb__rb-split-chevron'),
          )
          ?.getBoundingClientRect();
        const chevron = button
          .querySelector<HTMLElement>(':scope > .fc-tb__rb-split-chevron')
          ?.getBoundingClientRect();
        return {
          id: button.dataset.ribbonCommand,
          left: rect.left,
          right: rect.right,
          top: rect.top,
          bottom: rect.bottom,
          width: rect.width,
          height: rect.height,
          iconLeft: icon?.left ?? 0,
          iconTop: icon?.top ?? 0,
          iconWidth: icon?.width ?? 0,
          iconHeight: icon?.height ?? 0,
          labelLeft: label?.left ?? 0,
          labelRight: label?.right ?? 0,
          labelTop: label?.top ?? 0,
          labelBottom: label?.bottom ?? 0,
          chevronLeft: chevron?.left ?? 0,
          chevronRight: chevron?.right ?? 0,
          chevronTop: chevron?.top ?? 0,
          chevronBottom: chevron?.bottom ?? 0,
          hasMenu: button.getAttribute('aria-haspopup') === 'menu',
        };
      }),
    };
  });
  expect(singleLineGeometry.panelHeight).toBe(42);
  expect(singleLineGeometry.buttons).toHaveLength(10);
  const firstSingleLineButton = singleLineGeometry.buttons[0];
  const secondSingleLineButton = singleLineGeometry.buttons[1];
  expect(
    (secondSingleLineButton?.left ?? 0) - (firstSingleLineButton?.right ?? 0),
    'single-line Insert Function gap',
  ).toBeLessThanOrEqual(4);
  for (const button of singleLineGeometry.buttons) {
    expect(button.height, button.id).toBeGreaterThanOrEqual(29);
    expect(button.height, button.id).toBeLessThanOrEqual(31);
    expect(button.iconWidth, button.id).toBe(16);
    expect(button.iconHeight, button.id).toBe(16);
    const buttonCenter = (button.top + button.bottom) / 2;
    const iconCenter = (button.iconTop + button.iconTop + button.iconHeight) / 2;
    expect(Math.abs(iconCenter - buttonCenter), button.id).toBeLessThanOrEqual(3);
    const labelCenter = (button.labelTop + button.labelBottom) / 2;
    expect(Math.abs(labelCenter - buttonCenter), button.id).toBeLessThanOrEqual(3);
    if (button.hasMenu) {
      expect(button.chevronRight, button.id).toBeLessThanOrEqual(button.right + 1);
      expect(button.chevronLeft, button.id).toBeGreaterThanOrEqual(button.left - 1);
      expect(button.chevronTop, button.id).toBeGreaterThanOrEqual(button.top - 1);
      expect(button.chevronBottom, button.id).toBeLessThanOrEqual(button.bottom + 1);
      const chevronCenter = (button.chevronTop + button.chevronBottom) / 2;
      expect(Math.abs(chevronCenter - buttonCenter), button.id).toBeLessThanOrEqual(4);
    }
  }
  await page.locator('[data-ribbon-toggle]').click();
  await page.locator('[data-ribbon-display-option="full"]').click();
  await expect(page.locator('.fc-tb__ribbon-shell--full')).toHaveCount(1);
  // Compact groups must stay inside the ribbon in either display mode.
  for (const mode of ['full', 'singleLine']) {
    if (mode === 'singleLine') {
      await page.locator('[data-ribbon-toggle]').click();
      await page.locator('[data-ribbon-display-option="singleLine"]').click();
    }
    for (const width of [1440, 1024, 768, 390]) {
      await page.setViewportSize({ width, height: 900 });
      for (const id of tabs) {
        await page.locator(`[data-ribbon-tab="${id}"]`).click();
        const clipped = await page.locator(`[data-ribbon-panel="${id}"]`).evaluate((panel) => {
          const bounds = panel.getBoundingClientRect();
          return Array.from(
            panel.querySelectorAll<HTMLElement>(
              '.fc-tb__ribbon-tools > button, .fc-tb__ribbon-tools > .fc-tb__rb-dd, .fc-tb__ribbon-tools > .fc-tb__rb-color',
            ),
          )
            .filter((command) => {
              const rect = command.getBoundingClientRect();
              return (
                rect.height > 0 &&
                (rect.top < bounds.top - 1 ||
                  rect.bottom > bounds.bottom + 1 ||
                  Array.from(command.children).some((child) => {
                    const content = child.getBoundingClientRect();
                    return (
                      content.height > 0 &&
                      (content.top < bounds.top - 1 ||
                        content.bottom > bounds.bottom + 1 ||
                        content.left < rect.left - 1 ||
                        content.right > rect.right + 1)
                    );
                  }))
              );
            })
            .map((command) => command.dataset.ribbonCommand);
        });
        expect(clipped, `${id} controls clipped in ${mode} mode at ${width}px`).toEqual([]);
      }
    }
  }
}

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

export async function runMacViewShowScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await page.locator('[data-ribbon-tab="view"]').click();
  await sp.clickRibbon('mac.view.show');
  const menu = page.locator('#menu-mac-view-show');
  await expect(menu).toBeVisible();
  await menu.locator('[data-ribbon-command="mac.view.formulaBar"]').click();
  await expect(page.locator('.fc-host__formulabar')).toBeHidden();
  await sp.clickRibbon('mac.view.show');
  await menu.locator('[data-ribbon-command="mac.view.formulaBar"]').click();
  await expect(page.locator('.fc-host__formulabar')).toBeVisible();
  await expect(page.getByRole('dialog')).toHaveCount(0);
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

export async function runMacPageAndCalculationScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  const read = () =>
    page.evaluate(() => {
      const inst = (
        window as unknown as {
          __fcInst: {
            store: {
              getState(): import('../../../packages/formulon-cell/src/store/types.js').State;
            };
            workbook: { calcMode(): number };
          };
        }
      ).__fcInst;
      return {
        setup: inst.store.getState().pageSetup.setupBySheet.get(0),
        mode: inst.workbook.calcMode(),
      };
    });
  await page.locator('[data-ribbon-tab="pageLayout"]').click();
  await sp.clickRibbon('mac.page.orientation');
  await page
    .locator('#menu-mac-page-orientation [data-ribbon-command="mac.page.orientation.landscape"]')
    .click();
  await expect.poll(read).toMatchObject({ setup: { orientation: 'landscape' } });
  await sp.clickRibbon('mac.page.margins');
  await page
    .locator('#menu-mac-page-margins [data-ribbon-command="mac.page.margins.narrow"]')
    .click();
  await expect.poll(read).toMatchObject({ setup: { margins: { left: 0.25, right: 0.25 } } });
  await sp.goTo('C4');
  await sp.clickRibbon('mac.page.pageBreaks');
  await page
    .locator('[role=menu]:not([hidden]) [data-ribbon-command="mac.page.break.insert"]')
    .click();
  await expect
    .poll(read)
    .toMatchObject({ setup: { manualPageBreakRows: [3], manualPageBreakCols: [2] } });
  await sp.clickRibbon('mac.page.pageBreaks');
  await page
    .locator('[role=menu]:not([hidden]) [data-ribbon-command="mac.page.break.reset"]')
    .click();
  expect((await read()).setup?.manualPageBreakRows ?? []).toEqual([]);
  await page.locator('[data-ribbon-tab="formulas"]').click();
  await sp.clickRibbon('mac.formulas.calcOptions');
  await page
    .locator('[role=menu]:not([hidden]) [data-ribbon-command="mac.formulas.calc.manual"]')
    .click();
  await expect.poll(read).toMatchObject({ mode: 1 });
  await sp.clickRibbon('mac.formulas.calcOptions');
  await page.keyboard.press('Escape');
  await expect(page.locator('.fc-tb__menu--mac:not([hidden])')).toHaveCount(0);
  await expect(page.locator('[data-ribbon-command="mac.formulas.calcOptions"]')).toBeFocused();
  await sp.clickRibbon('mac.formulas.calcOptions');
  await page
    .locator('[role=menu]:not([hidden]) [data-ribbon-command="mac.formulas.calc.auto"]')
    .click();
  await expect.poll(read).toMatchObject({ mode: 0 });
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

export async function runMacFormulaAndAutomationScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.goTo('B1');
  await page.locator('[data-ribbon-tab="formulas"]').click();
  await sp.clickRibbon('mac.formulas.logical');
  await page
    .locator('.fc-tb__menu--mac:not([hidden]) [data-ribbon-command="mac.function.IF"]')
    .click();
  const dialog = page.locator('.fc-mac-formula-palette[role=complementary]');
  await expect(dialog).toBeVisible();
  const fields = dialog.locator('.fc-mac-formula-palette__argument input');
  await fields.nth(0).fill('1>0');
  await fields.nth(1).fill('42');
  await fields.nth(2).fill('0');
  await dialog.getByRole('button', { name: 'Done', exact: true }).click();
  await expect(dialog).toBeVisible();
  await dialog.getByRole('button', { name: 'Close', exact: true }).click();
  await expect(dialog).toBeHidden();
  await expect.poll(() => sp.readValue('B1')).toEqual({ kind: 'number', value: 42 });
  await sp.enter('A3', '1');
  await sp.goTo('A1:A3');
  await page.locator('[data-ribbon-tab="automate"]').click();
  await sp.clickRibbon('mac.automate.gallery');
  await page
    .locator('.fc-tb__menu--mac:not([hidden]) [data-ribbon-command="mac.automate.countEmptyRows"]')
    .click();
  const result = page.locator('.fc-macautomationresult[role=dialog]');
  await expect(result).toBeVisible();
  await expect(result.locator('.fc-macautomationresult__summary')).toContainText('1');
  await page.keyboard.press('Escape');
  await expect(result).toHaveCount(0);
  await page.locator('[data-ribbon-tab="review"]').click();
  await sp.clickRibbon('mac.review.stats');
  await expect(page.locator('.fc-macstatsdlg[role=dialog]')).toBeVisible();
  await sp.expectNoConsoleErrors();
}

export async function runMacHomeMenusScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.enter('A1', 'Formatted');
  await sp.goTo('A1');
  const underline = page.locator('.fc-tb__rb[data-ribbon-command="underline"]');
  await underline.locator('.fc-tb__rb-split-chevron').click();
  const menu = page.locator('#menu-underline');
  await expect(menu).toBeVisible();
  await menu.locator('[data-underline-action="double"]').click();
  const read = () =>
    page.evaluate(
      () =>
        (
          window as unknown as {
            __fcInst: {
              store: { getState(): { format: { formats: Map<string, { underline?: string }> } } };
            };
          }
        ).__fcInst.store
          .getState()
          .format.formats.get('0:0:0')?.underline,
    );
  await expect.poll(read).toBe('double');
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Meta+z');
  await expect.poll(read).not.toBe('double');
  await page.locator('.fc-tb__rb[data-ribbon-command="merge"] .fc-tb__rb-split-chevron').click();
  await expect(page.locator('#menu-merge')).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(page.locator('#menu-merge')).toBeHidden();
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

export async function runMacPageToggleAndZoomScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await page.locator('[data-ribbon-tab="pageLayout"]').click();
  for (const [action, initial] of [
    ['showGridlines', true],
    ['printGridlines', false],
    ['showHeadings', true],
    ['printHeadings', false],
  ] as const) {
    const button = page.locator(
      `[data-ribbon-panel="pageLayout"] [data-ribbon-command="mac.page.${action}"]`,
    );
    await expect(button).toHaveAttribute('aria-pressed', String(initial));
    await button.click();
    await expect(button).toHaveAttribute('aria-pressed', String(!initial));
  }
  await page.locator('[data-ribbon-tab="view"]').click();
  const panel = page.locator('[data-ribbon-panel="view"]');
  const views = panel
    .locator('section')
    .filter({ has: page.locator('[data-ribbon-select="sheetViewSelect"]') });
  await expect(views.locator('[data-ribbon-command="mac.view.standard"]')).toHaveCount(0);
  await sp.clickRibbon('mac.view.zoom');
  const dialog = page.getByRole('dialog');
  await expect(dialog).toBeVisible();
  await dialog.locator('input[type="number"]').fill('150');
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();
  const zoom = () =>
    page.evaluate(
      () =>
        (
          window as unknown as {
            __fcInst: { store: { getState(): { viewport: { zoom: number } } } };
          }
        ).__fcInst.store.getState().viewport.zoom,
    );
  await expect.poll(zoom).toBe(1.5);
  await sp.clickRibbon('mac.view.zoom100');
  await expect.poll(zoom).toBe(1);
  await page.locator('[data-ribbon-tab="pageLayout"]').click();
  await expect(page.locator('[data-ribbon-command="mac.page.printGridlines"]')).toHaveAttribute(
    'aria-pressed',
    'true',
  );
  await sp.expectNoConsoleErrors();
}

export async function runMacAuditMenusScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.enter('A1', '1');
  await sp.enter('B1', '=A1+1');
  await sp.enter('C1', '=B1+1');
  await sp.goTo('B1');
  await page.locator('[data-ribbon-tab="formulas"]').click();
  await sp.clickRibbon('mac.formulas.precedents');
  await sp.clickRibbon('mac.formulas.dependents');
  const traces = () =>
    page.evaluate(() =>
      (
        window as unknown as {
          __fcInst: { store: { getState(): { traces: { items: { kind: string }[] } } } };
        }
      ).__fcInst.store
        .getState()
        .traces.items.map((item) => item.kind),
    );
  await expect.poll(traces).toEqual(expect.arrayContaining(['precedent', 'dependent']));
  await sp.clickRibbon('mac.formulas.removeArrows');
  const remove = page.locator('#menu-mac-formulas-remove-arrows');
  await expect(remove).toBeVisible();
  await remove.locator('[data-ribbon-command="mac.formulas.removeArrows.precedents"]').click();
  await expect.poll(traces).toEqual(['dependent']);
  await sp.clickRibbon('mac.formulas.removeArrows');
  await expect(
    remove.locator('[data-ribbon-command="mac.formulas.removeArrows.precedents"]'),
  ).toBeDisabled();
  await remove.locator('[data-ribbon-command="mac.formulas.removeArrows.dependents"]').click();
  await expect.poll(traces).toEqual([]);
  await sp.clickRibbon('mac.formulas.precedents');
  await expect.poll(traces).toEqual(['precedent']);
  await sp.clickRibbon('mac.formulas.removeArrows');
  await remove.locator('[data-ribbon-command="mac.formulas.removeArrows.all"]').click();
  await expect.poll(traces).toEqual([]);
  await sp.enter('E1', '=A1/0');
  await sp.goTo('A1');
  await sp.clickRibbon('mac.formulas.errorCheck');
  const error = page.locator('#menu-mac-formulas-error-check');
  await expect(error).toBeVisible();
  await error.locator('[data-ribbon-command="mac.formulas.errorCheck.run"]').click();
  await expect(page.locator('.fc-host__formulabar-tag')).toHaveValue('E1');
  await sp.clickRibbon('mac.formulas.errorCheck');
  await expect(
    error.locator('[data-ribbon-command="mac.formulas.errorCheck.trace"]'),
  ).toBeEnabled();
  await error.locator('[data-ribbon-command="mac.formulas.errorCheck.trace"]').click();
  await expect.poll(traces).toEqual(['precedent']);
  await sp.clickRibbon('mac.formulas.errorCheck');
  await error.locator('[data-ribbon-command="mac.formulas.errorCheck.ignore"]').click();
  await expect
    .poll(() =>
      page.evaluate(() =>
        Array.from(
          (
            window as unknown as {
              __fcInst: {
                store: { getState(): { errorIndicators: { ignoredErrors: Set<string> } } };
              };
            }
          ).__fcInst.store.getState().errorIndicators.ignoredErrors,
        ),
      ),
    )
    .toContain('0:0:4');
  await sp.goTo('A1');
  await sp.clickRibbon('mac.formulas.errorCheck');
  await expect(
    error.locator('[data-ribbon-command="mac.formulas.errorCheck.ignore"]'),
  ).toBeDisabled();
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

export async function runMacDirectAutomationScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.enter('A1', '1');
  await sp.enter('A3', '3');
  await sp.goTo('A1:A3');
  await page.locator('[data-ribbon-tab="automate"]').click();
  const panel = page.locator('[data-ribbon-panel="automate"]');
  for (const id of [
    'allRowsColumns',
    'freezeSelection',
    'makeSubtable',
    'removeHyperlinks',
    'countEmptyRows',
    'tableToJson',
    'newPivotTable',
  ]) {
    await expect(
      panel.locator(`.fc-tb__ribbon-tools > [data-ribbon-command="mac.automate.${id}"]`),
    ).toHaveCount(1);
  }
  await panel
    .locator('.fc-tb__ribbon-tools > [data-ribbon-command="mac.automate.countEmptyRows"]')
    .click();
  const result = page.locator('.fc-macautomationresult[role=dialog]');
  await expect(result).toBeVisible();
  await expect(result.locator('.fc-macautomationresult__summary')).toContainText('1');
  await page.keyboard.press('Escape');
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

export async function runMacAutoSumMenuScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.enter('A1', '1');
  await sp.enter('A2', '3');
  await sp.goTo('A3');
  await page.locator('[data-ribbon-tab="formulas"]').click();
  await sp.clickRibbon('mac.formulas.autoSum');
  const menu = page.locator('#menu-mac-formulas-autosum');
  await expect(menu).toBeVisible();
  for (const fn of ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN']) {
    await expect(menu.locator(`[data-ribbon-command="mac.autosum.${fn}"]`)).toBeEnabled();
  }
  await menu.locator('[data-ribbon-command="mac.autosum.AVERAGE"]').click();
  await expect.poll(() => sp.readFormula('A3')).toBe('=AVERAGE(A1:A2)');
  await expect.poll(() => sp.readValue('A3')).toEqual({ kind: 'number', value: 2 });
  await sp.clickRibbon('mac.formulas.recent');
  await expect(
    page.locator('#menu-mac-formulas-recent [data-ribbon-command="mac.function.AVERAGE"]'),
  ).toBeVisible();
  await page.keyboard.press('Escape');
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Meta+z');
  await expect.poll(() => sp.readFormula('A3')).toBeNull();
  await sp.expectNoConsoleErrors();
}
