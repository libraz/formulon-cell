import { expect, type Page } from '@playwright/test';

import { UserJourneyPage } from '../pages/UserJourneyPage.js';

export async function runMacEditingPaletteScenario(
  page: Page,
  locale: 'en' | 'ja' = 'en',
): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac', locale });
  await sp.expectNoStub();
  await page.locator('[data-ribbon-tab="formulas"]').click();
  const palette = page.locator('.fc-mac-formula-palette[role=complementary]');
  const raw = '=1+SUM(2,3)*4';
  for (const source of ['inline', 'formulaBar'] as const) {
    await sp.enter('A1', '9');
    await sp.goTo('A1');
    await page.evaluate(() => {
      const inst = (window as Window & { __fcInst?: { history: { clear(): void } } }).__fcInst;
      if (!inst) throw new Error('Missing mounted instance.');
      inst.history.clear();
    });
    const edit =
      source === 'inline'
        ? page.locator('.fc-host__editor')
        : page.locator('.fc-host__formulabar-input');
    if (source === 'inline') await page.keyboard.press('Control+u');
    await edit.fill(raw);
    await edit.evaluate((element) => {
      (element as HTMLTextAreaElement).setSelectionRange(7, 10, 'backward');
    });
    const opener =
      source === 'inline'
        ? page.locator('[data-ribbon-command="mac.formulas.insertFunction"]')
        : page.locator('.fc-host__formulabar-fx');
    await opener.click();
    await expect(palette).toBeVisible();
    await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'number', value: 9 });
    await expect.poll(() => sp.readFormula('A1')).toBeNull();
    await palette.locator('[data-action="close"]').click();
    await expect(palette).toBeHidden();
    await expect(edit).toBeFocused();
    await expect(edit).toHaveValue(raw);
    expect(
      await edit.evaluate((element) => {
        const input = element as HTMLTextAreaElement;
        return [input.selectionStart, input.selectionEnd, input.selectionDirection];
      }),
    ).toEqual([7, 10, 'backward']);
    await opener.click();
    const fields = palette.locator('.fc-mac-formula-palette__argument input');
    await expect(fields).toHaveCount(2);
    await expect(fields.nth(0)).toHaveValue('2');
    await expect(fields.nth(1)).toHaveValue('3');
    await fields.nth(0).fill('5');
    await expect.poll(() => sp.formulaBarValue()).toBe('=1+SUM(5,3)*4');
    await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'number', value: 9 });
    await palette.locator('[data-action="done"]').click();
    await expect.poll(() => sp.readFormula('A1')).toBe('=1+SUM(5,3)*4');
    await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'number', value: 33 });
    await palette.locator('[data-action="close"]').click();
    await page.locator('.fc-host').focus();
    await page.keyboard.press('Meta+z');
    await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'number', value: 9 });
    await expect.poll(() => sp.readFormula('A1')).toBeNull();
    expect(
      await page.evaluate(() => {
        const inst = (window as Window & { __fcInst?: { history: { canUndo(): boolean } } })
          .__fcInst;
        if (!inst) throw new Error('Missing mounted instance.');
        return inst.history.canUndo();
      }),
    ).toBe(false);
    await page.keyboard.press('Meta+Shift+z');
    await expect.poll(() => sp.readFormula('A1')).toBe('=1+SUM(5,3)*4');

    await sp.goTo('A1');
    if (source === 'inline') await page.keyboard.press('Control+u');
    await edit.fill(raw);
    await opener.click();
    await expect(palette).toBeVisible();
    await sp.goTo('B2');
    await expect(palette).toBeHidden();
    await expect(page.locator('.fc-host__editor')).toHaveCount(0);
    await expect(page.locator('.fc-host__formulabar-tag')).toHaveValue('B2');
    await expect.poll(() => sp.readFormula('A1')).toBe('=1+SUM(5,3)*4');
  }
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
  // The new draft came from More Functions; Close restores its visible trigger.
  await expect(page.locator('[data-ribbon-command="mac.formulas.more"]')).toBeFocused();
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
