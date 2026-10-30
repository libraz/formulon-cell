import { expect, type Page } from '@playwright/test';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/mount/types.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

export type FormulaHelperSurface = 'inline' | 'bar';

export async function runFormulaKeyboardPointing(
  page: Page,
  surface: FormulaHelperSurface,
  platform: 'default' | 'mac',
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en', platform });
  await sheet.expectNoStub();
  await page.evaluate(() => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    instance.applyChanges([
      { addr: { sheet: 0, row: 3, col: 3 }, input: '=1' },
      { addr: { sheet: 0, row: 2, col: 2 }, input: '1' },
      { addr: { sheet: 0, row: 3, col: 2 }, input: '2' },
    ]);
  });
  const input = page
    .locator(surface === 'inline' ? '.fc-host__editor' : '.fc-host__formulabar-input')
    .first();
  const begin = async () => {
    await sheet.goTo('D4');
    if (surface === 'inline') await page.keyboard.press('F2');
    else await input.focus();
    await expect(input).toBeFocused();
  };
  const caret = () => input.evaluate((el) => (el as HTMLTextAreaElement).selectionStart);
  const selectCaret = (offset: number) =>
    input.evaluate(
      (el, index) => (el as HTMLTextAreaElement).setSelectionRange(index, index),
      offset,
    );

  await begin();
  await input.fill('=SUM(');
  await page.keyboard.press('ArrowLeft');
  await expect(input).toHaveValue('=SUM(');
  expect(await caret()).toBe(4);
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('F2');
  await page.keyboard.press('ArrowLeft');
  await expect(input).toHaveValue('=SUM(C4');
  await page.keyboard.press('Shift+ArrowUp');
  await expect(input).toHaveValue('=SUM(C3:C4');
  await page.keyboard.press('Shift+ArrowDown');
  await expect(input).toHaveValue('=SUM(C4');
  await page.keyboard.press('Shift+ArrowUp');
  await expect(input).toHaveValue('=SUM(C3:C4');
  expect(
    await page.evaluate(
      () =>
        (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.store.getState().selection
          .active,
    ),
  ).toEqual({ sheet: 0, row: 3, col: 3 });
  expect(await sheet.readFormula('D4')).toBe('=1');
  await page.keyboard.type(')');
  await page.keyboard.press('Enter');
  await expect.poll(() => sheet.readFormula('D4')).toBe('=SUM(C3:C4)');
  expect(await sheet.readValue('D4')).toEqual({ kind: 'number', value: 3 });

  await begin();
  await input.fill('=SU(C3:C4)');
  await selectCaret(3);
  const sum = page.locator('.fc-autocomplete__item').filter({
    has: page.locator('.fc-autocomplete__name', { hasText: /^SUM$/ }),
  });
  await expect(sum).toBeVisible();
  await sum.click();
  await page.keyboard.press('Tab');
  await expect(input).toHaveValue('=SUM(C3:C4)');
  expect(await caret()).toBe(5);
  await page.keyboard.press('ArrowLeft');
  await expect(input).toHaveValue('=SUM(C4)');
  await page.keyboard.press('Enter');
  await expect.poll(() => sheet.readFormula('D4')).toBe('=SUM(C4)');
  expect(await sheet.readValue('D4')).toEqual({ kind: 'number', value: 2 });

  await begin();
  await input.fill('=SU');
  await expect(page.locator('.fc-autocomplete')).toBeVisible();
  await page.keyboard.press('ArrowDown');
  await expect(input).toHaveValue('=SU');
  await page.keyboard.press('F2');
  await expect(page.locator('.fc-autocomplete')).toHaveCount(0);
  await page.keyboard.press('ArrowLeft');
  await expect(input).toHaveValue('=C4');
  await input.fill('="SUM"');
  await selectCaret(3);
  await page.keyboard.press('ArrowLeft');
  await expect(input).toHaveValue('="SUM"');
  expect(await caret()).toBe(2);
  await selectCaret(6);
  await page.keyboard.press('ArrowLeft');
  await expect(input).toHaveValue('="SUM"');
  expect(await caret()).toBe(5);
  await page.keyboard.press('Escape');
  expect(await sheet.readFormula('D4')).toBe('=SUM(C4)');

  await begin();
  await input.fill('=1+SUM(A1)+2');
  await input.evaluate((el) => (el as HTMLTextAreaElement).setSelectionRange(3, 10));
  await page.keyboard.press('F2');
  await page.keyboard.press('ArrowLeft');
  await expect(input).toHaveValue('=1+C4+2');
  await page.keyboard.press('Escape');
  expect(await sheet.readFormula('D4')).toBe('=SUM(C4)');

  if (surface === 'inline') {
    await sheet.goTo('D4');
    await page.keyboard.type('=');
    await page.keyboard.press('ArrowUp');
    await expect(input).toHaveValue('=D3');
    await page.keyboard.press('Shift+ArrowLeft');
    await expect(input).toHaveValue('=C3:D3');
    const before = await caret();
    await page.keyboard.press('F2');
    await page.keyboard.press('ArrowLeft');
    await expect(input).toHaveValue('=C3:D3');
    expect(await caret()).toBe(before - 1);
    await page.keyboard.press('Escape');
    expect(await sheet.readFormula('D4')).toBe('=SUM(C4)');
  }
  await sheet.expectNoConsoleErrors();
}

export async function runFormulaHelperInteraction(
  page: Page,
  surface: FormulaHelperSurface,
  platform: 'default' | 'mac',
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en', platform });
  await sheet.expectNoStub();
  const input = page
    .locator(surface === 'inline' ? '.fc-host__editor' : '.fc-host__formulabar-input')
    .first();
  const begin = async () => {
    await sheet.goTo('A1');
    if (surface === 'inline') await page.keyboard.press('F2');
    else await input.focus();
  };
  const source = () =>
    page.evaluate(() => {
      const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
      const addr = { sheet: 0, row: 0, col: 0 };
      return {
        formula: instance.workbook.cellFormula(addr),
        value: instance.workbook.getValue(addr),
      };
    });
  const list = page.locator('.fc-autocomplete');
  const options = list.locator('.fc-autocomplete__item');

  await begin();
  await input.fill('=SU');
  await expect(list).toBeVisible();
  expect(await options.count()).toBeGreaterThan(8);
  const scrolling = await list.locator('.fc-autocomplete__list').evaluate(async (element) => {
    element.scrollTop = element.scrollHeight;
    const requested = element.scrollTop;
    // Allow native scroll events and helper repositioning to complete.
    await new Promise<void>((resolve) =>
      requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
    );
    return { requested, actual: element.scrollTop };
  });
  expect(scrolling.requested).toBeGreaterThan(0);
  expect(scrolling.actual, 'manual list scrolling must not snap back to the active option').toBe(
    scrolling.requested,
  );
  await page.keyboard.press('End');
  const last = await options.last().getAttribute('id');
  await expect(input).toHaveAttribute('aria-activedescendant', last ?? '');
  await page.keyboard.press('Home');
  const first = await options.first().getAttribute('id');
  await expect(input).toHaveAttribute('aria-activedescendant', first ?? '');
  await page.keyboard.press('PageDown');
  expect(await input.getAttribute('aria-activedescendant')).not.toBe(first);
  await page.keyboard.press('PageUp');
  await expect(input).toHaveAttribute('aria-activedescendant', first ?? '');

  const sum = options.filter({ has: page.locator('.fc-autocomplete__name', { hasText: /^SUM$/ }) });
  await sum.click();
  await expect(input).toHaveValue('=SU');
  await expect(input).toBeFocused();
  await expect(page.locator('.fc-autocomplete__help')).toContainText('Adds');
  await sum.dblclick();
  await expect(input).toHaveValue('=SUM(');
  await expect(page.locator('.fc-arghelper')).toBeVisible();
  if (surface === 'inline') {
    expect(
      await page.evaluate(() => {
        const mode = (
          window as unknown as { __fcInst: SpreadsheetInstance }
        ).__fcInst.store.getState().ui.editor;
        return mode.kind === 'idle' ? null : mode.raw;
      }),
    ).toBe('=SUM(');
  }
  await page.keyboard.type('1,2)');
  await page.keyboard.press('Enter');
  await expect.poll(source).toEqual({ formula: '=SUM(1,2)', value: { kind: 'number', value: 3 } });

  await begin();
  await input.fill('=IF(TRUE,SU,7)');
  await input.evaluate((element) => (element as HTMLTextAreaElement).setSelectionRange(11, 11));
  await expect(list).toBeVisible();
  await sum.click();
  await page.keyboard.press('Tab');
  await expect(input).toHaveValue('=IF(TRUE,SUM(,7)');
  expect((await source()).formula).toBe('=SUM(1,2)');
  await page.keyboard.press('Escape');

  await begin();
  await input.fill('=SU');
  await page.keyboard.press('Escape');
  await expect(list).toHaveCount(0);
  await expect(input).toHaveValue('=SU');
  await page.keyboard.press('Alt+ArrowDown');
  await expect(list).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(list).toHaveCount(0);
  await page.keyboard.press('Escape');
  expect((await source()).formula).toBe('=SUM(1,2)');

  await begin();
  await input.fill('=SU');
  await page.keyboard.press('ArrowDown');
  const selected = await options
    .filter({ has: page.locator('.fc-autocomplete__name', { hasText: /^SUM$/ }) })
    .getAttribute('id');
  // Select SUM by pointer, then use the documented keyboard insertion trigger.
  await sum.click();
  await expect(input).toHaveAttribute('aria-activedescendant', selected ?? '');
  await page.keyboard.press('Tab');
  await expect(input).toHaveValue('=SUM(');
  await expect(list).toHaveCount(0);
  await page.keyboard.press('Escape');

  await begin();
  await input.fill('=SU');
  await page.keyboard.press('Enter');
  await expect.poll(async () => (await source()).formula).toBe('=SU');
  if (surface === 'inline') await expect(input).toHaveCount(0);

  await begin();
  await input.fill('=SU');
  await input.dispatchEvent('compositionstart');
  await input.fill('=SUM');
  await expect(list).toHaveCount(0);
  await expect(page.locator('.fc-arghelper')).toHaveCount(0);
  await input.dispatchEvent('compositionend');
  await expect(list).toBeVisible();
  await input.evaluate((element) => (element as HTMLTextAreaElement).setSelectionRange(0, 4));
  await expect(list).toHaveCount(0);
  await page.keyboard.press('Escape');

  await page.evaluate(() => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    if (!instance.workbook.setDefinedNameEntry('Rate.Factor', '=2')) {
      throw new Error('The real workbook rejected the defined-name fixture');
    }
    instance.formula.register('FOO', () => 0);
  });
  await begin();
  await input.fill('=Rate.F+1');
  await input.evaluate((element) => (element as HTMLTextAreaElement).setSelectionRange(7, 7));
  await expect(options.locator('.fc-autocomplete__name')).toHaveText(['Rate.Factor']);
  await options.first().click();
  await page.keyboard.press('Tab');
  await expect(input).toHaveValue('=Rate.Factor+1');
  await page.keyboard.press('Enter');
  await expect.poll(source).toEqual({
    formula: '=Rate.Factor+1',
    value: { kind: 'number', value: 3 },
  });
  await sheet.expectNoConsoleErrors();
}

export interface FormulaHelperVisualCase {
  surface: FormulaHelperSurface;
  platform: 'default' | 'mac';
  theme: 'paper' | 'ink' | 'contrast';
  locale: 'en' | 'ja';
  edge: 'start' | 'end';
  rtl: boolean;
}

// coverwise generate, strength=2, seed=42: all 70 parameter pairs covered.
export const formulaHelperVisualCases: readonly FormulaHelperVisualCase[] = [
  {
    surface: 'bar',
    platform: 'default',
    theme: 'contrast',
    locale: 'en',
    edge: 'start',
    rtl: false,
  },
  { surface: 'inline', platform: 'default', theme: 'ink', locale: 'ja', edge: 'end', rtl: true },
  { surface: 'inline', platform: 'mac', theme: 'ink', locale: 'en', edge: 'end', rtl: false },
  { surface: 'bar', platform: 'mac', theme: 'paper', locale: 'ja', edge: 'start', rtl: true },
  {
    surface: 'inline',
    platform: 'default',
    theme: 'paper',
    locale: 'en',
    edge: 'start',
    rtl: true,
  },
  {
    surface: 'inline',
    platform: 'default',
    theme: 'contrast',
    locale: 'ja',
    edge: 'end',
    rtl: false,
  },
  { surface: 'bar', platform: 'default', theme: 'paper', locale: 'ja', edge: 'end', rtl: false },
  { surface: 'bar', platform: 'mac', theme: 'contrast', locale: 'en', edge: 'end', rtl: true },
  { surface: 'inline', platform: 'default', theme: 'ink', locale: 'ja', edge: 'start', rtl: false },
  { surface: 'bar', platform: 'mac', theme: 'ink', locale: 'ja', edge: 'end', rtl: false },
];

export async function runFormulaHelperVisual(
  page: Page,
  scenario: FormulaHelperVisualCase,
  capture: string,
): Promise<void> {
  await page.setViewportSize({ width: 840, height: 620 });
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: scenario.locale, platform: scenario.platform });
  await sheet.expectNoStub();
  await page.evaluate(
    ({ theme, rtl, zoom }) => {
      const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
      instance.setTheme(theme);
      instance.store.setState((state) => ({
        ...state,
        ui: { ...state.ui, rightToLeft: rtl },
        viewport: { ...state.viewport, zoom },
      }));
    },
    { theme: scenario.theme, rtl: scenario.rtl, zoom: scenario.edge === 'end' ? 1.25 : 1 },
  );
  // Rendering counts include overscan; use actual geometry to target a fully
  // visible edge cell rather than assuming a fixed row/column fits at every zoom.
  const geometryUrl = `/@fs${new URL('../../../packages/formulon-cell/src/render/geometry.ts', import.meta.url).pathname}`;
  const addressUrl = `/@fs${new URL('../../../packages/formulon-cell/src/engine/address.ts', import.meta.url).pathname}`;
  const target =
    scenario.edge === 'end'
      ? await page.evaluate(
          async ({ geometryUrl, addressUrl }) => {
            const { cellRectUnclamped, layoutForView } = await import(geometryUrl);
            const { formatA1Cell } = await import(addressUrl);
            const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
            const state = instance.store.getState();
            const grid = document.querySelector('.fc-host__grid')?.getBoundingClientRect();
            if (!grid) throw new Error('Missing grid');
            let row = state.viewport.rowStart;
            let col = state.viewport.colStart;
            for (
              let r = state.viewport.rowStart;
              r < state.viewport.rowStart + state.viewport.rowCount;
              r += 1
            ) {
              for (
                let c = state.viewport.colStart;
                c < state.viewport.colStart + state.viewport.colCount;
                c += 1
              ) {
                const cell = cellRectUnclamped(layoutForView(state), state.viewport, r, c);
                if (
                  cell.x >= 0 &&
                  cell.y >= 0 &&
                  cell.x + cell.w <= grid.width &&
                  cell.y + cell.h <= grid.height
                ) {
                  row = r;
                  col = c;
                }
              }
            }
            return formatA1Cell(row, col);
          },
          { geometryUrl, addressUrl },
        )
      : 'A1';
  await sheet.goTo(target);
  const input = page
    .locator(scenario.surface === 'inline' ? '.fc-host__editor' : '.fc-host__formulabar-input')
    .first();
  if (scenario.surface === 'inline') await page.keyboard.press('F2');
  else await input.focus();
  await input.fill('=IF(TRUE,1,SU');
  const anchor = await input.boundingBox();
  expect(anchor, 'the placement fixture must anchor to a visible editor').not.toBeNull();
  if (!anchor) throw new Error('Missing editor anchor');
  expect(anchor.x + anchor.width).toBeGreaterThan(0);
  expect(anchor.y + anchor.height).toBeGreaterThan(0);
  expect(anchor.x).toBeLessThan(840);
  expect(anchor.y).toBeLessThan(620);
  const popup = page.locator('.fc-autocomplete');
  await expect(popup).toBeVisible();
  const candidateList = popup.locator('.fc-autocomplete__list');
  const listOrigin = await candidateList.boundingBox();
  await popup
    .locator('.fc-autocomplete__item')
    .filter({ has: page.getByText('SUBTOTAL', { exact: true }) })
    .hover();
  const withoutHelp = await candidateList.boundingBox();
  expect(withoutHelp?.x, 'candidate list must not jump when the ScreenTip is empty').toBe(
    listOrigin?.x,
  );
  expect(withoutHelp?.y).toBe(listOrigin?.y);
  await popup
    .locator('.fc-autocomplete__item')
    .filter({ has: page.getByText('SUBSTITUTE', { exact: true }) })
    .hover();
  const withHelp = await candidateList.boundingBox();
  expect(withHelp?.x).toBe(listOrigin?.x);
  expect(withHelp?.y).toBe(listOrigin?.y);
  const panels = page.locator('.fc-autocomplete, .fc-autocomplete__help, .fc-arghelper');
  await expect
    .poll(async () =>
      panels.evaluateAll((elements) =>
        elements.every((element) => {
          const rect = element.getBoundingClientRect();
          return (
            rect.left >= 0 &&
            rect.top >= 0 &&
            rect.right <= innerWidth &&
            rect.bottom <= innerHeight
          );
        }),
      ),
    )
    .toBe(true);
  const overlap = await page.evaluate(() => {
    const list = document.querySelector('.fc-autocomplete')?.getBoundingClientRect();
    const tip = document.querySelector('.fc-arghelper')?.getBoundingClientRect();
    return (
      !!list &&
      !!tip &&
      Math.min(list.right, tip.right) > Math.max(list.left, tip.left) &&
      Math.min(list.bottom, tip.bottom) > Math.max(list.top, tip.top)
    );
  });
  expect(overlap, 'candidate list and active-argument tip should remain separately readable').toBe(
    false,
  );
  await page.screenshot({ path: capture });
  await page.keyboard.press('Escape');
  await page.keyboard.press('Escape');
  await sheet.expectNoConsoleErrors();
}

/** Browser layout regression: the expected caret comes from full text width,
 * independently of the helper's DOM measurement implementation. */
export async function runFormulaHelperGeometry(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await page.setViewportSize({ width: 840, height: 620 });
  await sheet.mount({ locale: 'en' });
  const moduleUrl = `/@fs${new URL('../../../packages/formulon-cell/src/interact/formula-helper-position.ts', import.meta.url).pathname}`;
  const measured = await page.evaluate(async (moduleUrl) => {
    const { createFormulaCaretPositioner, positionFormulaHelper } = await import(moduleUrl);
    const textarea = document.createElement('textarea');
    textarea.style.cssText =
      'position:fixed;left:100px;top:100px;width:400px;height:80px;font:20px monospace;padding:4px;border:1px solid black;box-sizing:border-box';
    textarea.value = '1234567890';
    document.body.append(textarea);
    const positioner = createFormulaCaretPositioner(textarea);
    const canvas = document.createElement('canvas').getContext('2d');
    if (!canvas) throw new Error('Canvas is missing');
    const alignment = [];
    for (const align of ['center', 'right']) {
      textarea.style.textAlign = align;
      textarea.setSelectionRange(2, 2);
      const style = getComputedStyle(textarea);
      canvas.font = style.font;
      const full = canvas.measureText(textarea.value).width;
      const prefix = canvas.measureText(textarea.value.slice(0, 2)).width;
      const bounds = textarea.getBoundingClientRect();
      const available = bounds.width - 10;
      const expected =
        bounds.left + 5 + (available - full) * (align === 'center' ? 0.5 : 1) + prefix;
      alignment.push({ align, actual: positioner.caretRect().left, expected });
    }
    positioner.dispose();
    textarea.remove();
    const owner = document.createElement('div');
    owner.style.cssText =
      'position:fixed;left:20px;top:20px;width:400px;height:200px;transform:scale(2);transform-origin:0 0';
    const panel = document.createElement('div');
    panel.style.cssText = 'position:fixed;width:452px;height:200px;box-sizing:border-box';
    owner.append(panel);
    document.body.append(owner);
    positionFormulaHelper(panel, {
      left: 300,
      right: 300,
      top: 200,
      bottom: 220,
      width: 0,
      height: 20,
    });
    const placed = panel.getBoundingClientRect().toJSON();
    owner.remove();
    return { alignment, placed, width: innerWidth, height: innerHeight };
  }, moduleUrl);
  for (const item of measured.alignment) {
    // WebKit quantizes collapsed Range coordinates to CSS pixels while Canvas
    // reports fractional font advances. The old prefix-only mirror was 48px off.
    expect(item.actual, `${item.align} caret with a retained suffix`).toBeGreaterThan(
      item.expected - 1,
    );
    expect(item.actual, `${item.align} caret with a retained suffix`).toBeLessThan(
      item.expected + 1,
    );
  }
  expect(measured.placed.left).toBeGreaterThanOrEqual(0);
  expect(measured.placed.top).toBeGreaterThanOrEqual(0);
  expect(measured.placed.right).toBeLessThanOrEqual(measured.width);
  expect(measured.placed.bottom).toBeLessThanOrEqual(measured.height);
  await page.setViewportSize({ width: 480, height: 620 });
  await sheet.goTo('A1');
  await page.keyboard.press('F2');
  const input = page.locator('.fc-host__editor');
  await input.fill('=SU');
  const popup = page.locator('.fc-autocomplete');
  await expect(popup).toBeVisible();
  const origin = await popup.locator('.fc-autocomplete__list').boundingBox();
  await popup
    .locator('.fc-autocomplete__item')
    .filter({ has: page.getByText('SUBTOTAL', { exact: true }) })
    .hover();
  const after = await popup.locator('.fc-autocomplete__list').boundingBox();
  expect(after?.x).toBe(origin?.x);
  expect(after?.y, 'stacked ScreenTips must keep the list anchored').toBe(origin?.y);
  await page.keyboard.press('Escape');
  await page.keyboard.press('Escape');
  await sheet.expectNoConsoleErrors();
}
