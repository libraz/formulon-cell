import { expect, type Page } from '@playwright/test';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/mount/types.js';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** Blank-workbook defaults per platform: canonical row px, workbook font and
 *  painted text px. Mac renders 1 pt = 1 screen px. */
const STARTUP = {
  default: {
    ja: { rowHeight: 25, font: { name: '游ゴシック', size: 11 }, textPx: 11 * (96 / 72) },
    en: { rowHeight: 20, font: { name: 'Calibri', size: 11 }, textPx: 11 * (96 / 72) },
  },
  mac: {
    ja: { rowHeight: 80 / 3, font: { name: '游ゴシック', size: 12 }, textPx: 12 },
    en: { rowHeight: 64 / 3, font: { name: 'Aptos Narrow', size: 12 }, textPx: 12 },
  },
} as const;

type FontProbeWindow = Window & {
  __fcInst: SpreadsheetInstance;
  __startupText: Record<string, { font: string; width: number }>;
  __startupPreviousWorkbook?: SpreadsheetInstance['workbook'];
  __startupDisposeCount?: number;
};

export async function runStartupFont(
  page: Page,
  locale: 'ja' | 'en',
  platform: 'default' | 'mac',
): Promise<void> {
  await page.addInitScript(() => {
    const probe = window as unknown as FontProbeWindow;
    probe.__startupText = {};
    const original = CanvasRenderingContext2D.prototype.fillText;
    CanvasRenderingContext2D.prototype.fillText = function (text, x, y, maxWidth) {
      if (['item', 'London', 'Reykjavík', 'fahrenheit', 'Reykjavík ☼'].includes(text)) {
        probe.__startupText[text] = { font: this.font, width: this.measureText(text).width };
      }
      if (maxWidth === undefined) original.call(this, text, x, y);
      else original.call(this, text, x, y, maxWidth);
    };
  });
  const spreadsheet = new SpreadsheetPage(page);
  const errors = spreadsheet.collectConsoleErrors();
  await spreadsheet.mount({ fixture: null, locale, platform });
  await spreadsheet.expectNoStub();
  await page.waitForFunction(
    () => (window as unknown as FontProbeWindow).__startupText['Reykjavík'],
  );
  const expected = STARTUP[platform][locale];
  const fontName = expected.font.name;
  const fontSize = String(expected.font.size);
  const startup = await page.evaluate(() => {
    const probe = window as unknown as FontProbeWindow;
    const instance = probe.__fcInst;
    const state = instance.store.getState();
    return {
      font: instance.workbook.workbookDisplayFont,
      widths: [0, 1, 2, 3].map((col) => state.layout.colWidths.get(col)),
      rowHeight: state.layout.defaultRowHeight,
      draws: probe.__startupText,
      corner: { x: state.layout.headerColWidth + 12, y: state.layout.headerRowHeight + 8 },
    };
  });
  expect(startup.font).toEqual(expected.font);
  await expect(page.locator('[data-ribbon-select="fontSize"] .fc-tb__rb-dd__value')).toHaveText(
    fontSize,
  );
  await expect(page.locator('[data-ribbon-command="fontFamily"] button').first()).toContainText(
    fontName,
  );
  for (const [col, width] of [112, 80, 104, 128].entries()) {
    expect(startup.widths[col]).toBeCloseTo(width, 1);
  }
  expect(startup.rowHeight).toBeCloseTo(expected.rowHeight, 4);
  for (const [text, col] of [
    ['item', 0],
    ['London', 0],
    ['Reykjavík', 0],
    ['fahrenheit', 2],
    ['Reykjavík ☼', 3],
  ] as const) {
    const draw = startup.draws[text];
    expect(draw, `missing actual canvas draw: ${text}`).toBeDefined();
    expect(draw.font).toContain(fontName);
    expect(draw.font).toContain('sans-serif');
    expect(draw.width).toBeLessThan((startup.widths[col] ?? 0) - 8);
    expect(Number.parseFloat(draw.font)).toBeCloseTo(expected.textPx, 3);
  }
  await page.locator('.fc-host__grid').first().click({ position: startup.corner });
  await page.keyboard.press('F2');
  const editor = page.locator('.fc-host__editor');
  await expect(editor).toBeVisible();
  const editing = await editor.evaluate((el) => {
    const css = getComputedStyle(el);
    return { family: css.fontFamily, size: Number.parseFloat(css.fontSize) };
  });
  expect(editing.family).toContain(fontName);
  expect(editing.family).toContain('sans-serif');
  expect(editing.size).toBeCloseTo(expected.textPx, 3);
  await page.keyboard.press('Escape');

  // Drive the real demo callback through File > New after a UI locale change.
  // Merely passing a locale to a test-owned createDefault callback would not
  // detect the demo dropping the selected locale or leaking its initial owner.
  const nextLocale = locale === 'ja' ? 'en' : 'ja';
  const options = page.locator('.demo__account .demo__share[aria-pressed]').first();
  await options.click();
  const panel = page.locator('.demo__panel').first();
  await expect(panel).toBeVisible();
  await panel.getByRole('combobox').click();
  await panel
    .getByRole('option', { name: nextLocale === 'ja' ? '日本語' : 'English', exact: true })
    .click();
  await expect
    .poll(() => page.evaluate(() => (window as unknown as FontProbeWindow).__fcInst.locale))
    .toBe(nextLocale);
  await options.click();
  await expect(panel).toBeHidden();
  await page.evaluate(() => {
    const probe = window as unknown as FontProbeWindow;
    const previous = probe.__fcInst.workbook;
    probe.__startupPreviousWorkbook = previous;
    probe.__startupDisposeCount = 0;
    const dispose = previous.dispose.bind(previous);
    previous.dispose = () => {
      probe.__startupDisposeCount = (probe.__startupDisposeCount ?? 0) + 1;
      dispose();
    };
  });
  // Changing the UI language keeps authored fonts until a new book is created.
  expect(
    await page.evaluate(
      () => (window as unknown as FontProbeWindow).__fcInst.workbook.workbookDisplayFont,
    ),
  ).toEqual(expected.font);
  await page
    .locator(platform === 'mac' ? '.demo__quick .demo__brand-mark' : '[data-ribbon-tab="file"]')
    .click();
  const backstage = page.locator('.fc-tb__backstage').first();
  await expect(backstage).toBeVisible();
  await backstage
    .locator('.fc-tb__backstage-nav')
    .getByRole('button', { name: nextLocale === 'ja' ? '新規' : 'New', exact: true })
    .click();
  await expect(backstage).toBeHidden();
  await expect
    .poll(() =>
      page.evaluate(() => {
        const probe = window as unknown as FontProbeWindow;
        return probe.__fcInst.workbook !== probe.__startupPreviousWorkbook;
      }),
    )
    .toBe(true);
  const replacement = await page.evaluate(() => {
    const probe = window as unknown as FontProbeWindow;
    return {
      locale: probe.__fcInst.locale,
      font: probe.__fcInst.workbook.workbookDisplayFont,
      disposeCount: probe.__startupDisposeCount,
      value: probe.__fcInst.workbook.getValue({ sheet: 0, row: 0, col: 0 }),
    };
  });
  expect(replacement).toEqual({
    locale: nextLocale,
    font: STARTUP[platform][nextLocale].font,
    disposeCount: 1,
    value: { kind: 'blank' },
  });
  await expect(page.locator('[data-ribbon-select="fontSize"] .fc-tb__rb-dd__value')).toHaveText(
    String(STARTUP[platform][nextLocale].font.size),
  );
  await spreadsheet.expectNoStub();
  expect(errors.read()).toEqual([]);
}
