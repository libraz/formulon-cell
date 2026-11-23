import { expect, type Page } from '@playwright/test';
import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** Blank-workbook defaults per platform: canonical row px, workbook font and
 *  painted text px. Mac renders 1 pt = 1 screen px. */
const STARTUP = {
  default: {
    ja: { rowHeight: 25, font: { name: '游ゴシック', size: 11 }, textPx: 11 * (96 / 72) },
    en: { rowHeight: 20, font: { name: 'Aptos Narrow', size: 11 }, textPx: 11 * (96 / 72) },
  },
  mac: {
    ja: { rowHeight: 80 / 3, font: { name: '游ゴシック', size: 12 }, textPx: 12 },
    en: { rowHeight: 64 / 3, font: { name: 'Aptos Narrow', size: 12 }, textPx: 12 },
  },
} as const;

/** The demo's seeded column widths in px, which read back unchanged on every platform. */
const SEEDED_COLUMN_WIDTHS = [112, 80, 104, 128] as const;

/** A fresh book's default column and row px: the profile's measured points x 4/3. */
const FRESH_BOOK = {
  default: { ja: { colWidth: 72, rowHeight: 25 }, en: { colWidth: 64, rowHeight: 20 } },
  mac: { ja: { colWidth: 100, rowHeight: 80 / 3 }, en: { colWidth: 260 / 3, rowHeight: 64 / 3 } },
} as const;

type FontProbeWindow = Window & {
  __fcInst: SpreadsheetInstance;
  __fcProbes: E2eProbes;
  __startupText: Record<string, { font: string; width: number }>;
  __startupGlyphs: { ch: string; font: string; x: number; y: number; width: number }[];
  __startupFindRun: (text: string) => { font: string; width: number } | undefined;
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
    const words = ['item', 'London', 'Reykjavík', 'fahrenheit', 'Reykjavík ☼'];
    const chars = new Set(words.flatMap((word) => [...word]));
    probe.__startupText = {};
    probe.__startupGlyphs = [];
    // Mac cell text is drawn one glyph at a time; a word is a run of consecutive glyph draws.
    probe.__startupFindRun = (text) => {
      const whole = probe.__startupText[text];
      if (whole) return whole;
      const glyphs = probe.__startupGlyphs;
      const run = [...text];
      for (let i = 0; i + run.length <= glyphs.length; i++) {
        const slice = glyphs.slice(i, i + run.length);
        if (slice.every((glyph, k) => glyph.ch === run[k] && glyph.y === slice[0]?.y)) {
          const first = slice[0];
          const last = slice[slice.length - 1];
          if (first && last) {
            return { font: first.font, width: last.x + last.width - first.x };
          }
        }
      }
      return undefined;
    };
    const original = CanvasRenderingContext2D.prototype.fillText;
    CanvasRenderingContext2D.prototype.fillText = function (text, x, y, maxWidth) {
      if (words.includes(text)) {
        probe.__startupText[text] = { font: this.font, width: this.measureText(text).width };
      } else if ([...text].length === 1 && chars.has(text)) {
        probe.__startupGlyphs.push({
          ch: text,
          font: this.font,
          x,
          y,
          width: this.measureText(text).width,
        });
      }
      if (maxWidth === undefined) original.call(this, text, x, y);
      else original.call(this, text, x, y, maxWidth);
    };
  });
  const spreadsheet = new SpreadsheetPage(page);
  const errors = spreadsheet.collectConsoleErrors();
  await spreadsheet.mount({ fixture: null, locale, platform });
  await spreadsheet.expectNoStub();
  await page.waitForFunction(() =>
    (window as unknown as FontProbeWindow).__startupFindRun('Reykjavík'),
  );
  const expected = STARTUP[platform][locale];
  const fontName = expected.font.name;
  const fontSize = String(expected.font.size);
  const startup = await page.evaluate(() => {
    const probe = window as unknown as FontProbeWindow;
    const instance = probe.__fcInst;
    const state = probe.__fcProbes['mount/internals']
      .spreadsheetInternals(instance)
      .store.getState();
    return {
      font: probe.__fcProbes['api/workbook'].workbookHandleOf(instance.workbook)
        .workbookDisplayFont,
      widths: [0, 1, 2, 3].map((col) => state.layout.colWidths.get(col)),
      rowHeight: state.layout.defaultRowHeight,
      draws: Object.fromEntries(
        ['item', 'London', 'Reykjavík', 'fahrenheit', 'Reykjavík ☼'].map((word) => [
          word,
          probe.__startupFindRun(word),
        ]),
      ),
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
  for (const [col, width] of SEEDED_COLUMN_WIDTHS.entries()) {
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
    if (!draw) throw new Error(`missing actual canvas draw: ${text}`);
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
    .poll(() => page.evaluate(() => (window as unknown as FontProbeWindow).__fcInst.i18n.locale))
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
  // A UI language switch shows that language's theme face for the theme-linked
  // Normal font; its size stays as authored until a new book is created.
  expect(
    await page.evaluate(() => {
      const probe = window as unknown as FontProbeWindow;
      return probe.__fcProbes['api/workbook'].workbookHandleOf(probe.__fcInst.workbook)
        .workbookDisplayFont;
    }),
  ).toEqual({ name: STARTUP[platform][nextLocale].font.name, size: expected.font.size });
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
      locale: probe.__fcInst.i18n.locale,
      font: probe.__fcProbes['api/workbook'].workbookHandleOf(probe.__fcInst.workbook)
        .workbookDisplayFont,
      disposeCount: probe.__startupDisposeCount,
      value: probe.__fcInst.workbook.sheet(0).getValue(0, 0),
    };
  });
  expect(replacement).toEqual({
    locale: nextLocale,
    font: STARTUP[platform][nextLocale].font,
    disposeCount: 1,
    value: null,
  });
  await expect(page.locator('[data-ribbon-select="fontSize"] .fc-tb__rb-dd__value')).toHaveText(
    String(STARTUP[platform][nextLocale].font.size),
  );
  await spreadsheet.expectNoStub();
  expect(errors.read()).toEqual([]);
}

/** A fresh book takes the profile's default column width and row height. */
export async function runFreshBookDefaults(
  page: Page,
  locale: 'ja' | 'en',
  platform: 'default' | 'mac',
): Promise<void> {
  const spreadsheet = new SpreadsheetPage(page);
  const errors = spreadsheet.collectConsoleErrors();
  await spreadsheet.mount({ locale, platform });
  await spreadsheet.expectNoStub();
  const layout = await page.evaluate(() => {
    const probe = window as unknown as FontProbeWindow;
    const { layout } = probe.__fcProbes['mount/internals']
      .spreadsheetInternals(probe.__fcInst)
      .store.getState();
    return { colWidth: layout.defaultColWidth, rowHeight: layout.defaultRowHeight };
  });
  const expected = FRESH_BOOK[platform][locale];
  expect(layout.colWidth).toBeCloseTo(expected.colWidth, 4);
  expect(layout.rowHeight).toBeCloseTo(expected.rowHeight, 4);
  expect(errors.read()).toEqual([]);
}
