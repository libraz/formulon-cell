import { expect, type Page } from '@playwright/test';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/mount/types.js';
import { gridCellCenter, SpreadsheetPage } from '../pages/SpreadsheetPage.js';

type GeometryProbeWindow = Window & {
  __fcInst: SpreadsheetInstance;
  __geometryDraws: string[];
};

const PROBE_TEXT = 'PITCH';
/** Screen px slack: the default profile's 8.43-char column lands at 64.01 px. */
const PITCH_TOLERANCE = 0.05;

interface Expected {
  row: number;
  col?: number;
  font: { name: string; size: number };
  /** Painted cell text size in screen px. */
  textPx: number;
}

/** Blank-workbook screen geometry per platform at zoom 100%. Mac renders
 *  1 pt = 1 screen px; the default profile renders 96/72 px per point. */
const EXPECTED: Record<'default' | 'mac', Record<'ja' | 'en', Expected>> = {
  mac: {
    ja: { row: 20, col: 75, font: { name: '游ゴシック', size: 12 }, textPx: 12 },
    en: { row: 16, col: 65, font: { name: 'Aptos Narrow', size: 12 }, textPx: 12 },
  },
  default: {
    ja: { row: 25, font: { name: '游ゴシック', size: 11 }, textPx: 11 * (96 / 72) },
    en: { row: 20, col: 64, font: { name: 'Calibri', size: 11 }, textPx: 11 * (96 / 72) },
  },
};

const screenPitch = (page: Page) => {
  const moduleUrl = `/@fs${new URL('../../../packages/formulon-cell/src/render/geometry.ts', import.meta.url).pathname}`;
  return page.evaluate(async (nextModuleUrl) => {
    const instance = (window as unknown as GeometryProbeWindow).__fcInst;
    const { cellRectUnclamped } = await import(nextModuleUrl);
    const { layout, viewport, data } = instance.store.getState() as unknown as {
      layout: Parameters<typeof cellRectUnclamped>[0];
      viewport: Parameters<typeof cellRectUnclamped>[1] & { zoom: number };
      data: { sheetIndex: number };
    };
    const a1 = cellRectUnclamped(layout, viewport, 0, 0);
    const b2 = cellRectUnclamped(layout, viewport, 1, 1);
    return {
      sheet: data.sheetIndex,
      zoom: viewport.zoom,
      row: b2.y - a1.y,
      col: b2.x - a1.x,
      height: a1.h,
      width: a1.w,
    };
  }, moduleUrl);
};

export async function runPlatformGeometry(
  page: Page,
  locale: 'ja' | 'en',
  platform: 'default' | 'mac',
): Promise<void> {
  await page.addInitScript((probeText) => {
    const probe = window as unknown as GeometryProbeWindow;
    probe.__geometryDraws = [];
    const original = CanvasRenderingContext2D.prototype.fillText;
    CanvasRenderingContext2D.prototype.fillText = function (text, x, y, maxWidth) {
      if (text === probeText) probe.__geometryDraws.push(this.font);
      if (maxWidth === undefined) original.call(this, text, x, y);
      else original.call(this, text, x, y, maxWidth);
    };
  }, PROBE_TEXT);
  const expected = EXPECTED[platform][locale];
  const spreadsheet = new SpreadsheetPage(page);
  const errors = spreadsheet.collectConsoleErrors();
  await spreadsheet.mount({ fixture: 'empty', locale, platform });
  await spreadsheet.expectNoStub();

  const expectPitch = async (sheet: number) => {
    const pitch = await screenPitch(page);
    expect(pitch.sheet).toBe(sheet);
    expect(pitch.zoom).toBe(1);
    const label = `${locale}/${platform} sheet ${sheet + 1}`;
    const near = (value: number, target: number, what: string) =>
      expect(Math.abs(value - target), `${label} ${what}: ${value}`).toBeLessThanOrEqual(
        PITCH_TOLERANCE,
      );
    near(pitch.row, expected.row, 'row pitch');
    near(pitch.height, expected.row, 'row height');
    if (expected.col !== undefined) {
      near(pitch.col, expected.col, 'column pitch');
      near(pitch.width, expected.col, 'column width');
    }
  };
  await expectPitch(0);
  expect(
    await page.evaluate(
      () => (window as unknown as GeometryProbeWindow).__fcInst.workbook.workbookDisplayFont,
    ),
  ).toEqual(expected.font);

  const a1 = await gridCellCenter(page, 0, 0);
  await page.mouse.click(a1.x, a1.y);
  await page.keyboard.type(PROBE_TEXT);
  await page.keyboard.press('Enter');
  await expect
    .poll(() =>
      page.evaluate(() =>
        (window as unknown as GeometryProbeWindow).__fcInst.workbook.getValue({
          sheet: 0,
          row: 0,
          col: 0,
        }),
      ),
    )
    .toEqual({ kind: 'text', value: PROBE_TEXT });
  await page.waitForFunction(
    () => (window as unknown as GeometryProbeWindow).__geometryDraws.length > 0,
  );
  const drawFont = await page.evaluate(() => {
    const draws = (window as unknown as GeometryProbeWindow).__geometryDraws;
    return draws[draws.length - 1] ?? '';
  });
  expect(drawFont).toContain(expected.font.name);
  expect(Number.parseFloat(drawFont)).toBeCloseTo(expected.textPx, 3);

  const b2 = await gridCellCenter(page, 1, 1);
  await page.mouse.click(b2.x, b2.y);
  await page.keyboard.press('F2');
  const editor = page.locator('.fc-host__editor');
  await expect(editor).toBeVisible();
  const editorSize = await editor.evaluate((el) =>
    Number.parseFloat(getComputedStyle(el).fontSize),
  );
  expect(editorSize).toBeCloseTo(expected.textPx, 3);
  await page.keyboard.press('Escape');

  const tabs = page.locator('.fc-host__sheetbar-tab');
  await page.locator('.fc-host__sheetbar-add').click();
  await expect(tabs).toHaveCount(2);
  await expect(tabs.nth(1)).toHaveAttribute('aria-selected', 'true');
  await expectPitch(1);
  expect(errors.read()).toEqual([]);
}
