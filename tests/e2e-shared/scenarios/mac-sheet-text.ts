import { expect, type Page } from '@playwright/test';

import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type Draw = { text: string; px: number; family: string; x: number; y: number };
type DrawWindow = Window & {
  __fcInst: SpreadsheetInstance;
  __fcProbes: E2eProbes;
  __sheetTextDraws: Draw[];
};

/** Records every single-glyph canvas draw: its CSS px size, first family and position. Mac cell
 *  text is drawn one glyph at a time. */
const recordDraws = (page: Page) =>
  page.addInitScript(() => {
    const w = window as unknown as DrawWindow;
    w.__sheetTextDraws = [];
    const original = CanvasRenderingContext2D.prototype.fillText;
    CanvasRenderingContext2D.prototype.fillText = function (text, x, y, maxWidth) {
      const m = this.font.match(/(\d+(?:\.\d+)?)px\s+"?([^",]+)"?/);
      if (m && text.length === 1) {
        w.__sheetTextDraws.push({ text, px: Number(m[1]), family: m[2] ?? '', x, y });
      }
      if (maxWidth === undefined) original.call(this, text, x, y);
      else original.call(this, text, x, y, maxWidth);
    };
  });

/** Grid-space rectangle of a cell in the live layout. */
const cellRect = (page: Page, row: number, col: number) =>
  page.evaluate(
    ({ row, col }) => {
      const w = window as unknown as DrawWindow;
      const state = w.__fcProbes['mount/internals']
        .spreadsheetInternals(w.__fcInst)
        .store.getState();
      const { cellRectUnclamped, layoutForView } = w.__fcProbes['render/geometry'];
      return cellRectUnclamped(layoutForView(state), state.viewport, row, col);
    },
    { row, col },
  );

const draws = (page: Page) =>
  page.evaluate(() => (window as unknown as DrawWindow).__sheetTextDraws);

/** The latest draw of `text` inside the cell at `row`/`col`. */
const drawIn = async (page: Page, text: string, row: number, col: number) => {
  const r = await cellRect(page, row, col);
  const hit = (await draws(page)).filter(
    (d) => d.text === text && d.x >= r.x && d.x < r.x + r.w && d.y > r.y && d.y <= r.y + r.h,
  );
  return { rect: r, draw: hit.at(-1) };
};

/** The latest draw of a column-header letter, above the first row. */
const headerDraw = async (page: Page, letter: string) => {
  const top = await cellRect(page, 0, 0);
  return (await draws(page)).filter((d) => d.text === letter && d.y < top.y).at(-1);
};

const setZoom = async (sp: UserJourneyPage, page: Page, percent: number) => {
  await selectRibbonTab(page, 'view');
  await sp.clickRibbon('mac.view.zoom');
  const dialog = page.getByRole('dialog');
  await expect(dialog).toBeVisible();
  await dialog.locator('input[type="number"]').fill(String(percent));
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect
    .poll(() =>
      page.evaluate(() => (window as unknown as DrawWindow).__fcInst.viewport.zoomPercent),
    )
    .toBe(percent);
};

/** On the Mac profile the column-header letters take the cell text's painted size at every
 *  zoom (the 12 pt Normal font: 12 px at 100%, 18 px at 150%). */
export async function runMacHeaderLetterSizeScenario(page: Page): Promise<void> {
  await recordDraws(page);
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac', locale: 'en', fixture: 'empty' });
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.expectNoStub();
  await sp.enter('H5', 'Q');
  for (const [percent, px] of [
    [100, 12],
    [150, 18],
  ] as const) {
    if (percent !== 100) await setZoom(sp, page, percent);
    await expect
      .poll(async () => {
        const cell = (await drawIn(page, 'Q', 4, 7)).draw;
        const header = await headerDraw(page, 'H');
        return { cell: cell?.px, header: header?.px, family: header?.family };
      })
      .toEqual({ cell: px, header: px, family: 'Aptos Narrow' });
  }
  await sp.expectNoConsoleErrors();
}

/** Bottom-aligned Calibri and Aptos Narrow text sits with its baseline 4 pt above the cell's
 *  bottom edge at every font size, scaled with the zoom. */
export async function runMacBottomGapScenario(page: Page): Promise<void> {
  await recordDraws(page);
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac', locale: 'en', fixture: 'empty' });
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.expectNoStub();
  // [glyph, row, family, pt]: one cell per font and size, bottom-aligned by default.
  const cells = [
    ['Q', 1, 'Aptos Narrow', 12],
    ['W', 3, 'Aptos Narrow', 20],
    ['K', 5, 'Calibri', 11],
    ['Z', 7, 'Calibri', 24],
  ] as const;
  await page.evaluate((list) => {
    const w = window as unknown as DrawWindow;
    const { mutators } = w.__fcProbes['store/store'];
    const { store } = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst);
    for (const [glyph, row, family, pt] of list) {
      w.__fcInst.workbook.sheet(0).setText(row, 1, glyph);
      mutators.setCellFormat(
        store,
        { sheet: 0, row, col: 1 },
        { fontFamily: family, fontSize: (pt * 4) / 3 },
      );
    }
    w.__fcInst.history.clear();
  }, cells);
  for (const percent of [100, 150]) {
    if (percent !== 100) await setZoom(sp, page, percent);
    const gap = (4 * percent) / 100;
    for (const [glyph, row, family, pt] of cells) {
      await expect
        .poll(
          async () => {
            const { rect, draw } = await drawIn(page, glyph, row, 1);
            return draw ? Math.round((rect.y + rect.h - draw.y) * 2) / 2 : null;
          },
          { message: `${family} ${pt} pt at ${percent}%` },
        )
        .toBe(gap);
    }
  }
  await sp.expectNoConsoleErrors();
}
