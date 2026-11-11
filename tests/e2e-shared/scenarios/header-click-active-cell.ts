import { expect, type Page } from '@playwright/test';

import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { gridCellCenter } from '../pages/SpreadsheetPage.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type ProbeWindow = Window & { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };

const LAST_COL = 16383;
const LAST_ROW = 1048575;

/** The active cell as A1, the selected range, the name box text and the view's first row/column. */
const readState = (page: Page) =>
  page.evaluate(() => {
    const w = window as unknown as ProbeWindow;
    const { active, range } = w.__fcInst.selection.get();
    const { viewport } = w.__fcProbes['mount/internals']
      .spreadsheetInternals(w.__fcInst)
      .store.getState();
    const nameBox = document.querySelector<HTMLInputElement>('.fc-host__formulabar-tag');
    let col = '';
    for (let n = active.col + 1; n > 0; n = Math.floor((n - 1) / 26))
      col = String.fromCharCode(65 + ((n - 1) % 26)) + col;
    return {
      active: `${col}${active.row + 1}`,
      range,
      nameBox: nameBox?.value ?? null,
      view: { rowStart: viewport.rowStart, colStart: viewport.colStart },
    };
  });

/** Scrolls the body so `rowStart`/`colStart` (0-based) lead the scrolling pane. */
const scrollBody = (page: Page, rowStart: number, colStart: number) =>
  page.evaluate(
    ({ rowStart, colStart }) => {
      const w = window as unknown as ProbeWindow;
      const { store } = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst);
      store.setState((s) => ({ ...s, viewport: { ...s.viewport, rowStart, colStart } }));
    },
    { rowStart, colStart },
  );

/** Freezes the top `rows` rows and left `cols` columns, as the View freeze panes menu does. */
const setFreeze = (page: Page, rows: number, cols: number) =>
  page.evaluate(
    ({ rows, cols }) => {
      const w = window as unknown as ProbeWindow;
      const { store } = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst);
      store.setState((s) => ({
        ...s,
        layout: { ...s.layout, freezeRows: rows, freezeCols: cols },
      }));
    },
    { rows, cols },
  );

/** Page point of a header: `index` is a 0-based row or column, in a column of the view that shows it. */
const headerPoint = async (page: Page, axis: 'row' | 'col', index: number, along: number) => {
  const box = await page.locator('.fc-host__grid').first().boundingBox();
  if (!box) throw new Error('Missing grid');
  const cell = await gridCellCenter(
    page,
    axis === 'row' ? index : along,
    axis === 'col' ? index : along,
  );
  return axis === 'row' ? { x: box.x + 10, y: cell.y } : { x: cell.x, y: box.y + 10 };
};

const clickHeader = async (
  page: Page,
  axis: 'row' | 'col',
  index: number,
  along: number,
  modifier?: 'Shift' | 'Meta',
) => {
  const point = await headerPoint(page, axis, index, along);
  if (modifier) await page.keyboard.down(modifier);
  try {
    await page.mouse.click(point.x, point.y);
  } finally {
    if (modifier) await page.keyboard.up(modifier);
  }
};

const wholeRow = (row: number) => ({ firstRow: row, lastRow: row, firstCol: 0, lastCol: LAST_COL });
const wholeCol = (col: number) => ({ firstRow: 0, lastRow: LAST_ROW, firstCol: col, lastCol: col });

/** Expects `ref` as the active cell and in the name box, and the view left where it was. */
const expectActive = async (
  page: Page,
  ref: string,
  range: ReturnType<typeof wholeRow>,
  view: { rowStart: number; colStart: number },
) => {
  await expect.poll(async () => (await readState(page)).active).toBe(ref);
  expect(await readState(page)).toEqual({ active: ref, range, nameBox: ref, view });
};

/** A header click selects the whole row or column and puts the active cell on its first visible
 *  cell: the row's leftmost visible column, the column's top visible row, a frozen band counting as
 *  visible. The view does not scroll and the name box shows the active cell. Select-all, Shift-click,
 *  a drag and Shift+Space keep the active cell where it was; a Cmd-added row brings its own. */
export async function runHeaderClickActiveCellScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'ja', platform: 'mac' });
  await sheet.expectNoStub();
  const top = { rowStart: 0, colStart: 0 };

  await clickHeader(page, 'row', 2, 0);
  await expectActive(page, 'A3', wholeRow(2), top);

  await scrollBody(page, 0, 4);
  await clickHeader(page, 'row', 2, 4);
  await expectActive(page, 'E3', wholeRow(2), { rowStart: 0, colStart: 4 });

  await scrollBody(page, 9, 0);
  await clickHeader(page, 'col', 2, 9);
  await expectActive(page, 'C10', wholeCol(2), { rowStart: 9, colStart: 0 });

  // Freeze at C3: rows 1-2 and columns A-B stay put while the body scrolls to column E, then to
  // row 10. The frozen column A and frozen row 1 are the first visible cells.
  await scrollBody(page, 0, 0);
  await setFreeze(page, 2, 2);
  await scrollBody(page, 2, 4);
  await clickHeader(page, 'row', 2, 4);
  await expectActive(page, 'A3', wholeRow(2), { rowStart: 2, colStart: 4 });
  await scrollBody(page, 9, 4);
  const frozenView = { rowStart: 9, colStart: 4 };
  await clickHeader(page, 'row', 9, 4);
  await expectActive(page, 'A10', wholeRow(9), frozenView);
  await clickHeader(page, 'col', 5, 9);
  await expectActive(page, 'F1', wholeCol(5), frozenView);
  await setFreeze(page, 0, 0);
  await scrollBody(page, 0, 0);

  // Select-all keeps the active cell at A1.
  await sheet.goTo('D5');
  const box = await page.locator('.fc-host__grid').first().boundingBox();
  if (!box) throw new Error('Missing grid');
  await page.mouse.click(box.x + 5, box.y + 5);
  await expect
    .poll(async () => (await readState(page)).range)
    .toEqual({ firstRow: 0, lastRow: LAST_ROW, firstCol: 0, lastCol: LAST_COL });
  expect((await readState(page)).active).toBe('A1');
  expect((await readState(page)).nameBox).toBe('A1');

  // Shift-click and a drag keep the anchor row's active cell.
  await clickHeader(page, 'row', 1, 0);
  await clickHeader(page, 'row', 4, 0, 'Shift');
  await expectActive(page, 'A2', { firstRow: 1, lastRow: 4, firstCol: 0, lastCol: LAST_COL }, top);
  const from = await headerPoint(page, 'row', 1, 0);
  const to = await headerPoint(page, 'row', 4, 0);
  await clickHeader(page, 'row', 7, 0);
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(to.x, to.y, { steps: 5 });
  await page.mouse.up();
  await expectActive(page, 'A2', { firstRow: 1, lastRow: 4, firstCol: 0, lastCol: LAST_COL }, top);

  // A Cmd-added row brings its own active cell.
  await clickHeader(page, 'row', 1, 0);
  await clickHeader(page, 'row', 3, 0, 'Meta');
  await expect.poll(async () => (await readState(page)).active).toBe('A4');
  expect((await readState(page)).nameBox).toBe('A4');

  // Shift+Space selects the row and keeps the active cell.
  await sheet.goTo('D5');
  await page.keyboard.press('Shift+Space');
  await expectActive(page, 'D5', wholeRow(4), top);
  await sheet.expectNoConsoleErrors();
}
