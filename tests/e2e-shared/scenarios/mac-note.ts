import { expect, type Locator, type Page } from '@playwright/test';

import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type ProbeWindow = Window & { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };

/** The Mac new-note box in pt: size, gap right of the cell, lift above the cell top, and the
 *  rise from the cell bottom when the cell is in the last row. */
const NOTE = { w: 104, h: 76, font: 10, gap: 15, lift: 10, flipUp: 88 };

/** Page-space rectangle of the cell at `row`/`col` in the live layout. */
const cellBox = async (page: Page, row: number, col: number) => {
  const grid = await page.locator('.fc-host__grid').first().boundingBox();
  if (!grid) throw new Error('Missing grid');
  const r = await page.evaluate(
    ({ row, col }) => {
      const w = window as unknown as ProbeWindow;
      const state = w.__fcProbes['mount/internals']
        .spreadsheetInternals(w.__fcInst)
        .store.getState();
      const { cellRectUnclamped, layoutForView } = w.__fcProbes['render/geometry'];
      return cellRectUnclamped(layoutForView(state), state.viewport, row, col);
    },
    { row, col },
  );
  return { x: grid.x + r.x, y: grid.y + r.y, w: r.w, h: r.h };
};

const noteMenuRow = async (sp: UserJourneyPage, page: Page, ref: string): Promise<Locator> => {
  await sp.goTo(ref);
  await selectRibbonTab(page, 'review');
  await sp.clickRibbon('mac.review.notes');
  return page.locator('.fc-tb__menu [data-ribbon-command="mac.review.newNote"]').first();
};

const openNewNote = async (sp: UserJourneyPage, page: Page, ref: string): Promise<Locator> => {
  const row = await noteMenuRow(sp, page, ref);
  await expect(row).toHaveText('New Note');
  await row.click();
  const note = page.locator('.fc-cmtnote.fc-cmtnote--mac-note');
  await expect(note).toBeVisible();
  return note;
};

/** The first row of Review > Notes at `ref`: New Note on a plain cell, Edit Note on a note. */
const expectNoteRow = async (sp: UserJourneyPage, page: Page, ref: string, text: string) => {
  const row = await noteMenuRow(sp, page, ref);
  await expect(row).toHaveText(text);
  await page.keyboard.press('Escape');
  await expect(row).toBeHidden();
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
      page.evaluate(() => (window as unknown as ProbeWindow).__fcInst.viewport.zoomPercent),
    )
    .toBe(percent);
};

/** The box takes the Mac note size and font at `zoom`, with no header or footer. */
const expectShape = async (note: Locator, zoom: number) => {
  const px = (pt: number) => `${Math.round(pt * zoom * 100) / 100}px`;
  const style = await note.evaluate((el) => {
    const s = (el as HTMLElement).style;
    return { width: s.width, height: s.height, fontSize: s.fontSize };
  });
  expect(style).toEqual({ width: px(NOTE.w), height: px(NOTE.h), fontSize: px(NOTE.font) });
  await expect(note.locator('.fc-cmtnote__head')).toBeHidden();
  await expect(note.locator('.fc-cmtnote__footer')).toBeHidden();
  await expect(note.locator('textarea')).toBeFocused();
};

const near = (actual: number, target: number, what: string) =>
  expect(Math.abs(actual - target), `${what}: ${actual} vs ${target}`).toBeLessThanOrEqual(1);

/** A new note on the Mac profile is a bare 104x76 pt box in 10 pt type, 15 pt right of the cell
 *  and 10 pt above its top, scaled with the zoom. Clicking outside or Cmd+Enter keeps it,
 *  Escape drops it, and at the sheet's last cell it opens to the left and rises from the bottom. */
export async function runMacNewNoteScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac', locale: 'en', fixture: 'empty' });
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.expectNoStub();

  let note = await openNewNote(sp, page, 'C3');
  await expectShape(note, 1);
  let cell = await cellBox(page, 2, 2);
  let box = await note.boundingBox();
  if (!box) throw new Error('Missing note box');
  near(box.x, cell.x + cell.w + NOTE.gap, 'note left');
  near(box.y, cell.y - NOTE.lift, 'note top');
  await note.locator('textarea').fill('Cmd+Enter keeps it');
  await page.keyboard.press('Meta+Enter');
  await expect(note).toBeHidden();
  await expectNoteRow(sp, page, 'C3', 'Edit Note');

  note = await openNewNote(sp, page, 'E8');
  await note.locator('textarea').fill('Escape drops it');
  await page.keyboard.press('Escape');
  await expect(note).toBeHidden();
  await expectNoteRow(sp, page, 'E8', 'New Note');

  await setZoom(sp, page, 150);
  note = await openNewNote(sp, page, 'D4');
  await expectShape(note, 1.5);
  cell = await cellBox(page, 3, 3);
  box = await note.boundingBox();
  if (!box) throw new Error('Missing note box');
  near(box.x, cell.x + cell.w + NOTE.gap * 1.5, 'zoomed note left');
  near(box.y, cell.y - NOTE.lift * 1.5, 'zoomed note top');
  await page.keyboard.press('Escape');
  await expect(note).toBeHidden();

  await setZoom(sp, page, 100);
  note = await openNewNote(sp, page, 'XFD1048576');
  await expectShape(note, 1);
  cell = await cellBox(page, 1048575, 16383);
  box = await note.boundingBox();
  if (!box) throw new Error('Missing note box');
  near(box.x + box.width, cell.x - NOTE.gap, 'last-column note right');
  near(box.y, cell.y + cell.h - NOTE.flipUp, 'last-row note top');
  await page.keyboard.press('Escape');
  await expect(note).toBeHidden();

  // A click on another cell keeps the note on the cell it was opened for.
  note = await openNewNote(sp, page, 'B2');
  await note.locator('textarea').fill('Outside keeps it');
  cell = await cellBox(page, 11, 1);
  await page.mouse.click(cell.x + cell.w / 2, cell.y + cell.h / 2);
  await expect(note).toBeHidden();
  await expectNoteRow(sp, page, 'B2', 'Edit Note');
  await expectNoteRow(sp, page, 'B12', 'New Note');
  await sp.expectNoConsoleErrors();
}
