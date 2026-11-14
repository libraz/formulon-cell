import { expect, type Page } from '@playwright/test';

import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type Platform = 'default' | 'mac';
type StrokeWindow = Window & {
  __fcInst: SpreadsheetInstance;
  __fcProbes: E2eProbes;
  __staleStrokes: { x: number; y: number }[];
};

const MANUAL = 1 as const;

/** Records each path's start point, so a strikeout line inside a cell can be found. */
const recordStrokes = (page: Page) =>
  page.addInitScript(() => {
    const w = window as unknown as StrokeWindow;
    w.__staleStrokes = [];
    const original = CanvasRenderingContext2D.prototype.moveTo;
    CanvasRenderingContext2D.prototype.moveTo = function (x, y) {
      w.__staleStrokes.push({ x, y });
      original.call(this, x, y);
    };
  });

/** What marks B2 (=A2*2) as stale on screen: the engine's stale flag, the error-check triangle
 *  and a strikeout line drawn inside the cell, away from its gridlines. */
const staleMarksAtB2 = (page: Page) =>
  page.evaluate(() => {
    const w = window as unknown as StrokeWindow;
    const handle = w.__fcProbes['api/workbook'].workbookHandleOf(w.__fcInst.workbook);
    const grid = document.querySelector<HTMLElement>('.fc-host__grid');
    const state = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst).store.getState();
    const { cellRectUnclamped, layoutForView } = w.__fcProbes['render/geometry'];
    const r = cellRectUnclamped(layoutForView(state), state.viewport, 1, 1);
    const struck = w.__staleStrokes.some(
      (p) => p.x >= r.x && p.x < r.x + r.w && p.y > r.y + 2 && p.y < r.y + r.h - 2,
    );
    const triangle = grid
      ? w.__fcProbes['render/grid/hit-state']
          .gridHitState(grid)
          .errorTriangles.some(({ addr }) => addr.sheet === 0 && addr.row === 1 && addr.col === 1)
      : false;
    return {
      mode: handle.calcMode(),
      stale: handle.isCellStale({ sheet: 0, row: 1, col: 1 }),
      triangle,
      struck,
    };
  });

/** Clears the recorded strokes so the next paint is read on its own. */
const resetStrokes = (page: Page) =>
  page.evaluate(() => {
    (window as unknown as StrokeWindow).__staleStrokes = [];
  });

/** Opens `path` through File > Open: the File tab on the default chrome, the brand mark on the
 *  Mac chrome, which has no File tab. */
const openSaved = async (page: Page, path: string): Promise<void> => {
  const fileTab = page.locator('[data-ribbon-tab="file"]');
  if ((await fileTab.count()) > 0) await fileTab.first().click();
  else await page.locator('.demo__brand-mark').first().click();
  const backstage = page.locator('.fc-tb__backstage[role="dialog"]').first();
  await expect(backstage).toBeVisible();
  const [chooser] = await Promise.all([
    page.waitForEvent('filechooser'),
    backstage.getByRole('button', { name: 'Open', exact: true }).first().click(),
  ]);
  await chooser.setFiles(path);
  if (await backstage.isVisible())
    await backstage.getByRole('button', { name: 'Close', exact: true }).first().click();
  await expect(backstage).toBeHidden();
};

/** Reads `<calcPr calcOnSave>` through the workbook. */
const calcOnSave = (page: Page) =>
  page.evaluate(() => {
    const w = window as unknown as StrokeWindow;
    return w.__fcProbes['api/workbook'].workbookHandleOf(w.__fcInst.workbook).calcSettings()
      ?.calcOnSave;
  });

/** Turns off "Recalculate workbook before saving": in Preferences > Calculation on the Mac, in
 *  Formulas > Calculation Options > Iterative Calculation elsewhere. */
const turnOffRecalcBeforeSave = async (
  page: Page,
  sp: UserJourneyPage,
  platform: Platform,
): Promise<void> => {
  if (platform === 'mac') {
    await page.locator('.demo__brand-mark').first().click();
    const backstage = page.locator('.fc-tb__backstage[role="dialog"]').first();
    await expect(backstage).toBeVisible();
    await backstage.getByRole('button', { name: 'Options', exact: true }).first().click();
    const prefs = page.locator('.fc-prefs[role="dialog"]:not([hidden])');
    await expect(prefs).toBeVisible();
    await prefs.getByRole('button', { name: 'Show All', exact: true }).click();
    await prefs.locator('.fc-prefs__pane-button[data-pane="calculation"]').click();
    const box = prefs.locator('[data-calc-option="calcOnSave"]');
    await expect(box).toBeChecked();
    await expect(box).toBeEnabled();
    await box.uncheck();
    await prefs.locator('.fc-fmtdlg__close').first().click();
    await expect(prefs).toBeHidden();
    // Preferences opens over the File backstage, which stays up after it closes.
    if (await backstage.isVisible()) await page.keyboard.press('Escape');
    await expect(backstage).toBeHidden();
  } else {
    await selectRibbonTab(page, 'formulas');
    await sp.clickRibbon('calcOptions');
    const menu = page.locator('#menu-calc-options');
    await expect(menu).toBeVisible();
    await menu.locator('[data-calc-option="iterative"]').click();
    const dialog = page
      .getByRole('dialog')
      .filter({ has: page.locator('[data-calc-option="calcOnSave"]') });
    await expect(dialog).toBeVisible();
    const box = dialog.locator('[data-calc-option="calcOnSave"]');
    await expect(box).toBeChecked();
    await expect(box).toBeEnabled();
    await box.uncheck();
    await dialog.getByRole('button', { name: 'OK', exact: true }).click();
    await expect(dialog).toBeHidden();
  }
  await expect.poll(() => calcOnSave(page)).toBe(false);
};

/** Seeds B2 = A2*2 over A2 = 1 under manual calculation. */
const seedManualB2 = (page: Page) =>
  page.evaluate((manual) => {
    const w = window as unknown as StrokeWindow;
    const sheet = w.__fcInst.workbook.sheet(0);
    w.__fcInst.workbook.batch(() => {
      sheet.setNumber(1, 0, 1);
      sheet.setFormula(1, 1, '=A2*2');
    });
    w.__fcProbes['api/workbook'].workbookHandleOf(w.__fcInst.workbook).setCalcMode(manual);
    w.__fcInst.history.clear();
  }, MANUAL);

/** Edits A2 to 5, which leaves B2 stale with its old value. */
const makeB2Stale = async (page: Page, sp: UserJourneyPage): Promise<void> => {
  await sp.enter('A2', '5');
  await expect
    .poll(() => staleMarksAtB2(page))
    .toEqual({ mode: MANUAL, stale: true, triangle: true, struck: true });
  await expect.poll(() => sp.readValue('B2')).toEqual({ kind: 'number', value: 2 });
};

/** Opens `path` on a fresh page of `platform` and hands it to `check`. */
const reopen = async (
  page: Page,
  platform: Platform,
  path: string,
  check: (fresh: Page, reopened: UserJourneyPage) => Promise<void>,
): Promise<void> => {
  const fresh = await page.context().newPage();
  await recordStrokes(fresh);
  const reopened = new UserJourneyPage(fresh);
  try {
    await reopened.mount({ platform, fixture: 'empty' });
    await fresh.setViewportSize({ width: 1440, height: 900 });
    await reopened.expectNoStub();
    await openSaved(fresh, path);
    await expect.poll(() => reopened.readValue('A2')).toEqual({ kind: 'number', value: 5 });
    await resetStrokes(fresh);
    await reopened.goTo('D4');
    await check(fresh, reopened);
    await reopened.expectNoConsoleErrors();
  } finally {
    await fresh.close();
  }
};

/** Under manual calculation an edit leaves its dependants stale. With "Recalculate workbook
 *  before saving" on (the default) a save computes them first, so the saved and the live book
 *  are current; with it off the save writes the stale values, and the reopened book is still
 *  manual and draws the dependant stale with its old value. */
export async function runStaleValueSaveScenario(
  page: Page,
  platform: Platform,
  recalcBeforeSave: boolean,
): Promise<void> {
  await recordStrokes(page);
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform, fixture: 'empty' });
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.expectNoStub();
  await expect.poll(() => calcOnSave(page)).toBe(true);
  await seedManualB2(page);
  // The option is only offered under manual calculation.
  if (!recalcBeforeSave) await turnOffRecalcBeforeSave(page, sp, platform);
  await makeB2Stale(page, sp);
  const path = await sp.saveDownload();
  if (recalcBeforeSave) {
    const current = { mode: MANUAL, stale: false, triangle: false };
    const marks = async (p: Page) => {
      const m = await staleMarksAtB2(p);
      return { mode: m.mode, stale: m.stale, triangle: m.triangle };
    };
    await expect.poll(() => marks(page)).toEqual(current);
    await expect.poll(() => sp.readValue('B2')).toEqual({ kind: 'number', value: 10 });
    await reopen(page, platform, path, async (fresh, reopened) => {
      await expect.poll(() => marks(fresh)).toEqual(current);
      await expect.poll(() => reopened.readValue('B2')).toEqual({ kind: 'number', value: 10 });
    });
  } else {
    const stale = { mode: MANUAL, stale: true, triangle: true, struck: true };
    await expect.poll(() => staleMarksAtB2(page)).toEqual(stale);
    await reopen(page, platform, path, async (fresh, reopened) => {
      await expect.poll(() => staleMarksAtB2(fresh)).toEqual(stale);
      await expect.poll(() => reopened.readValue('B2')).toEqual({ kind: 'number', value: 2 });
      await expect.poll(() => calcOnSave(fresh)).toBe(false);
    });
  }
  await sp.expectNoConsoleErrors();
}

/** The Mac status bar offers Calculate while the active sheet holds stale values; it
 *  recalculates, stays in manual mode and hides again. The default profile has no such button. */
export async function runStatusBarRecalcScenario(
  page: Page,
  platform: Platform,
  locale: 'en' | 'ja' = 'en',
): Promise<void> {
  await recordStrokes(page);
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform, locale, fixture: 'empty' });
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.expectNoStub();
  await page.evaluate((manual) => {
    const w = window as unknown as StrokeWindow;
    const sheet = w.__fcInst.workbook.sheet(0);
    w.__fcInst.workbook.batch(() => {
      sheet.setNumber(1, 0, 1);
      sheet.setFormula(1, 1, '=A2*2');
    });
    w.__fcProbes['api/workbook'].workbookHandleOf(w.__fcInst.workbook).setCalcMode(manual);
    w.__fcInst.history.clear();
  }, MANUAL);
  const recalc = page.locator('.fc-host__statusbar-recalc');
  await expect(recalc).toBeHidden();
  await sp.enter('A2', '5');
  await expect.poll(async () => (await staleMarksAtB2(page)).stale).toBe(true);
  if (platform === 'default') {
    await expect(recalc).toBeHidden();
    await sp.expectNoConsoleErrors();
    return;
  }
  await expect(recalc).toBeVisible();
  await expect(recalc).toHaveText(locale === 'ja' ? '再計算' : 'Calculate');
  await recalc.click();
  await expect.poll(() => sp.readValue('B2')).toEqual({ kind: 'number', value: 10 });
  await expect
    .poll(async () => {
      const m = await staleMarksAtB2(page);
      return { mode: m.mode, stale: m.stale, triangle: m.triangle };
    })
    .toEqual({ mode: MANUAL, stale: false, triangle: false });
  await expect(recalc).toBeHidden();
  await sp.enter('A2', '6');
  await expect(recalc).toBeVisible();
  await sp.expectNoConsoleErrors();
}

/** Hovering the error triangle of a stale value shows the Stale Value tip; a plain cell shows
 *  none. */
export async function runStaleHoverTipScenario(page: Page, platform: Platform): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform, locale: 'en', fixture: 'empty' });
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.expectNoStub();
  await page.evaluate((manual) => {
    const w = window as unknown as StrokeWindow;
    const sheet = w.__fcInst.workbook.sheet(0);
    w.__fcInst.workbook.batch(() => {
      sheet.setNumber(1, 0, 1);
      sheet.setFormula(1, 1, '=A2*2');
    });
    w.__fcProbes['api/workbook'].workbookHandleOf(w.__fcInst.workbook).setCalcMode(manual);
    w.__fcInst.history.clear();
  }, MANUAL);
  await sp.enter('A2', '5');
  const triangle = () =>
    page.evaluate(() => {
      const w = window as unknown as StrokeWindow;
      const grid = document.querySelector<HTMLElement>('.fc-host__grid');
      if (!grid) return null;
      const hit = w.__fcProbes['render/grid/hit-state']
        .gridHitState(grid)
        .errorTriangles.find(({ addr }) => addr.sheet === 0 && addr.row === 1 && addr.col === 1);
      if (!hit) return null;
      const box = grid.getBoundingClientRect();
      return { x: box.x + hit.rect.x + hit.rect.w / 2, y: box.y + hit.rect.y + hit.rect.h / 2 };
    });
  await expect.poll(triangle).not.toBeNull();
  const at = await triangle();
  if (!at) throw new Error('B2 has no error-check triangle');
  const tip = page.locator('.fc-hover-tip');
  await page.mouse.move(at.x, at.y);
  await expect(tip).toBeVisible();
  expect(await tip.evaluate((el) => el.textContent)).toBe(
    'Stale Value\nThe value in this cell is stale. Calculate the workbook to see an up-to-date value.',
  );
  await page.mouse.move(at.x - 120, at.y + 80);
  await expect(tip).toBeHidden();
  await sp.expectNoConsoleErrors();
}
