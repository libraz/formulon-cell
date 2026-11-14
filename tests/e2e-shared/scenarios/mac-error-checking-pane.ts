import { expect, type Page } from '@playwright/test';

import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

/** Preferences opens over the File backstage; closing it leaves the backstage up. */
const closeBackstage = async (page: Page): Promise<void> => {
  const backstage = page.locator('.fc-tb__backstage[role="dialog"]').first();
  if (await backstage.isVisible()) await page.keyboard.press('Escape');
  await expect(backstage).toBeHidden();
};

/** Error Checking pane rows per UI language: the master switch, then the eleven rules. */
const ERROR_CHECKING_ROWS = {
  en: [
    'Turn on background error checking',
    'Formulas that result in an error',
    'Years in 2 digits',
    'Numbers formatted as text',
    "Formulas that don't match nearby formulas",
    'Formulas omitting cells in a region',
    'Unlocked cells containing formulas',
    'Formulas referring to empty cells',
    'Inconsistent calculated column formula in tables',
    'Misleading number formats',
    "Cells containing data types that couldn't refresh",
    'Cells containing stale values',
  ],
  ja: [
    'バックグラウンドでのエラー チェックをオンにする',
    '結果がエラーになる数式',
    '2 桁で示された年',
    '文字列形式の数値',
    '隣接する数式に一致していない数式',
    '領域内のセルを除外している数式',
    '数式を含むロックされていないセル',
    '空白セルを参照している数式',
    'テーブル内の矛盾した集計列の数式',
    '誤解を招く数値の書式',
    '更新できないデータの種類が含まれているセル',
    '古い値を含むセル',
  ],
} as const;
const ERROR_CHECKING_TEXT = {
  en: {
    pane: 'Error Checking',
    options: 'Options',
    reset: 'Reset Ignored Errors',
    showFor: 'Show Errors For',
  },
  ja: {
    pane: 'エラー チェック',
    options: 'オプション',
    reset: '無視したエラーのリセット',
    showFor: '表示するエラー',
  },
} as const;
/** [checked, enabled] per row: every rule is on except empty cells, and the rules no setting
 *  backs yet are greyed at their defaults. */
const ERROR_CHECKING_STATE: readonly (readonly [boolean, boolean])[] = [
  [true, true],
  [true, true],
  [true, false],
  [true, true],
  [true, true],
  [true, true],
  [true, true],
  [false, false],
  [true, false],
  [true, false],
  [true, false],
  [true, true],
];

/** Preferences > Error Checking lists the master switch, Reset Ignored Errors, the error colour
 *  and eleven rules in the desktop's order. Turning background checking off hides every
 *  triangle; Reset Ignored Errors brings back an ignored one and undoes; the stale rule is the
 *  Format Stale Values setting. */
export async function runMacErrorCheckingPaneScenario(
  page: Page,
  locale: keyof typeof ERROR_CHECKING_ROWS,
): Promise<void> {
  const text = ERROR_CHECKING_TEXT[locale];
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac', locale, fixture: 'empty' });
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.expectNoStub();
  await sp.enter('B3', '=1/0');
  type ProbeWindow = Window & { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
  const triangleB3 = () =>
    page.evaluate(() => {
      const w = window as unknown as ProbeWindow;
      const grid = document.querySelector<HTMLElement>('.fc-host__grid');
      if (!grid) return null;
      const hit = w.__fcProbes['render/grid/hit-state']
        .gridHitState(grid)
        .errorTriangles.find(({ addr }) => addr.sheet === 0 && addr.row === 2 && addr.col === 1);
      if (!hit) return null;
      const box = grid.getBoundingClientRect();
      return { x: box.x + hit.rect.x + hit.rect.w / 2, y: box.y + hit.rect.y + hit.rect.h / 2 };
    });
  const errorTag = () =>
    page.evaluate(() => {
      const w = window as unknown as ProbeWindow;
      const grid = document.querySelector<HTMLElement>('.fc-host__grid');
      const tag = grid ? w.__fcProbes['render/grid/hit-state'].gridHitState(grid).errorTag : null;
      if (!tag || !grid) return null;
      const box = grid.getBoundingClientRect();
      return { x: box.x + tag.rect.x + tag.rect.w / 2, y: box.y + tag.rect.y + tag.rect.h / 2 };
    });
  const formatStale = () =>
    page.evaluate(() => {
      const w = window as unknown as ProbeWindow;
      return w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst).store.getState().ui
        .formatStaleValues;
    });
  await expect.poll(triangleB3).not.toBeNull();

  const prefs = page.locator('.fc-prefs[role="dialog"]:not([hidden])');
  const pane = prefs.locator('.fc-prefs__pane[data-pane="errorChecking"]');
  const openPane = async () => {
    await page.locator('.demo__brand-mark').first().click();
    const backstage = page.locator('.fc-tb__backstage[role="dialog"]').first();
    await expect(backstage).toBeVisible();
    await backstage.getByRole('button', { name: text.options, exact: true }).first().click();
    await expect(prefs).toBeVisible();
    await prefs.locator('[role="toolbar"] .fc-prefs__heading').waitFor();
    await prefs.locator('[role="toolbar"] button').nth(2).click();
    await prefs.locator('.fc-prefs__pane-button[data-pane="errorChecking"]').click();
    await expect(pane).toBeVisible();
    await expect(prefs.locator('.fc-prefs__heading')).toHaveText(text.pane);
  };
  const closePrefs = async () => {
    await prefs.locator('.fc-fmtdlg__close').first().click();
    await expect(prefs).toBeHidden();
    await closeBackstage(page);
  };
  const rows = pane.locator('.fc-prefs__check');
  const box = (index: number) => rows.nth(index).locator('input[type="checkbox"]');

  await openPane();
  await expect(rows).toHaveText([...ERROR_CHECKING_ROWS[locale]]);
  const states = await rows.evaluateAll((labels) =>
    labels.map((label) => {
      const input = label.querySelector('input');
      return [!!input?.checked, !input?.disabled] as const;
    }),
  );
  expect(states).toEqual(ERROR_CHECKING_STATE);
  await expect(pane.locator('.fc-prefs__subheading')).toHaveText(text.showFor);
  await expect(pane.locator('[data-prefs-action="reset"]')).toHaveText(text.reset);
  await expect(pane.locator('[data-prefs-action="color"]')).toBeDisabled();

  // The stale rule and Format Stale Values are one setting.
  await box(11).uncheck();
  await expect.poll(formatStale).toBe(false);
  await box(11).check();
  await expect.poll(formatStale).toBe(true);

  // Background checking off hides the triangle; on brings it back.
  await box(0).uncheck();
  await closePrefs();
  await expect.poll(triangleB3).toBeNull();
  await openPane();
  await expect(box(0)).not.toBeChecked();
  await box(0).check();
  await closePrefs();
  await expect.poll(triangleB3).not.toBeNull();

  // Ignore Error hides it; Reset Ignored Errors restores it, and undo ignores it again.
  const triangle = await triangleB3();
  if (!triangle) throw new Error('B3 has no error-check triangle');
  await page.mouse.click(triangle.x, triangle.y);
  await expect.poll(errorTag).not.toBeNull();
  const tag = await errorTag();
  if (!tag) throw new Error('B3 has no error smart tag');
  await page.mouse.click(tag.x, tag.y);
  const menu = page.locator('.fc-errmenu:visible');
  await expect(menu).toBeVisible();
  await menu.locator('[data-fc-action="ignore"]').click();
  await expect.poll(triangleB3).toBeNull();
  await openPane();
  await pane.locator('[data-prefs-action="reset"]').click();
  await closePrefs();
  await expect.poll(triangleB3).not.toBeNull();
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Meta+z');
  await expect.poll(triangleB3).toBeNull();
  await sp.expectNoConsoleErrors();
}
