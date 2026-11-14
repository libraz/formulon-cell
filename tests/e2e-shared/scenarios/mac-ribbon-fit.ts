import { expect, type Page } from '@playwright/test';

import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

export type MacRibbonFitLocale = 'fr' | 'de' | 'en' | 'ja';

/** `reduce` maps a group position to the desktop's drawn width of that group at 1440 pt, after it
 *  gave up width; `collapse` lists the positions folded into one button. */
type Fit = { reduce: Readonly<Record<number, number>>; collapse: readonly number[] };

const planned = (reduce: Record<number, number> = {}, collapse: number[] = []): Fit => ({
  reduce,
  collapse,
});
const WHOLE = planned();
const UNCHANGED = { draw: WHOLE, pageLayout: WHOLE, review: WHOLE, automate: WHOLE };

/** Groups the desktop narrows or folds at a 1440 pt window, by position within each tab panel. */
const MEASURED: Record<MacRibbonFitLocale, Record<string, Fit>> = {
  fr: {
    ...UNCHANGED,
    home: planned({ 0: 84, 2: 199, 6: 161 }),
    insert: planned({}, [0, 3, 10, 11]),
    formulas: planned({ 4: 74 }),
    data: planned({ 0: 191, 1: 78.5, 4: 252 }, [6]),
    view: planned({ 0: 102 }, [2]),
  },
  de: {
    ...UNCHANGED,
    home: planned({ 0: 84.5, 2: 199, 4: 249.5, 6: 158 }),
    insert: planned({ 3: 70 }, [10, 11]),
    formulas: planned({ 0: 354.5, 4: 140.5 }),
    data: planned({ 0: 168, 1: 92, 4: 229 }, [6]),
    view: planned({ 0: 102 }, [2]),
  },
  en: {
    ...UNCHANGED,
    home: planned({ 0: 84, 4: 161.5, 6: 134 }),
    insert: planned({ 3: 70, 10: 127.5 }),
    formulas: WHOLE,
    data: planned({ 0: 122.5, 2: 240, 4: 196, 6: 109.5 }),
    view: planned({}, [2, 5]),
  },
  ja: {
    ...UNCHANGED,
    home: planned({ 0: 84, 4: 162.5, 6: 138.5 }),
    insert: planned({ 3: 70, 10: 140.5 }),
    formulas: WHOLE,
    data: planned({ 0: 135.5, 2: 166, 6: 124.5 }),
    view: planned({}, [2, 5]),
  },
};

const TABS = [
  'home',
  'insert',
  'draw',
  'pageLayout',
  'formulas',
  'data',
  'review',
  'view',
  'automate',
] as const;

/** The panel's overflow and, per group position, its width and whether it is collapsed or
 *  carries a narrowing stage. */
const readFit = (page: Page, tab: string) =>
  page.locator(`[data-ribbon-panel="${tab}"]`).evaluate((panel) => ({
    overflow: panel.scrollWidth - panel.clientWidth,
    groups: Array.from(panel.querySelectorAll(':scope > section.fc-tb__ribbon-group')).map(
      (group) => {
        const has = (stage: string) => group.classList.contains(`fc-tb__ribbon-group--${stage}`);
        return {
          width: group.getBoundingClientRect().width,
          collapsed: has('collapsed'),
          reduced: has('narrow') || has('condensed') || has('icon-only'),
        };
      },
    ),
  }));

/** At a 1440 wide Mac window every tab fits without scrolling; exactly the desktop's groups
 *  collapse; no group the desktop keeps whole is narrowed; and each group the desktop narrows
 *  is narrowed too, unless it already draws no wider than the desktop's narrowed group. */
export async function runMacRibbonFitScenario(
  page: Page,
  locale: MacRibbonFitLocale,
): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac', locale });
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.expectNoStub();
  for (const tab of TABS) {
    await selectRibbonTab(page, tab);
    await expect(page.locator(`[data-ribbon-panel="${tab}"]`)).toBeVisible();
    const plan = MEASURED[locale][tab] ?? WHOLE;
    await expect
      .soft(async () => {
        const read = await readFit(page, tab);
        const where = `${locale} ${tab}`;
        expect(read.overflow, `${where} overflow`).toBeLessThanOrEqual(0);
        const positions = (pick: (g: (typeof read.groups)[number]) => boolean) =>
          read.groups.flatMap((g, i) => (pick(g) ? [i] : []));
        expect(
          positions((g) => g.collapsed),
          `${where} collapsed groups`,
        ).toEqual([...plan.collapse]);
        expect(
          positions((g) => g.reduced && !g.collapsed).filter((i) => !(i in plan.reduce)),
          `${where} unplanned narrowed groups`,
        ).toEqual([]);
        const tooWide = Object.entries(plan.reduce).flatMap(([at, desktop]) => {
          const g = read.groups[Number(at)];
          return g && !g.reduced && g.width > desktop + 0.5
            ? [`${at}: ${g.width.toFixed(1)} > ${desktop}`]
            : [];
        });
        expect(tooWide, `${where} planned groups left wider than the desktop`).toEqual([]);
      })
      .toPass({ timeout: 3_000 });
  }
  await sp.expectNoConsoleErrors();
}
