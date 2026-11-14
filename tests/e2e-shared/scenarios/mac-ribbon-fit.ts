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

type Engine = 'chromium' | 'webkit';

/** Accepted differences from the desktop at 1440, keyed `locale/tab/group`: our drawn width per
 *  engine (a group missing for an engine matches the desktop there) and the desktop's width.
 *  `extra` rows are narrowed where the desktop keeps them whole; `wide` rows are planned to narrow
 *  but take no stage and stay wider than the desktop's narrowed group. The widths are pinned, so
 *  any further drift fails. */
const ACCEPTED: Readonly<
  Record<string, { kind: 'extra' | 'wide'; ours: Partial<Record<Engine, number>>; desktop: number }>
> = {
  'fr/home/4': { kind: 'extra', ours: { chromium: 203.5, webkit: 205.5 }, desktop: 220.5 },
  'fr/home/5': { kind: 'extra', ours: { chromium: 109.2, webkit: 109.2 }, desktop: 112 },
  'fr/formulas/2': { kind: 'extra', ours: { chromium: 116, webkit: 116 }, desktop: 244 },
  'fr/formulas/3': { kind: 'extra', ours: { chromium: 215, webkit: 215 }, desktop: 325 },
  'fr/view/5': { kind: 'extra', ours: { chromium: 42, webkit: 42 }, desktop: 48 },
  'de/formulas/3': { kind: 'extra', ours: { chromium: 97, webkit: 97 }, desktop: 406.5 },
  'de/view/5': { kind: 'extra', ours: { chromium: 42, webkit: 42 }, desktop: 48 },
  'ja/insert/1': { kind: 'extra', ours: { chromium: 64.5 }, desktop: 48 },
  'ja/insert/11': { kind: 'extra', ours: { chromium: 89.3 }, desktop: 77 },
  'de/formulas/0': { kind: 'wide', ours: { chromium: 816, webkit: 816 }, desktop: 354.5 },
  'en/data/0': { kind: 'wide', ours: { chromium: 142, webkit: 142 }, desktop: 122.5 },
  'ja/data/0': { kind: 'wide', ours: { chromium: 166, webkit: 155 }, desktop: 135.5 },
  'ja/insert/10': { kind: 'wide', ours: { webkit: 206 }, desktop: 140.5 },
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

/** At a 1440 wide Mac window every tab fits without scrolling and exactly the desktop's groups
 *  collapse. No group the desktop keeps whole is narrowed, and each group the desktop narrows is
 *  narrowed too unless it already draws no wider than the desktop's narrowed group, except for
 *  the accepted differences, which must hold at their pinned widths. */
export async function runMacRibbonFitScenario(
  page: Page,
  locale: MacRibbonFitLocale,
): Promise<void> {
  const engine: Engine =
    page.context().browser()?.browserType().name() === 'webkit' ? 'webkit' : 'chromium';
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
        // Each deviation from the plan must be an accepted row at its pinned width, and each
        // accepted row must still deviate.
        const deviations = new Map<number, 'extra' | 'wide'>();
        for (const i of positions((g) => g.reduced && !g.collapsed))
          if (!(i in plan.reduce)) deviations.set(i, 'extra');
        for (const [at, desktop] of Object.entries(plan.reduce)) {
          const g = read.groups[Number(at)];
          if (g && !g.reduced && g.width > desktop + 0.5) deviations.set(Number(at), 'wide');
        }
        const found = [...deviations].map(([i, kind]) => {
          const accepted = ACCEPTED[`${locale}/${tab}/${i}`];
          const pinned = accepted?.kind === kind ? accepted.ours[engine] : undefined;
          const width = read.groups[i]?.width ?? 0;
          return pinned !== undefined && Math.abs(width - pinned) <= 0.5
            ? `${i} ${kind}`
            : `${i} ${kind} at ${width.toFixed(1)} (not accepted)`;
        });
        const accepted = Object.entries(ACCEPTED).flatMap(([key, row]) =>
          key.startsWith(`${locale}/${tab}/`) && row.ours[engine] !== undefined
            ? [`${key.split('/')[2]} ${row.kind}`]
            : [],
        );
        expect(found.sort(), `${where} deviations from the desktop`).toEqual(accepted.sort());
      })
      .toPass({ timeout: 3_000 });
  }
  await sp.expectNoConsoleErrors();
}
