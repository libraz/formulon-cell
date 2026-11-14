/**
 * Which Mac ribbon groups give up width at a 1440 pt desktop window, per locale and tab.
 *
 * Entries are group positions within the tab panel. `reduce` groups run every fit
 * stage they allow (their labels drop); `collapse` groups fold into one named button.
 * A tab listed with no entries stays whole at that width. Widths between 1440 and 1920
 * are unmeasured, so the fit applies a plan as one step once the panel overflows.
 */

import { builtInLocaleFor } from '../../../i18n/locales.js';

export interface MacRibbonFitPlan {
  readonly reduce: readonly number[];
  readonly collapse: readonly number[];
}

const plan = (reduce: readonly number[] = [], collapse: readonly number[] = []): MacRibbonFitPlan =>
  Object.freeze({ reduce, collapse });

const WHOLE = plan();

/** Tabs measured whole in every locale. */
const UNCHANGED = { draw: WHOLE, pageLayout: WHOLE, review: WHOLE, automate: WHOLE };

// Home: 0 Clipboard, 2 Alignment, 4 Styles, 6 Editing. Insert: 0 Tables, 3 Illustrations,
// 10 Text, 11 Symbols. Formulas: 0 Function Library, 4 Calculation. Data: 0 Get & Transform,
// 1 Queries, 2 Data Types, 4 Data Tools, 6 Outline. View: 0 Sheet View, 2 Show, 5 Macros.
const MAC_RIBBON_FIT_1440: Readonly<Record<string, Readonly<Record<string, MacRibbonFitPlan>>>> = {
  fr: {
    ...UNCHANGED,
    home: plan([0, 2, 6]),
    insert: plan([], [0, 3, 10, 11]),
    formulas: plan([4]),
    data: plan([0, 1, 4], [6]),
    view: plan([0], [2]),
  },
  de: {
    ...UNCHANGED,
    home: plan([0, 2, 4, 6]),
    insert: plan([3], [10, 11]),
    formulas: plan([0, 4]),
    data: plan([0, 1, 4], [6]),
    view: plan([0], [2]),
  },
  en: {
    ...UNCHANGED,
    home: plan([0, 4, 6]),
    insert: plan([3, 10]),
    formulas: WHOLE,
    data: plan([0, 2, 4, 6]),
    view: plan([], [2, 5]),
  },
  ja: {
    ...UNCHANGED,
    home: plan([0, 4, 6]),
    insert: plan([3, 10]),
    formulas: WHOLE,
    data: plan([0, 2, 6]),
    view: plan([], [2, 5]),
  },
};

/** The measured plan for `tab` in `lang`, or undefined where none was measured. */
export const macRibbonFitPlan = (lang: string, tab: string): MacRibbonFitPlan | undefined => {
  const locale = builtInLocaleFor(lang);
  return locale ? MAC_RIBBON_FIT_1440[locale]?.[tab] : undefined;
};
