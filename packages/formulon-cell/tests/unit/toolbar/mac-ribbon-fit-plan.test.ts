import { describe, expect, it } from 'vitest';
import { macRibbonFitPlan } from '../../../src/toolbar/ribbon/mac/fit-plan.js';
import { buildRibbonModel } from '../../../src/toolbar/ribbon-model.js';

type Expected = { reduce?: readonly string[]; collapse?: readonly string[] };

/** Groups the desktop reduces at a 1440 pt window, by English group title. */
const MEASURED_1440: Record<string, Partial<Record<string, Expected>>> = {
  fr: {
    home: { reduce: ['Clipboard', 'Alignment', 'Editing'] },
    insert: { collapse: ['Tables', 'Illustrations', 'Text', 'Symbols'] },
    formulas: { reduce: ['Calculation'] },
    data: {
      reduce: ['Get & Transform Data', 'Queries & Connections', 'Data Tools'],
      collapse: ['Outline'],
    },
    view: { reduce: ['Sheet View'], collapse: ['Show'] },
  },
  de: {
    home: { reduce: ['Clipboard', 'Alignment', 'Styles', 'Editing'] },
    insert: { reduce: ['Illustrations'], collapse: ['Text', 'Symbols'] },
    formulas: { reduce: ['Function Library', 'Calculation'] },
    data: {
      reduce: ['Get & Transform Data', 'Queries & Connections', 'Data Tools'],
      collapse: ['Outline'],
    },
    view: { reduce: ['Sheet View'], collapse: ['Show'] },
  },
  en: {
    home: { reduce: ['Clipboard', 'Styles', 'Editing'] },
    insert: { reduce: ['Illustrations', 'Text'] },
    data: { reduce: ['Get & Transform Data', 'Data Types', 'Data Tools', 'Outline'] },
    view: { collapse: ['Show', 'Macros'] },
  },
  ja: {
    home: { reduce: ['Clipboard', 'Styles', 'Editing'] },
    insert: { reduce: ['Illustrations', 'Text'] },
    data: { reduce: ['Get & Transform Data', 'Data Types', 'Outline'] },
    view: { collapse: ['Show', 'Macros'] },
  },
};

const MEASURED_TABS = [
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

const macTabs = buildRibbonModel('en', { profile: 'excel365Mac' });

describe('Mac ribbon fit plan at 1440', () => {
  for (const [lang, tabs] of Object.entries(MEASURED_1440)) {
    for (const tabId of MEASURED_TABS) {
      it(`${lang} ${tabId} reduces only the measured groups`, () => {
        const tab = macTabs.find((candidate) => candidate.id === tabId);
        const plan = macRibbonFitPlan(lang, tabId);
        expect(plan).toBeDefined();
        const titles = (indices: readonly number[] = []) =>
          indices.map((index) => tab?.groups[index]?.title);
        expect({ reduce: titles(plan?.reduce), collapse: titles(plan?.collapse) }).toEqual({
          reduce: tabs[tabId]?.reduce ?? [],
          collapse: tabs[tabId]?.collapse ?? [],
        });
      });
    }
  }

  it('collapses only groups of several commands', () => {
    for (const lang of Object.keys(MEASURED_1440)) {
      for (const tabId of MEASURED_TABS) {
        const tab = macTabs.find((candidate) => candidate.id === tabId);
        for (const index of macRibbonFitPlan(lang, tabId)?.collapse ?? []) {
          const commands = tab?.groups[index]?.commands.filter(
            (command) => command.kind !== 'break' && command.kind !== 'label',
          );
          expect(commands?.length ?? 0).toBeGreaterThan(1);
        }
      }
    }
  });

  it('resolves region tags and has no plan for unmeasured locales or tabs', () => {
    expect(macRibbonFitPlan('fr-FR', 'formulas')).toEqual(macRibbonFitPlan('fr', 'formulas'));
    expect(macRibbonFitPlan('th', 'home')).toBeUndefined();
    expect(macRibbonFitPlan('fr', 'developer')).toBeUndefined();
  });
});
