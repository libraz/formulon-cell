import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { buildRibbonModel, type ToolbarLang } from '../../../src/toolbar/ribbon-model.js';

/** Group titles the desktop app draws on each Mac tab, in order, per locale: the app's own
 *  group labels with accelerators stripped. `groups` names each slot's ribbon group. */
const measured = JSON.parse(
  readFileSync(join(__dirname, '../../fixtures/mac-ribbon-group-titles.json'), 'utf8'),
) as { groups: Record<string, string[]>; titles: Record<ToolbarLang, Record<string, string[]>> };

/** Groups whose content differs from the desktop's, so their slots are compared once the tab
 *  carries the desktop group: ours by English title, the desktop's by ribbon group. */
const OURS_PENDING = new Set(['Queries & Connections']);
const DESKTOP_PENDING = new Set(['ChunkDataConnections']);

const titles = (lang: ToolbarLang): Map<string, string[]> =>
  new Map(
    buildRibbonModel(lang, {
      profile: 'excel365Mac',
      contextualTabs: [
        'chartDesign',
        'chartFormat',
        'tableDesign',
        'pivotAnalyze',
        'pivotDesign',
        'sparklineDesign',
      ],
    }).map((tab) => [tab.id, tab.groups.map((group) => group.title)]),
  );

describe('Mac ribbon group titles', () => {
  const english = titles('en');

  it.each(Object.keys(measured.titles) as ToolbarLang[])(
    'draw the desktop group labels in order in %s',
    (lang) => {
      const ours = titles(lang);
      for (const [tab, slots] of Object.entries(measured.groups)) {
        const keys = english.get(tab) ?? [];
        const drawn = (ours.get(tab) ?? []).filter((_, i) => !OURS_PENDING.has(keys[i] ?? ''));
        const desktop = measured.titles[lang][tab]?.filter(
          (_, i) => !DESKTOP_PENDING.has(slots[i] ?? ''),
        );
        expect(drawn, `${lang} ${tab}`).toEqual(desktop);
      }
    },
  );
});
