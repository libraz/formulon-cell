import { describe, expect, it } from 'vitest';
import {
  evaluateCfFromEngine,
  hydrateConditionalRulesFromEngine,
} from '../../../../src/engine/cf-sync.js';
import {
  conditionalRuleToEngineInput,
  rewriteLoadedConditionalRule,
  syncConditionalRulesToEngine,
} from '../../../../src/engine/cf-writeback.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { type ConditionalRule, createSpreadsheetStore } from '../../../../src/store/store.js';
import { zipEntries } from '../zip-entries.js';

import { canLoadWasm } from './fixtures.js';

type IconRule = Extract<ConditionalRule, { kind: 'icon-set' }>;

const column = (col: number) => ({ sheet: 0, r0: 0, c0: col, r1: 4, c1: col });
const pct = (value: number) => ({ kind: 'percent' as const, value });

const authored: IconRule[] = [
  {
    kind: 'icon-set',
    range: column(0),
    icons: 'arrows5',
    thresholds: [pct(20), pct(40), pct(60), pct(80)],
  },
  { kind: 'icon-set', range: column(1), icons: 'triangles3', thresholds: [pct(33), pct(67)] },
  {
    kind: 'icon-set',
    range: column(2),
    icons: 'traffic3',
    thresholds: [pct(33), pct(67)],
    customIcons: [{ set: 'stars3', slot: 2 }, null, { set: 'redToBlack4', slot: 3 }],
  },
];

const iconRules = (wb: WorkbookHandle): IconRule[] => {
  const store = createSpreadsheetStore();
  hydrateConditionalRulesFromEngine(wb, store, 0);
  return store
    .getState()
    .conditional.rules.filter((rule): rule is IconRule => rule.kind === 'icon-set');
};

describe.skipIf(!canLoadWasm())('real xlsx round-trip of icon sets', () => {
  it('saves every family and per-level custom icons and reads them back', async () => {
    const first = await WorkbookHandle.createDefault();
    try {
      if (!first.capabilities.conditionalFormatMutate) return;
      for (let row = 0; row < 5; row += 1) {
        for (let col = 0; col < 3; col += 1) first.setNumber({ sheet: 0, row, col }, row + 1);
      }
      expect(syncConditionalRulesToEngine(first, authored, 0)).toEqual({ written: 3, skipped: 0 });

      const bytes = first.save();
      const sheetXml = new TextDecoder().decode(zipEntries(bytes).get('xl/worksheets/sheet1.xml'));
      expect(sheetXml).toContain('iconSet="5Arrows"');
      expect(sheetXml).toContain('<x14:cfIcon iconSet="3Triangles" iconId="2"/>');
      expect(sheetXml).toContain('<x14:cfIcon iconSet="NoIcons" iconId="0"/>');

      const loaded = await WorkbookHandle.loadBytes(bytes);
      try {
        const rules = iconRules(loaded);
        expect(rules.map(({ icons, customIcons }) => ({ icons, customIcons }))).toEqual([
          { icons: 'arrows5', customIcons: undefined },
          { icons: 'triangles3', customIcons: undefined },
          {
            icons: 'traffic3',
            customIcons: [{ set: 'stars3', slot: 2 }, null, { set: 'redToBlack4', slot: 3 }],
          },
        ]);

        const overlay = evaluateCfFromEngine(loaded, 0, 0, 0, 4, 2);
        expect(overlay.get('0:0:2')).toMatchObject({ iconKind: 'stars3', iconSlot: 2 });
        expect(overlay.get('0:2:2')?.iconKind).toBeUndefined();
        expect(overlay.get('0:4:2')).toMatchObject({ iconKind: 'redToBlack4', iconSlot: 3 });
        expect(overlay.get('0:4:1')).toMatchObject({ iconKind: 'triangles3', iconSlot: 2 });

        const custom = rules[2];
        if (!custom) throw new Error('Expected the custom icon rule.');
        expect(rewriteLoadedConditionalRule(loaded, custom, { ...custom, showValue: false })).toBe(
          true,
        );
        const rewritten = loaded
          .getConditionalFormats(0)
          .find((entry) => entry.id === custom.engineId);
        expect(rewritten?.iconSet).toMatchObject({
          name: 3,
          showValue: false,
          icons: [
            { iconSet: '3Stars', iconId: 2 },
            { iconSet: 'NoIcons', iconId: 0 },
            { iconSet: '4RedToBlack', iconId: 3 },
          ],
        });
      } finally {
        loaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('keeps an engine set the gallery does not offer under its own name', async () => {
    const first = await WorkbookHandle.createDefault();
    try {
      if (!first.capabilities.conditionalFormatMutate) return;
      expect(
        first.addConditionalFormat(0, {
          sqref: [{ firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 0 }],
          type: 4,
          iconSet: {
            name: 8,
            thresholds: [25, 50, 75].map((value) => ({ type: 1, value: String(value) })),
          },
        }),
      ).toBeGreaterThanOrEqual(0);
      const loaded = await WorkbookHandle.loadBytes(first.save());
      try {
        const [rule] = iconRules(loaded);
        expect(rule?.icons).toBe('arrows4');
        expect(rule && conditionalRuleToEngineInput(rule)?.iconSet?.name).toBe(8);
      } finally {
        loaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });
});
