import { describe, expect, it } from 'vitest';
import { hydrateConditionalRulesFromEngine } from '../../../src/engine/cf-sync.js';
import {
  conditionalRuleToEngineInput,
  inputWithDxf,
  rewriteLoadedConditionalRule,
  type SyncedConditionalRuleMap,
  syncTrackedConditionalRulesToEngine,
} from '../../../src/engine/cf-writeback.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import type { ConditionalRule } from '../../../src/store/types.js';

const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };
type BlanksRule = Extract<ConditionalRule, { apply: unknown }>;
const blanks = (apply: BlanksRule['apply']): BlanksRule => ({ kind: 'blanks', range, apply });

describe('conditional format manager persistence', () => {
  it('keeps two identical session rules as two engine rules', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const tracked: SyncedConditionalRuleMap = new Map();
      const twin = blanks({ bold: true });
      syncTrackedConditionalRulesToEngine(wb, [twin, { ...twin }], 0, { tracked });
      expect(wb.getConditionalFormats(0)).toHaveLength(2);

      const restored = await WorkbookHandle.loadBytes(wb.save());
      try {
        expect(restored.getConditionalFormats(0)).toHaveLength(2);
      } finally {
        restored.dispose();
      }

      // Re-syncing is stable, and dropping one twin removes exactly one rule.
      syncTrackedConditionalRulesToEngine(wb, [twin, { ...twin }], 0, { tracked });
      expect(wb.getConditionalFormats(0)).toHaveLength(2);
      syncTrackedConditionalRulesToEngine(wb, [twin], 0, { tracked });
      expect(wb.getConditionalFormats(0)).toHaveLength(1);
    } finally {
      wb.dispose();
    }
  });

  it('edits a loaded rule in place, keeping its id and priority', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      const first = blanks({ bold: true });
      const second: ConditionalRule = { kind: 'duplicates', range, apply: { italic: true } };
      for (const rule of [first, second]) {
        const input = conditionalRuleToEngineInput(rule);
        if (!input) throw new Error('untranslatable rule');
        expect(wb.addConditionalFormat(0, inputWithDxf(wb, input, rule))).toBeGreaterThanOrEqual(0);
      }
      hydrateConditionalRulesFromEngine(wb, store, 0);
      const loaded = store
        .getState()
        .conditional.rules.find((rule): rule is BlanksRule => rule.kind === 'blanks');
      expect(loaded?.engineId).toBeTruthy();
      const before = wb.getConditionalFormats(0).find((entry) => entry.id === loaded?.engineId);

      if (!loaded) throw new Error('rule did not hydrate');
      const edited: BlanksRule = { ...loaded, apply: { italic: true } };
      expect(rewriteLoadedConditionalRule(wb, loaded, edited)).toBe(true);

      const after = wb.getConditionalFormats(0);
      expect(after).toHaveLength(2);
      const entry = after.find((candidate) => candidate.id === loaded.engineId);
      expect(entry?.priority).toBe(before?.priority);
      expect(wb.getDxf(entry?.dxfId ?? -1)?.font?.italic).toBe(true);

      const restored = await WorkbookHandle.loadBytes(wb.save());
      try {
        expect(restored.getConditionalFormats(0)).toHaveLength(2);
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });
});
