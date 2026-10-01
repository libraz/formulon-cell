import { describe, expect, it } from 'vitest';

import {
  conditionalRuleToEngineInput,
  syncConditionalRulesToEngine,
} from '../../../src/engine/cf-writeback.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { mutators } from '../../../src/store/store.js';
import { mountStubSheet } from '../../test-utils/mount.js';

const range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };

describe('conditional formatting across workbook replacement', () => {
  it.each([false, true])(
    'replaces old session rules (replacement has rules: %s)',
    async (hasRules) => {
      const original = await WorkbookHandle.createDefault();
      const next = await WorkbookHandle.createDefault();
      expect(original.isStub).toBe(false);
      expect(next.isStub).toBe(false);
      if (hasRules) {
        const input = conditionalRuleToEngineInput({
          kind: 'data-bar',
          range,
          color: '#638ec6',
          min: { kind: 'number', value: 20 },
          max: { kind: 'number', value: 80 },
          direction: 'right-to-left',
        });
        if (!input) throw new Error('Data bar must be representable by the engine');
        expect(next.addConditionalFormat(0, input)).toBeGreaterThanOrEqual(0);
      }
      const sheet = await mountStubSheet({ workbook: original });
      try {
        mutators.addConditionalRule(sheet.instance.store, {
          kind: 'cell-value',
          range,
          op: '>',
          a: 100,
          apply: { fill: '#ff0000' },
        });
        expect(sheet.instance.store.getState().conditional.rules).toHaveLength(1);
        expect(
          syncConditionalRulesToEngine(
            original,
            sheet.instance.store.getState().conditional.rules,
            0,
          ),
        ).toEqual({ written: 1, skipped: 0 });
        expect(original.getConditionalFormats(0)).toHaveLength(1);
        await sheet.instance.setWorkbook(next);
        const rules = sheet.instance.store.getState().conditional.rules;
        expect(rules).toHaveLength(hasRules ? 1 : 0);
        if (hasRules) {
          expect(rules[0]).toMatchObject({
            kind: 'data-bar',
            min: { kind: 'number', value: 20 },
            max: { kind: 'number', value: 80 },
            direction: 'right-to-left',
          });
          expect(rules[0]?.engineId).toBeDefined();
        }
        expect(next.getConditionalFormats(0)).toHaveLength(hasRules ? 1 : 0);
        expect(original.getConditionalFormats(0)).toHaveLength(1);
      } finally {
        sheet.dispose();
        original.dispose();
        next.dispose();
      }
    },
  );
});
