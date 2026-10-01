import { describe, expect, it } from 'vitest';
import { hydrateConditionalRulesFromEngine } from '../../../src/engine/cf-sync.js';
import {
  conditionalRuleToEngineInput,
  type SyncedConditionalRuleMap,
  syncTrackedConditionalRulesToEngine,
} from '../../../src/engine/cf-writeback.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import type { ConditionalRule } from '../../../src/store/types.js';

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';

const iconRange = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 };
const barRange = { sheet: 0, r0: 0, c0: 1, r1: 2, c1: 1 };
const importedRange = [{ firstRow: 4, firstCol: 0, lastRow: 4, lastCol: 0 }];
type ConditionalFormatEntry = ReturnType<WorkbookHandle['getConditionalFormats']>[number];

const iconInput = (formats: readonly ConditionalFormatEntry[]) =>
  formats.find((entry) => entry.type === 4);

const dataBarInput = (formats: readonly ConditionalFormatEntry[]) =>
  formats.find((entry) => entry.type === 3);

describe.skipIf(!canLoadWasm())('tracked conditional-format round-trip', () => {
  it.each([
    {
      name: 'positive-only border',
      bothBorders: false,
      ruleField: 'borderColor',
      engineField: 'border',
      otherRuleField: 'negativeBorderColor',
      otherEngineField: 'negativeBorder',
    },
    {
      name: 'negative-only border',
      bothBorders: false,
      ruleField: 'negativeBorderColor',
      engineField: 'negativeBorder',
      otherRuleField: 'borderColor',
      otherEngineField: 'border',
    },
    {
      name: 'equal positive and negative borders',
      bothBorders: true,
      ruleField: 'borderColor',
      engineField: 'border',
      otherRuleField: 'negativeBorderColor',
      otherEngineField: 'negativeBorder',
    },
  ] as const)('preserves $name through hydration and XLSX', async (fields) => {
    const rule: ConditionalRule = {
      kind: 'data-bar',
      range: barRange,
      color: '#638ec6',
      [fields.ruleField]: '#1f1f1f',
      ...(fields.bothBorders ? { negativeBorderColor: '#1f1f1f' } : {}),
    };
    const expectedColor = { a: 255, r: 31, g: 31, b: 31 };
    const assertBorders = (
      bar:
        | Pick<NonNullable<ConditionalFormatEntry['dataBar']>, 'border' | 'negativeBorder'>
        | undefined,
    ): void => {
      expect(bar?.[fields.engineField]).toEqual(expectedColor);
      expect(bar?.[fields.otherEngineField]).toEqual(
        fields.bothBorders ? expectedColor : undefined,
      );
    };
    assertBorders(conditionalRuleToEngineInput(rule)?.dataBar);
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.isStub).toBe(false);
      expect(syncTrackedConditionalRulesToEngine(wb, [rule], 0, { tracked: new Map() })).toEqual({
        written: 1,
        skipped: 0,
        removed: 0,
      });
      assertBorders(dataBarInput(wb.getConditionalFormats(0))?.dataBar);
      const store = createSpreadsheetStore();
      hydrateConditionalRulesFromEngine(wb, store, 0);
      expect(store.getState().conditional.rules).toHaveLength(1);
      const restoredRule = store.getState().conditional.rules[0];
      if (restoredRule?.kind !== 'data-bar') throw new Error('Expected hydrated data bar');
      expect(restoredRule[fields.ruleField]).toBe('rgb(31, 31, 31)');
      expect(restoredRule[fields.otherRuleField]).toBe(
        fields.bothBorders ? 'rgb(31, 31, 31)' : undefined,
      );
      assertBorders(conditionalRuleToEngineInput(restoredRule)?.dataBar);
      const loaded = await WorkbookHandle.loadBytes(wb.save());
      try {
        expect(loaded.isStub).toBe(false);
        assertBorders(dataBarInput(loaded.getConditionalFormats(0))?.dataBar);
      } finally {
        loaded.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('reconciles edited icon and data-bar rules without duplicating them', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.isStub).toBe(false);
      expect(wb.capabilities.conditionalFormatMutate).toBe(true);

      expect(
        wb.addConditionalFormat(0, {
          sqref: importedRange,
          type: 1,
          op: 5,
          formula1: '999',
        }),
      ).toBeGreaterThanOrEqual(0);

      const firstRules: ConditionalRule[] = [
        {
          kind: 'icon-set',
          range: iconRange,
          icons: 'traffic3',
          thresholds: [
            { kind: 'percent', value: 33, gte: false },
            { kind: 'percent', value: 67, gte: false },
          ],
          showValue: true,
        },
        {
          kind: 'data-bar',
          range: barRange,
          color: '#0078d4',
          axisPosition: 'middle',
          negativeColor: '#c00000',
          borderColor: '#1f1f1f',
          negativeBorderColor: '#7f0000',
          axisColor: '#404040',
          min: { kind: 'number', value: 20 },
          max: { kind: 'number', value: 80 },
          direction: 'right-to-left',
        },
      ];
      const tracked: SyncedConditionalRuleMap = new Map();

      expect(syncTrackedConditionalRulesToEngine(wb, firstRules, 0, { tracked })).toEqual({
        written: 2,
        skipped: 0,
        removed: 0,
      });
      expect(wb.getConditionalFormats(0)).toHaveLength(3);
      expect(dataBarInput(wb.getConditionalFormats(0))?.dataBar).toMatchObject({
        axisPosition: 1,
        negativeFill: { a: 255, r: 192, g: 0, b: 0 },
        border: { a: 255, r: 31, g: 31, b: 31 },
        negativeBorder: { a: 255, r: 127, g: 0, b: 0 },
        axisColor: { a: 255, r: 64, g: 64, b: 64 },
      });

      const secondRules: ConditionalRule[] = [
        {
          kind: 'icon-set',
          range: iconRange,
          icons: 'traffic3',
          thresholds: [
            { kind: 'number', value: 20 },
            { kind: 'number', value: 40 },
          ],
          showValue: true,
        },
        {
          kind: 'data-bar',
          range: barRange,
          color: '#0078d4',
          axisPosition: 'none',
          negativeColor: '#c00000',
          borderColor: '#1f1f1f',
          negativeBorderColor: '#7f0000',
          axisColor: '#404040',
          min: { kind: 'number', value: 30 },
          max: { kind: 'number', value: 70 },
          direction: 'left-to-right',
        },
      ];

      expect(syncTrackedConditionalRulesToEngine(wb, secondRules, 0, { tracked })).toEqual({
        written: 2,
        skipped: 0,
        removed: 2,
      });
      const edited = wb.getConditionalFormats(0);
      expect(edited).toHaveLength(3);
      expect(edited.some((entry) => entry.type === 1 && entry.formula1 === '999')).toBe(true);
      expect(iconInput(edited)?.iconSet).toMatchObject({
        thresholds: [
          { type: 0, value: '20' },
          { type: 0, value: '40' },
        ],
      });
      expect(dataBarInput(edited)?.dataBar).toMatchObject({
        min: { type: 0, value: '30' },
        max: { type: 0, value: '70' },
        axisPosition: 2,
        negativeFill: { a: 255, r: 192, g: 0, b: 0 },
        border: { a: 255, r: 31, g: 31, b: 31 },
        negativeBorder: { a: 255, r: 127, g: 0, b: 0 },
        axisColor: { a: 255, r: 64, g: 64, b: 64 },
        direction: 1,
      });

      const loaded = await WorkbookHandle.loadBytes(wb.save());
      try {
        expect(loaded.isStub).toBe(false);
        const restored = loaded.getConditionalFormats(0);
        expect(restored).toHaveLength(3);
        expect(restored.some((entry) => entry.type === 1 && entry.formula1 === '999')).toBe(true);
        expect(iconInput(restored)?.iconSet).toMatchObject({
          thresholds: [
            { type: 0, value: '20' },
            { type: 0, value: '40' },
          ],
        });
        expect(dataBarInput(restored)?.dataBar).toMatchObject({
          min: { type: 0, value: '30' },
          max: { type: 0, value: '70' },
          axisPosition: 2,
          negativeFill: { a: 255, r: 192, g: 0, b: 0 },
          border: { a: 255, r: 31, g: 31, b: 31 },
          negativeBorder: { a: 255, r: 127, g: 0, b: 0 },
          axisColor: { a: 255, r: 64, g: 64, b: 64 },
          direction: 1,
        });
      } finally {
        loaded.dispose();
      }
    } finally {
      wb.dispose();
    }
  });
});
