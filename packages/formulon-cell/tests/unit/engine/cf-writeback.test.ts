import { describe, expect, it, vi } from 'vitest';
import {
  conditionalRuleToEngineInput,
  type SyncedConditionalRuleMap,
  syncConditionalRulesToEngine,
  syncTrackedConditionalRulesToEngine,
} from '../../../src/engine/cf-writeback.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import type { ConditionalRule } from '../../../src/store/types.js';

const range = { sheet: 0, r0: 1, c0: 2, r1: 3, c1: 4 };
const sqref = [{ firstRow: 1, firstCol: 2, lastRow: 3, lastCol: 4 }];

describe('conditionalRuleToEngineInput (H-13)', () => {
  it('translates a cell-value rule with operator + threshold', () => {
    const rule: ConditionalRule = {
      kind: 'cell-value',
      range,
      op: '>',
      a: 10,
      apply: { bold: true },
    };
    expect(conditionalRuleToEngineInput(rule)).toEqual({
      sqref,
      type: 1,
      op: 5,
      formula1: '10',
    });
    expect(
      conditionalRuleToEngineInput({
        kind: 'cell-value',
        range,
        op: '=',
        a: 'Done',
        apply: {},
      }),
    ).toEqual({ sqref, type: 1, op: 2, formula1: 'Done' });
  });

  it('emits formula2 for a between rule', () => {
    const rule: ConditionalRule = {
      kind: 'cell-value',
      range,
      op: 'between',
      a: 1,
      b: 9,
      apply: {},
    };
    const input = conditionalRuleToEngineInput(rule);
    expect(input).toMatchObject({ type: 1, op: 6, formula1: '1', formula2: '9' });
  });

  it('translates an =-prefixed formula rule to an expression rule', () => {
    const rule: ConditionalRule = {
      kind: 'formula',
      range,
      formula: '=A1>B1',
      apply: {},
    };
    expect(conditionalRuleToEngineInput(rule)).toEqual({
      sqref,
      type: 0,
      formula1: 'A1>B1',
    });
  });

  it('skips a comparator-prefix formula rule (not a standalone expression)', () => {
    const rule: ConditionalRule = { kind: 'formula', range, formula: '>10', apply: {} };
    expect(conditionalRuleToEngineInput(rule)).toBeNull();
  });

  it('translates text / blanks / errors / duplicate predicate rules', () => {
    expect(
      conditionalRuleToEngineInput({ kind: 'text-contains', range, text: 'hi', apply: {} }),
    ).toEqual({ sqref, type: 7, text: 'hi' });
    expect(
      conditionalRuleToEngineInput({
        kind: 'text-contains',
        range,
        text: 'hi',
        mode: 'begins-with',
        apply: {},
      }),
    ).toEqual({ sqref, type: 9, text: 'hi' });
    expect(
      conditionalRuleToEngineInput({
        kind: 'text-contains',
        range,
        text: 'hi',
        mode: 'ends-with',
        apply: {},
      }),
    ).toEqual({ sqref, type: 10, text: 'hi' });
    expect(
      conditionalRuleToEngineInput({
        kind: 'text-contains',
        range,
        text: 'hi',
        mode: 'not-contains',
        apply: {},
      }),
    ).toEqual({ sqref, type: 8, text: 'hi' });
    expect(conditionalRuleToEngineInput({ kind: 'blanks', range, apply: {} })).toEqual({
      sqref,
      type: 11,
    });
    expect(conditionalRuleToEngineInput({ kind: 'errors', range, apply: {} })).toEqual({
      sqref,
      type: 13,
    });
    expect(conditionalRuleToEngineInput({ kind: 'duplicates', range, apply: {} })).toEqual({
      sqref,
      type: 16,
    });
    expect(conditionalRuleToEngineInput({ kind: 'unique', range, apply: {} })).toEqual({
      sqref,
      type: 17,
    });
  });

  it('translates top-bottom and average rules', () => {
    expect(
      conditionalRuleToEngineInput({
        kind: 'top-bottom',
        range,
        mode: 'bottom',
        n: 5,
        percent: true,
        apply: {},
      }),
    ).toEqual({ sqref, type: 5, rank: 5, percent: true, bottom: true });
    expect(
      conditionalRuleToEngineInput({
        kind: 'average',
        range,
        mode: 'equal-or-above',
        apply: {},
      }),
    ).toEqual({ sqref, type: 6, aboveAverage: true, equalAverage: true });
    expect(
      conditionalRuleToEngineInput({
        kind: 'average',
        range,
        mode: 'below-std-dev',
        stdDev: 2,
        apply: {},
      }),
    ).toEqual({ sqref, type: 6, aboveAverage: false, equalAverage: false, stdDev: 2 });
  });

  it('translates visual rules into engine payloads', () => {
    expect(
      conditionalRuleToEngineInput({
        kind: 'color-scale',
        range,
        stops: ['#fff', '#000'],
        thresholds: [{ kind: 'min' }, { kind: 'max' }],
      }),
    ).toEqual({
      sqref,
      type: 2,
      colorScale: {
        thresholds: [{ type: 3 }, { type: 4 }],
        colors: [
          { a: 255, r: 255, g: 255, b: 255 },
          { a: 255, r: 0, g: 0, b: 0 },
        ],
      },
    });
    expect(
      conditionalRuleToEngineInput({
        kind: 'data-bar',
        range,
        color: '#0078d4',
        showValue: false,
      }),
    ).toEqual({
      sqref,
      type: 3,
      dataBar: {
        min: { type: 3 },
        max: { type: 4 },
        fill: { a: 255, r: 0, g: 120, b: 212 },
        showValue: false,
      },
    });
    expect(
      conditionalRuleToEngineInput({
        kind: 'icon-set',
        range,
        icons: 'traffic3',
        thresholds: [
          { kind: 'percent', value: 33 },
          { kind: 'percent', value: 67 },
        ],
        reverseOrder: true,
        showValue: false,
      }),
    ).toEqual({
      sqref,
      type: 4,
      iconSet: {
        name: 3,
        thresholds: [
          { type: 1, value: '33' },
          { type: 1, value: '67' },
        ],
        reverse: true,
        showValue: false,
        percent: true,
      },
    });
  });

  it('writes data-bar direction and gradient plus an independent icon floor', () => {
    expect(
      conditionalRuleToEngineInput({
        kind: 'data-bar',
        range,
        color: '#0078d4',
        gradient: false,
        direction: 'right-to-left',
      }),
    ).toMatchObject({
      type: 3,
      dataBar: {
        gradient: false,
        direction: 2,
      },
    });
    expect(
      conditionalRuleToEngineInput({
        kind: 'icon-set',
        range,
        icons: 'traffic3',
        thresholds: [{ kind: 'percent', value: 67 }],
        floor: { kind: 'number', value: 10 },
      }),
    ).toMatchObject({
      type: 4,
      iconSet: {
        thresholds: [{ type: 1, value: '67' }],
        floor: { type: 0, value: '10' },
      },
    });
  });

  it('writes data-bar axis and sign-specific appearance fields', () => {
    expect(
      conditionalRuleToEngineInput({
        kind: 'data-bar',
        range,
        color: '#0078d4',
        axisPosition: 'middle',
        negativeColor: '#c00000',
        borderColor: '#1f1f1f',
        negativeBorderColor: '#7f0000',
        axisColor: '#404040',
      }),
    ).toMatchObject({
      type: 3,
      dataBar: {
        axisPosition: 1,
        fill: { a: 255, r: 0, g: 120, b: 212 },
        negativeFill: { a: 255, r: 192, g: 0, b: 0 },
        border: { a: 255, r: 31, g: 31, b: 31 },
        negativeBorder: { a: 255, r: 127, g: 0, b: 0 },
        axisColor: { a: 255, r: 64, g: 64, b: 64 },
      },
    });
  });

  it('writes custom data-bar endpoints with their scale-point kinds', () => {
    expect(
      conditionalRuleToEngineInput({
        kind: 'data-bar',
        range,
        color: '#0078d4',
        min: { kind: 'number', value: 20, gte: false },
        max: { kind: 'percentile', value: 80, gte: true },
        direction: 'right-to-left',
      }),
    ).toMatchObject({
      type: 3,
      dataBar: {
        min: { type: 0, value: '20', gte: false },
        max: { type: 2, value: '80', gte: true },
        direction: 2,
      },
    });
  });

  it('preserves strict icon floor and threshold comparisons', () => {
    expect(
      conditionalRuleToEngineInput({
        kind: 'icon-set',
        range,
        icons: 'traffic3',
        thresholds: [
          { kind: 'number', value: 20, gte: false },
          { kind: 'number', value: 40, gte: false },
        ],
        floor: { kind: 'number', value: 15, gte: false },
      }),
    ).toMatchObject({
      type: 4,
      iconSet: {
        thresholds: [
          { type: 0, value: '20', gte: false },
          { type: 0, value: '40', gte: false },
        ],
        floor: { type: 0, value: '15', gte: false },
      },
    });
  });

  it('returns null for date-occurring rules the engine cannot author', () => {
    expect(
      conditionalRuleToEngineInput({
        kind: 'date-occurring',
        range,
        period: 'today',
        apply: {},
      }),
    ).toBeNull();
  });
});

describe('syncConditionalRulesToEngine (H-13)', () => {
  const fakeWb = (mutate: boolean) => {
    const added: Array<{ sheet: number; type: number; dxfId?: number }> = [];
    const wb = {
      capabilities: { conditionalFormatMutate: mutate },
      addConditionalFormat: vi.fn((sheet: number, rule: { type: number; dxfId?: number }) => {
        added.push({
          sheet,
          type: rule.type,
          ...(rule.dxfId !== undefined ? { dxfId: rule.dxfId } : {}),
        });
        return added.length - 1;
      }),
    } as unknown as WorkbookHandle;
    return { wb, added };
  };

  it('writes representable rules and counts skips for the rest', () => {
    const { wb, added } = fakeWb(true);
    const rules: ConditionalRule[] = [
      { kind: 'cell-value', range, op: '>', a: 10, apply: {} },
      { kind: 'data-bar', range, color: '#0078d4' },
      { kind: 'text-contains', range, text: 'x', apply: {} },
      { kind: 'cell-value', range: { ...range, sheet: 1 }, op: '<', a: 0, apply: {} }, // other sheet
    ];
    const result = syncConditionalRulesToEngine(wb, rules, 0);
    expect(result).toEqual({ written: 3, skipped: 0 });
    expect(added).toEqual([
      { sheet: 0, type: 1 },
      { sheet: 0, type: 3 },
      { sheet: 0, type: 7 },
    ]);
  });

  it('is a no-op when the engine cannot author conditional formats', () => {
    const { wb, added } = fakeWb(false);
    const result = syncConditionalRulesToEngine(
      wb,
      [{ kind: 'cell-value', range, op: '>', a: 10, apply: {} }],
      0,
    );
    expect(result).toEqual({ written: 0, skipped: 0 });
    expect(added).toHaveLength(0);
  });

  it('can skip rules already synced by a mounted session', () => {
    const { wb, added } = fakeWb(true);
    const seenKeys = new Set<string>();
    const rules: ConditionalRule[] = [{ kind: 'cell-value', range, op: '>', a: 10, apply: {} }];

    expect(syncConditionalRulesToEngine(wb, rules, 0, { seenKeys })).toEqual({
      written: 1,
      skipped: 0,
    });
    expect(syncConditionalRulesToEngine(wb, rules, 0, { seenKeys })).toEqual({
      written: 0,
      skipped: 0,
    });
    expect(added).toHaveLength(1);
  });

  it('does not add rules hydrated from existing engine conditional formats', () => {
    const { wb, added } = fakeWb(true);
    const rules: ConditionalRule[] = [
      { engineId: 'imported', kind: 'cell-value', range, op: '>', a: 10, apply: {} },
      { kind: 'text-contains', range, text: 'x', apply: {} },
    ];

    expect(syncConditionalRulesToEngine(wb, rules, 0)).toEqual({ written: 1, skipped: 0 });
    expect(added).toEqual([{ sheet: 0, type: 7 }]);
  });

  it('does not add imported rules whose engine id is the empty sentinel', () => {
    const { wb, added } = fakeWb(true);
    const rules: ConditionalRule[] = [
      { engineId: '', kind: 'cell-value', range, op: '>', a: 10, apply: {} },
      { kind: 'text-contains', range, text: 'x', apply: {} },
    ];

    expect(syncConditionalRulesToEngine(wb, rules, 0)).toEqual({ written: 1, skipped: 0 });
    expect(added).toEqual([{ sheet: 0, type: 7 }]);
  });

  it('persists apply formatting through a dxfId when the engine supports dxfs', () => {
    const addedDxf: unknown[] = [];
    const addedRules: unknown[] = [];
    const wb = {
      capabilities: { conditionalFormatMutate: true, conditionalFormatDxf: true },
      addDxf: vi.fn((record: unknown) => {
        addedDxf.push(record);
        return 7;
      }),
      addConditionalFormat: vi.fn((_: number, rule: unknown) => {
        addedRules.push(rule);
        return addedRules.length - 1;
      }),
    } as unknown as WorkbookHandle;

    const result = syncConditionalRulesToEngine(
      wb,
      [
        {
          kind: 'cell-value',
          range,
          op: '>',
          a: 10,
          apply: { fill: '#e2f0d9', color: '#006100', bold: true },
        },
      ],
      0,
    );

    expect(result).toEqual({ written: 1, skipped: 0 });
    expect(addedRules[0]).toMatchObject({ type: 1, op: 5, formula1: '10', dxfId: 7 });
    expect(addedDxf[0]).toMatchObject({
      fill: { pattern: 1, fgArgb: 0xffe2f0d9 },
      font: { bold: true, colorArgb: 0xff006100 },
    });
  });
});

describe('syncTrackedConditionalRulesToEngine (H-13)', () => {
  const trackedWb = () => {
    const formats: Array<{
      id: string;
      type: number;
      priority: number;
      stopIfTrue: boolean;
      sqref: typeof sqref;
      formula1?: string;
      formula2?: string;
      op?: number;
      text?: string;
      dxfId?: number;
    }> = [
      {
        id: 'imported',
        type: 1,
        priority: 1,
        stopIfTrue: false,
        sqref,
        op: 0,
        formula1: '0',
      },
    ];
    const wb = {
      capabilities: { conditionalFormatMutate: true, conditionalFormatDxf: true },
      addDxf: vi.fn(() => formats.length + 10),
      getConditionalFormats: vi.fn(() => formats.map((entry) => ({ ...entry }))),
      addConditionalFormat: vi.fn(
        (
          _: number,
          rule: { type: number; op?: number; formula1?: string; text?: string; dxfId?: number },
        ) => {
          formats.push({
            id: `added-${formats.length}`,
            type: rule.type,
            priority: formats.length + 1,
            stopIfTrue: false,
            sqref,
            ...(rule.op !== undefined ? { op: rule.op } : {}),
            ...(rule.formula1 !== undefined ? { formula1: rule.formula1 } : {}),
            ...(rule.text !== undefined ? { text: rule.text } : {}),
            ...(rule.dxfId !== undefined ? { dxfId: rule.dxfId } : {}),
          });
          return formats.length - 1;
        },
      ),
      removeConditionalFormatAt: vi.fn((_: number, index: number) => {
        if (index < 0 || index >= formats.length) return false;
        formats.splice(index, 1);
        return true;
      }),
    } as unknown as WorkbookHandle;
    return { wb, formats };
  };

  it('removes stale tracked rules by stable id after engine payload normalization', () => {
    const rule: ConditionalRule = {
      kind: 'data-bar',
      range,
      color: '#0078d4',
      min: { kind: 'number', value: 20 },
      max: { kind: 'number', value: 80 },
      direction: 'right-to-left',
    };
    const input = conditionalRuleToEngineInput(rule);
    if (!input) throw new Error('Expected data-bar input');
    const formats = [
      {
        id: 'session-data-bar',
        type: 3,
        priority: 1,
        stopIfTrue: false,
        sqref,
        dataBar: {
          min: { type: 0, value: '20', gte: true },
          max: { type: 0, value: '80', gte: true },
          fill: { a: 255, r: 0, g: 120, b: 212 },
          showValue: true,
          minLengthPct: 10,
          maxLengthPct: 90,
          gradient: true,
          direction: 2,
        },
      },
    ];
    const removeConditionalFormatAt = vi.fn((_: number, index: number) => {
      if (index < 0 || index >= formats.length) return false;
      formats.splice(index, 1);
      return true;
    });
    const wb = {
      capabilities: { conditionalFormatMutate: true },
      getConditionalFormats: vi.fn(() => formats.slice()),
      removeConditionalFormatAt,
    } as unknown as WorkbookHandle;
    const tracked: SyncedConditionalRuleMap = new Map([
      ['stale-key', { sheet: 0, index: 0, id: 'session-data-bar', input }],
    ]);

    expect(syncTrackedConditionalRulesToEngine(wb, [], 0, { tracked })).toEqual({
      written: 0,
      skipped: 0,
      removed: 1,
    });
    expect(removeConditionalFormatAt).toHaveBeenCalledWith(0, 0);
    expect(formats).toHaveLength(0);
  });

  it('does not remove an untracked rule when a tracked entry has no stable id', () => {
    const rule: ConditionalRule = {
      kind: 'cell-value',
      range,
      op: '>',
      a: 10,
      apply: {},
    };
    const input = conditionalRuleToEngineInput(rule);
    if (!input) throw new Error('Expected cell-value input');
    const formats = [
      {
        id: 'untracked',
        type: 1,
        priority: 1,
        stopIfTrue: false,
        sqref,
        op: 5,
        formula1: '10',
      },
    ];
    const removeConditionalFormatAt = vi.fn(() => true);
    const wb = {
      capabilities: { conditionalFormatMutate: true },
      getConditionalFormats: vi.fn(() => formats.slice()),
      removeConditionalFormatAt,
    } as unknown as WorkbookHandle;
    const tracked: SyncedConditionalRuleMap = new Map([
      ['stale-key', { sheet: 0, index: 0, input }],
    ]);

    expect(syncTrackedConditionalRulesToEngine(wb, [], 0, { tracked })).toEqual({
      written: 0,
      skipped: 1,
      removed: 0,
    });
    expect(removeConditionalFormatAt).not.toHaveBeenCalled();
    expect(formats).toHaveLength(1);
  });

  it('removes only session-tracked engine rules that disappeared from the store', () => {
    const { wb, formats } = trackedWb();
    const tracked: SyncedConditionalRuleMap = new Map();
    const rules: ConditionalRule[] = [
      { kind: 'cell-value', range, op: '>', a: 10, apply: {} },
      { kind: 'text-contains', range, text: 'x', apply: {} },
    ];
    const remainingRules = rules.slice(1);

    expect(syncTrackedConditionalRulesToEngine(wb, rules, 0, { tracked })).toEqual({
      written: 2,
      skipped: 0,
      removed: 0,
    });
    expect(formats.map((entry) => entry.id)).toEqual(['imported', 'added-1', 'added-2']);

    expect(syncTrackedConditionalRulesToEngine(wb, remainingRules, 0, { tracked })).toEqual({
      written: 0,
      skipped: 0,
      removed: 1,
    });
    expect(formats.map((entry) => entry.id)).toEqual(['imported', 'added-2']);
    expect(wb.removeConditionalFormatAt).toHaveBeenCalledWith(0, 1);
  });

  it('clears session-tracked engine rules while leaving imported rules intact', () => {
    const { wb, formats } = trackedWb();
    const tracked: SyncedConditionalRuleMap = new Map();
    const rules: ConditionalRule[] = [
      { kind: 'cell-value', range, op: '>', a: 10, apply: {} },
      { kind: 'text-contains', range, text: 'x', apply: {} },
    ];

    syncTrackedConditionalRulesToEngine(wb, rules, 0, { tracked });
    expect(syncTrackedConditionalRulesToEngine(wb, [], 0, { tracked })).toEqual({
      written: 0,
      skipped: 0,
      removed: 2,
    });
    expect(formats.map((entry) => entry.id)).toEqual(['imported']);
    expect(wb.clearConditionalFormats).toBeUndefined();
  });

  it('uses the read-back engine id when a tracked index shifts before removal', () => {
    const { wb, formats } = trackedWb();
    const tracked: SyncedConditionalRuleMap = new Map();
    const rules: ConditionalRule[] = [
      { kind: 'cell-value', range, op: '>', a: 10, apply: {} },
      { kind: 'text-contains', range, text: 'x', apply: {} },
    ];
    const remainingRules = rules.slice(1);

    syncTrackedConditionalRulesToEngine(wb, rules, 0, { tracked });
    formats.unshift({
      id: 'external-prepend',
      type: 1,
      priority: 0,
      stopIfTrue: false,
      sqref,
      op: 2,
      formula1: '999',
    });

    expect(syncTrackedConditionalRulesToEngine(wb, remainingRules, 0, { tracked })).toEqual({
      written: 0,
      skipped: 0,
      removed: 1,
    });
    expect(formats.map((entry) => entry.id)).toEqual(['external-prepend', 'imported', 'added-2']);
    expect(wb.removeConditionalFormatAt).toHaveBeenCalledWith(0, 2);
  });

  it('does not track or duplicate imported engine-hydrated store rules', () => {
    const { wb, formats } = trackedWb();
    const tracked: SyncedConditionalRuleMap = new Map();
    const rules: ConditionalRule[] = [
      { engineId: '', kind: 'cell-value', range, op: '<', a: 0, apply: {} },
      { kind: 'text-contains', range, text: 'x', apply: {} },
    ];

    expect(syncTrackedConditionalRulesToEngine(wb, rules, 0, { tracked })).toEqual({
      written: 1,
      skipped: 0,
      removed: 0,
    });
    expect(formats.map((entry) => entry.id)).toEqual(['imported', 'added-1']);
    expect(tracked.size).toBe(1);
  });

  it('attaches dxf only when adding a new tracked rule', () => {
    const { wb, formats } = trackedWb();
    const tracked: SyncedConditionalRuleMap = new Map();
    const rules: ConditionalRule[] = [
      { kind: 'cell-value', range, op: '>', a: 10, apply: { fill: '#e2f0d9' } },
    ];

    expect(syncTrackedConditionalRulesToEngine(wb, rules, 0, { tracked })).toEqual({
      written: 1,
      skipped: 0,
      removed: 0,
    });
    expect(formats[1]).toMatchObject({ id: 'added-1', dxfId: 11 });
    expect(wb.addDxf).toHaveBeenCalledTimes(1);

    expect(syncTrackedConditionalRulesToEngine(wb, rules, 0, { tracked })).toEqual({
      written: 0,
      skipped: 0,
      removed: 0,
    });
    expect(wb.addDxf).toHaveBeenCalledTimes(1);
  });
});
