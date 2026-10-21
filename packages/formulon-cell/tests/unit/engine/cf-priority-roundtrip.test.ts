import { describe, expect, it } from 'vitest';
import { hydrateConditionalRulesFromEngine } from '../../../src/engine/cf-sync.js';
import {
  rewriteEngineConditionalFormat,
  type SyncedConditionalRuleMap,
  syncTrackedConditionalRulesToEngine,
} from '../../../src/engine/cf-writeback.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import type { ConditionalRule } from '../../../src/store/types.js';

const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };
const sqref = [{ firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 0 }];

const blanks = (stopIfTrue?: boolean): ConditionalRule => ({
  kind: 'blanks',
  range,
  apply: { bold: true },
  ...(stopIfTrue ? { stopIfTrue } : {}),
});
const dupes: ConditionalRule = { kind: 'duplicates', range, apply: { italic: true } };

const summary = (wb: WorkbookHandle) =>
  wb
    .getConditionalFormats(0)
    .map((e) => ({ type: e.type, priority: e.priority, stop: e.stopIfTrue }));

describe('conditional formatting priority and stopIfTrue persistence', () => {
  it('writes stopIfTrue and store-order priorities for session rules', async () => {
    const wb = await WorkbookHandle.createDefault();
    let loaded: WorkbookHandle | null = null;
    try {
      const tracked: SyncedConditionalRuleMap = new Map();
      syncTrackedConditionalRulesToEngine(wb, [blanks(true), dupes], 0, { tracked });
      const written = summary(wb);
      expect(written).toEqual([
        { type: 11, priority: 1, stop: true },
        { type: 16, priority: 2, stop: false },
      ]);
      loaded = await WorkbookHandle.loadBytes(wb.save());
      expect(summary(loaded)).toEqual(written);
    } finally {
      loaded?.dispose();
      wb.dispose();
    }
  });

  it('reordering the store reorders the engine priorities', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const tracked: SyncedConditionalRuleMap = new Map();
      syncTrackedConditionalRulesToEngine(wb, [blanks(), dupes], 0, { tracked });
      syncTrackedConditionalRulesToEngine(wb, [dupes, blanks()], 0, { tracked });
      const byPriority = [...summary(wb)].sort((a, b) => a.priority - b.priority);
      expect(byPriority.map((e) => e.type)).toEqual([16, 11]);
      expect(summary(wb)).toHaveLength(2);
    } finally {
      wb.dispose();
    }
  });

  it('ranks session rules after rules the mount did not write', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.addConditionalFormat(0, { sqref, type: 17, priority: 7 })).toBeGreaterThanOrEqual(
        0,
      );
      syncTrackedConditionalRulesToEngine(wb, [blanks()], 0, { tracked: new Map() });
      expect(summary(wb).map((e) => [e.type, e.priority])).toEqual([
        [17, 7],
        [11, 8],
      ]);
    } finally {
      wb.dispose();
    }
  });

  it('rewrites a loaded rule keeping its id, priority and stopIfTrue', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      wb.addConditionalFormat(0, { sqref, type: 17, priority: 3, stopIfTrue: true });
      wb.addConditionalFormat(0, { sqref, type: 16, priority: 4 });
      const [first] = wb.getConditionalFormats(0);
      if (!first) throw new Error('missing entry');
      expect(
        rewriteEngineConditionalFormat(wb, 0, first.id, {
          input: { sqref, type: 11 },
        }),
      ).toBe(true);
      const after = wb.getConditionalFormats(0).find((e) => e.id === first.id);
      expect(after).toMatchObject({ type: 11, priority: 3, stopIfTrue: true });

      expect(rewriteEngineConditionalFormat(wb, 0, first.id, { priority: 9 })).toBe(true);
      expect(wb.getConditionalFormats(0).find((e) => e.id === first.id)).toMatchObject({
        type: 11,
        priority: 9,
        stopIfTrue: true,
      });
    } finally {
      wb.dispose();
    }
  });

  it('round-trips a date-occurring rule through the file', async () => {
    const wb = await WorkbookHandle.createDefault();
    let loaded: WorkbookHandle | null = null;
    try {
      const rule: ConditionalRule = {
        kind: 'date-occurring',
        range,
        period: 'last-week',
        apply: { bold: true },
      };
      expect(syncTrackedConditionalRulesToEngine(wb, [rule], 0, { tracked: new Map() })).toEqual({
        written: 1,
        skipped: 0,
        removed: 0,
      });
      loaded = await WorkbookHandle.loadBytes(wb.save());
      const store = createSpreadsheetStore();
      hydrateConditionalRulesFromEngine(loaded, store, 0);
      expect(store.getState().conditional.rules).toMatchObject([
        { kind: 'date-occurring', period: 'last-week', engineId: expect.any(String) },
      ]);
    } finally {
      loaded?.dispose();
      wb.dispose();
    }
  });
});
