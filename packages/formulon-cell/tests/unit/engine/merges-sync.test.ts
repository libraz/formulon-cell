import { beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { applyMerge, applyUnmerge } from '../../../src/commands/merge.js';
import { hydrateMergesFromEngine } from '../../../src/engine/merges-sync.js';
import type { Range } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

interface FakeMergeRange {
  firstRow: number;
  firstCol: number;
  lastRow: number;
  lastCol: number;
}

interface MergeCalls {
  add: { sheet: number; range: FakeMergeRange }[];
  remove: { sheet: number; range: FakeMergeRange }[];
  clear: number[];
}

/** Stand-in WorkbookHandle that simulates `capabilities.merges = true` and
 *  records every merge engine call. Keeps an in-memory store so `getMerges`
 *  reflects the running state. */
const makeFake = (
  opts: { merges?: boolean; initial?: Range[] } = {},
): { wb: WorkbookHandle; calls: MergeCalls } => {
  const calls: MergeCalls = { add: [], remove: [], clear: [] };
  const caps = { merges: opts.merges ?? true };
  const live = new Map<number, FakeMergeRange[]>();
  for (const r of opts.initial ?? []) {
    if (!live.has(r.sheet)) live.set(r.sheet, []);
    live.get(r.sheet)?.push({
      firstRow: r.r0,
      firstCol: r.c0,
      lastRow: r.r1,
      lastCol: r.c1,
    });
  }
  const fake = {
    capabilities: caps,
    engineAddMerge(sheet: number, range: Range): boolean {
      if (!caps.merges) return false;
      const m: FakeMergeRange = {
        firstRow: range.r0,
        firstCol: range.c0,
        lastRow: range.r1,
        lastCol: range.c1,
      };
      calls.add.push({ sheet, range: m });
      if (!live.has(sheet)) live.set(sheet, []);
      live.get(sheet)?.push(m);
      return true;
    },
    engineRemoveMerge(sheet: number, range: Range): boolean {
      if (!caps.merges) return false;
      const m: FakeMergeRange = {
        firstRow: range.r0,
        firstCol: range.c0,
        lastRow: range.r1,
        lastCol: range.c1,
      };
      calls.remove.push({ sheet, range: m });
      // Remove every entry overlapping the range — same semantics as the
      // upstream removeMerge call.
      const list = live.get(sheet) ?? [];
      live.set(
        sheet,
        list.filter(
          (e) =>
            e.lastRow < m.firstRow ||
            e.firstRow > m.lastRow ||
            e.lastCol < m.firstCol ||
            e.firstCol > m.lastCol,
        ),
      );
      return true;
    },
    engineClearMerges(sheet: number): boolean {
      if (!caps.merges) return false;
      calls.clear.push(sheet);
      live.set(sheet, []);
      return true;
    },
    getMerges(sheet: number): Range[] {
      if (!caps.merges) return [];
      return (live.get(sheet) ?? []).map((m) => ({
        sheet,
        r0: m.firstRow,
        c0: m.firstCol,
        r1: m.lastRow,
        c1: m.lastCol,
      }));
    },
  };
  return { wb: fake as unknown as WorkbookHandle, calls };
};

const fakeSetBlank = (): void => {};
const wbWithBlankWriter = (wb: WorkbookHandle): WorkbookHandle => {
  // applyMerge calls wb.setBlank for non-anchor cells; the fake doesn't need
  // engine-side blanks, so just stub the method.
  (wb as unknown as { setBlank: typeof fakeSetBlank }).setBlank = fakeSetBlank;
  return wb;
};

describe('engine merge wrappers via commands/merge.ts', () => {
  let store: SpreadsheetStore;
  let history: History;

  beforeEach(() => {
    store = createSpreadsheetStore();
    history = new History();
  });

  it('applyMerge clears then re-adds the post-mutation set on the engine', () => {
    const { wb, calls } = makeFake({ merges: true });
    wbWithBlankWriter(wb);
    const r: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    expect(applyMerge(store, wb, history, r)).toBe(true);
    // recordMergesChangeWithEngine syncs by clearing then re-adding every
    // anchor — apply emits exactly one clear and one add.
    expect(calls.clear).toEqual([0]);
    expect(calls.add).toEqual([
      { sheet: 0, range: { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 } },
    ]);
  });

  it('applyUnmerge clears the engine when the last merge is dropped', () => {
    const { wb, calls } = makeFake({ merges: true });
    wbWithBlankWriter(wb);
    const r: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    applyMerge(store, wb, history, r);
    calls.clear.length = 0;
    calls.add.length = 0;
    expect(applyUnmerge(store, wb, history, r)).toBe(true);
    // Unmerge: clear (always), then no adds since the slice is empty.
    expect(calls.clear).toEqual([0]);
    expect(calls.add).toEqual([]);
  });

  it('undo / redo round-trips engine state', () => {
    const { wb, calls } = makeFake({ merges: true });
    wbWithBlankWriter(wb);
    const r: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    applyMerge(store, wb, history, r);
    calls.clear.length = 0;
    calls.add.length = 0;
    history.undo();
    expect(calls.clear).toEqual([0]);
    expect(calls.add).toEqual([]);
    calls.clear.length = 0;
    history.redo();
    expect(calls.clear).toEqual([0]);
    expect(calls.add).toEqual([
      { sheet: 0, range: { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 } },
    ]);
  });

  it('capability off: applyMerge does NOT touch the engine', () => {
    const { wb, calls } = makeFake({ merges: false });
    wbWithBlankWriter(wb);
    const r: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    applyMerge(store, wb, history, r);
    expect(calls.clear).toEqual([]);
    expect(calls.add).toEqual([]);
    // Store still mutates — JS-only path stays intact.
    expect(store.getState().merges.byAnchor.size).toBe(1);
  });
});

describe('hydrateMergesFromEngine', () => {
  it('seeds the store from wb.getMerges when capability is on', () => {
    const initial: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 2 };
    const { wb } = makeFake({ merges: true, initial: [initial] });
    const store = createSpreadsheetStore();
    hydrateMergesFromEngine(wb, store, 0);
    const anchors = Array.from(store.getState().merges.byAnchor.values());
    expect(anchors).toEqual([initial]);
  });

  it('replaces existing merges on the same sheet (without touching others)', () => {
    const { wb } = makeFake({
      merges: true,
      initial: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 }],
    });
    const store = createSpreadsheetStore();
    // Seed a stale merge on sheet 0 and one on sheet 1.
    store.setState((s) => ({
      ...s,
      merges: {
        byAnchor: new Map([
          ['0:5:5', { sheet: 0, r0: 5, c0: 5, r1: 5, c1: 6 }],
          ['1:0:0', { sheet: 1, r0: 0, c0: 0, r1: 0, c1: 1 }],
        ]),
        byCell: new Map(),
      },
    }));
    hydrateMergesFromEngine(wb, store, 0);
    const anchors = Array.from(store.getState().merges.byAnchor.values());
    // Sheet 0 should now match the engine; sheet 1 untouched.
    expect(anchors.filter((r) => r.sheet === 0)).toEqual([
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
    ]);
    expect(anchors.filter((r) => r.sheet === 1)).toEqual([
      { sheet: 1, r0: 0, c0: 0, r1: 0, c1: 1 },
    ]);
  });

  it('no-op when capability is off (store keeps its prior state)', () => {
    const { wb } = makeFake({
      merges: false,
      initial: [{ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 }],
    });
    const store = createSpreadsheetStore();
    hydrateMergesFromEngine(wb, store, 0);
    expect(store.getState().merges.byAnchor.size).toBe(0);
  });
});
