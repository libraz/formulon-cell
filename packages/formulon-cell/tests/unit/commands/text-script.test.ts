import { describe, expect, it, vi } from 'vitest';
import { applyTextScriptToRange } from '../../../src/commands/text-script.js';
import type { Addr } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const key = (addr: Addr): string => `${addr.sheet}:${addr.row}:${addr.col}`;

const fakeWorkbook = (writes: Map<string, string | null>): WorkbookHandle =>
  ({
    withBatchedRecalc: <T>(fn: () => T): T => fn(),
    setText: (addr: Addr, value: string) => writes.set(key(addr), value),
    setBlank: (addr: Addr) => writes.set(key(addr), null),
  }) as unknown as WorkbookHandle;

describe('commands/text-script', () => {
  it('applies text transforms only to text cells in the target range', () => {
    const store = createSpreadsheetStore();
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'text', value: ' alpha ' }, null);
    mutators.setCell(store, { sheet: 0, row: 0, col: 1 }, { kind: 'number', value: 12 }, null);
    const writes = new Map<string, string | null>();

    const count = applyTextScriptToRange(
      store.getState(),
      fakeWorkbook(writes),
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      'trim',
    );

    expect(count).toBe(1);
    expect(writes).toEqual(new Map([['0:0:0', 'alpha']]));
  });

  it('applies lowercase and uppercase transforms across the selected text range', () => {
    const store = createSpreadsheetStore();
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'text', value: 'MiXeD' }, null);
    mutators.setCell(store, { sheet: 0, row: 0, col: 1 }, { kind: 'text', value: 'second' }, null);
    const writes = new Map<string, string | null>();

    expect(
      applyTextScriptToRange(
        store.getState(),
        fakeWorkbook(writes),
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
        'lowercase',
      ),
    ).toBe(1);
    expect(writes).toEqual(new Map([['0:0:0', 'mixed']]));

    writes.clear();
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'text', value: 'mixed' }, null);
    expect(
      applyTextScriptToRange(
        store.getState(),
        fakeWorkbook(writes),
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
        'uppercase',
      ),
    ).toBe(2);
    expect(writes).toEqual(
      new Map([
        ['0:0:0', 'MIXED'],
        ['0:0:1', 'SECOND'],
      ]),
    );
  });

  it('clears all populated cells while respecting protected locked cells', () => {
    const store = createSpreadsheetStore();
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'text', value: 'locked' }, null);
    mutators.setCell(store, { sheet: 0, row: 0, col: 1 }, { kind: 'text', value: 'open' }, null);
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 1 }, { locked: false });
    mutators.setSheetProtected(store, 0, true);
    const writes = new Map<string, string | null>();

    const count = applyTextScriptToRange(
      store.getState(),
      fakeWorkbook(writes),
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      'clear',
    );

    expect(count).toBe(1);
    expect(writes).toEqual(new Map([['0:0:1', null]]));
  });

  it('applies huge whole-column text scripts by visiting only materialized cells', () => {
    const store = createSpreadsheetStore();
    mutators.setCell(store, { sheet: 0, row: 8, col: 2 }, { kind: 'text', value: ' Mixed ' }, null);
    mutators.setCell(store, { sheet: 0, row: 8, col: 3 }, { kind: 'text', value: 'Outside' }, null);
    const writes = new Map<string, string | null>();

    const count = applyTextScriptToRange(
      store.getState(),
      fakeWorkbook(writes),
      { sheet: 0, r0: 0, c0: 2, r1: 1048575, c1: 2 },
      'trim',
    );

    expect(count).toBe(1);
    expect(writes).toEqual(new Map([['0:8:2', 'Mixed']]));
  });

  it('wraps every write in one batched recalc', () => {
    const store = createSpreadsheetStore();
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'text', value: 'a' }, null);
    mutators.setCell(store, { sheet: 0, row: 0, col: 1 }, { kind: 'text', value: 'b' }, null);
    const writes = new Map<string, string | null>();
    const wb = fakeWorkbook(writes);
    const batched = vi.fn(<T>(fn: () => T): T => {
      expect(writes.size).toBe(0);
      return fn();
    });
    (wb as unknown as { withBatchedRecalc: typeof batched }).withBatchedRecalc = batched;

    applyTextScriptToRange(
      store.getState(),
      wb,
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      'uppercase',
    );

    expect(batched).toHaveBeenCalledTimes(1);
    expect(writes.size).toBe(2);
  });
});
