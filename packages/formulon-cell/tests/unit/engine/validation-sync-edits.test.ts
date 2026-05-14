import { describe, expect, it } from 'vitest';

import { syncValidationsToEngine } from '../../../src/engine/validation-sync.js';
import { addrKey, type WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { type CellFormat, createSpreadsheetStore } from '../../../src/store/store.js';

interface AddedEntry {
  sheet: number;
  // The encoded engine payload shape (loose typing matches the sync's output).
  payload: Record<string, unknown>;
}

const makeEngine = (): {
  wb: WorkbookHandle;
  added: AddedEntry[];
  cleared: number[];
} => {
  const added: AddedEntry[] = [];
  const cleared: number[] = [];
  return {
    added,
    cleared,
    wb: {
      capabilities: { dataValidation: true } as never,
      clearValidations(sheet: number): boolean {
        cleared.push(sheet);
        return true;
      },
      addValidationEntry(sheet: number, payload: unknown): boolean {
        added.push({ sheet, payload: payload as Record<string, unknown> });
        return true;
      },
    } as unknown as WorkbookHandle,
  };
};

const setValidation = (
  store: ReturnType<typeof createSpreadsheetStore>,
  cells: { sheet: number; row: number; col: number; validation: CellFormat['validation'] }[],
) => {
  store.setState((s) => {
    const formats = new Map(s.format.formats);
    for (const c of cells) {
      const k = addrKey({ sheet: c.sheet, row: c.row, col: c.col });
      const prev = formats.get(k) ?? {};
      formats.set(k, { ...prev, validation: c.validation });
    }
    return { ...s, format: { formats } };
  });
};

const clearValidation = (
  store: ReturnType<typeof createSpreadsheetStore>,
  cells: { sheet: number; row: number; col: number }[],
) => {
  store.setState((s) => {
    const formats = new Map(s.format.formats);
    for (const c of cells) {
      const k = addrKey({ sheet: c.sheet, row: c.row, col: c.col });
      const prev = formats.get(k);
      if (!prev) continue;
      const { validation: _drop, ...rest } = prev;
      if (Object.keys(rest).length === 0) formats.delete(k);
      else formats.set(k, rest);
    }
    return { ...s, format: { formats } };
  });
};

describe('engine/validation-sync — edits', () => {
  it('adding a rule on a new cell pushes one entry to the engine', () => {
    const { wb, added, cleared } = makeEngine();
    const store = createSpreadsheetStore();
    setValidation(store, [
      {
        sheet: 0,
        row: 0,
        col: 0,
        validation: { kind: 'whole', op: 'between', a: 1, b: 10 },
      },
    ]);

    syncValidationsToEngine(wb, store, 0);
    expect(cleared).toEqual([0]);
    expect(added).toHaveLength(1);
    expect(added[0]?.payload).toMatchObject({
      type: 1, // DV_TYPE whole
      formula1: '1',
      formula2: '10',
    });
  });

  it('coalesces identical rules across cells into one entry with N ranges', () => {
    const { wb, added } = makeEngine();
    const store = createSpreadsheetStore();
    setValidation(store, [
      { sheet: 0, row: 0, col: 0, validation: { kind: 'list', source: ['A', 'B'] } },
      { sheet: 0, row: 0, col: 1, validation: { kind: 'list', source: ['A', 'B'] } },
      { sheet: 0, row: 1, col: 1, validation: { kind: 'list', source: ['A', 'B'] } },
    ]);

    syncValidationsToEngine(wb, store, 0);
    expect(added).toHaveLength(1);
    const ranges = (added[0]?.payload.ranges as unknown[]) ?? [];
    expect(ranges).toHaveLength(3);
  });

  it('splits distinct rules into separate entries', () => {
    const { wb, added } = makeEngine();
    const store = createSpreadsheetStore();
    setValidation(store, [
      { sheet: 0, row: 0, col: 0, validation: { kind: 'list', source: ['A'] } },
      { sheet: 0, row: 0, col: 1, validation: { kind: 'list', source: ['B'] } },
    ]);

    syncValidationsToEngine(wb, store, 0);
    expect(added).toHaveLength(2);
  });

  it('replacing a rule on an existing cell shows up as a single entry on the next sync', () => {
    const { wb, added } = makeEngine();
    const store = createSpreadsheetStore();
    setValidation(store, [
      { sheet: 0, row: 0, col: 0, validation: { kind: 'list', source: ['Yes', 'No'] } },
    ]);
    syncValidationsToEngine(wb, store, 0);
    added.length = 0;

    // Mutate to a different rule kind.
    setValidation(store, [
      {
        sheet: 0,
        row: 0,
        col: 0,
        validation: { kind: 'whole', op: '>', a: 0 },
      },
    ]);
    syncValidationsToEngine(wb, store, 0);

    expect(added).toHaveLength(1);
    expect(added[0]?.payload).toMatchObject({ type: 1, formula1: '0' });
  });

  it('removing the validation field drops the entry on the next sync', () => {
    const { wb, added } = makeEngine();
    const store = createSpreadsheetStore();
    setValidation(store, [
      { sheet: 0, row: 0, col: 0, validation: { kind: 'list', source: ['A', 'B'] } },
    ]);
    syncValidationsToEngine(wb, store, 0);
    added.length = 0;

    clearValidation(store, [{ sheet: 0, row: 0, col: 0 }]);
    syncValidationsToEngine(wb, store, 0);

    // clearValidations is still invoked (always), but no new entry was added.
    expect(added).toHaveLength(0);
  });

  it('extending the range of a rule emits more ranges next sync', () => {
    const { wb, added } = makeEngine();
    const store = createSpreadsheetStore();
    setValidation(store, [
      { sheet: 0, row: 0, col: 0, validation: { kind: 'list', source: ['A'] } },
    ]);
    syncValidationsToEngine(wb, store, 0);
    added.length = 0;

    // Apply the same rule to two more cells.
    setValidation(store, [
      { sheet: 0, row: 0, col: 1, validation: { kind: 'list', source: ['A'] } },
      { sheet: 1, row: 0, col: 0, validation: { kind: 'list', source: ['A'] } },
    ]);
    syncValidationsToEngine(wb, store, 0);

    expect(added).toHaveLength(1);
    const ranges = added[0]?.payload.ranges as unknown[];
    expect(ranges).toHaveLength(2); // sheet 0 row 0 col 0 + sheet 0 row 0 col 1
  });

  it('mutations on other sheets do not push to this sheet', () => {
    const { wb, added } = makeEngine();
    const store = createSpreadsheetStore();
    setValidation(store, [
      { sheet: 1, row: 0, col: 0, validation: { kind: 'list', source: ['A'] } },
    ]);
    syncValidationsToEngine(wb, store, 0);
    expect(added).toHaveLength(0);
  });

  it('partial mutation (a→b only on between) regenerates the formula payload', () => {
    const { wb, added } = makeEngine();
    const store = createSpreadsheetStore();
    setValidation(store, [
      {
        sheet: 0,
        row: 0,
        col: 0,
        validation: { kind: 'whole', op: 'between', a: 1, b: 10 },
      },
    ]);
    syncValidationsToEngine(wb, store, 0);
    added.length = 0;

    setValidation(store, [
      {
        sheet: 0,
        row: 0,
        col: 0,
        validation: { kind: 'whole', op: 'between', a: 1, b: 50 },
      },
    ]);
    syncValidationsToEngine(wb, store, 0);

    expect(added).toHaveLength(1);
    expect(added[0]?.payload).toMatchObject({ formula1: '1', formula2: '50' });
  });

  it('capability off: no engine calls regardless of store mutations', () => {
    const { added, cleared } = makeEngine();
    const wb = {
      capabilities: { dataValidation: false } as never,
      clearValidations: () => true,
      addValidationEntry: () => true,
    } as unknown as WorkbookHandle;

    const store = createSpreadsheetStore();
    setValidation(store, [
      { sheet: 0, row: 0, col: 0, validation: { kind: 'list', source: ['A'] } },
    ]);
    syncValidationsToEngine(wb, store, 0);

    expect(added).toEqual([]);
    expect(cleared).toEqual([]);
  });

  it('error message / errorStyle changes round through encodeMeta on each sync', () => {
    const { wb, added } = makeEngine();
    const store = createSpreadsheetStore();
    setValidation(store, [
      {
        sheet: 0,
        row: 0,
        col: 0,
        validation: {
          kind: 'list',
          source: ['A'],
          errorStyle: 'warning',
          errorMessage: 'pick from list',
        },
      },
    ]);
    syncValidationsToEngine(wb, store, 0);
    expect(added[0]?.payload).toMatchObject({
      errorStyle: 1, // warning
      errorMessage: 'pick from list',
    });

    added.length = 0;
    setValidation(store, [
      {
        sheet: 0,
        row: 0,
        col: 0,
        validation: {
          kind: 'list',
          source: ['A'],
          errorStyle: 'stop',
          errorMessage: 'no',
        },
      },
    ]);
    syncValidationsToEngine(wb, store, 0);
    expect(added[0]?.payload).toMatchObject({
      errorStyle: 0, // stop
      errorMessage: 'no',
    });
  });
});
