import { beforeEach, describe, expect, it } from 'vitest';
import { captureSnapshot } from '../../../../src/commands/clipboard/snapshot.js';
import { addrKey } from '../../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';

describe('captureSnapshot', () => {
  let store: SpreadsheetStore;

  beforeEach(() => {
    store = createSpreadsheetStore();
  });

  it('returns null for an inverted range', () => {
    expect(captureSnapshot(store.getState(), { sheet: 0, r0: 5, c0: 5, r1: 4, c1: 4 })).toBeNull();
  });

  it('returns null when the requested area exceeds the cap', () => {
    const huge = { sheet: 0, r0: 0, c0: 0, r1: 1_048_575, c1: 16_383 };
    expect(captureSnapshot(store.getState(), huge)).toBeNull();
  });

  it('captures values, formulas and formats', () => {
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      cells.set(addrKey({ sheet: 0, row: 0, col: 0 }), {
        value: { kind: 'number', value: 5 },
        formula: null,
      });
      cells.set(addrKey({ sheet: 0, row: 0, col: 1 }), {
        value: { kind: 'text', value: 'hi' },
        formula: '="hi"',
      });
      return { ...s, data: { ...s.data, cells } };
    });
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });

    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    expect(snap).not.toBeNull();
    expect(snap?.rows).toBe(1);
    expect(snap?.cols).toBe(2);
    expect(snap?.cells[0]?.[0]).toEqual({
      value: { kind: 'number', value: 5 },
      formula: null,
      format: { bold: true },
    });
    expect(snap?.cells[0]?.[1]).toMatchObject({
      value: { kind: 'text', value: 'hi' },
      formula: '="hi"',
    });
  });

  it('represents missing cells as blank entries with undefined format', () => {
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    expect(snap?.cells.flat()).toEqual([
      { value: { kind: 'blank' }, formula: null, format: undefined },
      { value: { kind: 'blank' }, formula: null, format: undefined },
      { value: { kind: 'blank' }, formula: null, format: undefined },
      { value: { kind: 'blank' }, formula: null, format: undefined },
    ]);
  });

  it('detaches borders from the live state', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { borders: { top: true, left: true } },
    );
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const captured = snap?.cells[0]?.[0]?.format?.borders;
    expect(captured).toEqual({ top: true, left: true });

    // Mutate live borders — the snapshot must not change.
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { borders: { top: false, right: true } },
    );
    expect(snap?.cells[0]?.[0]?.format?.borders).toEqual({ top: true, left: true });
  });
});
