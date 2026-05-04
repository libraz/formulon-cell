import { describe, expect, it } from 'vitest';
import type { Range } from '../../../src/engine/types.js';
import {
  hydrateValidationsFromEngine,
  syncValidationsToEngine,
} from '../../../src/engine/validation-sync.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const DV_TYPE_LIST = 3;
const DV_TYPE_WHOLE = 1;

interface FakeValidationEntry {
  ranges: Range[];
  type: number;
  op: number;
  errorStyle: number;
  allowBlank: boolean;
  showInputMessage: boolean;
  showErrorMessage: boolean;
  formula1: string;
  formula2: string;
  errorTitle: string;
  errorMessage: string;
  promptTitle: string;
  promptMessage: string;
}

const partial = (
  p: Partial<FakeValidationEntry> & Pick<FakeValidationEntry, 'type' | 'ranges'>,
): FakeValidationEntry => ({
  op: 0,
  errorStyle: 0,
  allowBlank: true,
  showInputMessage: false,
  showErrorMessage: false,
  formula1: '',
  formula2: '',
  errorTitle: '',
  errorMessage: '',
  promptTitle: '',
  promptMessage: '',
  ...p,
});

const makeFake = (
  opts: {
    dataValidation?: boolean;
    entries?: FakeValidationEntry[];
    onAdd?: (sheet: number, input: unknown) => boolean;
    onClear?: (sheet: number) => boolean;
  } = {},
): WorkbookHandle => {
  const caps = { dataValidation: opts.dataValidation ?? true };
  const fake = {
    capabilities: caps,
    getValidationsForSheet(_sheet: number): FakeValidationEntry[] {
      if (!caps.dataValidation) return [];
      return opts.entries ?? [];
    },
    addValidationEntry(sheet: number, input: unknown): boolean {
      return opts.onAdd ? opts.onAdd(sheet, input) : true;
    },
    clearValidations(sheet: number): boolean {
      return opts.onClear ? opts.onClear(sheet) : true;
    },
  };
  return fake as unknown as WorkbookHandle;
};

describe('hydrateValidationsFromEngine', () => {
  it('seeds format.validation for cells inside a list-type range', () => {
    const wb = makeFake({
      entries: [
        partial({
          ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 }],
          type: DV_TYPE_LIST,
          formula1: '"Yes,No,Maybe"',
        }),
      ],
    });
    const store = createSpreadsheetStore();
    hydrateValidationsFromEngine(wb, store, 0);
    const a1 = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    const a2 = store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 0 }));
    expect(a1?.validation).toEqual({ kind: 'list', source: ['Yes', 'No', 'Maybe'] });
    expect(a2?.validation).toEqual({ kind: 'list', source: ['Yes', 'No', 'Maybe'] });
  });

  it('hydrates list validations whose formula1 is a range reference', () => {
    const wb = makeFake({
      entries: [
        partial({
          ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }],
          type: DV_TYPE_LIST,
          formula1: 'Sheet1!$A$1:$A$10',
        }),
      ],
    });
    const store = createSpreadsheetStore();
    hydrateValidationsFromEngine(wb, store, 0);
    const a1 = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(a1?.validation).toEqual({
      kind: 'list',
      source: { ref: 'Sheet1!$A$1:$A$10' },
    });
  });

  it('hydrates whole-number validation with op + bounds', () => {
    const wb = makeFake({
      entries: [
        partial({
          ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }],
          type: DV_TYPE_WHOLE,
          formula1: '1',
          formula2: '100',
          op: 0,
        }),
      ],
    });
    const store = createSpreadsheetStore();
    hydrateValidationsFromEngine(wb, store, 0);
    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.validation).toEqual({ kind: 'whole', op: 'between', a: 1, b: 100 });
  });

  it('preserves existing format fields when adding validation', () => {
    const wb = makeFake({
      entries: [
        partial({
          ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }],
          type: DV_TYPE_LIST,
          formula1: 'A,B',
        }),
      ],
    });
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      format: {
        formats: new Map([[addrKey({ sheet: 0, row: 0, col: 0 }), { bold: true, fill: '#ffe' }]]),
      },
    }));
    hydrateValidationsFromEngine(wb, store, 0);
    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.bold).toBe(true);
    expect(fmt?.fill).toBe('#ffe');
    expect(fmt?.validation).toEqual({ kind: 'list', source: ['A', 'B'] });
  });

  it('no-op when capability is off', () => {
    const wb = makeFake({
      dataValidation: false,
      entries: [
        partial({
          ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }],
          type: DV_TYPE_LIST,
          formula1: 'A,B',
        }),
      ],
    });
    const store = createSpreadsheetStore();
    hydrateValidationsFromEngine(wb, store, 0);
    expect(store.getState().format.formats.size).toBe(0);
  });

  it('no-op when engine returns no entries', () => {
    const wb = makeFake({ entries: [] });
    const store = createSpreadsheetStore();
    hydrateValidationsFromEngine(wb, store, 0);
    expect(store.getState().format.formats.size).toBe(0);
  });

  it('handles unquoted comma list (Excel allows both forms)', () => {
    const wb = makeFake({
      entries: [
        partial({
          ranges: [{ sheet: 0, r0: 5, c0: 5, r1: 5, c1: 5 }],
          type: DV_TYPE_LIST,
          formula1: 'High, Medium, Low',
        }),
      ],
    });
    const store = createSpreadsheetStore();
    hydrateValidationsFromEngine(wb, store, 0);
    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 5, col: 5 }));
    expect(fmt?.validation).toEqual({ kind: 'list', source: ['High', 'Medium', 'Low'] });
  });
});

describe('syncValidationsToEngine', () => {
  it('clears + writes back per-cell validations as list rules', () => {
    type LogEntry =
      | { kind: 'clear'; sheet: number }
      | { kind: 'add'; sheet: number; input: unknown };
    const log: LogEntry[] = [];
    const wb = makeFake({
      onClear: (sheet) => {
        log.push({ kind: 'clear', sheet });
        return true;
      },
      onAdd: (sheet, input) => {
        log.push({ kind: 'add', sheet, input });
        return true;
      },
    });
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      format: {
        formats: new Map([
          [
            addrKey({ sheet: 0, row: 0, col: 0 }),
            { validation: { kind: 'list' as const, source: ['Yes', 'No'] } },
          ],
          [
            addrKey({ sheet: 0, row: 1, col: 0 }),
            { validation: { kind: 'list' as const, source: ['Yes', 'No'] } },
          ],
          [
            addrKey({ sheet: 0, row: 2, col: 0 }),
            { validation: { kind: 'list' as const, source: ['A', 'B', 'C'] } },
          ],
        ]),
      },
    }));
    syncValidationsToEngine(wb, store, 0);
    expect(log[0]).toEqual({ kind: 'clear', sheet: 0 });
    // Two add calls — one bucket per source signature, ranges coalesced.
    const adds = log.filter((c) => c.kind === 'add');
    expect(adds).toHaveLength(2);
  });

  it('skips cells from other sheets', () => {
    let added = 0;
    const wb = makeFake({
      onAdd: () => {
        added += 1;
        return true;
      },
    });
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      format: {
        formats: new Map([
          [
            addrKey({ sheet: 1, row: 0, col: 0 }),
            { validation: { kind: 'list' as const, source: ['X'] } },
          ],
        ]),
      },
    }));
    syncValidationsToEngine(wb, store, 0);
    expect(added).toBe(0);
  });

  it('no-op when capability is off', () => {
    let touched = false;
    const wb = makeFake({
      dataValidation: false,
      onClear: () => {
        touched = true;
        return true;
      },
      onAdd: () => {
        touched = true;
        return true;
      },
    });
    const store = createSpreadsheetStore();
    syncValidationsToEngine(wb, store, 0);
    expect(touched).toBe(false);
  });

  it('encodes range-source list rules with a leading = on formula1', () => {
    type AddCall = { sheet: number; input: { formula1?: string } };
    const adds: AddCall[] = [];
    const wb = makeFake({
      onClear: () => true,
      onAdd: (sheet, input) => {
        adds.push({ sheet, input: input as { formula1?: string } });
        return true;
      },
    });
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      format: {
        formats: new Map([
          [
            addrKey({ sheet: 0, row: 0, col: 0 }),
            { validation: { kind: 'list' as const, source: { ref: 'Sheet1!$A$1:$A$10' } } },
          ],
        ]),
      },
    }));
    syncValidationsToEngine(wb, store, 0);
    expect(adds).toHaveLength(1);
    expect(adds[0]?.input.formula1).toBe('=Sheet1!$A$1:$A$10');
  });
});
