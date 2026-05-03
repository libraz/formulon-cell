import { describe, expect, it } from 'vitest';
import type { Range } from '../../../src/engine/types.js';
import { hydrateValidationsFromEngine } from '../../../src/engine/validation-sync.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

interface FakeValidationEntry {
  ranges: Range[];
  type: string;
  op: string;
  formula1: string;
  formula2: string;
  errorMessage: string;
}

const makeFake = (
  opts: { dataValidation?: boolean; entries?: FakeValidationEntry[] } = {},
): WorkbookHandle => {
  const caps = { dataValidation: opts.dataValidation ?? true };
  const fake = {
    capabilities: caps,
    getValidationsForSheet(_sheet: number): FakeValidationEntry[] {
      if (!caps.dataValidation) return [];
      return opts.entries ?? [];
    },
  };
  return fake as unknown as WorkbookHandle;
};

describe('hydrateValidationsFromEngine', () => {
  it('seeds format.validation for cells inside a list-type range', () => {
    const wb = makeFake({
      entries: [
        {
          ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 }],
          type: 'list',
          op: '',
          formula1: '"Yes,No,Maybe"',
          formula2: '',
          errorMessage: '',
        },
      ],
    });
    const store = createSpreadsheetStore();
    hydrateValidationsFromEngine(wb, store, 0);
    const a1 = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    const a2 = store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 0 }));
    expect(a1?.validation).toEqual({ kind: 'list', source: ['Yes', 'No', 'Maybe'] });
    expect(a2?.validation).toEqual({ kind: 'list', source: ['Yes', 'No', 'Maybe'] });
  });

  it('drops validations whose formula1 is a range reference (cannot expand)', () => {
    const wb = makeFake({
      entries: [
        {
          ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }],
          type: 'list',
          formula1: 'Sheet1!$A$1:$A$10',
          op: '',
          formula2: '',
          errorMessage: '',
        },
      ],
    });
    const store = createSpreadsheetStore();
    hydrateValidationsFromEngine(wb, store, 0);
    expect(store.getState().format.formats.size).toBe(0);
  });

  it('skips non-list validation types (no UI today)', () => {
    const wb = makeFake({
      entries: [
        {
          ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }],
          type: 'whole',
          formula1: '1',
          formula2: '100',
          op: 'between',
          errorMessage: '',
        },
      ],
    });
    const store = createSpreadsheetStore();
    hydrateValidationsFromEngine(wb, store, 0);
    expect(store.getState().format.formats.size).toBe(0);
  });

  it('preserves existing format fields when adding validation', () => {
    const wb = makeFake({
      entries: [
        {
          ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }],
          type: 'list',
          formula1: 'A,B',
          op: '',
          formula2: '',
          errorMessage: '',
        },
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
        {
          ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }],
          type: 'list',
          formula1: 'A,B',
          op: '',
          formula2: '',
          errorMessage: '',
        },
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
        {
          ranges: [{ sheet: 0, r0: 5, c0: 5, r1: 5, c1: 5 }],
          type: 'list',
          formula1: 'High, Medium, Low',
          op: '',
          formula2: '',
          errorMessage: '',
        },
      ],
    });
    const store = createSpreadsheetStore();
    hydrateValidationsFromEngine(wb, store, 0);
    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 5, col: 5 }));
    expect(fmt?.validation).toEqual({ kind: 'list', source: ['High', 'Medium', 'Low'] });
  });
});
