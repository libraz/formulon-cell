import { describe, expect, it, vi } from 'vitest';
import { textToColumns } from '../../../src/commands/text-to-columns.js';
import type { Addr } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const key = (addr: Addr): string => `${addr.sheet}:${addr.row}:${addr.col}`;

const fakeWorkbook = (
  profile: string | null = null,
): WorkbookHandle & {
  recalcAuto: ReturnType<typeof vi.fn>;
  writes: Map<string, unknown>;
} => {
  const writes = new Map<string, unknown>();
  return {
    writes,
    setBlank: vi.fn((addr: Addr) => writes.set(key(addr), { kind: 'blank' })),
    setBool: vi.fn((addr: Addr, value: boolean) => writes.set(key(addr), { kind: 'bool', value })),
    setFormula: vi.fn((addr: Addr, formula: string) =>
      writes.set(key(addr), { kind: 'formula', formula }),
    ),
    setNumber: vi.fn((addr: Addr, value: number) =>
      writes.set(key(addr), { kind: 'number', value }),
    ),
    setText: vi.fn((addr: Addr, value: string) => writes.set(key(addr), { kind: 'text', value })),
    spreadsheetProfileId: vi.fn(() => profile),
    withBatchedRecalc: vi.fn(<T>(fn: () => T): T => fn()),
    recalcAuto: vi.fn(),
  } as unknown as WorkbookHandle & {
    recalcAuto: ReturnType<typeof vi.fn>;
    writes: Map<string, unknown>;
  };
};

describe('textToColumns', () => {
  it('splits huge whole-column selections by visiting only materialized text cells', () => {
    const store = createSpreadsheetStore();
    mutators.setCell(
      store,
      { sheet: 0, row: 8, col: 2 },
      { kind: 'text', value: 'alpha,beta' },
      null,
    );
    mutators.setCell(
      store,
      { sheet: 0, row: 8, col: 3 },
      { kind: 'text', value: 'outside,ignored' },
      null,
    );
    mutators.setCellFormat(store, { sheet: 0, row: 8, col: 2 }, { bold: true });
    const wb = fakeWorkbook();

    const count = textToColumns(
      store.getState(),
      store,
      wb,
      { sheet: 0, r0: 0, c0: 2, r1: 1048575, c1: 2 },
      ',',
    );

    expect(count).toBe(2);
    expect(wb.writes).toEqual(
      new Map([
        ['0:8:2', { kind: 'text', value: 'alpha' }],
        ['0:8:3', { kind: 'text', value: 'beta' }],
      ]),
    );
    expect(store.getState().format.formats.get('0:8:2')).toEqual({ bold: true });
    expect(store.getState().format.formats.get('0:8:3')).toEqual({ bold: true });
    expect(wb.recalcAuto).toHaveBeenCalledTimes(1);
  });

  it('uses the workbook profile when coercing localized tokens', () => {
    const store = createSpreadsheetStore();
    mutators.setCell(
      store,
      { sheet: 0, row: 0, col: 0 },
      { kind: 'text', value: '1,25;2,50' },
      null,
    );
    const wb = fakeWorkbook('windows-de_DE');

    expect(
      textToColumns(store.getState(), store, wb, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, ';'),
    ).toBe(2);
    expect(wb.writes).toEqual(
      new Map([
        ['0:0:0', { kind: 'number', value: 1.25 }],
        ['0:0:1', { kind: 'number', value: 2.5 }],
      ]),
    );
  });

  it('splits as Excel does with the double-quote qualifier', () => {
    // Excel (comma, qualifier "): 'a,b,c' -> a|b|c; '1,"x,y",3' -> 1|x,y|3;
    // '007,2026-11-03,1,234' -> 7 | 2026-11-03 as a date | 1 | 234; 'a,,c' -> a|blank|c.
    const store = createSpreadsheetStore();
    const rows = ['a,b,c', '1,"x,y",3', '007,2026-11-03,1,234', 'a,,c'];
    rows.forEach((value, row) => {
      mutators.setCell(store, { sheet: 0, row, col: 0 }, { kind: 'text', value }, null);
    });
    const wb = fakeWorkbook();
    textToColumns(store.getState(), store, wb, { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 }, ',');
    const at = (row: number, col: number) => wb.writes.get(`0:${row}:${col}`);
    expect([at(0, 0), at(0, 1), at(0, 2)]).toEqual([
      { kind: 'text', value: 'a' },
      { kind: 'text', value: 'b' },
      { kind: 'text', value: 'c' },
    ]);
    expect([at(1, 0), at(1, 1), at(1, 2), at(1, 3)]).toEqual([
      { kind: 'number', value: 1 },
      { kind: 'text', value: 'x,y' },
      { kind: 'number', value: 3 },
      undefined,
    ]);
    expect([at(2, 0), at(2, 1), at(2, 2), at(2, 3)]).toEqual([
      { kind: 'number', value: 7 },
      { kind: 'number', value: 46329 },
      { kind: 'number', value: 1 },
      { kind: 'number', value: 234 },
    ]);
    expect([at(3, 0), at(3, 1), at(3, 2)]).toEqual([
      { kind: 'text', value: 'a' },
      { kind: 'blank' },
      { kind: 'text', value: 'c' },
    ]);
  });
});
