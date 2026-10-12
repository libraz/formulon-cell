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
});
