import { describe, expect, it, vi } from 'vitest';
import { axisEditSplitsDataTable } from '../../../src/commands/data-table-guard.js';
import type { Addr } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';

const at = (row: number, col: number): Addr => ({ sheet: 0, row, col });

/** A1 = 5 is the input; B3:B5 hold 1..3; C2 = A1*2; result block C3:C5. */
async function oneVariable(): Promise<WorkbookHandle> {
  const wb = await WorkbookHandle.createDefault();
  wb.setNumber(at(0, 0), 5);
  for (const i of [1, 2, 3]) wb.setNumber(at(1 + i, 1), i);
  wb.setFormula(at(1, 2), '=A1*2');
  const ok = wb.addDataTable({
    range: { sheet: 0, r0: 2, c0: 2, r1: 4, c1: 2 },
    rowInput: null,
    colInput: at(0, 0),
  });
  expect(ok).toBe(true);
  return wb;
}

const num = (wb: WorkbookHandle, row: number, col: number): unknown => {
  const v = wb.getValue(at(row, col));
  return v.kind === 'number' ? v.value : v;
};

describe.skipIf(!canLoadWasm())('data tables', () => {
  it('computes a one-variable table on creation', async () => {
    const wb = await oneVariable();
    try {
      expect([num(wb, 2, 2), num(wb, 3, 2), num(wb, 4, 2)]).toEqual([2, 4, 6]);
      expect(wb.dataTableAt(at(3, 2))).toEqual({
        range: { sheet: 0, r0: 2, c0: 2, r1: 4, c1: 2 },
        rowInput: null,
        colInput: at(0, 0),
        inputDeleted: false,
      });
      expect(wb.dataTableAt(at(1, 2))).toBeNull();
    } finally {
      wb.dispose();
    }
  });

  it('computes a two-variable table', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      wb.setNumber(at(0, 0), 1);
      wb.setNumber(at(0, 1), 1);
      wb.setFormula(at(2, 2), '=A1*B1');
      for (const i of [1, 2]) {
        wb.setNumber(at(2 + i, 2), i);
        wb.setNumber(at(2, 2 + i), 10 * i);
      }
      expect(
        wb.addDataTable({
          range: { sheet: 0, r0: 3, c0: 3, r1: 4, c1: 4 },
          rowInput: at(0, 0),
          colInput: at(0, 1),
        }),
      ).toBe(true);
      expect(num(wb, 4, 4)).toBe(40);
      expect(wb.dataTableAt(at(3, 3))?.rowInput).toEqual(at(0, 0));
    } finally {
      wb.dispose();
    }
  });

  it('rejects an input cell inside the result block', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(
        wb.addDataTable({
          range: { sheet: 0, r0: 3, c0: 3, r1: 4, c1: 4 },
          rowInput: at(3, 3),
          colInput: null,
        }),
      ).toBe(false);
    } finally {
      wb.dispose();
    }
  });

  it('refuses writes into the result block', async () => {
    const wb = await oneVariable();
    try {
      expect(() => wb.setNumber(at(3, 2), 9)).toThrow(/data table/);
      expect(num(wb, 3, 2)).toBe(4);
    } finally {
      wb.dispose();
    }
  });

  it('leaves tables to Calculate Now under Automatic Except for Data Tables', async () => {
    const wb = await oneVariable();
    try {
      wb.setCalcMode(2);
      wb.setFormula(at(1, 2), '=A1*3');
      expect(num(wb, 1, 2)).toBe(15);
      expect(num(wb, 3, 2)).toBe(4);
      wb.recalc();
      expect(num(wb, 3, 2)).toBe(6);
    } finally {
      wb.dispose();
    }
  });

  it('keeps tables live under Automatic', async () => {
    const wb = await oneVariable();
    try {
      wb.setFormula(at(1, 2), '=A1*3');
      expect(num(wb, 3, 2)).toBe(6);
    } finally {
      wb.dispose();
    }
  });

  it('lists the tables of a sheet and of a range', async () => {
    const wb = await oneVariable();
    try {
      expect(wb.dataTables(0).map((t) => t.range)).toEqual([
        { sheet: 0, r0: 2, c0: 2, r1: 4, c1: 2 },
      ]);
      expect(wb.dataTablesInRange({ sheet: 0, r0: 4, c0: 0, r1: 9, c1: 9 })).toHaveLength(1);
      expect(wb.dataTablesInRange({ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 9 })).toEqual([]);
    } finally {
      wb.dispose();
    }
  });

  it('removes a table and blanks its result cells', async () => {
    const wb = await oneVariable();
    try {
      expect(wb.removeDataTable(at(4, 2))).toBe(true);
      expect(wb.getValue(at(3, 2))).toEqual({ kind: 'blank' });
      expect(wb.dataTables(0)).toEqual([]);
    } finally {
      wb.dispose();
    }
  });

  it('round-trips through save and load', async () => {
    const wb = await oneVariable();
    try {
      const restored = await WorkbookHandle.loadBytes(wb.save());
      try {
        expect(restored.dataTables(0)).toEqual([
          {
            range: { sheet: 0, r0: 2, c0: 2, r1: 4, c1: 2 },
            rowInput: null,
            colInput: at(0, 0),
            inputDeleted: false,
          },
        ]);
        expect(num(restored, 4, 2)).toBe(6);
        expect(() => restored.setNumber(at(3, 2), 1)).toThrow(/data table/);
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('looks each table up once, however many result cells it has', async () => {
    const wb = await oneVariable();
    try {
      const engine = (wb as unknown as { wb: { getDataTable: () => unknown } }).wb;
      const lookups = vi.spyOn(engine, 'getDataTable');
      expect(wb.dataTablesInRange({ sheet: 0, r0: 0, c0: 0, r1: 9, c1: 9 })).toHaveLength(1);
      expect(lookups).toHaveBeenCalledTimes(1);
    } finally {
      wb.dispose();
    }
  });

  it('checks a row edit against the edited lines only', async () => {
    const wb = await oneVariable();
    try {
      const scans = vi.spyOn(wb, 'cellsInRange');
      expect(axisEditSplitsDataTable(wb, 0, 'row', 'insert', 40, 1)).toBe(false);
      expect(axisEditSplitsDataTable(wb, 0, 'row', 'delete', 1, 1)).toBe(true);
      const scanned = scans.mock.calls.map(([range]) => [range.r0, range.r1]);
      expect(scanned).toEqual([
        [40, 40],
        [1, 2],
      ]);
    } finally {
      wb.dispose();
    }
  });

  it('asks the engine for the tables only where its own pass skips them', async () => {
    const wb = await oneVariable();
    try {
      const engine = (wb as unknown as { wb: { recalcDataTables: () => unknown } }).wb;
      const tables = vi.spyOn(engine, 'recalcDataTables');
      wb.recalc();
      expect(tables).not.toHaveBeenCalled();
      wb.setCalcMode(1);
      wb.setFormula(at(1, 2), '=A1*4');
      expect(num(wb, 3, 2)).toBe(4);
      wb.recalc();
      expect(num(wb, 3, 2)).toBe(8);
      expect(tables).not.toHaveBeenCalled();
      wb.setCalcMode(2);
      tables.mockClear();
      wb.recalc();
      expect(tables).toHaveBeenCalledTimes(1);
    } finally {
      wb.dispose();
    }
  });
});
