import { describe, expect, it } from 'vitest';
import {
  type SpillCellAt,
  type SpillOccupantSource,
  spillBlockers,
  spillOccupancy,
} from '../../../src/engine/spill-blockers.js';
import type { CellValue, Range } from '../../../src/engine/types.js';

const rect = (r0: number, c0: number, r1: number, c1: number): Range => ({
  sheet: 0,
  r0,
  c0,
  r1,
  c1,
});

const source = (over: Partial<SpillOccupantSource> = {}): SpillOccupantSource => ({
  getMerges: () => [],
  getTables: () => [],
  pivotBounds: () => [],
  dataTables: () => [],
  ...over,
});

const cells =
  (entries: Record<string, { value: CellValue; formula?: string | null }>): SpillCellAt =>
  (row, col) => {
    const cell = entries[`${row}:${col}`];
    return cell ? { value: cell.value, formula: cell.formula ?? null } : undefined;
  };

const anchor = { sheet: 0, row: 0, col: 0 };
const target = rect(0, 0, 3, 1);

describe('spill blockers', () => {
  it('counts a value and an empty-text formula, not a blank cell or the anchor', () => {
    const occupied = spillOccupancy(
      source(),
      0,
      cells({
        '0:0': { value: { kind: 'error', code: 8, text: '#SPILL!' }, formula: '=SEQUENCE(4)' },
        '1:0': { value: { kind: 'text', value: 'X' } },
        '2:1': { value: { kind: 'text', value: '' }, formula: '=""' },
        '3:0': { value: { kind: 'blank' } },
      }),
    );
    expect(spillBlockers(target, anchor, occupied)).toEqual([
      { sheet: 0, row: 1, col: 0 },
      { sheet: 0, row: 2, col: 1 },
    ]);
  });

  it('counts every cell of an empty merge, a table and a data table', () => {
    const occupied = spillOccupancy(
      source({
        getMerges: () => [rect(1, 0, 1, 1)],
        getTables: () => [
          { sheetIndex: 0, ref: 'A3' },
          { sheetIndex: 1, ref: 'B3:B4' },
        ],
        dataTables: () => [{ range: rect(3, 0, 3, 1) }],
      }),
      0,
      cells({}),
    );
    expect(spillBlockers(target, anchor, occupied)).toEqual([
      { sheet: 0, row: 1, col: 0 },
      { sheet: 0, row: 1, col: 1 },
      { sheet: 0, row: 2, col: 0 },
      { sheet: 0, row: 3, col: 0 },
      { sheet: 0, row: 3, col: 1 },
    ]);
  });

  it("counts another report's cells but not the report's own", () => {
    const occupied = spillOccupancy(
      source({
        pivotBounds: () => [
          { pivotIndex: 0, range: rect(0, 0, 0, 0) },
          { pivotIndex: 1, range: rect(2, 1, 4, 2) },
        ],
      }),
      0,
      cells({}),
      0,
    );
    expect(spillBlockers(target, anchor, occupied)).toEqual([
      { sheet: 0, row: 2, col: 1 },
      { sheet: 0, row: 3, col: 1 },
    ]);
  });

  it('reads only the cells on the given rows and columns', () => {
    const occupied = spillOccupancy(
      source(),
      0,
      cells({
        '1:0': { value: { kind: 'number', value: 1 } },
        '3:1': { value: { kind: 'number', value: 2 } },
      }),
    );
    expect(spillBlockers(target, anchor, occupied, { rows: [0, 3, 9], cols: [1, 5] })).toEqual([
      { sheet: 0, row: 3, col: 1 },
    ]);
  });
});

describe('a blocked formula spill on the engine', () => {
  it('reports the rectangle it attempted until the blocking cell is cleared', async () => {
    const { WorkbookHandle } = await import('../../../src/engine/workbook-handle.js');
    const wb = await WorkbookHandle.createDefault();
    wb.setFormula({ sheet: 0, row: 0, col: 0 }, '=SEQUENCE(3)');
    wb.setText({ sheet: 0, row: 1, col: 0 }, 'X');
    wb.recalc();
    expect(wb.blockedSpillRange(0, 0, 0)).toEqual(rect(0, 0, 2, 0));
    expect(wb.spillInfo(0, 0, 0)).toBeNull();
    const occupied = spillOccupancy(wb, 0, (row, col) => {
      const addr = { sheet: 0, row, col };
      return { value: wb.getValue(addr), formula: wb.cellFormula(addr) };
    });
    expect(spillBlockers(rect(0, 0, 2, 0), anchor, occupied)).toEqual([
      { sheet: 0, row: 1, col: 0 },
    ]);
    wb.setBlank({ sheet: 0, row: 1, col: 0 });
    wb.recalc();
    expect(wb.blockedSpillRange(0, 0, 0)).toBeNull();
    wb.dispose();
  });
});
