import { beforeEach, describe, expect, it } from 'vitest';
import {
  findMatchingRanges,
  type GoToSpecialKind,
  type GoToSpecialOptions,
} from '../../../src/commands/goto-special.js';
import { MAX_COL, MAX_ROW } from '../../../src/engine/address.js';
import type { Range } from '../../../src/engine/types.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

// Measured with Range.SpecialCells in the desktop app on this fixture: row 3
// and column E hidden, a note on B2, validation on B4:B5, a rule on C5:C6.
const FIXTURE: Record<string, string | number | boolean> = {
  A1: 1,
  B1: 't',
  C1: true,
  D1: '=1/0',
  A2: '=A1+1',
  B2: '=B1&"x"',
  C2: '=C1',
  D2: '=NA()',
  E2: '=ISBLANK(E1)',
  A3: 3,
  C3: '=A3>2',
  D3: '="a"&"b"',
  A4: '=SUM(A1:A3)',
  A5: '',
  F5: 9,
  B6: 't"x"("x")',
};

const cell = (a1: string): { row: number; col: number } => {
  const m = /^([A-Z])(\d+)$/.exec(a1);
  if (!m) throw new Error(a1);
  return { row: Number(m[2]) - 1, col: (m[1] ?? 'A').charCodeAt(0) - 65 };
};

/** Every cell of `A1:B2`-style areas, as sorted A1 refs. */
const cellsOf = (areas: string): string[] => {
  if (areas === '') return [];
  const out: string[] = [];
  for (const area of areas.split(',')) {
    const [from, to = from] = area.split(':') as [string, string?];
    const a = cell(from);
    const b = cell(to);
    for (let row = a.row; row <= b.row; row += 1)
      for (let col = a.col; col <= b.col; col += 1)
        out.push(`${String.fromCharCode(65 + col)}${row + 1}`);
  }
  return out.sort();
};

const cellsOfRanges = (ranges: readonly Range[]): string[] =>
  cellsOf(
    ranges
      .map(
        (r) =>
          `${String.fromCharCode(65 + r.c0)}${r.r0 + 1}:${String.fromCharCode(65 + r.c1)}${r.r1 + 1}`,
      )
      .join(','),
  );

type Case = [kind: GoToSpecialKind, options: GoToSpecialOptions, areas: string];

const ALL: Case[] = [
  ['constants', {}, 'A1:C1,A3,F5,B6'],
  ['constants', { filters: { numbers: true } }, 'A1,A3,F5'],
  ['constants', { filters: { text: true } }, 'B1,B6'],
  ['constants', { filters: { logical: true } }, 'C1'],
  ['constants', { filters: { errors: true } }, ''],
  ['formulas', {}, 'D1,A2:E2,C3:D3,A4'],
  ['formulas', { filters: { numbers: true } }, 'A2,A4'],
  ['formulas', { filters: { text: true } }, 'B2,D3'],
  ['formulas', { filters: { logical: true } }, 'E2,C2:C3'],
  ['formulas', { filters: { errors: true } }, 'D1:D2'],
  ['blanks', {}, 'E1,F1:F2,B3,E3:F3,B4:F4,B5:E5,A5:A6,C6:F6'],
  ['last-cell', {}, 'F6'],
  ['conditional-format', {}, 'C5:C6'],
  ['data-validation', {}, 'B4:B5'],
  ['comments', {}, 'B2'],
  ['conditional-format', { same: true }, ''],
];

const SUB: Case[] = [
  ['constants', {}, 'B6'],
  ['blanks', {}, 'B5'],
  ['visible-cells', {}, 'B5:B6'],
  ['data-validation', {}, 'B5'],
  ['formulas', {}, ''],
  ['conditional-format', {}, ''],
  ['comments', {}, ''],
];

describe('Go To Special matches the desktop app', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  const select = (r0: number, c0: number, r1 = r0, c1 = c0): void => {
    store.setState((s) => ({
      ...s,
      selection: {
        active: { sheet: 0, row: r0, col: c0 },
        anchor: { sheet: 0, row: r0, col: c0 },
        range: { sheet: 0, r0, c0, r1, c1 },
        extraRanges: [],
      },
    }));
  };
  const found = (kind: GoToSpecialKind, options: GoToSpecialOptions): string[] => {
    const result = findMatchingRanges(wb, store, kind, options);
    if (result.status === 'too-large') throw new Error('too-large');
    return result.status === 'found' ? cellsOfRanges(result.ranges) : [];
  };

  beforeEach(async () => {
    wb = await WorkbookHandle.createDefault();
    for (const [ref, input] of Object.entries(FIXTURE)) {
      const addr = { sheet: 0, ...cell(ref) };
      if (typeof input === 'number') wb.setNumber(addr, input);
      else if (typeof input === 'boolean') wb.setBool(addr, input);
      else if (input.startsWith('=')) wb.setFormula(addr, input);
      else wb.setText(addr, input);
    }
    wb.recalc();
    store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      data: {
        ...s.data,
        cells: new Map([...wb.cells(0)].map((e) => [addrKey(e.addr), e])),
      },
      layout: { ...s.layout, hiddenRows: new Set([2]), hiddenCols: new Set([4]) },
    }));
    mutators.setCellFormat(store, { sheet: 0, ...cell('B2') }, { comment: 'note' });
    const rule = { kind: 'list' as const, source: ['1'] };
    mutators.setCellFormat(store, { sheet: 0, ...cell('B4') }, { validation: rule });
    mutators.setCellFormat(store, { sheet: 0, ...cell('B5') }, { validation: { ...rule } });
    mutators.addConditionalRule(store, {
      kind: 'data-bar',
      range: { sheet: 0, r0: 4, c0: 2, r1: 5, c1: 2 },
      color: '#638ec6',
    });
  });

  it.each(ALL)('%s %o over A1:F6', (kind, options, areas) => {
    select(0, 0, 5, 5);
    expect(found(kind, options)).toEqual(cellsOf(areas));
  });

  it.each(ALL)('%s %o from a single cell', (kind, options, areas) => {
    select(0, 0);
    expect(found(kind, options)).toEqual(cellsOf(areas));
  });

  it.each(SUB)('%s %o over B5:B6', (kind, options, areas) => {
    select(4, 1, 5, 1);
    expect(found(kind, options)).toEqual(cellsOf(areas));
  });

  it('visible cells skip the hidden row and column', () => {
    select(0, 0, 5, 5);
    expect(found('visible-cells', {})).toEqual(cellsOf('A1:D2,F1:F2,A4:D6,F4:F6'));
  });

  it('visible cells from a single cell reach the sheet edges', () => {
    select(0, 0);
    const result = findMatchingRanges(wb, store, 'visible-cells');
    expect(result.status).toBe('found');
    const ranges = result.status === 'found' ? result.ranges : [];
    const covers = (row: number, col: number) =>
      ranges.some((r) => row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1);
    expect(covers(1, MAX_COL)).toBe(true);
    expect(covers(MAX_ROW, 0)).toBe(true);
    expect(covers(MAX_ROW, MAX_COL)).toBe(true);
    expect(covers(2, 0)).toBe(false);
    expect(covers(10, 4)).toBe(false);
  });

  it('row and column differences compare against the active cell line', () => {
    // H1:K4 = [[1,2,3,4],[1,5,3,6],[7,2,3,4],[1,2,9,4]], comparison cell H1.
    const grid = [
      [1, 2, 3, 4],
      [1, 5, 3, 6],
      [7, 2, 3, 4],
      [1, 2, 9, 4],
    ];
    grid.forEach((line, row) => {
      line.forEach((v, i) => {
        wb.setNumber({ sheet: 0, row, col: 7 + i }, v);
      });
    });
    wb.recalc();
    store.setState((s) => ({
      ...s,
      data: { ...s.data, cells: new Map([...wb.cells(0)].map((e) => [addrKey(e.addr), e])) },
    }));
    select(0, 7, 3, 10);
    expect(found('row-differences', {})).toEqual(cellsOf('I1:K4'));
    expect(found('column-differences', {})).toEqual(cellsOf('I2,K2,H3,J4'));
  });

  it('precedents and dependents stay on the sheet', () => {
    select(3, 0);
    expect(found('precedents', {})).toEqual(cellsOf('A1:A3'));
    select(0, 0);
    expect(found('dependents', {})).toEqual(cellsOf('A2,A4'));
  });
});
