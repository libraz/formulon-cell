// @vitest-environment node

import { afterEach, describe, expect, it } from 'vitest';
import { canShiftMerges, insertCells } from '../../../src/commands/cell-shift.js';
import { insertCopiedCellsFromTSV } from '../../../src/commands/clipboard/insert-copied-cells.js';
import { History } from '../../../src/commands/history.js';
import { conditionalRuleToEngineInput } from '../../../src/engine/cf-writeback.js';
import { hydrateMergesFromEngine } from '../../../src/engine/merges-sync.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

// Expectations are measured against the desktop spreadsheet's Insert Cells.

const MAX_ROW = 1_048_575;
const MAX_COL = 16_383;

type Direction = 'down' | 'right';
type Store = ReturnType<typeof createSpreadsheetStore>;

const BAND_MERGE: Range = { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 1 };

const a1 = (row: number, col: number): string => `${String.fromCharCode(65 + col)}${row + 1}`;
/** Cells covered by `ranges`, as a sorted A1 list (piece order is irrelevant). */
const cellsOf = (ranges: readonly Range[]): string[] => {
  const out = new Set<string>();
  for (const r of ranges)
    for (let row = r.r0; row <= r.r1; row += 1)
      for (let col = r.c0; col <= r.c1; col += 1) out.add(a1(row, col));
  return [...out].sort();
};
const cells = (...refs: string[]): string[] =>
  cellsOf(
    refs.map((ref) => {
      const [from = '', to = from] = ref.split(':');
      const at = (x: string) => ({ row: Number(x.slice(1)) - 1, col: x.charCodeAt(0) - 65 });
      const f = at(from);
      const t = at(to);
      return { sheet: 0, r0: f.row, c0: f.col, r1: t.row, c1: t.col };
    }),
  );

const cfCells = (wb: WorkbookHandle): string[] =>
  cellsOf(
    wb.getConditionalFormats(0).flatMap((rule) =>
      rule.ranges.map((r) => ({
        sheet: 0,
        r0: r.firstRow,
        c0: r.firstCol,
        r1: r.lastRow,
        c1: r.lastCol,
      })),
    ),
  );
const dvCells = (wb: WorkbookHandle): string[] =>
  cellsOf(wb.getValidationsForSheet(0).flatMap((v) => v.ranges));
const names = (wb: WorkbookHandle): Record<string, string> =>
  Object.fromEntries(Array.from(wb.definedNames()).map((n) => [n.name, n.formula]));

/** Cells, formulas, merges, rules and names: what an undo must restore exactly. */
function dump(wb: WorkbookHandle, store: Store) {
  wb.recalc();
  return {
    cells: [0, 1].map((sheet) =>
      Array.from(wb.physicalCells(sheet))
        .map((c) => ({ r: c.addr.row, c: c.addr.col, v: c.value, f: c.formula }))
        .sort((x, y) => x.r - y.r || x.c - y.c),
    ),
    merges: wb
      .getMerges(0)
      .map((m) => a1(m.r0, m.c0) + a1(m.r1, m.c1))
      .sort(),
    storeMerges: [...store.getState().merges.byAnchor.values()]
      .map((m) => a1(m.r0, m.c0) + a1(m.r1, m.c1))
      .sort(),
    cf: cfCells(wb),
    dv: dvCells(wb),
    names: names(wb),
  };
}

const live: WorkbookHandle[] = [];
afterEach(() => {
  for (const wb of live.splice(0)) wb.dispose();
});

async function fixture(merge: Range = BAND_MERGE) {
  const store = createSpreadsheetStore();
  const wb = await WorkbookHandle.createDefault();
  live.push(wb);
  expect(wb.isStub).toBe(false);
  wb.addSheet('Other');
  const s0 = wb.sheetName(0);
  const num = (row: number, col: number, v: number, sheet = 0) =>
    wb.setNumber({ sheet, row, col }, v);
  const fx = (row: number, col: number, f: string, sheet = 0) =>
    wb.setFormula({ sheet, row, col }, f);

  // Band area B2:C5 plus formulas that sit left of, above, inside and
  // beyond it, and formulas on the other sheet.
  num(1, 1, 1);
  num(2, 1, 2);
  num(3, 1, 3);
  num(1, 2, 10);
  num(2, 2, 20);
  fx(4, 1, '=B2+1');
  fx(4, 2, '=$C$2');
  num(8, 1, 7);
  num(1, 4, 4);
  num(1, 5, 5);
  fx(0, 0, '=SUM(B1:B7)'); // straddles a down band
  fx(1, 0, '=B3'); // single ref into the band
  fx(0, 3, '=SUM(B2:C3)'); // fully inside the band columns
  fx(2, 5, '=SUM(A2:F2)'); // straddles a right band
  fx(2, 6, '=E2'); // beyond a right band
  fx(0, 0, `=SUM(B1:B7)`);
  fx(0, 0, `=SUM(B1:B7)`, 1);
  fx(1, 0, `=${s0}!B3`, 1);
  fx(2, 0, `=SUM(${s0}!A2:F2)`, 1);
  fx(3, 0, `=SUM(${s0}!B1:B7)`, 1);
  fx(4, 0, `=${s0}!$C$2`, 1);
  expect(wb.setDefinedNameEntry('Band', `${s0}!$B$2:$B$4`)).toEqual({ ok: true });
  expect(wb.setDefinedNameEntry('Cell', `${s0}!$C$3`)).toEqual({ ok: true });
  expect(wb.setDefinedNameEntry('BelowBand', `${s0}!$B$9`)).toEqual({ ok: true });
  expect(wb.engineAddMerge(0, merge)).toBe(true);
  expect(wb.engineAddMerge(0, { sheet: 0, r0: 20, c0: 4, r1: 20, c1: 5 })).toBe(true);
  const bar = conditionalRuleToEngineInput({
    kind: 'data-bar',
    range: { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 1 },
    color: '#638ec6',
    min: { kind: 'number', value: 0 },
    max: { kind: 'number', value: 10 },
    direction: 'left-to-right',
  });
  if (!bar) throw new Error('data bar must be representable');
  expect(wb.addConditionalFormat(0, bar)).toBeGreaterThanOrEqual(0);
  expect(
    wb.addValidationEntry(0, {
      ranges: [{ sheet: 0, r0: 1, c0: 1, r1: 3, c1: 1 }],
      type: 1,
      op: 0,
      formula1: '1',
      formula2: '10',
    }),
  ).toBe(true);
  wb.recalc();
  mutators.replaceCells(store, wb.cells(0));
  hydrateMergesFromEngine(wb, store, 0);
  const history = new History();
  wb.attachHistory(history);
  return { store, wb, history };
}

/** The band an insert of `range` shifts. */
const bandOf = (range: Range, direction: Direction): Range =>
  direction === 'down' ? { ...range, r1: MAX_ROW } : { ...range, c1: MAX_COL };

interface Case {
  name: string;
  range: Range;
  direction: Direction;
  merge: Range;
  names: Record<string, string>;
  cf: string[];
  dv: string[];
  moved: Record<string, number>;
}

const cases: Case[] = [
  {
    name: 'down 2x2 at B2',
    range: { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 },
    direction: 'down',
    merge: BAND_MERGE,
    names: { Band: 'Sheet1!$B$4:$B$6', Cell: 'Sheet1!$C$5', BelowBand: 'Sheet1!$B$11' },
    cf: cells('B4:B6'),
    dv: cells('B4:B6'),
    moved: { B4: 1, B5: 2, B6: 3, C4: 10, C5: 20, E2: 4 },
  },
  {
    name: 'right 2x2 at B2',
    range: { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 },
    direction: 'right',
    merge: BAND_MERGE,
    names: { Band: 'Sheet1!$B$2:$B$4', Cell: 'Sheet1!$E$3', BelowBand: 'Sheet1!$B$9' },
    cf: cells('B4', 'D2:D3'),
    dv: cells('B4', 'D2:D3'),
    moved: { D2: 1, D3: 2, B4: 3, E2: 10, E3: 20, G2: 4, H2: 5 },
  },
  {
    name: 'right 1x2 at B3',
    range: { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 2 },
    direction: 'right',
    merge: { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 2 },
    names: { Cell: 'Sheet1!$E$3' },
    cf: cells('B2', 'B4', 'D3'),
    dv: cells('B2', 'B4', 'D3'),
    moved: { B2: 1, D3: 2, B4: 3, E3: 20 },
  },
  {
    name: 'down 1x1 at B3',
    range: { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
    direction: 'down',
    merge: { sheet: 0, r0: 2, c0: 1, r1: 3, c1: 1 },
    names: { Band: 'Sheet1!$B$2:$B$5', Cell: 'Sheet1!$C$3', BelowBand: 'Sheet1!$B$10' },
    cf: cells('B2:B5'),
    dv: cells('B2:B5'),
    moved: { B2: 1, B4: 2, B5: 3, C3: 20 },
  },
];

const numberAt = (wb: WorkbookHandle, ref: string) => {
  const v = wb.getValue({ sheet: 0, row: Number(ref.slice(1)) - 1, col: ref.charCodeAt(0) - 65 });
  return v.kind === 'number' ? v.value : v;
};

describe('Insert Cells shift through the engine', () => {
  it.each(cases)(
    '$name: names, rules and cells move as measured; undo and redo are exact',
    async (c) => {
      const { store, wb, history } = await fixture(c.merge);
      const before = dump(wb, store);
      expect(insertCells(store, wb, history, c.range, c.direction)).toBe(true);
      const after = dump(wb, store);
      expect(names(wb)).toMatchObject(c.names);
      expect(cfCells(wb)).toEqual(c.cf);
      expect(dvCells(wb)).toEqual(c.dv);
      for (const [ref, value] of Object.entries(c.moved)) expect(numberAt(wb, ref)).toBe(value);
      // Merges move with their band, in the engine and in the store.
      expect(after.merges).toEqual(after.storeMerges);
      expect(history.undo()).toBe(true);
      expect(dump(wb, store)).toEqual(before);
      expect(history.redo()).toBe(true);
      expect(dump(wb, store)).toEqual(after);
    },
  );

  it('down 2x2 at B2: formulas that reference the band follow it, on every sheet', async () => {
    const { store, wb, history } = await fixture();
    expect(insertCells(store, wb, history, cases[0]?.range as Range, 'down')).toBe(true);
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=SUM(B1:B9)');
    expect(wb.cellFormula({ sheet: 0, row: 1, col: 0 })).toBe('=B5');
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 3 })).toBe('=SUM(B4:C5)');
    expect(wb.cellFormula({ sheet: 1, row: 1, col: 0 })).toBe('=Sheet1!B5');
    expect(wb.cellFormula({ sheet: 1, row: 3, col: 0 })).toBe('=SUM(Sheet1!B1:B9)');
  });
});

describe('Insert Copied Cells uses the same shift, then pastes', () => {
  it.each(['down', 'right'] as const)(
    '%s: the pasted block lands and undo restores the prior state',
    async (direction) => {
      const { store, wb, history } = await fixture();
      store.setState((s) => ({
        ...s,
        selection: {
          ...s.selection,
          active: { sheet: 0, row: 1, col: 1 },
          anchor: { sheet: 0, row: 1, col: 1 },
          range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
        },
      }));
      const before = dump(wb, store);
      const result = insertCopiedCellsFromTSV(store, wb, history, '5\t6\n7\t8', direction);
      expect(result.range).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 });
      expect([
        numberAt(wb, 'B2'),
        numberAt(wb, 'C2'),
        numberAt(wb, 'B3'),
        numberAt(wb, 'C3'),
      ]).toEqual([5, 6, 7, 8]);
      const expected = cases[direction === 'down' ? 0 : 1];
      expect(names(wb)).toMatchObject(expected?.names ?? {});
      const after = dump(wb, store);
      expect(history.undo()).toBe(true);
      expect(dump(wb, store)).toEqual(before);
      expect(history.redo()).toBe(true);
      expect(dump(wb, store)).toEqual(after);
    },
  );
});

describe('Insert Copied Cells refuses a merge the band cuts through', () => {
  it.each(['down', 'right'] as const)('%s', async (direction) => {
    const { store } = await fixture();
    const range: Range =
      direction === 'down'
        ? { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 }
        : { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 2 };
    expect(canShiftMerges(store.getState(), bandOf(range, direction), direction)).toBe(false);
  });
});
