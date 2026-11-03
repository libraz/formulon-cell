import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { type FindOptions, findAll, replaceAll } from '../../../src/commands/find.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

// Measured with Range.Find / Range.Replace in the desktop app (ja UI) on this
// fixture; every hit list is in by-rows order from the first cell.
const FIXTURE: Record<string, string | number> = {
  A1: 'apple',
  A2: 'Apple',
  A3: 'a*b',
  A4: 'a?b',
  A5: 'axxb',
  A6: 'a~b',
  A7: 'ＡＢＣ',
  A8: 'ABC',
  A9: 'ｱｲｳ',
  A10: 'アイウ',
  A11: 'あいう',
  A12: '=1+2',
  A13: 'apple pie',
  A14: 'ｶﾞ',
  A15: 'ガ',
  B1: 'apple',
  B2: 3,
  B3: '3.0',
  C3: 'apple',
  D1: '="ap"&"ple"',
};

const parse = (a1: string) => {
  const m = /^([A-Z])(\d+)$/.exec(a1);
  if (!m) throw new Error(a1);
  return { sheet: 0, row: Number(m[2]) - 1, col: (m[1] ?? 'A').charCodeAt(0) - 65 };
};
const a1 = (addr: { row: number; col: number }) =>
  `${String.fromCharCode(65 + addr.col)}${addr.row + 1}`;

type Case = [query: string, opts: Partial<FindOptions>, hits: string[]];

const formulas = { lookIn: 'formulas' as const };
const values = { lookIn: 'values' as const };
const whole = { matchWhole: true };
const byte = { matchWidth: true };

const FIND: Case[] = [
  ['apple', formulas, ['A1', 'B1', 'A2', 'C3', 'A13']],
  ['apple', { ...formulas, ...whole }, ['A1', 'B1', 'A2', 'C3']],
  ['apple', { ...formulas, caseSensitive: true }, ['A1', 'B1', 'C3', 'A13']],
  ['apple', { ...values, ...whole }, ['A1', 'B1', 'D1', 'A2', 'C3']],
  ['ple', formulas, ['A1', 'B1', 'D1', 'A2', 'C3', 'A13']],
  ['ple', values, ['A1', 'B1', 'D1', 'A2', 'C3', 'A13']],
  ['3', formulas, ['B2', 'B3']],
  ['3', values, ['B2', 'B3', 'A12']],
  ['3', { ...formulas, ...whole }, ['B2']],
  ['=1+2', formulas, ['A12']],
  ['1+2', formulas, ['A12']],
  ['a?b', formulas, ['A3', 'A4', 'A6']],
  ['a?b', { ...formulas, ...whole }, ['A3', 'A4', 'A6']],
  ['a*b', formulas, ['A3', 'A4', 'A5', 'A6', 'A7', 'A8']],
  ['a~*b', formulas, ['A3']],
  ['a~?b', formulas, ['A4']],
  ['a~~b', formulas, ['A6']],
  ['?', { ...formulas, ...whole }, ['B2', 'A15']],
  [
    '*',
    { ...formulas, ...whole },
    [
      'A1',
      'B1',
      'D1',
      'A2',
      'B2',
      'A3',
      'B3',
      'C3',
      'A4',
      'A5',
      'A6',
      'A7',
      'A8',
      'A9',
      'A10',
      'A11',
      'A12',
      'A13',
      'A14',
      'A15',
    ],
  ],
  ['ABC', formulas, ['A7', 'A8']],
  ['ABC', { ...formulas, ...byte }, ['A7', 'A8']],
  ['ＡＢＣ', formulas, ['A7', 'A8']],
  ['ＡＢＣ', { ...formulas, ...byte }, ['A7', 'A8']],
  ['ｱｲｳ', formulas, ['A9', 'A10', 'A11']],
  ['ｱｲｳ', { ...formulas, ...byte }, ['A9', 'A10', 'A11']],
  ['アイウ', formulas, ['A9', 'A10', 'A11']],
  ['アイウ', { ...formulas, ...byte }, ['A9', 'A10', 'A11']],
  ['あいう', formulas, ['A9', 'A10', 'A11']],
  ['あいう', { ...formulas, ...byte }, ['A9', 'A10', 'A11']],
  ['ガ', formulas, ['A14', 'A15']],
  ['ガ', { ...formulas, ...byte }, ['A15']],
  ['ｶﾞ', formulas, ['A14', 'A15']],
  ['ｶﾞ', { ...formulas, ...byte }, ['A14']],
  ['apple', { ...formulas, searchBy: 'columns' }, ['A1', 'A2', 'A13', 'B1', 'C3']],
];

type ReplaceCase = [
  query: string,
  replacement: string,
  opts: Partial<FindOptions>,
  changed: Record<string, string>,
];

const REPLACE: ReplaceCase[] = [
  ['apple', 'X', formulas, { A1: 'X', B1: 'X', A2: 'X', C3: 'X', A13: 'X pie' }],
  ['apple', 'X', { ...formulas, ...whole }, { A1: 'X', B1: 'X', A2: 'X', C3: 'X' }],
  ['apple', 'X', { ...formulas, caseSensitive: true }, { A1: 'X', B1: 'X', C3: 'X', A13: 'X pie' }],
  [
    'ple',
    'Q',
    formulas,
    { A1: 'apQ', B1: 'apQ', A2: 'ApQ', C3: 'apQ', A13: 'apQ pie', D1: '="ap"&"Q"' },
  ],
  ['a~*b', 'Q', formulas, { A3: 'Q' }],
  ['a?b', 'Q', { ...formulas, ...whole }, { A3: 'Q', A4: 'Q', A6: 'Q' }],
  [
    '*',
    'Z',
    { ...formulas, ...whole },
    Object.fromEntries(Object.keys(FIXTURE).map((ref) => [ref, 'Z'])),
  ],
];

describe('Find matches the desktop app', () => {
  let wb: WorkbookHandle;
  let store: SpreadsheetStore;

  beforeEach(async () => {
    wb = await WorkbookHandle.createDefault();
    for (const [ref, input] of Object.entries(FIXTURE)) {
      const addr = parse(ref);
      if (typeof input === 'number') wb.setNumber(addr, input);
      else if (input.startsWith('=')) wb.setFormula(addr, input);
      else wb.setText(addr, input);
    }
    wb.recalc();
    store = createSpreadsheetStore();
    const cells = [...wb.cells(0)];
    store.setState((s) => ({
      ...s,
      data: { ...s.data, cells: new Map(cells.map((entry) => [addrKey(entry.addr), entry])) },
    }));
  });

  afterEach(() => wb.dispose());

  it.each(FIND)('finds %s with %o', (query, opts, hits) => {
    const found = findAll(store.getState(), { query, workbook: wb, ...opts });
    expect(found.map((m) => a1(m.addr))).toEqual(hits);
  });

  it.each(REPLACE)('replaces %s with %s (%o)', (query, replacement, opts, changed) => {
    const before = Object.fromEntries(Object.keys(FIXTURE).map((ref) => [ref, content(wb, ref)]));
    const count = replaceAll(store.getState(), wb, { query, workbook: wb, ...opts }, replacement);
    expect(count).toBe(Object.keys(changed).length);
    for (const ref of Object.keys(FIXTURE)) {
      expect(content(wb, ref), ref).toBe(changed[ref] ?? before[ref]);
    }
  });
});

const content = (wb: WorkbookHandle, ref: string): string => {
  const addr = parse(ref);
  const formula = wb.cellFormula(addr);
  if (formula) return formula;
  const value = wb.getValue(addr);
  return value.kind === 'blank' ? '' : String('value' in value ? value.value : '');
};
