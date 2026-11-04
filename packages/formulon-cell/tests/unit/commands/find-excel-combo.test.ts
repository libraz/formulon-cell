import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { type FindOptions, findAll, replaceAll } from '../../../src/commands/find.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

// Results measured with Find / Replace All in the desktop app on this fixture.
const FIXTURE: Record<string, string | number> = {
  A1: 'Alpha',
  B1: 1234,
  C1: 1234.5,
  A2: 'alpha',
  B2: 12345,
  C2: 1200,
  A3: 'ALPHA',
  B3: 12,
  C3: 0.255,
  A4: 'Alphabet',
  B4: 123.45,
  A5: 'a*b',
  C5: 12345.678,
  B6: 120,
  F1: '=1/0',
  F4: '=IF(B3>0,1/0,0)',
  E1: '=CONCATENATE("Alpha","x")',
  E8: '=TEXT(B1,"#,##0")',
};

const parse = (a1: string) => {
  const m = /^([A-Z])(\d+)$/.exec(a1);
  if (!m) throw new Error(a1);
  return { sheet: 0, row: Number(m[2]) - 1, col: (m[1] ?? 'A').charCodeAt(0) - 65 };
};
const a1 = (addr: { row: number; col: number }) =>
  `${String.fromCharCode(65 + addr.col)}${addr.row + 1}`;

describe('Find and Replace All match the desktop app', () => {
  let wb: WorkbookHandle;
  let store: SpreadsheetStore;

  const select = (range: string) => {
    const [from = 'A1', to = from] = range.split(':');
    const f = parse(from);
    const t = parse(to);
    store.setState((st) => ({
      ...st,
      selection: {
        active: f,
        anchor: f,
        range: { sheet: 0, r0: f.row, c0: f.col, r1: t.row, c1: t.col },
      },
    }));
  };
  const find = (opts: Partial<FindOptions> & { query: string }) =>
    findAll(store.getState(), { workbook: wb, locale: 'en-US', listFromActive: true, ...opts }).map(
      (m) => a1(m.addr),
    );
  const content = (ref: string): string => {
    const addr = parse(ref);
    const formula = wb.cellFormula(addr);
    if (formula) return formula;
    const value = wb.getValue(addr);
    return value.kind === 'blank' ? '' : String('value' in value ? value.value : '');
  };

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

  it('searches only a multi-cell selection, starting after its first cell', () => {
    select('A2:B4');
    expect(find({ query: 'alpha', lookIn: 'values' })).toEqual(['A3', 'A4', 'A2']);
    expect(find({ query: '12?', lookIn: 'values', caseSensitive: true })).toEqual(['B2', 'B4']);
  });

  it('searches the whole sheet for a single-cell selection', () => {
    select('A2');
    expect(find({ query: 'alpha', lookIn: 'values' })).toEqual(['A3', 'A4', 'A1', 'E1', 'A2']);
  });

  it('lists the active cell last', () => {
    expect(find({ query: 'alpha', lookIn: 'values' })).toEqual(['E1', 'A2', 'A3', 'A4', 'A1']);
  });

  it('matches the stored text of a constant under Look in formulas, not its display', () => {
    expect(find({ query: '1,2?0', lookIn: 'formulas' })).toEqual([]);
    expect(find({ query: '1,2*0', lookIn: 'formulas', caseSensitive: true })).toEqual([]);
    expect(find({ query: '12~3', lookIn: 'formulas' })).toEqual(['B1', 'C1', 'B2', 'B4', 'C5']);
  });

  it('replaces through a tilde that escapes nothing', () => {
    replaceAll(store.getState(), wb, { query: '12~3', lookIn: 'formulas', workbook: wb }, '9');
    expect(content('B1')).toBe('94');
    expect(content('C1')).toBe('94.5');
    expect(content('C5')).toBe('945.678');
  });

  it('replaces the shortest text a star spans and skips results that are not formulas', () => {
    const opts = { query: '1*0', lookIn: 'formulas' as const, workbook: wb };
    const n = replaceAll(store.getState(), wb, opts, '0');
    expect(content('F1')).toBe('=0');
    expect(content('F4')).toBe('=IF(B3>0,0,0)');
    expect(content('C2')).toBe('0');
    expect(content('B6')).toBe('0');
    expect(content('E8')).toBe('=TEXT(B1,"#,##0")');
    expect(n).toBe(4);
  });

  it('refuses a replacement that breaks a formula without throwing', () => {
    const opts = { query: 'A*x', lookIn: 'formulas' as const, caseSensitive: true, workbook: wb };
    expect(replaceAll(store.getState(), wb, opts, 'Q')).toBe(0);
    expect(content('E1')).toBe('=CONCATENATE("Alpha","x")');
  });
});
