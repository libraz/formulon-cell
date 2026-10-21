import { describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  createDefinedNamesFromSelection,
  definedNameDisplayFormula,
  definedNameExists,
  definedNameValueText,
  deleteDefinedName,
  insertDefinedNameFormula,
  isValidDefinedName,
  listDefinedNames,
  recordDefinedNamesChange,
  upsertDefinedName,
} from '../../../src/commands/named-ranges.js';
import type { Addr } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

interface MutableWb {
  capabilities: { definedNameMutate: boolean };
  definedNames(): IterableIterator<{ name: string; formula: string; localSheetId: number }>;
  setDefinedNameEntry(name: string, formula: string, localSheetId?: number): boolean;
  setFormula?(addr: Addr, formula: string): void;
  recalc(): void;
  recalcAuto(): void;
}

const makeWb = (
  canMutate = true,
  writeOk = true,
): {
  wb: WorkbookHandle;
  calls: { name: string; formula: string; localSheetId: number }[];
  formulas: Map<string, string>;
  recalcs: () => number;
} => {
  const registry = new Map([['-1:TaxRate', { name: 'TaxRate', formula: '=Sheet1!$A$1' }]]);
  const calls: { name: string; formula: string; localSheetId: number }[] = [];
  const formulas = new Map<string, string>();
  let recalcCount = 0;
  const fake: MutableWb = {
    capabilities: { definedNameMutate: canMutate },
    *definedNames() {
      for (const [key, entry] of registry) {
        const localSheetId = Number.parseInt(key.split(':', 1)[0] ?? '-1', 10);
        yield { name: entry.name, formula: entry.formula, localSheetId };
      }
    },
    setDefinedNameEntry(name, formula, localSheetId = -1) {
      calls.push({ name, formula, localSheetId });
      if (!writeOk) return false;
      const key = `${localSheetId}:${name}`;
      if (formula === '') registry.delete(key);
      else registry.set(key, { name, formula });
      return true;
    },
    setFormula(addr, formula) {
      formulas.set(`${addr.sheet}:${addr.row}:${addr.col}`, formula);
    },
    recalc() {
      recalcCount += 1;
    },
    recalcAuto() {
      recalcCount += 1;
    },
  };
  return { wb: fake as unknown as WorkbookHandle, calls, formulas, recalcs: () => recalcCount };
};

const seedText = (store: SpreadsheetStore, row: number, col: number, value: string): void => {
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    cells.set(`0:${row}:${col}`, { value: { kind: 'text', value }, formula: null });
    return { ...s, data: { ...s.data, cells } };
  });
};

const setRange = (
  store: SpreadsheetStore,
  r0: number,
  c0: number,
  r1: number,
  c1: number,
): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      active: { sheet: 0, row: r0, col: c0 },
      anchor: { sheet: 0, row: r0, col: c0 },
      range: { sheet: 0, r0, c0, r1, c1 },
    },
  }));
};

describe('named range commands', () => {
  it('lists defined names from the workbook handle', () => {
    const { wb } = makeWb();
    expect(listDefinedNames(wb)).toEqual([
      { name: 'TaxRate', formula: '=Sheet1!$A$1', localSheetId: -1 },
    ]);
  });

  it('displays stored defined-name formulas with one leading =', () => {
    expect(definedNameDisplayFormula('Sheet1!$A$1')).toBe('=Sheet1!$A$1');
    expect(definedNameDisplayFormula('=Sheet1!$A$1')).toBe('=Sheet1!$A$1');
    expect(definedNameDisplayFormula('')).toBe('');
  });

  it('adds or replaces a defined name', () => {
    const { wb, calls } = makeWb();
    const result = upsertDefinedName(wb, ' Total ', ' =Sheet1!$B$1 ');

    expect(result).toEqual({
      ok: true,
      entry: { name: 'Total', formula: 'Sheet1!$B$1', localSheetId: -1 },
    });
    expect(calls).toEqual([{ name: 'Total', formula: '=Sheet1!$B$1', localSheetId: -1 }]);
  });

  it('recalculates dependents after add/delete so values are not stale (H-38)', () => {
    const added = makeWb();
    upsertDefinedName(added.wb, 'Total', '=Sheet1!$B$1');
    expect(added.recalcs()).toBe(1);

    const removed = makeWb();
    deleteDefinedName(removed.wb, 'TaxRate');
    expect(removed.recalcs()).toBe(1);

    // Failed / no-op mutations must not recalc.
    const failed = makeWb(true, false);
    upsertDefinedName(failed.wb, 'X', '=A1');
    expect(failed.recalcs()).toBe(0);
  });

  it('validates empty names and references before calling the engine', () => {
    const { wb, calls } = makeWb();

    expect(upsertDefinedName(wb, ' ', '=A1')).toEqual({ ok: false, reason: 'empty-name' });
    expect(upsertDefinedName(wb, 'Name', ' ')).toEqual({
      ok: false,
      reason: 'empty-formula',
    });
    expect(calls).toEqual([]);
  });

  it('rejects names that break spreadsheet naming rules before the engine (H-33)', () => {
    const { wb, calls } = makeWb();

    // Cell-reference collisions (A1 and R1C1 style) and reserved single letters.
    for (const bad of ['A1', '$A$1', 'XFD1', 'B$2', 'R1C1', 'RC', 'R', 'C']) {
      expect(upsertDefinedName(wb, bad, '=Sheet1!$A$1')).toEqual({
        ok: false,
        reason: 'invalid-name',
      });
    }
    // Illegal leading character / illegal body characters / spaces.
    for (const bad of ['1Name', 'a-b', 'a b', 'a!']) {
      expect(upsertDefinedName(wb, bad, '=Sheet1!$A$1')).toEqual({
        ok: false,
        reason: 'invalid-name',
      });
    }
    expect(calls).toEqual([]);
  });

  it('accepts valid names that merely resemble references', () => {
    const { wb } = makeWb();
    // Trailing letters keep these out of the A1 grid, so they are valid names.
    for (const good of ['Sales', '_2026', 'Revenue', 'Q1_2026', 'C3PO', 'Region.North']) {
      expect(upsertDefinedName(wb, good, '=Sheet1!$A$1').ok).toBe(true);
    }
  });

  it('isValidDefinedName mirrors the engine-facing rules', () => {
    expect(isValidDefinedName('Sales')).toBe(true);
    expect(isValidDefinedName('_hidden')).toBe(true);
    expect(isValidDefinedName('A1')).toBe(false);
    expect(isValidDefinedName('R')).toBe(false);
    expect(isValidDefinedName('a b')).toBe(false);
    expect(isValidDefinedName('a'.repeat(256))).toBe(false);
  });

  it.each([
    '売上',
    '税率',
    'a?b',
    'ｱｲｳ',
    '売上1',
    '_x.y',
    '\\ab',
    'a\\b',
    '①',
    'ー長',
    'Ⅳ',
    '々',
    '々々',
    'A1売',
    'é',
    'ä1',
    '×a',
    '→a',
    '😀a',
    'a😀',
    'ａｂｃ',
    'Sheet1',
    '「a',
    '・a',
    'a・b',
    '・',
    '“a',
    '…a',
    '†a',
    '『a',
    'ａ．ｂ.',
    'a。b',
    '  売上\u3000',
    '売上１',
    'abcd１',
    'ａ１ｂ',
    'x１y',
    '_１',
  ])('accepts %j as a defined name', (name) => {
    expect(isValidDefinedName(name)).toBe(true);
  });

  it.each([
    '?ab',
    'Ａ１',
    '1売上',
    '売 上',
    '.ab',
    'ＲＣ',
    'r',
    '売上!',
    'a\u3000b',
    '、a',
    'a、b',
    '＋a',
    'a＋b',
    '１a',
    'a１',
    'Tax２０２６',
    '€a',
    'a-b',
    'a$b',
    'a@b',
    'a%b',
    '～a',
    '。a',
    '¥a',
    'a¥',
    '£a',
    '〜a',
    '￥a',
    'FALSE',
    'true',
    'ab\u3000c',
    '$a',
    '（a',
    'a）',
    '',
    '   ',
  ])('rejects %j as a defined name', (name) => {
    expect(isValidDefinedName(name)).toBe(false);
  });

  it('keeps Japanese labels when creating names from the selection', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeWb();
    seedText(store, 0, 0, '売上 合計');
    seedText(store, 0, 1, 'TRUE');
    seedText(store, 0, 2, '1月');
    seedText(store, 0, 3, 'a-b');
    setRange(store, 0, 0, 2, 3);
    const result = createDefinedNamesFromSelection(store.getState(), wb, 'top-row');
    expect(result.ok && result.entries.map((e) => e.name)).toEqual([
      '売上_合計',
      '_TRUE',
      '_1月',
      'a_b',
    ]);
    if (result.ok) for (const e of result.entries) expect(isValidDefinedName(e.name)).toBe(true);
  });

  it('detects duplicate names within a scope ignoring case and width', () => {
    const { wb } = makeWb();
    expect(definedNameExists(wb, 'taxrate')).toBe(true);
    expect(definedNameExists(wb, 'ＴａｘＲａｔｅ')).toBe(true);
    expect(definedNameExists(wb, 'TaxRate', 0)).toBe(false);
    expect(definedNameExists(wb, 'Other')).toBe(false);
  });

  it('renders the Value text the way Name Manager does', () => {
    const cells: Record<string, unknown> = {
      '0:0:0': { kind: 'text', value: 'a' },
      '0:0:1': { kind: 'number', value: 2 },
      '0:1:0': { kind: 'text', value: 'c' },
      '0:1:1': { kind: 'number', value: 4 },
      '1:0:0': { kind: 'number', value: 9 },
    };
    const wb = {
      capabilities: {},
      sheetCount: 2,
      sheetName: (i: number) => (i === 0 ? 'Sheet1' : 'My Sheet'),
      getValue: (a: Addr) => cells[`${a.sheet}:${a.row}:${a.col}`] ?? { kind: 'blank' },
    } as unknown as WorkbookHandle;
    const text = (formula: string, localSheetId = -1): string =>
      definedNameValueText(wb, { name: 'N', formula, localSheetId });
    expect(text('=0.08')).toBe('0.08');
    expect(text('="abc"')).toBe('"abc"');
    expect(text('=true')).toBe('TRUE');
    expect(text('=Sheet1!$A$1')).toBe('a');
    expect(text('=Sheet1!$A$1:$B$2')).toBe('{"a",2;"c",4}');
    expect(text("='My Sheet'!$A$1")).toBe('9');
    expect(text('=$A$1', 1)).toBe('9');
    expect(text('=SUM(Sheet1!A1)')).toBe('');
    expect(text('=Sheet1!$A$1:$Z$100').endsWith('…')).toBe(true);
  });

  it('sanitizes header labels that collide with cell references (H-33)', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeWb();
    // A header of "A1" would sanitize to the valid identifier "A1", which is a
    // cell reference — the create path must prefix it so the write succeeds.
    seedText(store, 0, 0, 'A1');
    setRange(store, 0, 0, 2, 0);

    const result = createDefinedNamesFromSelection(store.getState(), wb, 'top-row');
    expect(result).toEqual({
      ok: true,
      entries: [{ name: '_A1', formula: '$A$2:$A$3', localSheetId: -1 }],
    });
  });

  it('reports unsupported and failed writes', () => {
    expect(upsertDefinedName(makeWb(false).wb, 'Name', '=A1')).toEqual({
      ok: false,
      reason: 'unsupported',
    });
    expect(upsertDefinedName(makeWb(true, false).wb, 'Name', '=A1')).toEqual({
      ok: false,
      reason: 'engine-failed',
    });
  });

  it('deletes a defined name via the workbook empty-formula convention', () => {
    const { wb, calls } = makeWb();
    const result = deleteDefinedName(wb, ' TaxRate ');

    expect(result).toEqual({
      ok: true,
      entry: { name: 'TaxRate', formula: '', localSheetId: -1 },
    });
    expect(calls).toEqual([{ name: 'TaxRate', formula: '', localSheetId: -1 }]);
  });

  it('creates defined names from the top row of the selection', () => {
    const store = createSpreadsheetStore();
    const { wb, calls } = makeWb();
    seedText(store, 0, 0, 'Sales Total');
    seedText(store, 0, 1, '2026 Rate');
    setRange(store, 0, 0, 2, 1);

    const result = createDefinedNamesFromSelection(store.getState(), wb, 'top-row');

    expect(result).toEqual({
      ok: true,
      entries: [
        { name: 'Sales_Total', formula: '$A$2:$A$3', localSheetId: -1 },
        { name: '_2026_Rate', formula: '$B$2:$B$3', localSheetId: -1 },
      ],
    });
    expect(calls).toEqual([
      { name: 'Sales_Total', formula: '=$A$2:$A$3', localSheetId: -1 },
      { name: '_2026_Rate', formula: '=$B$2:$B$3', localSheetId: -1 },
    ]);
  });

  it('creates defined names from the bottom row of the selection', () => {
    const store = createSpreadsheetStore();
    const { wb, calls } = makeWb();
    seedText(store, 2, 0, 'Sales Total');
    seedText(store, 2, 1, 'Tax Rate');
    setRange(store, 0, 0, 2, 1);

    const result = createDefinedNamesFromSelection(store.getState(), wb, 'bottom-row');

    expect(result).toEqual({
      ok: true,
      entries: [
        { name: 'Sales_Total', formula: '$A$1:$A$2', localSheetId: -1 },
        { name: 'Tax_Rate', formula: '$B$1:$B$2', localSheetId: -1 },
      ],
    });
    expect(calls).toEqual([
      { name: 'Sales_Total', formula: '=$A$1:$A$2', localSheetId: -1 },
      { name: 'Tax_Rate', formula: '=$B$1:$B$2', localSheetId: -1 },
    ]);
  });

  it('creates defined names from the left and right columns of the selection', () => {
    const store = createSpreadsheetStore();
    const left = makeWb();
    seedText(store, 0, 0, 'North');
    seedText(store, 1, 0, 'South');
    setRange(store, 0, 0, 1, 2);

    expect(createDefinedNamesFromSelection(store.getState(), left.wb, 'left-column')).toEqual({
      ok: true,
      entries: [
        { name: 'North', formula: '$B$1:$C$1', localSheetId: -1 },
        { name: 'South', formula: '$B$2:$C$2', localSheetId: -1 },
      ],
    });

    const right = makeWb();
    seedText(store, 0, 2, 'West');
    seedText(store, 1, 2, 'East');

    expect(createDefinedNamesFromSelection(store.getState(), right.wb, 'right-column')).toEqual({
      ok: true,
      entries: [
        { name: 'West', formula: '$A$1:$B$1', localSheetId: -1 },
        { name: 'East', formula: '$A$2:$B$2', localSheetId: -1 },
      ],
    });
  });

  it('refuses to create defined names from huge whole-row or whole-column selections', () => {
    const store = createSpreadsheetStore();
    const top = makeWb();
    setRange(store, 0, 0, 1, 16_383);

    expect(createDefinedNamesFromSelection(store.getState(), top.wb, 'top-row')).toEqual({
      ok: false,
      reason: 'too-large',
    });
    expect(top.calls).toEqual([]);

    const left = makeWb();
    setRange(store, 0, 0, 1_048_575, 1);

    expect(createDefinedNamesFromSelection(store.getState(), left.wb, 'left-column')).toEqual({
      ok: false,
      reason: 'too-large',
    });
    expect(left.calls).toEqual([]);
  });

  it('records defined-name mutations as one undoable history action', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeWb();
    const history = new History();
    seedText(store, 0, 0, 'Sales Total');
    seedText(store, 0, 1, '2026 Rate');
    setRange(store, 0, 0, 2, 1);

    recordDefinedNamesChange(history, wb, () => {
      createDefinedNamesFromSelection(store.getState(), wb, 'top-row');
    });

    expect([...wb.definedNames()].map((entry) => entry.name)).toEqual([
      'TaxRate',
      'Sales_Total',
      '_2026_Rate',
    ]);
    expect(history.canUndo()).toBe(true);

    history.undo();
    expect([...wb.definedNames()]).toEqual([
      { name: 'TaxRate', formula: '=Sheet1!$A$1', localSheetId: -1 },
    ]);

    history.redo();
    expect([...wb.definedNames()].map((entry) => entry.name)).toEqual([
      'TaxRate',
      'Sales_Total',
      '_2026_Rate',
    ]);
  });

  it('undo removes every defined name created from an initially empty registry', () => {
    const registry = new Map<string, string>();
    const wb = {
      capabilities: { definedNameMutate: true },
      *definedNames() {
        for (const [name, formula] of registry) yield { name, formula, localSheetId: -1 };
      },
      setDefinedNameEntry(name: string, formula: string) {
        if (formula === '') registry.delete(name);
        else registry.set(name, formula);
        return true;
      },
      recalc() {},
      recalcAuto() {},
    } as unknown as WorkbookHandle;
    const history = new History();

    recordDefinedNamesChange(history, wb, () => {
      wb.setDefinedNameEntry('Net_Sales', '=$A$56:$A$57');
      wb.setDefinedNameEntry('Tax_Rate', '=$B$56:$B$57');
    });

    expect([...wb.definedNames()].map((entry) => entry.name)).toEqual(['Net_Sales', 'Tax_Rate']);
    expect(history.undo()).toBe(true);
    expect([...wb.definedNames()]).toEqual([]);
    expect(history.redo()).toBe(true);
    expect([...wb.definedNames()].map((entry) => entry.name)).toEqual(['Net_Sales', 'Tax_Rate']);
  });

  it('inserts a defined name as a formula into the active cell', () => {
    const store = createSpreadsheetStore();
    const { wb, formulas } = makeWb();
    setRange(store, 4, 2, 4, 2);

    const result = insertDefinedNameFormula(store.getState(), wb, ' taxrate ');

    expect(result).toEqual({ addr: { sheet: 0, row: 4, col: 2 }, formula: '=TaxRate' });
    expect(formulas.get('0:4:2')).toBe('=TaxRate');
  });

  it('prefers sheet-scoped defined names over workbook-scoped names when inserting formulas', () => {
    const store = createSpreadsheetStore();
    const { wb, formulas } = makeWb();
    wb.setDefinedNameEntry('taxrate', '=Sheet1!$B$2', 0);
    setRange(store, 4, 2, 4, 2);

    const result = insertDefinedNameFormula(store.getState(), wb, ' taxrate ');

    expect(result).toEqual({ addr: { sheet: 0, row: 4, col: 2 }, formula: '=taxrate' });
    expect(formulas.get('0:4:2')).toBe('=taxrate');
    expect([...wb.definedNames()]).toContainEqual({
      name: 'taxrate',
      formula: '=Sheet1!$B$2',
      localSheetId: 0,
    });
  });

  it('does not insert formulas for names that do not exist in the workbook', () => {
    const store = createSpreadsheetStore();
    const { wb, formulas } = makeWb();
    setRange(store, 4, 2, 4, 2);

    expect(insertDefinedNameFormula(store.getState(), wb, 'MissingName')).toBeNull();
    expect(formulas.size).toBe(0);
  });

  it('blocks defined-name formula insertion into locked protected cells', () => {
    const store = createSpreadsheetStore();
    const { wb, formulas } = makeWb();
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    setRange(store, 4, 2, 4, 2);
    mutators.setSheetProtected(store, 0, true);

    try {
      const result = insertDefinedNameFormula(store.getState(), wb, 'TaxRate', store);

      expect(result).toBeNull();
      expect(formulas.size).toBe(0);
      expect(warn).toHaveBeenCalled();
    } finally {
      warn.mockRestore();
    }
  });
});
