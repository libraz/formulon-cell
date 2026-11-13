import { describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  definedNameDisplayFormula,
  definedNameExists,
  definedNameValueText,
  deleteDefinedName,
  insertDefinedNameFormula,
  isValidDefinedName,
  listDefinedNames,
  planDefinedNamesFromSelection,
  recordDefinedNamesChange,
  sanitizeName,
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
  sheetName(sheet: number): string;
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
    sheetName: (sheet) => ['Sheet1', 'My Sheet'][sheet] ?? '',
  };
  return { wb: fake as unknown as WorkbookHandle, calls, formulas, recalcs: () => recalcCount };
};

const seedText = (
  store: SpreadsheetStore,
  row: number,
  col: number,
  value: string,
  sheet = 0,
): void => {
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    cells.set(`${sheet}:${row}:${col}`, { value: { kind: 'text', value }, formula: null });
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
      sheetIndexByName: (name: string) =>
        ({ sheet1: 0, 'my sheet': 1 })[name.toLowerCase()] ?? null,
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

  it('records defined-name mutations as one undoable history action', () => {
    const { wb } = makeWb();
    const history = new History();

    recordDefinedNamesChange(history, wb, () => {
      upsertDefinedName(wb, 'Sales_Total', '=Sheet1!$A$2:$A$3');
      upsertDefinedName(wb, '_2026_Rate', '=Sheet1!$B$2:$B$3');
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

// Labels and the names the desktop's Create from Selection gave them, measured over every edge.
const MEASURED_LABELS: readonly (readonly [string, string])[] = [
  ['Q1', 'Q1_'],
  ['A1', 'A1_'],
  ['XFD1048576', 'XFD1048576_'],
  ['XFE1', 'XFE1'],
  ['R1C1', '_R1C1'],
  ['R', 'R_'],
  ['C', 'C_'],
  ['r', 'r_'],
  ['2020', '_2020'],
  ['1abc', '_1abc'],
  ['a b', 'a_b'],
  ['Sales-Total', 'Sales_Total'],
  ['Sales.Total', 'Sales.Total'],
  ['TRUE', 'TRUE_'],
  ['_x', '_x'],
  ['Ä1', 'Ä1'],
  ['売上', '売上'],
  ['Q1 ', 'Q1_'],
  ['=x', 'x'],
];

describe('Create from Selection names', () => {
  it.each(MEASURED_LABELS)('names the text label %j as %j', (label, name) => {
    expect(sanitizeName(label)).toBe(name);
    expect(isValidDefinedName(name)).toBe(true);
  });

  it('gives no name for a blank label', () => {
    expect(sanitizeName('')).toBe('');
    expect(sanitizeName('   ')).toBe('');
  });

  it.each([
    ['top', { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 }],
    ['left', { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 }],
  ] as const)('names every measured label the same way on the %s edge', (edge, range) => {
    const edges = { top: edge === 'top', left: edge === 'left', bottom: false, right: false };
    for (const [label, name] of MEASURED_LABELS) {
      const store = createSpreadsheetStore();
      seedText(store, 0, 0, label);
      const planned = planDefinedNamesFromSelection(store.getState(), makeWb().wb, range, edges);
      expect(planned.ok && planned.names.map((n) => n.name), label).toEqual([name]);
    }
  });

  it('defines the measured header row: later labels win, numbers, booleans and blanks give none', () => {
    const store = createSpreadsheetStore();
    const headers = ['Q1', 'A1', 'XFD1048576', 'XFE1', 'R1C1', 'R', 'C', 'r', 2020, '1abc', 'a b'];
    const more = [
      'Sales-Total',
      'Sales.Total',
      true,
      null,
      '_x',
      'Ä1',
      '売上',
      'Q1 ',
      'Dup',
      'Dup',
    ];
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      for (const [col, label] of [...headers, ...more].entries()) {
        if (label === null) continue;
        const value =
          typeof label === 'number'
            ? { kind: 'number' as const, value: label }
            : typeof label === 'boolean'
              ? { kind: 'bool' as const, value: label }
              : { kind: 'text' as const, value: label };
        cells.set(`0:0:${col}`, { value, formula: null });
      }
      return { ...s, data: { ...s.data, cells } };
    });
    const planned = planDefinedNamesFromSelection(
      store.getState(),
      makeWb().wb,
      { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 20 },
      { top: true, left: false, bottom: false, right: false },
    );
    const measured = [
      ['_1abc', '=Sheet1!$J$2:$J$3'],
      ['_R1C1', '=Sheet1!$E$2:$E$3'],
      ['_x', '=Sheet1!$P$2:$P$3'],
      ['a_b', '=Sheet1!$K$2:$K$3'],
      ['Ä1', '=Sheet1!$Q$2:$Q$3'],
      ['A1_', '=Sheet1!$B$2:$B$3'],
      ['C_', '=Sheet1!$G$2:$G$3'],
      ['Dup', '=Sheet1!$U$2:$U$3'],
      ['Q1_', '=Sheet1!$S$2:$S$3'],
      ['r_', '=Sheet1!$H$2:$H$3'],
      ['Sales_Total', '=Sheet1!$L$2:$L$3'],
      ['Sales.Total', '=Sheet1!$M$2:$M$3'],
      ['XFD1048576_', '=Sheet1!$C$2:$C$3'],
      ['XFE1', '=Sheet1!$D$2:$D$3'],
      ['売上', '=Sheet1!$R$2:$R$3'],
    ];
    const byName = (a: string[], b: string[]) => (a[0] ?? '').localeCompare(b[0] ?? '');
    expect(planned.ok && planned.names.map((n) => [n.name, n.formula]).sort(byName)).toEqual(
      measured.sort(byName),
    );
  });

  it('leaves the other checked edges out of each name, qualified by the sheet', () => {
    const store = createSpreadsheetStore();
    seedText(store, 0, 1, 'Qty', 1);
    seedText(store, 0, 2, 'Price', 1);
    seedText(store, 1, 0, 'North', 1);
    seedText(store, 2, 0, 'South', 1);
    const planned = planDefinedNamesFromSelection(
      store.getState(),
      makeWb().wb,
      { sheet: 1, r0: 0, c0: 0, r1: 2, c1: 2 },
      { top: true, left: true, bottom: false, right: false },
    );
    expect(planned).toEqual({
      ok: true,
      names: [
        { name: 'Qty', formula: "='My Sheet'!$B$2:$B$3" },
        { name: 'Price', formula: "='My Sheet'!$C$2:$C$3" },
        { name: 'North', formula: "='My Sheet'!$B$2:$C$2" },
        { name: 'South', formula: "='My Sheet'!$B$3:$C$3" },
      ],
    });
  });

  it('names from the bottom row and right column, and gives a single cell its own reference', () => {
    const store = createSpreadsheetStore();
    seedText(store, 1, 0, 'Total');
    seedText(store, 0, 1, 'East');
    const range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    const wb = makeWb().wb;
    expect(
      planDefinedNamesFromSelection(store.getState(), wb, range, {
        top: false,
        left: false,
        bottom: true,
        right: false,
      }),
    ).toEqual({ ok: true, names: [{ name: 'Total', formula: '=Sheet1!$A$1' }] });
    expect(
      planDefinedNamesFromSelection(store.getState(), wb, range, {
        top: false,
        left: false,
        bottom: false,
        right: true,
      }),
    ).toEqual({ ok: true, names: [{ name: 'East', formula: '=Sheet1!$A$1' }] });
  });

  it('resolves a name two edges give to the later edge', () => {
    const store = createSpreadsheetStore();
    seedText(store, 0, 1, 'Total');
    seedText(store, 1, 0, 'TOTAL');
    const planned = planDefinedNamesFromSelection(
      store.getState(),
      makeWb().wb,
      { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
      { top: true, left: true, bottom: false, right: false },
    );
    expect(planned).toEqual({ ok: true, names: [{ name: 'TOTAL', formula: '=Sheet1!$B$2' }] });
  });

  it('refuses huge whole-row or whole-column selections', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeWb();
    expect(
      planDefinedNamesFromSelection(
        store.getState(),
        wb,
        { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 16_383 },
        { top: true, left: false, bottom: false, right: false },
      ),
    ).toEqual({ ok: false, reason: 'too-large' });
    expect(
      planDefinedNamesFromSelection(
        store.getState(),
        wb,
        { sheet: 0, r0: 0, c0: 0, r1: 1_048_575, c1: 1 },
        { top: false, left: true, bottom: false, right: false },
      ),
    ).toEqual({ ok: false, reason: 'too-large' });
  });
});
