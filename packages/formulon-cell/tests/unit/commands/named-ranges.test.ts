import { describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  createDefinedNamesFromSelection,
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
  definedNames(): IterableIterator<{ name: string; formula: string }>;
  setDefinedNameEntry(name: string, formula: string): boolean;
  setFormula?(addr: Addr, formula: string): void;
  recalc(): void;
}

const makeWb = (
  canMutate = true,
  writeOk = true,
): {
  wb: WorkbookHandle;
  calls: { name: string; formula: string }[];
  formulas: Map<string, string>;
  recalcs: () => number;
} => {
  const registry = new Map([['TaxRate', '=Sheet1!$A$1']]);
  const calls: { name: string; formula: string }[] = [];
  const formulas = new Map<string, string>();
  let recalcCount = 0;
  const fake: MutableWb = {
    capabilities: { definedNameMutate: canMutate },
    *definedNames() {
      for (const [name, formula] of registry) yield { name, formula };
    },
    setDefinedNameEntry(name, formula) {
      calls.push({ name, formula });
      if (!writeOk) return false;
      if (formula === '') registry.delete(name);
      else registry.set(name, formula);
      return true;
    },
    setFormula(addr, formula) {
      formulas.set(`${addr.sheet}:${addr.row}:${addr.col}`, formula);
    },
    recalc() {
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
    expect(listDefinedNames(wb)).toEqual([{ name: 'TaxRate', formula: '=Sheet1!$A$1' }]);
  });

  it('adds or replaces a defined name', () => {
    const { wb, calls } = makeWb();
    const result = upsertDefinedName(wb, ' Total ', ' =Sheet1!$B$1 ');

    expect(result).toEqual({ ok: true, entry: { name: 'Total', formula: '=Sheet1!$B$1' } });
    expect(calls).toEqual([{ name: 'Total', formula: '=Sheet1!$B$1' }]);
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
      entries: [{ name: '_A1', formula: '=$A$2:$A$3' }],
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

    expect(result).toEqual({ ok: true, entry: { name: 'TaxRate', formula: '' } });
    expect(calls).toEqual([{ name: 'TaxRate', formula: '' }]);
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
        { name: 'Sales_Total', formula: '=$A$2:$A$3' },
        { name: '_2026_Rate', formula: '=$B$2:$B$3' },
      ],
    });
    expect(calls).toEqual([
      { name: 'Sales_Total', formula: '=$A$2:$A$3' },
      { name: '_2026_Rate', formula: '=$B$2:$B$3' },
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
        { name: 'Sales_Total', formula: '=$A$1:$A$2' },
        { name: 'Tax_Rate', formula: '=$B$1:$B$2' },
      ],
    });
    expect(calls).toEqual([
      { name: 'Sales_Total', formula: '=$A$1:$A$2' },
      { name: 'Tax_Rate', formula: '=$B$1:$B$2' },
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
        { name: 'North', formula: '=$B$1:$C$1' },
        { name: 'South', formula: '=$B$2:$C$2' },
      ],
    });

    const right = makeWb();
    seedText(store, 0, 2, 'West');
    seedText(store, 1, 2, 'East');

    expect(createDefinedNamesFromSelection(store.getState(), right.wb, 'right-column')).toEqual({
      ok: true,
      entries: [
        { name: 'West', formula: '=$A$1:$B$1' },
        { name: 'East', formula: '=$A$2:$B$2' },
      ],
    });
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
    expect([...wb.definedNames()]).toEqual([{ name: 'TaxRate', formula: '=Sheet1!$A$1' }]);

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
        for (const [name, formula] of registry) yield { name, formula };
      },
      setDefinedNameEntry(name: string, formula: string) {
        if (formula === '') registry.delete(name);
        else registry.set(name, formula);
        return true;
      },
      recalc() {},
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
