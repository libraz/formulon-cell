import { describe, expect, it } from 'vitest';
import {
  deleteDefinedName,
  listDefinedNames,
  upsertDefinedName,
} from '../../../src/commands/named-ranges.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

interface MutableWb {
  capabilities: { definedNameMutate: boolean };
  definedNames(): IterableIterator<{ name: string; formula: string }>;
  setDefinedNameEntry(name: string, formula: string): boolean;
}

const makeWb = (
  canMutate = true,
  writeOk = true,
): { wb: WorkbookHandle; calls: { name: string; formula: string }[] } => {
  const registry = new Map([['TaxRate', '=Sheet1!$A$1']]);
  const calls: { name: string; formula: string }[] = [];
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
  };
  return { wb: fake as unknown as WorkbookHandle, calls };
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

  it('validates empty names and references before calling the engine', () => {
    const { wb, calls } = makeWb();

    expect(upsertDefinedName(wb, ' ', '=A1')).toEqual({ ok: false, reason: 'empty-name' });
    expect(upsertDefinedName(wb, 'Name', ' ')).toEqual({
      ok: false,
      reason: 'empty-formula',
    });
    expect(calls).toEqual([]);
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
});
