import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  buildSelectionInputBatch,
  writeSelectionInput,
} from '../../../src/interact/selection-input.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/index.js';

const A1 = { sheet: 0, row: 0, col: 0 };
const A2 = { sheet: 0, row: 1, col: 0 };
const D4 = { sheet: 0, row: 3, col: 3 };
const E4 = { sheet: 0, row: 3, col: 4 };
const stopFive = {
  kind: 'whole' as const,
  op: '=' as const,
  a: 5,
  errorStyle: 'stop' as const,
  errorMessage: 'Value must be five.',
  showErrorMessage: true,
};

describe('interact/selection-input — writeSelectionInput', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet({ workbook: await WorkbookHandle.createDefault() });
  });

  afterEach(() => sheet.dispose());

  const run = (raw: string) => {
    const store = sheet.instance.store;
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 });
    return writeSelectionInput(sheet.workbook, store, store.getState(), raw, A1);
  };

  it('writes nothing when a later cell rejects the input', () => {
    const store = sheet.instance.store;
    mutators.setCellFormat(store, A2, { validation: stopFive });
    const result = run('4');
    expect(result).toEqual({
      status: 'rejected',
      outcome: expect.objectContaining({ severity: 'stop', message: 'Value must be five.' }),
    });
    expect(sheet.workbook.getValue(A1)).toEqual({ kind: 'blank' });
    expect(sheet.workbook.getValue(A2)).toEqual({ kind: 'blank' });
  });

  it('writes every cell when the input satisfies the rules', () => {
    mutators.setCellFormat(sheet.instance.store, A2, { validation: stopFive });
    expect(run('5')).toEqual({ status: 'applied' });
    expect(sheet.workbook.getValue(A1)).toEqual({ kind: 'number', value: 5 });
    expect(sheet.workbook.getValue(A2)).toEqual({ kind: 'number', value: 5 });
  });

  it('keeps R1C1 formulas relative to each destination in Ctrl+Enter batches', () => {
    const store = sheet.instance.store;
    mutators.setR1C1(store, true);
    mutators.setRange(store, { sheet: 0, r0: D4.row, c0: D4.col, r1: E4.row, c1: E4.col });

    expect(buildSelectionInputBatch(store.getState(), '=RC5', D4)?.changes).toEqual([
      { addr: D4, input: '=RC5' },
      { addr: E4, input: '=RC5' },
    ]);

    expect(writeSelectionInput(sheet.workbook, store, store.getState(), '=RC5', D4)).toEqual({
      status: 'applied',
    });
    expect(sheet.workbook.cellFormula(D4)).toBe('=$E4');
    expect(sheet.workbook.cellFormula(E4)).toBe('=$E4');
  });

  it('shifts a relative R1C1 formula across Ctrl+Enter destinations', () => {
    const store = sheet.instance.store;
    mutators.setR1C1(store, true);
    mutators.setRange(store, { sheet: 0, r0: D4.row, c0: D4.col, r1: E4.row, c1: E4.col });

    expect(writeSelectionInput(sheet.workbook, store, store.getState(), '=RC[-1]', D4)).toEqual({
      status: 'applied',
    });
    expect(sheet.workbook.cellFormula(D4)).toBe('=C4');
    expect(sheet.workbook.cellFormula(E4)).toBe('=D4');
  });

  it('retains relative offsets when an invalid anchor reference becomes valid at the next target', () => {
    mutators.setR1C1(sheet.instance.store, true);

    expect(run('=R[-1]C')).toEqual({ status: 'applied' });
    expect(sheet.workbook.cellFormula(A1)).toBe('=#REF!');
    expect(sheet.workbook.cellFormula(A2)).toBe('=A1');
  });

  it.each([
    ['windows-de_DE', '=SUMME(1,5;A1)'],
    ['windows-fr_FR', '=SOMME(1,5;A1)'],
  ] as const)(
    'canonicalizes localized formulas before shifting a multi-cell fill (%s)',
    (profile, raw) => {
      expect(sheet.workbook.isStub).toBe(false);
      expect(sheet.workbook.capabilities.formulaLocale).toBe(true);
      expect(sheet.workbook.capabilities.localeFacts).toBe(true);
      expect(sheet.workbook.setSpreadsheetProfileId(profile)).toBe(true);
      const store = sheet.instance.store;
      mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 });

      expect(writeSelectionInput(sheet.workbook, store, store.getState(), raw, A1)).toEqual({
        status: 'applied',
      });
      expect(sheet.workbook.cellFormula(A1)).toBe('=SUM(1.5,A1)');
      expect(sheet.workbook.cellFormula(A2)).toBe('=SUM(1.5,A2)');
    },
  );

  it('skips a protected cell instead of rejecting the fill', () => {
    const store = sheet.instance.store;
    mutators.setCellFormat(store, A1, { locked: false });
    mutators.setCellFormat(store, A2, { validation: stopFive });
    setProtectedSheet(store, 0, true, { workbook: sheet.workbook });
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    expect(run('4')).toEqual({ status: 'applied' });
    warn.mockRestore();
    expect(sheet.workbook.getValue(A1)).toEqual({ kind: 'number', value: 4 });
    expect(sheet.workbook.getValue(A2)).toEqual({ kind: 'blank' });
  });
});
