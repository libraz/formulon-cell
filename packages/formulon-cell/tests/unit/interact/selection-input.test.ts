import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { writeSelectionInput } from '../../../src/interact/selection-input.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/index.js';

const A1 = { sheet: 0, row: 0, col: 0 };
const A2 = { sheet: 0, row: 1, col: 0 };
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
