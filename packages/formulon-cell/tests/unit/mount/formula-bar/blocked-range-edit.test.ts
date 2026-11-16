import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { addAllowedEditRange, setProtectedSheet } from '../../../../src/commands/protection.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/index.js';
import { attachFormulaBarHarness, selectRange } from './fixtures.js';

describe('mount/formula-bar - edit blocked by an allowed edit range', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
  });
  afterEach(() => sheet.dispose());

  it('settles a blocked selection commit so a later commit still runs', () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const store = sheet.instance.store;
    addAllowedEditRange(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      { password: 'secret', title: 'Guarded' },
    );
    setProtectedSheet(store, 0, true);
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    selectRange(sheet, { sheet: 0, row: 0, col: 0 }, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.value = 'x';
    harness.fxInput.dispatchEvent(new Event('input'));
    const commit = (): void => {
      harness.fxInput.dispatchEvent(
        new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, cancelable: true }),
      );
    };
    commit();
    expect(warn).toHaveBeenCalled();
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });

    setProtectedSheet(store, 0, false);
    commit();
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'x',
    });
    warn.mockRestore();
    harness.detach();
  });
});
