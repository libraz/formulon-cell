import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { setProtectedSheet } from '../../../../src/commands/protection.js';
import { addrKey } from '../../../../src/engine/address.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/index.js';
import { attachFormulaBarHarness } from './fixtures.js';

describe('mount/formula-bar — native commit result', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    const workbook = await WorkbookHandle.createDefault();
    sheet = await mountStubSheet({ workbook });
    expect(sheet.workbook.isStub).toBe(false);
  });

  afterEach(() => sheet.dispose());

  it('returns true after a native write and keeps the cell format', () => {
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(sheet.instance.store, addr);
    mutators.setCellFormat(sheet.instance.store, addr, { bold: true });
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.value = '42';
    harness.fxInput.dispatchEvent(new Event('input'));

    expect(harness.controller.commitFx('none')).toBe(true);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 42 });
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')?.bold).toBe(true);
    expect(harness.controller.isEditing()).toBe(false);
    harness.detach();
  });

  it('returns false for a protected default-mounted cell and keeps the edit state intact', () => {
    const onValidation = vi.fn();
    const harness = attachFormulaBarHarness(sheet, onValidation);
    const addr = { sheet: 0, row: 0, col: 0 };
    expect(sheet.instance.commands.policy).toBeUndefined();
    setProtectedSheet(sheet.instance.store, 0, true, { workbook: sheet.workbook });
    mutators.setActive(sheet.instance.store, addr);
    mutators.setCellFormat(sheet.instance.store, addr, { italic: true });
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.value = 'denied';
    harness.fxInput.dispatchEvent(new Event('input'));

    expect(harness.controller.commitFx('none')).toBe(false);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'blank' });
    expect(harness.controller.isEditing()).toBe(true);
    expect(document.activeElement).toBe(harness.fxInput);
    expect(sheet.instance.store.getState().format.formats.get(addrKey(addr))).toMatchObject({
      italic: true,
    });
    expect(onValidation).toHaveBeenCalledWith(expect.objectContaining({ severity: 'stop' }));
    harness.detach();
  });

  it('refuses typing into a locked cell of a protected sheet before the edit starts', () => {
    // Excel: the first keystroke in the formula bar raises the protected-cell alert.
    const onValidation = vi.fn();
    const harness = attachFormulaBarHarness(sheet, onValidation);
    const addr = { sheet: 0, row: 0, col: 0 };
    setProtectedSheet(sheet.instance.store, 0, true, { workbook: sheet.workbook });
    mutators.setActive(sheet.instance.store, addr);
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    const typed = new InputEvent('beforeinput', {
      inputType: 'insertText',
      data: 'a',
      cancelable: true,
    });
    harness.fxInput.dispatchEvent(typed);

    expect(typed.defaultPrevented).toBe(true);
    expect(onValidation).toHaveBeenCalledWith(
      expect.objectContaining({ severity: 'stop', code: 'protectedCell' }),
    );
    expect(harness.controller.isEditing()).toBe(false);
    expect(harness.fxInput.value).toBe('');

    // An unlocked cell still takes the keystroke.
    onValidation.mockClear();
    mutators.setCellFormat(sheet.instance.store, addr, { locked: false });
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    const allowed = new InputEvent('beforeinput', {
      inputType: 'insertText',
      data: 'a',
      cancelable: true,
    });
    harness.fxInput.dispatchEvent(allowed);
    expect(allowed.defaultPrevented).toBe(false);
    expect(onValidation).not.toHaveBeenCalled();
    harness.detach();
  });

  it('returns false on stop validation without a controller and preserves the value and edit state', () => {
    const store = createSpreadsheetStore();
    const onValidation = vi.fn();
    const harness = attachFormulaBarHarness(sheet, onValidation, store);
    const addr = { sheet: 0, row: 0, col: 0 };
    sheet.workbook.setNumber(addr, 9);
    mutators.setActive(store, addr);
    mutators.setCellFormat(store, addr, {
      validation: {
        kind: 'whole',
        op: '=',
        a: 5,
        errorStyle: 'stop',
        errorMessage: 'Value must be five.',
        showErrorMessage: true,
      },
    });
    mutators.setCellFormat(store, addr, { underline: true });
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.value = '4';
    harness.fxInput.dispatchEvent(new Event('input'));

    expect(harness.controller.commitFx('none')).toBe(false);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 9 });
    expect(harness.controller.isEditing()).toBe(true);
    expect(harness.fxInput.value).toBe('4');
    expect(store.getState().format.formats.get(addrKey(addr))).toMatchObject({ underline: true });
    expect(onValidation).toHaveBeenCalledWith(
      expect.objectContaining({ severity: 'stop', message: 'Value must be five.' }),
    );
    harness.detach();
  });
});
