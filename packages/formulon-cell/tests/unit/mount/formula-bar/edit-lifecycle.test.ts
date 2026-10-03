import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { defaultStrings } from '../../../../src/i18n/strings.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/index.js';
import { attachFormulaBarHarness } from './fixtures.js';

/**
 * Unit: formula-bar controller — drives editing lifecycle (focus → input →
 * commit / cancel) on the fxInput textarea. Exercised through the mounted
 * sheet because the controller relies on the wider host + autocomplete +
 * arg-helper plumbing that mount.ts wires together.
 */
describe('mount/formula-bar — edit lifecycle', () => {
  let sheet: MountedStubSheet;
  let fxInput: HTMLTextAreaElement;
  let fxCancel: HTMLButtonElement;
  let fxAccept: HTMLButtonElement;
  let formulabar: HTMLDivElement;

  beforeEach(async () => {
    sheet = await mountStubSheet();
    fxInput = sheet.host.querySelector('.fc-host__formulabar-input') as HTMLTextAreaElement;
    fxCancel = sheet.host.querySelector('.fc-host__formulabar-action--cancel') as HTMLButtonElement;
    fxAccept = sheet.host.querySelector('.fc-host__formulabar-action--accept') as HTMLButtonElement;
    formulabar = sheet.host.querySelector('.fc-host__formulabar') as HTMLDivElement;
  });

  afterEach(() => sheet.dispose());

  it('idle: cancel + accept buttons are disabled, fcEditing=0', () => {
    expect(fxCancel.disabled).toBe(true);
    expect(fxAccept.disabled).toBe(true);
    expect(fxCancel.dataset.disabledReason).toBe(defaultStrings.a11y.cancelFormulaEditUnavailable);
    expect(fxAccept.dataset.disabledReason).toBe(defaultStrings.a11y.enterFormulaUnavailable);
    expect(fxCancel.getAttribute('aria-description')).toBe(
      defaultStrings.a11y.cancelFormulaEditUnavailable,
    );
    expect(fxAccept.title).toBe(
      `${defaultStrings.a11y.enterFormula}\n${defaultStrings.a11y.enterFormulaUnavailable}`,
    );
    expect(formulabar.dataset.fcEditing).toBe('0');
  });

  it('keeps the Excel-style formula bar slot order and icon buttons', () => {
    expect(
      Array.from(formulabar.children).map((child) => (child as HTMLElement).className),
    ).toEqual([
      'fc-host__formulabar-tag',
      'fc-host__formulabar-action fc-host__formulabar-action--cancel',
      'fc-host__formulabar-action fc-host__formulabar-action--accept',
      'fc-host__formulabar-fx',
      'fc-host__formulabar-input',
      'fc-host__formulabar-expand',
    ]);
    expect(fxCancel.querySelector('svg')).not.toBeNull();
    expect(fxAccept.querySelector('svg')).not.toBeNull();
    expect(formulabar.querySelector('.fc-host__formulabar-expand svg')).not.toBeNull();
  });

  it('focus starts editing — cancel enables, accept stays disabled until dirty', () => {
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    expect(formulabar.dataset.fcEditing).toBe('1');
    expect(fxCancel.disabled).toBe(false);
    expect(fxAccept.disabled).toBe(true);
    expect(fxCancel.dataset.disabledReason).toBeUndefined();
    expect(fxCancel.hasAttribute('aria-description')).toBe(false);
    expect(fxCancel.title).toBe(defaultStrings.a11y.cancelFormulaEdit);
    expect(fxAccept.dataset.disabledReason).toBe(defaultStrings.a11y.enterFormulaNoChanges);
  });

  it('changing the value flips accept on (dirty); Escape rolls it back', () => {
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = '=A1+1';
    fxInput.dispatchEvent(new Event('input'));
    expect(fxAccept.disabled).toBe(false);
    expect(fxAccept.dataset.disabledReason).toBeUndefined();
    expect(fxAccept.hasAttribute('aria-description')).toBe(false);
    expect(fxAccept.title).toBe(defaultStrings.a11y.enterFormula);

    fxInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
    expect(fxInput.value).toBe('');
    expect(formulabar.dataset.fcEditing).toBe('0');
    expect(fxAccept.disabled).toBe(true);
    expect(fxAccept.dataset.disabledReason).toBe(defaultStrings.a11y.enterFormulaUnavailable);
  });

  it('Enter commits and advances the active cell down', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = '42';
    fxInput.dispatchEvent(new Event('input'));
    fxInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }));

    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'number',
      value: 42,
    });
    expect(sheet.instance.store.getState().selection.active).toEqual({
      sheet: 0,
      row: 1,
      col: 0,
    });
    expect(formulabar.dataset.fcEditing).toBe('0');
  });

  it('Enter applies pending empty-cell format to the committed formula-bar value', () => {
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(sheet.instance.store, addr);
    mutators.setPendingFormat(sheet.instance.store, {
      addr,
      format: { bold: true, numFmt: { kind: 'text' } },
    });
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = 'typed';
    fxInput.dispatchEvent(new Event('input'));
    fxInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }));

    expect(sheet.instance.store.getState().format.formats.get('0:0:0')?.bold).toBe(true);
    expect(sheet.instance.store.getState().ui.pendingFormat).toBeNull();
  });

  it('Enter preserves numeric-looking input as text for cells formatted as Text', () => {
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(sheet.instance.store, addr);
    mutators.setCellFormat(sheet.instance.store, addr, { numFmt: { kind: 'text' } });
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = '00123';
    fxInput.dispatchEvent(new Event('input'));
    fxInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }));

    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'text', value: '00123' });
  });

  it('Tab commits and advances the active cell right', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = 'hello';
    fxInput.dispatchEvent(new Event('input'));
    fxInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab' }));

    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'hello',
    });
    expect(sheet.instance.store.getState().selection.active).toEqual({
      sheet: 0,
      row: 0,
      col: 1,
    });
  });

  it('default-platform Shift+Tab commits without moving the active cell', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 1 });
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = 'left';
    fxInput.dispatchEvent(new Event('input'));
    fxInput.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Tab', shiftKey: true, cancelable: true }),
    );

    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
      kind: 'text',
      value: 'left',
    });
    expect(sheet.instance.store.getState().selection.active).toEqual({
      sheet: 0,
      row: 0,
      col: 1,
    });
  });

  it('default-platform Shift+Enter stays in edit mode for a literal newline', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 2, col: 0 });
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = 'up';
    fxInput.dispatchEvent(new Event('input'));
    const event = new KeyboardEvent('keydown', {
      key: 'Enter',
      shiftKey: true,
      cancelable: true,
    });
    fxInput.dispatchEvent(event);

    expect(event.defaultPrevented).toBe(false);
    expect(sheet.workbook.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({ kind: 'blank' });
    expect(sheet.instance.store.getState().selection.active).toEqual({
      sheet: 0,
      row: 2,
      col: 0,
    });
    expect(formulabar.dataset.fcEditing).toBe('1');
  });

  it('clicking the cancel button reverts a dirty edit', () => {
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = '=BAD';
    fxInput.dispatchEvent(new Event('input'));
    expect(fxAccept.disabled).toBe(false);

    fxCancel.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(fxInput.value).toBe('');
    expect(fxCancel.disabled).toBe(true);
    expect(formulabar.dataset.fcEditing).toBe('0');
  });

  it('clicking the accept button commits without changing the active cell', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 1, col: 1 });
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = '7';
    fxInput.dispatchEvent(new Event('input'));

    fxAccept.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({
      kind: 'number',
      value: 7,
    });
    // No advance for "accept" — selection stays where it was.
    expect(sheet.instance.store.getState().selection.active).toEqual({
      sheet: 0,
      row: 1,
      col: 1,
    });
  });

  it('commitFx reports success after an ordinary controller-backed write', () => {
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(sheet.instance.store, addr);
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.value = '42';
    harness.fxInput.dispatchEvent(new Event('input'));

    expect(harness.controller.commitFx('none')).toBe(true);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 42 });
    expect(harness.controller.isEditing()).toBe(false);
    harness.detach();
  });

  it('commitFx reports controller rejection and leaves formula-bar editing active', () => {
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(sheet.instance.store, addr);
    sheet.instance.setPolicy({ readOnly: true });
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.value = 'denied';
    harness.fxInput.dispatchEvent(new Event('input'));

    expect(harness.controller.commitFx('none')).toBe(false);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'blank' });
    expect(harness.controller.isEditing()).toBe(true);
    harness.detach();
  });

  it('commitFx reports controller exceptions as rejection', () => {
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(sheet.instance.store, addr);
    const execute = vi.spyOn(sheet.instance.commands, 'execute').mockImplementation(() => {
      throw new Error('controller failed');
    });
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.value = '42';
    harness.fxInput.dispatchEvent(new Event('input'));

    expect(harness.controller.commitFx('none')).toBe(false);
    expect(harness.controller.isEditing()).toBe(true);
    execute.mockRestore();
    warn.mockRestore();
    harness.detach();
  });

  it('commitFx reports stop-validation rejection and leaves the value unwritten', () => {
    const store = createSpreadsheetStore();
    const onValidation = vi.fn();
    const harness = attachFormulaBarHarness(sheet, onValidation, store);
    const addr = { sheet: 0, row: 0, col: 0 };
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
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.value = '4';
    harness.fxInput.dispatchEvent(new Event('input'));

    expect(harness.controller.commitFx('none')).toBe(false);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'blank' });
    expect(harness.controller.isEditing()).toBe(true);
    expect(onValidation).toHaveBeenCalledWith(
      expect.objectContaining({ severity: 'stop', message: 'Value must be five.' }),
    );
    harness.detach();
  });

  it('commitFx reports warning validation as accepted after writing', () => {
    const store = createSpreadsheetStore();
    const onValidation = vi.fn();
    const harness = attachFormulaBarHarness(sheet, onValidation, store);
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(store, addr);
    mutators.setCellFormat(store, addr, {
      validation: {
        kind: 'whole',
        op: '=',
        a: 5,
        errorStyle: 'warning',
        errorMessage: 'Value is not five.',
        showErrorMessage: true,
      },
    });
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.value = '4';
    harness.fxInput.dispatchEvent(new Event('input'));

    expect(harness.controller.commitFx('none')).toBe(true);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 4 });
    expect(harness.controller.isEditing()).toBe(false);
    expect(onValidation).toHaveBeenCalledWith(
      expect.objectContaining({ severity: 'warning', message: 'Value is not five.' }),
    );
    harness.detach();
  });

  it('commitFx reports write exceptions as rejection and keeps editing active', () => {
    const store = createSpreadsheetStore();
    const onValidation = vi.fn();
    const harness = attachFormulaBarHarness(sheet, onValidation, store);
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(store, addr);
    const setNumber = vi.spyOn(sheet.workbook, 'setNumber').mockImplementation(() => {
      throw new Error('engine write failed');
    });
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.value = '42';
    harness.fxInput.dispatchEvent(new Event('input'));

    expect(harness.controller.commitFx('none')).toBe(false);
    expect(harness.controller.isEditing()).toBe(true);
    expect(onValidation).toHaveBeenCalledWith(
      expect.objectContaining({
        severity: 'stop',
        message: 'The formula-bar value could not be written.',
      }),
    );
    setNumber.mockRestore();
    harness.detach();
  });

  it('blur with a dirty edit commits (matches click-elsewhere behaviour)', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = 'blur-commit';
    fxInput.dispatchEvent(new Event('input'));
    fxInput.dispatchEvent(new FocusEvent('blur'));

    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'blur-commit',
    });
    expect(formulabar.dataset.fcEditing).toBe('0');
  });

  it('blur without changes just leaves edit-mode (no commit)', () => {
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.dispatchEvent(new FocusEvent('blur'));
    expect(formulabar.dataset.fcEditing).toBe('0');
    // Active cell stayed blank.
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });
  });
});
