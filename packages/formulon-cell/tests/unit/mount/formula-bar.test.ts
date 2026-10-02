import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { setProtectedSheet } from '../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { defaultStrings } from '../../../src/i18n/strings.js';
import { attachFormulaBarController } from '../../../src/mount/formula-bar.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/index.js';

const attachFormulaBarHarness = (
  sheet: MountedStubSheet,
  onValidation: (outcome: {
    severity: 'stop' | 'warning' | 'information';
    title?: string;
    message: string;
  }) => void,
  store: SpreadsheetStore = sheet.instance.store,
) => {
  const formulabar = document.createElement('div');
  const fxInput = document.createElement('textarea');
  const fxCancel = document.createElement('button');
  const fxAccept = document.createElement('button');
  formulabar.append(fxCancel, fxAccept, fxInput);
  sheet.host.appendChild(formulabar);
  const autocomplete = {
    isOpen: () => false,
    move: () => {},
    acceptHighlighted: () => false,
    close: () => {},
    refresh: () => {},
  };
  const argHelper = { refresh: () => {}, close: () => {} };
  const controller = attachFormulaBarController({
    formulabar,
    fxAccept,
    fxCancel,
    fxInput,
    getArgHelper: () => argHelper,
    getAutocomplete: () => autocomplete,
    getStrings: () => defaultStrings,
    cancelBindingEditor: () => {},
    host: sheet.host,
    onValidation,
    store,
    updateChrome: () => {},
    wb: () => sheet.workbook,
  });
  return { controller, fxInput, detach: controller.detach };
};

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
    mutators.setPendingFormat(sheet.instance.store, { addr, format: { bold: true } });
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

  it('F4 rotates the ref under the caret', () => {
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = '=A1';
    fxInput.dispatchEvent(new Event('input'));
    fxInput.setSelectionRange(3, 3);
    fxInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'F4' }));
    // First rotation: A1 → $A$1.
    expect(fxInput.value).toBe('=$A$1');
  });
});

describe('mount/formula-bar — native commit result', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    const workbook = await WorkbookHandle.createDefault();
    sheet = await mountStubSheet({ workbook });
    expect(sheet.workbook.isStub).toBe(false);
  });

  afterEach(() => sheet.dispose());

  it('returns true after a native write and applies a matching pending format', () => {
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(sheet.instance.store, addr);
    mutators.setPendingFormat(sheet.instance.store, { addr, format: { bold: true } });
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.value = '42';
    harness.fxInput.dispatchEvent(new Event('input'));

    expect(harness.controller.commitFx('none')).toBe(true);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 42 });
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')?.bold).toBe(true);
    expect(sheet.instance.store.getState().ui.pendingFormat).toBeNull();
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
    mutators.setPendingFormat(sheet.instance.store, { addr, format: { italic: true } });
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.value = 'denied';
    harness.fxInput.dispatchEvent(new Event('input'));

    expect(harness.controller.commitFx('none')).toBe(false);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'blank' });
    expect(harness.controller.isEditing()).toBe(true);
    expect(document.activeElement).toBe(harness.fxInput);
    expect(sheet.instance.store.getState().ui.pendingFormat).toEqual({
      addr,
      format: { italic: true },
    });
    expect(onValidation).toHaveBeenCalledWith(expect.objectContaining({ severity: 'stop' }));
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
    mutators.setPendingFormat(store, { addr, format: { underline: true } });
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.value = '4';
    harness.fxInput.dispatchEvent(new Event('input'));

    expect(harness.controller.commitFx('none')).toBe(false);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 9 });
    expect(harness.controller.isEditing()).toBe(true);
    expect(harness.fxInput.value).toBe('4');
    expect(store.getState().ui.pendingFormat).toEqual({
      addr,
      format: { underline: true },
    });
    expect(onValidation).toHaveBeenCalledWith(
      expect.objectContaining({ severity: 'stop', message: 'Value must be five.' }),
    );
    harness.detach();
  });
});
