import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import type { Addr } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { defaultStrings } from '../../../src/i18n/strings.js';
import type { FormulaEditLease } from '../../../src/interact/formula-edit-lease.js';
import {
  attachFormulaBarController,
  type FormulaBarController,
} from '../../../src/mount/formula-bar.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/index.js';

const selectRange = (
  sheet: MountedStubSheet,
  active: { sheet: number; row: number; col: number },
  range: { sheet: number; r0: number; c0: number; r1: number; c1: number },
  extraRanges: { sheet: number; r0: number; c0: number; r1: number; c1: number }[] = [],
): void => {
  mutators.setActive(sheet.instance.store, active);
  mutators.setRange(sheet.instance.store, range);
  sheet.instance.store.setState((state) => ({
    ...state,
    selection: {
      ...state.selection,
      extraRanges: extraRanges.map((extra) => ({ ...extra })),
    },
  }));
};

const attachFormulaBarHarness = (
  sheet: MountedStubSheet,
  onValidation: (outcome: {
    severity: 'stop' | 'warning' | 'information';
    title?: string;
    message: string;
  }) => void,
  store: SpreadsheetStore = sheet.instance.store,
  getWorkbook: () => WorkbookHandle = () => sheet.workbook,
  cancelBindingEditor: () => void = () => {},
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
  const argHelper = { close: () => {}, refresh: () => {} };
  const controller = attachFormulaBarController({
    formulabar,
    fxAccept,
    fxCancel,
    fxInput,
    getArgHelper: () => argHelper,
    getAutocomplete: () => autocomplete,
    getStrings: () => defaultStrings,
    cancelBindingEditor,
    host: sheet.host,
    onValidation,
    store,
    updateChrome: () => {},
    wb: getWorkbook,
  });
  return { controller, fxInput, detach: controller.detach };
};

interface ExternalDraftTestHandle {
  readonly anchor: Addr;
  value(): string;
  setValue(raw: string, caret?: number): void;
  commit(): boolean;
  cancel(): void;
  discard(): void;
  subscribe(fn: (raw: string) => void): () => void;
}

type ExternalDraftTestController = FormulaBarController & {
  beginExternalDraft(
    anchor: Addr,
    seed: string,
    hooks: {
      onFinish(outcome: 'committed' | 'cancelled', restoredFocusTarget?: HTMLElement | null): void;
    },
    options?: { lease?: FormulaEditLease },
  ): ExternalDraftTestHandle | null;
};

const asExternalDraftController = (controller: FormulaBarController): ExternalDraftTestController =>
  controller as ExternalDraftTestController;

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

  it('Cmd+T rotates a ref in the formula bar on Mac', () => {
    sheet.host.dataset.fcPlatform = 'mac';
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = '=A1';
    fxInput.dispatchEvent(new Event('input'));
    fxInput.setSelectionRange(3, 3);
    fxInput.dispatchEvent(
      new KeyboardEvent('keydown', { key: 't', metaKey: true, cancelable: true }),
    );

    expect(fxInput.value).toBe('=$A$1');
    expect(formulabar.dataset.fcEditing).toBe('1');
  });

  it.each([
    { key: 'Enter', shiftKey: false, start: { row: 0, col: 0 }, next: { row: 1, col: 0 } },
    { key: 'Enter', shiftKey: true, start: { row: 1, col: 0 }, next: { row: 0, col: 0 } },
    { key: 'Tab', shiftKey: false, start: { row: 0, col: 0 }, next: { row: 0, col: 1 } },
    { key: 'Tab', shiftKey: true, start: { row: 0, col: 1 }, next: { row: 0, col: 0 } },
  ])(
    'Mac $key commits and traverses the selected rectangle while preserving its shape',
    ({ key, shiftKey, start, next }) => {
      sheet.host.dataset.fcPlatform = 'mac';
      const range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
      const extraRanges = [{ sheet: 0, r0: 3, c0: 3, r1: 3, c1: 3 }];
      const active = { sheet: 0, ...start };
      selectRange(sheet, active, range, extraRanges);
      fxInput.focus();
      fxInput.dispatchEvent(new FocusEvent('focus'));
      fxInput.value = 'selected';
      fxInput.dispatchEvent(new Event('input'));

      fxInput.dispatchEvent(new KeyboardEvent('keydown', { key, shiftKey, cancelable: true }));

      const selection = sheet.instance.store.getState().selection;
      expect(sheet.workbook.getValue(active)).toEqual({ kind: 'text', value: 'selected' });
      expect(selection.active).toEqual({ sheet: 0, ...next });
      expect(selection.anchor).toEqual(active);
      expect(selection.range).toEqual(range);
      expect(selection.extraRanges).toEqual(extraRanges);
    },
  );

  it('Mac Return exits a one-cell merge with merge-aware motion inside navigation bounds', () => {
    sheet.host.dataset.fcPlatform = 'mac';
    sheet.instance.setViewportOptions({
      range: { sheet: 0, r0: 0, c0: 0, r1: 6, c1: 4 },
    });
    const merge = { sheet: 0, r0: 3, c0: 2, r1: 4, c1: 2 };
    mutators.mergeRange(sheet.instance.store, merge);
    const anchor = { sheet: 0, row: merge.r0, col: merge.c0 };
    mutators.setActive(sheet.instance.store, anchor);
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = 'merged';
    fxInput.dispatchEvent(new Event('input'));
    fxInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', cancelable: true }));

    expect(sheet.instance.store.getState().selection.active).toEqual({
      sheet: 0,
      row: 5,
      col: 2,
    });
  });

  it('default-platform Return steps past a merged active cell instead of into its body', () => {
    const merge = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 };
    mutators.mergeRange(sheet.instance.store, merge);
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = 'merged';
    fxInput.dispatchEvent(new Event('input'));
    fxInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', cancelable: true }));

    expect(sheet.instance.store.getState().selection.active).toEqual({ sheet: 0, row: 2, col: 0 });
  });

  it('Mac Return clamps a merged-cell step at the configured range edge', () => {
    sheet.host.dataset.fcPlatform = 'mac';
    sheet.instance.setViewportOptions({
      range: { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 4 },
    });
    const merge = { sheet: 0, r0: 3, c0: 2, r1: 4, c1: 2 };
    mutators.mergeRange(sheet.instance.store, merge);
    const anchor = { sheet: 0, row: merge.r0, col: merge.c0 };
    mutators.setActive(sheet.instance.store, anchor);
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = 'edge';
    fxInput.dispatchEvent(new Event('input'));
    fxInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', cancelable: true }));

    expect(sheet.instance.store.getState().selection.active).toEqual(anchor);
  });

  it('Mac Tab follows explicit navigation stops before the selected rectangle', () => {
    sheet.host.dataset.fcPlatform = 'mac';
    sheet.instance.setViewportOptions({
      range: { sheet: 0, r0: 0, c0: 0, r1: 5, c1: 5 },
    });
    const active = { sheet: 0, row: 2, col: 2 };
    const selectedRange = { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 };
    selectRange(sheet, active, selectedRange);
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = 'tab-stop';
    fxInput.dispatchEvent(new Event('input'));
    fxInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', cancelable: true }));

    expect(sheet.instance.store.getState().selection.active).toEqual({
      sheet: 0,
      row: 2,
      col: 3,
    });
  });

  it('Mac Option+Return leaves the textarea newline available to the browser', () => {
    sheet.host.dataset.fcPlatform = 'mac';
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = 'first line';
    fxInput.dispatchEvent(new Event('input'));
    const key = new KeyboardEvent('keydown', {
      key: 'Enter',
      altKey: true,
      cancelable: true,
    });

    fxInput.dispatchEvent(key);

    expect(key.defaultPrevented).toBe(false);
    expect(fxInput.value).toBe('first line');
    expect(formulabar.dataset.fcEditing).toBe('1');
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });
  });

  it.each([
    { label: 'Command+Return', metaKey: true, ctrlKey: false },
    { label: 'Control+Return', metaKey: false, ctrlKey: true },
  ])(
    'Mac $label atomically fills formulas, preserves selection, and uses one undo entry',
    ({ metaKey, ctrlKey }) => {
      sheet.host.dataset.fcPlatform = 'mac';
      const active = { sheet: 0, row: 1, col: 1 };
      const range = { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 };
      const extraRanges = [{ sheet: 0, r0: 1, c0: 3, r1: 1, c1: 3 }];
      selectRange(sheet, active, range, extraRanges);
      sheet.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 11);
      sheet.workbook.setNumber({ sheet: 0, row: 0, col: 1 }, 12);
      sheet.workbook.setNumber({ sheet: 0, row: 1, col: 0 }, 21);
      sheet.workbook.setNumber(active, 22);
      sheet.workbook.setNumber({ sheet: 0, row: 0, col: 2 }, 31);
      sheet.instance.history.clear();
      fxInput.focus();
      fxInput.dispatchEvent(new FocusEvent('focus'));
      fxInput.value = '=A1';
      fxInput.dispatchEvent(new Event('input'));

      fxInput.dispatchEvent(
        new KeyboardEvent('keydown', { key: 'Enter', metaKey, ctrlKey, cancelable: true }),
      );

      expect(sheet.workbook.getValue(active)).toEqual({ kind: 'number', value: 11 });
      expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 2 })).toEqual({
        kind: 'number',
        value: 12,
      });
      expect(sheet.workbook.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({
        kind: 'number',
        value: 21,
      });
      expect(sheet.workbook.getValue({ sheet: 0, row: 2, col: 2 })).toEqual({
        kind: 'number',
        value: 11,
      });
      expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 3 })).toEqual({
        kind: 'number',
        value: 31,
      });
      expect(sheet.instance.store.getState().selection).toMatchObject({
        active,
        anchor: active,
        range,
        extraRanges,
      });
      expect(formulabar.dataset.fcEditing).toBe('0');
      expect(sheet.instance.history.canUndo()).toBe(true);

      expect(sheet.instance.undo()).toBe(true);
      expect(sheet.workbook.getValue(active)).toEqual({ kind: 'number', value: 22 });
      expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 2 })).toEqual({ kind: 'blank' });
      expect(sheet.workbook.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({ kind: 'blank' });
      expect(sheet.workbook.getValue({ sheet: 0, row: 2, col: 2 })).toEqual({ kind: 'blank' });
      expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 3 })).toEqual({ kind: 'blank' });
      expect(sheet.instance.history.canUndo()).toBe(false);
    },
  );

  it('default-platform Control+Return fills the selection while Command+Return does not', () => {
    const active = { sheet: 0, row: 0, col: 0 };
    selectRange(sheet, active, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = 'meta';
    fxInput.dispatchEvent(new Event('input'));
    fxInput.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Enter', metaKey: true, cancelable: true }),
    );
    expect(sheet.workbook.getValue(active)).toEqual({ kind: 'blank' });
    expect(formulabar.dataset.fcEditing).toBe('1');

    fxInput.value = 'both';
    fxInput.dispatchEvent(new Event('input'));
    const key = new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, cancelable: true });
    fxInput.dispatchEvent(key);

    expect(key.defaultPrevented).toBe(true);
    expect(sheet.workbook.getValue(active)).toEqual({ kind: 'text', value: 'both' });
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
      kind: 'text',
      value: 'both',
    });
    expect(sheet.instance.store.getState().selection.active).toEqual(active);
    expect(formulabar.dataset.fcEditing).toBe('0');
  });

  it('Mac fill policy rejection leaves the full selection untouched and editing open', () => {
    sheet.host.dataset.fcPlatform = 'mac';
    const onValidation = vi.fn();
    const harness = attachFormulaBarHarness(sheet, onValidation);
    const active = { sheet: 0, row: 0, col: 0 };
    const range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 };
    selectRange(sheet, active, range);
    sheet.instance.setPolicy({ readOnly: true });
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.value = 'denied';
    harness.fxInput.dispatchEvent(new Event('input'));
    harness.fxInput.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Enter', metaKey: true, cancelable: true }),
    );

    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'blank' });
    expect(sheet.instance.store.getState().selection).toMatchObject({ active, range });
    expect(harness.controller.isEditing()).toBe(true);
    expect(onValidation).toHaveBeenCalledWith(
      expect.objectContaining({ severity: 'stop', message: expect.stringContaining('read-only') }),
    );
    harness.detach();
  });

  it('Mac fill validation rejection does not partially write or leave edit mode', () => {
    sheet.host.dataset.fcPlatform = 'mac';
    const onValidation = vi.fn();
    const harness = attachFormulaBarHarness(sheet, onValidation);
    const active = { sheet: 0, row: 0, col: 0 };
    const range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 };
    selectRange(sheet, active, range);
    mutators.setCellFormat(
      sheet.instance.store,
      { sheet: 0, row: 0, col: 1 },
      {
        validation: {
          kind: 'list',
          source: ['accepted'],
          errorMessage: 'Choose an accepted value.',
          showErrorMessage: true,
        },
      },
    );
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.value = 'rejected';
    harness.fxInput.dispatchEvent(new Event('input'));
    harness.fxInput.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, cancelable: true }),
    );

    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'blank' });
    expect(harness.controller.isEditing()).toBe(true);
    expect(onValidation).toHaveBeenCalledWith(
      expect.objectContaining({ severity: 'stop', message: 'Choose an accepted value.' }),
    );
    harness.detach();
  });

  it('Mac fill over the unique-cell limit keeps the edit open and explains the limit', () => {
    sheet.host.dataset.fcPlatform = 'mac';
    const onValidation = vi.fn();
    const harness = attachFormulaBarHarness(sheet, onValidation);
    const active = { sheet: 0, row: 0, col: 0 };
    const range = { sheet: 0, r0: 0, c0: 0, r1: 100_000, c1: 0 };
    sheet.instance.store.setState((state) => ({
      ...state,
      selection: { ...state.selection, active, anchor: active, range, extraRanges: [] },
    }));
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.value = 'too many';
    harness.fxInput.dispatchEvent(new Event('input'));
    harness.fxInput.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Enter', metaKey: true, cancelable: true }),
    );

    expect(harness.controller.isEditing()).toBe(true);
    expect(onValidation).toHaveBeenCalledWith(
      expect.objectContaining({
        severity: 'stop',
        message: expect.stringContaining('100,000 unique cells'),
      }),
    );
    expect(sheet.workbook.getValue(active)).toEqual({ kind: 'blank' });
    harness.detach();
  });

  it('formula-bar reference insertion replaces a partial token and refreshes helpers', () => {
    const detached = document.createElement('div');
    const detachedInput = document.createElement('textarea');
    const detachedCancel = document.createElement('button');
    const detachedAccept = document.createElement('button');
    const autocomplete = {
      isOpen: () => false,
      move: () => {},
      acceptHighlighted: () => false,
      close: () => {},
      refresh: () => {},
    };
    const argHelper = { close: () => {}, refresh: () => {} };
    const autocompleteRefresh = vi.spyOn(autocomplete, 'refresh');
    const argHelperRefresh = vi.spyOn(argHelper, 'refresh');
    const controller = attachFormulaBarController({
      formulabar: detached,
      fxAccept: detachedAccept,
      fxCancel: detachedCancel,
      fxInput: detachedInput,
      getArgHelper: () => argHelper,
      getAutocomplete: () => autocomplete,
      getStrings: () => defaultStrings,
      cancelBindingEditor: () => {},
      host: sheet.host,
      store: sheet.instance.store,
      updateChrome: () => {},
      wb: () => sheet.workbook,
    });
    detachedInput.value = '=SUM(A';
    detachedInput.focus();
    detachedInput.dispatchEvent(new FocusEvent('focus'));
    expect(controller.isFormulaEdit()).toBe(true);

    controller.insertRefAtCaret('B2');
    controller.insertRefAtCaret('C3');

    expect(detachedInput.value).toBe('=SUM(C3');
    expect(detachedInput.selectionStart).toBe(detachedInput.value.length);
    expect(sheet.instance.store.getState().ui.editorRefs).toHaveLength(1);
    expect(autocompleteRefresh).toHaveBeenCalled();
    expect(argHelperRefresh).toHaveBeenCalled();
    controller.detach();
  });

  it('Mac formula-bar range insertion leaves the caret ready for more formula text', () => {
    sheet.host.dataset.fcPlatform = 'mac';
    const active = { sheet: 0, row: 0, col: 2 };
    mutators.setActive(sheet.instance.store, active);
    sheet.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 2);
    sheet.workbook.setNumber({ sheet: 0, row: 0, col: 1 }, 4);
    sheet.workbook.setNumber({ sheet: 0, row: 1, col: 0 }, 3);
    sheet.workbook.setNumber({ sheet: 0, row: 1, col: 1 }, 5);
    sheet.workbook.recalc();
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));

    const harness = attachFormulaBarHarness(sheet, vi.fn());
    const { controller, fxInput: input } = harness;
    input.focus();
    input.dispatchEvent(new FocusEvent('focus'));
    input.value = '=SUM(';
    input.dispatchEvent(new Event('input'));
    input.setSelectionRange(input.value.length, input.value.length);

    controller.insertRefAtCaret('A1');
    controller.insertRefAtCaret('A1:B2');

    expect(document.activeElement).toBe(input);
    expect(input.value).toBe('=SUM(A1:B2');
    expect(input.selectionStart).toBe(input.value.length);

    input.setRangeText(')', input.selectionStart, input.selectionEnd, 'end');
    input.dispatchEvent(new Event('input'));
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', cancelable: true }));

    expect(sheet.workbook.cellFormula(active)).toBe('=SUM(A1:B2)');
    expect(sheet.workbook.getValue(active)).toEqual({ kind: 'number', value: 14 });
    harness.detach();
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

  it('commits an external draft once at its fixed anchor without blur or navigation', () => {
    const onValidation = vi.fn();
    const harness = attachFormulaBarHarness(
      sheet,
      onValidation,
      sheet.instance.store,
      () => sheet.instance.workbook,
    );
    const addr = { sheet: 0, row: 0, col: 0 };
    const moved = { sheet: 0, row: 2, col: 2 };
    sheet.workbook.setNumber(addr, 1);
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, addr);
    sheet.instance.history.clear();
    harness.fxInput.value = '1';
    const hostFocus = vi.spyOn(sheet.host, 'focus').mockImplementation(() => {});
    const outcomes: string[] = [];
    const draft = asExternalDraftController(harness.controller).beginExternalDraft(addr, '=', {
      onFinish: (outcome) => outcomes.push(outcome),
    });
    if (!draft) throw new Error('External draft did not start.');

    const values: string[] = [];
    const unsubscribe = draft.subscribe((raw) => values.push(raw));
    expect(draft.anchor).toEqual(addr);
    expect(draft.value()).toBe('=');
    draft.setValue('42', 2);
    expect(draft.value()).toBe('42');
    expect(harness.fxInput.selectionStart).toBe(2);
    expect(values).toEqual(['42']);

    mutators.setActive(sheet.instance.store, moved);
    mutators.setPendingFormat(sheet.instance.store, { addr, format: { bold: true } });
    harness.fxInput.dispatchEvent(new FocusEvent('blur'));
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 1 });
    expect(sheet.instance.history.canUndo()).toBe(false);
    expect(harness.controller.isEditing()).toBe(true);

    harness.fxInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', cancelable: true }));
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 42 });
    expect(sheet.instance.store.getState().selection.active).toEqual(moved);
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')?.bold).toBe(true);
    expect(sheet.instance.store.getState().ui.pendingFormat).toBeNull();
    expect(sheet.instance.history.canUndo()).toBe(true);
    expect(harness.controller.isEditing()).toBe(false);
    expect(outcomes).toEqual(['committed']);
    expect(hostFocus).not.toHaveBeenCalled();
    expect(draft.commit()).toBe(false);
    draft.cancel();
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 42 });
    expect(onValidation).not.toHaveBeenCalled();
    unsubscribe();
    hostFocus.mockRestore();
    harness.detach();
  });

  it('refuses to suspend the formula bar during IME composition', () => {
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    harness.fxInput.value = '=1';
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new CompositionEvent('compositionstart'));
    expect(
      harness.controller.suspendForFormulaPalette({
        getLocale: () => 'ja',
        contextCurrent: () => true,
      }),
    ).toBeNull();
    expect(harness.controller.isEditing()).toBe(true);
    expect(harness.fxInput.value).toBe('=1');
    expect(sheet.instance.history.canUndo()).toBe(false);
    harness.fxInput.dispatchEvent(new CompositionEvent('compositionend'));
    const lease = harness.controller.suspendForFormulaPalette({
      getLocale: () => 'ja',
      contextCurrent: () => true,
    });
    expect(lease).not.toBeNull();
    expect(harness.controller.commitFx('none')).toBe(false);
    harness.fxInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', cancelable: true }));
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBeNull();
    harness.controller.cancelFx();
    expect(lease?.userCancel()).toBeNull();
    harness.detach();
  });

  it('does not clear a newer editor when a rejected commit discards its adopted draft', () => {
    let contextCurrent = true;
    let draft: ExternalDraftTestHandle | null = null;
    const newerPending = {
      addr: { sheet: 0, row: 3, col: 3 },
      format: { italic: true },
    };
    const newerRefs = [{ r0: 3, c0: 3, r1: 3, c1: 3, colorIndex: 1 }];
    const onValidation = vi.fn(() => {
      contextCurrent = false;
      mutators.setPendingFormat(sheet.instance.store, newerPending);
      mutators.setEditorRefs(sheet.instance.store, newerRefs);
      draft?.discard();
      draft?.cancel();
    });
    const harness = attachFormulaBarHarness(sheet, onValidation);
    const anchor = { sheet: 0, row: 0, col: 0 };
    harness.fxInput.value = '=1';
    harness.fxInput.focus();
    const lease = harness.controller.suspendForFormulaPalette({
      getLocale: () => 'en',
      contextCurrent: () => contextCurrent,
    });
    expect(lease).not.toBeNull();
    const onFinish = vi.fn();
    draft = harness.controller.beginExternalDraft(
      anchor,
      '=',
      { onFinish },
      { lease: lease ?? undefined },
    );
    if (!draft) throw new Error('Adopted draft did not start.');
    sheet.instance.setPolicy({ readOnly: true });
    expect(draft.commit()).toBe(false);
    expect(onValidation).toHaveBeenCalledOnce();
    expect(sheet.instance.store.getState().ui.pendingFormat).toEqual(newerPending);
    expect(sheet.instance.store.getState().ui.editorRefs).toEqual(newerRefs);
    expect(harness.controller.isEditing()).toBe(false);
    expect(onFinish).not.toHaveBeenCalled();
    harness.fxInput.dispatchEvent(new FocusEvent('blur'));
    harness.fxInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', cancelable: true }));
    expect(sheet.workbook.cellFormula(anchor)).toBeNull();
    expect(sheet.instance.history.canUndo()).toBe(false);
    harness.detach();
  });

  it('adopts a suspended formula-bar owner and commits its fixed anchor once', () => {
    const harness = attachFormulaBarHarness(
      sheet,
      vi.fn(),
      sheet.instance.store,
      () => sheet.instance.workbook,
    );
    const anchor = { sheet: 0, row: 0, col: 0 };
    const moved = { sheet: 0, row: 4, col: 4 };
    mutators.setActive(sheet.instance.store, anchor);
    mutators.setPendingFormat(sheet.instance.store, { addr: anchor, format: { bold: true } });
    sheet.instance.history.clear();
    harness.fxInput.value = '=1';
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.setSelectionRange(1, 1, 'backward');
    harness.fxInput.dispatchEvent(new Event('input'));
    const lease = harness.controller.suspendForFormulaPalette({
      getLocale: () => 'en',
      contextCurrent: () => true,
    });
    expect(lease).not.toBeNull();

    const outcomes: [string, HTMLElement | null | undefined][] = [];
    const draft = asExternalDraftController(harness.controller).beginExternalDraft(
      anchor,
      'ignored-seed',
      {
        onFinish: (outcome, restored) => outcomes.push([outcome, restored]),
      },
      { lease: lease ?? undefined },
    );
    if (!draft) throw new Error('Adopted formula-bar draft did not start.');
    expect(draft.value()).toBe('=1');
    draft.setValue('=ACOS(1)');
    sheet.instance.store.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        active: moved,
        anchor: moved,
        range: { sheet: 0, r0: moved.row, c0: moved.col, r1: moved.row, c1: moved.col },
      },
    }));
    expect(draft.commit()).toBe(true);
    expect(sheet.workbook.cellFormula(anchor)).toBe('=ACOS(1)');
    expect(sheet.workbook.cellFormula(moved)).toBeNull();
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')?.bold).toBe(true);
    expect(sheet.instance.history.canUndo()).toBe(true);
    expect(outcomes).toEqual([['committed', undefined]]);
    expect(sheet.instance.history.undo()).toBe(true);
    expect(sheet.workbook.cellFormula(anchor)).toBeNull();
    expect(sheet.instance.history.canUndo()).toBe(false);
    expect(sheet.instance.history.redo()).toBe(true);
    expect(sheet.workbook.cellFormula(anchor)).toBe('=ACOS(1)');
    expect(draft.commit()).toBe(false);
    harness.detach();
  });

  it('keeps an adopted draft open after policy rejection without focusing the suspended bar', () => {
    const harness = attachFormulaBarHarness(
      sheet,
      vi.fn(),
      sheet.instance.store,
      () => sheet.instance.workbook,
    );
    const anchor = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(sheet.instance.store, anchor);
    harness.fxInput.value = '=1';
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    const lease = harness.controller.suspendForFormulaPalette({
      getLocale: () => 'en',
      contextCurrent: () => true,
    });
    expect(lease).not.toBeNull();
    const draft = asExternalDraftController(harness.controller).beginExternalDraft(
      anchor,
      '=',
      { onFinish: vi.fn() },
      { lease: lease ?? undefined },
    );
    if (!draft) throw new Error('Adopted formula-bar draft did not start.');
    draft.setValue('=1');
    sheet.instance.setPolicy({ readOnly: true });
    const focusedBefore = document.activeElement;
    expect(draft.commit()).toBe(false);
    expect(draft.value()).toBe('=1');
    expect(document.activeElement).toBe(focusedBefore);
    expect(sheet.instance.history.canUndo()).toBe(false);
    sheet.instance.setPolicy(undefined);
    expect(draft.commit()).toBe(true);
    expect(sheet.workbook.cellFormula(anchor)).toBe('=1');
    expect(sheet.instance.history.undo()).toBe(true);
    expect(sheet.instance.history.canUndo()).toBe(false);
    draft.discard();
    expect(draft.commit()).toBe(false);
    harness.detach();
  });

  it('restores a valid leased owner on user cancellation and discards after context loss', () => {
    const harness = attachFormulaBarHarness(
      sheet,
      vi.fn(),
      sheet.instance.store,
      () => sheet.instance.workbook,
    );
    const anchor = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(sheet.instance.store, anchor);
    harness.fxInput.value = '=SUM(A1)';
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.setSelectionRange(2, 5, 'backward');
    const lease = harness.controller.suspendForFormulaPalette({
      getLocale: () => 'ja',
      contextCurrent: () => true,
    });
    expect(lease).not.toBeNull();
    const outcomes: [string, HTMLElement | null | undefined][] = [];
    const draft = asExternalDraftController(harness.controller).beginExternalDraft(
      anchor,
      '=',
      { onFinish: (outcome, restored) => outcomes.push([outcome, restored]) },
      { lease: lease ?? undefined },
    );
    if (!draft) throw new Error('Adopted formula-bar draft did not start.');
    draft.setValue('=SUM(A1)+2');
    draft.cancel();
    expect(outcomes[0]?.[0]).toBe('cancelled');
    expect(outcomes[0]?.[1]).toBe(harness.fxInput);
    expect(harness.fxInput.value).toBe('=SUM(A1)');
    expect(harness.fxInput.selectionStart).toBe(2);
    expect(harness.fxInput.selectionEnd).toBe(5);
    expect(harness.controller.isEditing()).toBe(true);
    harness.controller.cancelFx();

    // A context-invalid token must not resurrect the old bar or overwrite a
    // newer value. The lease itself is the sole owner of this validity check.
    harness.fxInput.value = '=9';
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    let contextCurrent = true;
    const staleLease = harness.controller.suspendForFormulaPalette({
      getLocale: () => 'en',
      contextCurrent: () => contextCurrent,
    });
    expect(staleLease).not.toBeNull();
    contextCurrent = false;
    expect(staleLease?.userCancel()).toBeNull();
    expect(harness.fxInput.value).toBe('=9');
    harness.fxInput.dispatchEvent(new FocusEvent('blur'));
    harness.fxInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', cancelable: true }));
    expect(sheet.workbook.cellFormula(anchor)).toBeNull();

    harness.fxInput.value = '=10';
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    const discardLease = harness.controller.suspendForFormulaPalette({
      getLocale: () => 'en',
      contextCurrent: () => true,
    });
    expect(discardLease).not.toBeNull();
    const discarded = asExternalDraftController(harness.controller).beginExternalDraft(
      anchor,
      '=',
      { onFinish: vi.fn() },
      { lease: discardLease ?? undefined },
    );
    if (!discarded) throw new Error('Discardable formula-bar draft did not start.');
    discarded.setValue('=11');
    discarded.discard();
    harness.fxInput.dispatchEvent(new FocusEvent('blur'));
    harness.fxInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', cancelable: true }));
    expect(sheet.workbook.cellFormula(anchor)).toBeNull();
    harness.detach();
  });

  it('notifies external draft subscribers for pointer insertion and both reference rotations', () => {
    sheet.host.dataset.fcPlatform = 'mac';
    const onValidation = vi.fn();
    const harness = attachFormulaBarHarness(
      sheet,
      onValidation,
      sheet.instance.store,
      () => sheet.instance.workbook,
    );
    const addr = { sheet: 0, row: 0, col: 0 };
    sheet.workbook.setNumber(addr, 1);
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, addr);
    sheet.instance.history.clear();
    const draft = asExternalDraftController(harness.controller).beginExternalDraft(addr, '=SUM(', {
      onFinish: vi.fn(),
    });
    if (!draft) throw new Error('External draft did not start.');

    const updates: string[] = [];
    draft.subscribe((raw) => updates.push(raw));
    harness.controller.insertRefAtCaret('A1');
    expect(draft.value()).toBe('=SUM(A1');

    harness.fxInput.dispatchEvent(
      new KeyboardEvent('keydown', { key: 't', metaKey: true, cancelable: true }),
    );
    expect(draft.value()).toBe('=SUM($A$1');

    harness.fxInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'F4', cancelable: true }));
    expect(draft.value()).toBe('=SUM(A$1');
    expect(updates).toEqual(['=SUM(A1', '=SUM($A$1', '=SUM(A$1']);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 1 });
    expect(sheet.instance.history.canUndo()).toBe(false);
    expect(onValidation).not.toHaveBeenCalled();

    draft.cancel();
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 1 });
    harness.detach();
  });

  it('fails closed when committed or cancelled handles target a newer draft', () => {
    const harness = attachFormulaBarHarness(
      sheet,
      vi.fn(),
      sheet.instance.store,
      () => sheet.instance.workbook,
    );
    const addr = { sheet: 0, row: 0, col: 0 };
    sheet.workbook.setNumber(addr, 1);
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, addr);
    sheet.instance.history.clear();

    const committed = asExternalDraftController(harness.controller).beginExternalDraft(addr, '=', {
      onFinish: vi.fn(),
    });
    if (!committed) throw new Error('Committed draft did not start.');
    committed.setValue('=ACOS(1)');
    expect(committed.commit()).toBe(true);

    const cancelled = asExternalDraftController(harness.controller).beginExternalDraft(
      addr,
      '=ACOS(1)',
      { onFinish: vi.fn() },
    );
    if (!cancelled) throw new Error('Cancelled draft did not start.');
    cancelled.setValue('=ACOS(0)');
    cancelled.cancel();

    sheet.instance.history.clear();
    const active = asExternalDraftController(harness.controller).beginExternalDraft(addr, '=', {
      onFinish: vi.fn(),
    });
    if (!active) throw new Error('Active draft did not start.');
    const updates: string[] = [];
    active.subscribe((raw) => updates.push(raw));

    committed.setValue('stale committed', 4);
    cancelled.setValue('stale cancelled', 4);

    expect(active.value()).toBe('=');
    expect(updates).toEqual([]);
    expect(sheet.instance.store.getState().ui.editorRefs).toEqual([]);
    expect(sheet.workbook.cellFormula(addr)).toBe('=ACOS(1)');
    expect(sheet.instance.history.canUndo()).toBe(false);
    active.cancel();
    harness.detach();
  });

  it.each([
    { label: 'enter', mode: { kind: 'enter', raw: 'inline' } },
    { label: 'edit', mode: { kind: 'edit', raw: 'inline', caret: 6 } },
  ] as const)(
    'does not replace an active inline $label editor with an external draft',
    ({ mode }) => {
      const cancelBindingEditor = vi.fn();
      const harness = attachFormulaBarHarness(
        sheet,
        vi.fn(),
        sheet.instance.store,
        () => sheet.instance.workbook,
        cancelBindingEditor,
      );
      mutators.setEditor(sheet.instance.store, mode);
      const beforeValue = harness.fxInput.value;

      const draft = asExternalDraftController(harness.controller).beginExternalDraft(
        { sheet: 0, row: 0, col: 0 },
        '=ACOS(',
        { onFinish: vi.fn() },
      );

      expect(draft).toBeNull();
      expect(cancelBindingEditor).not.toHaveBeenCalled();
      expect(harness.fxInput.value).toBe(beforeValue);
      expect(sheet.instance.store.getState().ui.editor).toEqual(mode);
      harness.detach();
    },
  );

  it('leaves browser Tab and Shift+Tab noncommitting during an external draft', () => {
    sheet.host.dataset.fcPlatform = 'mac';
    const harness = attachFormulaBarHarness(
      sheet,
      vi.fn(),
      sheet.instance.store,
      () => sheet.instance.workbook,
    );
    const addr = { sheet: 0, row: 0, col: 0 };
    sheet.workbook.setNumber(addr, 1);
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, addr);
    sheet.instance.history.clear();
    const draft = asExternalDraftController(harness.controller).beginExternalDraft(addr, '=', {
      onFinish: vi.fn(),
    });
    if (!draft) throw new Error('External draft did not start.');
    draft.setValue('42');

    for (const shiftKey of [false, true]) {
      const event = new KeyboardEvent('keydown', {
        key: 'Tab',
        shiftKey,
        cancelable: true,
      });
      harness.fxInput.dispatchEvent(event);
      expect(event.defaultPrevented).toBe(false);
      expect(draft.value()).toBe('42');
      expect(harness.controller.isEditing()).toBe(true);
      expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 1 });
      expect(sheet.instance.history.canUndo()).toBe(false);
    }

    draft.cancel();
    harness.detach();
  });

  it('commits ACOS at its fixed anchor as one undoable native entry', () => {
    const harness = attachFormulaBarHarness(
      sheet,
      vi.fn(),
      sheet.instance.store,
      () => sheet.instance.workbook,
    );
    const addr = { sheet: 0, row: 0, col: 0 };
    sheet.workbook.setNumber(addr, 1);
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, addr);
    sheet.instance.history.clear();
    const draft = asExternalDraftController(harness.controller).beginExternalDraft(addr, '=', {
      onFinish: vi.fn(),
    });
    if (!draft) throw new Error('External draft did not start.');
    draft.setValue('=ACOS(1)');

    expect(draft.commit()).toBe(true);
    expect(sheet.workbook.cellFormula(addr)).toBe('=ACOS(1)');
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 0 });
    expect(sheet.instance.history.canUndo()).toBe(true);
    expect(sheet.instance.history.canRedo()).toBe(false);

    expect(sheet.instance.undo()).toBe(true);
    expect(sheet.workbook.cellFormula(addr)).toBeNull();
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 1 });
    expect(sheet.instance.history.canUndo()).toBe(false);
    expect(sheet.instance.history.canRedo()).toBe(true);

    expect(sheet.instance.redo()).toBe(true);
    expect(sheet.workbook.cellFormula(addr)).toBe('=ACOS(1)');
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 0 });
    expect(sheet.instance.history.canUndo()).toBe(true);
    expect(sheet.instance.history.canRedo()).toBe(false);
    harness.detach();
  });

  it('settles one native commit while a controller subscriber reenters every draft route', () => {
    sheet.host.dataset.fcPlatform = 'mac';
    const harness = attachFormulaBarHarness(
      sheet,
      vi.fn(),
      sheet.instance.store,
      () => sheet.instance.workbook,
    );
    const anchor = { sheet: 0, row: 0, col: 0 };
    const other = { sheet: 0, row: 2, col: 2 };
    sheet.workbook.setNumber(anchor, 1);
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, other);
    sheet.instance.history.clear();
    const outcomes: string[] = [];
    const draft = asExternalDraftController(harness.controller).beginExternalDraft(anchor, '=', {
      onFinish: (outcome) => outcomes.push(outcome),
    });
    if (!draft) throw new Error('External draft did not start.');
    draft.setValue('=ACOS(1)');

    let reentered = false;
    const unsubscribe = sheet.instance.commands.subscribe(() => {
      if (reentered) return;
      reentered = true;
      harness.controller.commitFx('none');
      harness.controller.acceptFx();
      harness.controller.cancelFx();
      harness.controller.syncFxRefs();
      harness.controller.insertRefAtCaret('C3');
      harness.fxInput.dispatchEvent(new Event('input'));
      harness.fxInput.dispatchEvent(new FocusEvent('focus'));
      harness.fxInput.dispatchEvent(new FocusEvent('blur'));
      harness.fxInput.dispatchEvent(
        new KeyboardEvent('keydown', { key: 'Enter', metaKey: true, cancelable: true }),
      );
      harness.fxInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', cancelable: true }));
      harness.fxInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
      harness.fxInput.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter' }));
    });

    expect(draft.commit()).toBe(true);
    expect(reentered).toBe(true);
    expect(sheet.workbook.cellFormula(anchor)).toBe('=ACOS(1)');
    expect(sheet.workbook.getValue(anchor)).toEqual({ kind: 'number', value: 0 });
    expect(sheet.workbook.getValue(other)).toEqual({ kind: 'blank' });
    expect(sheet.instance.history.canUndo()).toBe(true);
    expect(outcomes).toEqual(['committed']);
    expect(draft.commit()).toBe(false);
    expect(sheet.instance.undo()).toBe(true);
    expect(sheet.workbook.getValue(anchor)).toEqual({ kind: 'number', value: 1 });
    expect(sheet.workbook.getValue(other)).toEqual({ kind: 'blank' });
    expect(sheet.instance.history.canUndo()).toBe(false);
    unsubscribe();
    harness.detach();
  });

  it('commits a rejected draft only once when a controller subscriber sends Escape', () => {
    const harness = attachFormulaBarHarness(
      sheet,
      vi.fn(),
      sheet.instance.store,
      () => sheet.instance.workbook,
    );
    const addr = { sheet: 0, row: 0, col: 0 };
    sheet.workbook.setNumber(addr, 1);
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, addr);
    sheet.instance.setPolicy({ readOnly: true });
    sheet.instance.history.clear();
    const outcomes: string[] = [];
    const draft = asExternalDraftController(harness.controller).beginExternalDraft(addr, '=', {
      onFinish: (outcome) => outcomes.push(outcome),
    });
    if (!draft) throw new Error('External draft did not start.');
    draft.setValue('denied');
    let rejected = false;
    const unsubscribe = sheet.instance.commands.subscribe((result) => {
      if (rejected || result.status !== 'rejected') return;
      rejected = true;
      harness.fxInput.dispatchEvent(
        new KeyboardEvent('keydown', { key: 'Escape', cancelable: true }),
      );
    });

    expect(draft.commit()).toBe(false);
    expect(rejected).toBe(true);
    expect(outcomes).toEqual(['cancelled']);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 1 });
    expect(sheet.instance.history.canUndo()).toBe(false);
    expect(harness.controller.isEditing()).toBe(false);
    expect(harness.fxInput.value).toBe('');
    unsubscribe();
    sheet.instance.setPolicy();
    harness.detach();
  });

  it('suppresses detached refocus when a rejected controller subscriber detaches', () => {
    const onValidation = vi.fn();
    const harness = attachFormulaBarHarness(
      sheet,
      onValidation,
      sheet.instance.store,
      () => sheet.instance.workbook,
    );
    const addr = { sheet: 0, row: 0, col: 0 };
    sheet.workbook.setNumber(addr, 1);
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, addr);
    sheet.instance.setPolicy({ readOnly: true });
    sheet.instance.history.clear();
    const outcomes: string[] = [];
    const draft = asExternalDraftController(harness.controller).beginExternalDraft(addr, '=', {
      onFinish: (outcome) => outcomes.push(outcome),
    });
    if (!draft) throw new Error('External draft did not start.');
    draft.setValue('denied');
    let detached = false;
    const unsubscribe = sheet.instance.commands.subscribe((result) => {
      if (detached || result.status !== 'rejected') return;
      detached = true;
      harness.controller.detach();
    });

    expect(draft.commit()).toBe(false);
    expect(detached).toBe(true);
    expect(outcomes).toEqual(['cancelled']);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 1 });
    expect(sheet.instance.history.canUndo()).toBe(false);
    expect(document.activeElement).not.toBe(harness.fxInput);
    expect(onValidation).toHaveBeenCalledWith(expect.objectContaining({ severity: 'stop' }));
    expect(
      asExternalDraftController(harness.controller).beginExternalDraft(addr, '=', {
        onFinish: vi.fn(),
      }),
    ).toBeNull();
    unsubscribe();
  });

  it('finishes an external commit once when detach reenters during controller notification', () => {
    const harness = attachFormulaBarHarness(
      sheet,
      vi.fn(),
      sheet.instance.store,
      () => sheet.instance.workbook,
    );
    const addr = { sheet: 0, row: 0, col: 0 };
    sheet.workbook.setNumber(addr, 1);
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, addr);
    sheet.instance.history.clear();
    const outcomes: string[] = [];
    const draft = asExternalDraftController(harness.controller).beginExternalDraft(addr, '=', {
      onFinish: (outcome) => outcomes.push(outcome),
    });
    if (!draft) throw new Error('External draft did not start.');
    draft.setValue('=ACOS(1)');
    let detached = false;
    const unsubscribe = sheet.instance.commands.subscribe(() => {
      if (detached) return;
      detached = true;
      harness.controller.detach();
    });

    expect(draft.commit()).toBe(true);
    expect(outcomes).toEqual(['committed']);
    expect(sheet.workbook.cellFormula(addr)).toBe('=ACOS(1)');
    expect(sheet.instance.history.canUndo()).toBe(true);
    expect(draft.commit()).toBe(false);
    expect(
      asExternalDraftController(harness.controller).beginExternalDraft(addr, '=', {
        onFinish: vi.fn(),
      }),
    ).toBeNull();
    unsubscribe();
  });

  it('contains validation callback throws without wedging warning writes or stop retries', () => {
    const addr = { sheet: 0, row: 0, col: 0 };
    sheet.workbook.setNumber(addr, 1);
    sheet.instance.history.clear();
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});

    const warningStore = createSpreadsheetStore();
    mutators.setActive(warningStore, addr);
    mutators.setCellFormat(warningStore, addr, {
      validation: {
        kind: 'whole',
        op: '=',
        a: 5,
        errorStyle: 'warning',
        errorMessage: 'Value is not five.',
        showErrorMessage: true,
      },
    });
    const warningValidation = vi.fn(() => {
      throw new Error('warning observer failed');
    });
    const warningHarness = attachFormulaBarHarness(
      sheet,
      warningValidation,
      warningStore,
      () => sheet.workbook,
    );
    const warningOutcomes: string[] = [];
    const warningDraft = asExternalDraftController(warningHarness.controller).beginExternalDraft(
      addr,
      '=',
      {
        onFinish: (outcome) => warningOutcomes.push(outcome),
      },
    );
    if (!warningDraft) throw new Error('Warning draft did not start.');
    warningDraft.setValue('4');

    expect(() => warningDraft.commit()).not.toThrow();
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 4 });
    expect(warningValidation).toHaveBeenCalledTimes(1);
    expect(warningOutcomes).toEqual(['committed']);
    expect(warningHarness.controller.isEditing()).toBe(false);
    expect(sheet.instance.history.canUndo()).toBe(true);
    expect(sheet.instance.undo()).toBe(true);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 1 });
    expect(sheet.instance.history.canUndo()).toBe(false);
    warningHarness.detach();

    const stopStore = createSpreadsheetStore();
    mutators.setActive(stopStore, addr);
    mutators.setCellFormat(stopStore, addr, {
      validation: {
        kind: 'whole',
        op: '=',
        a: 5,
        errorStyle: 'stop',
        errorMessage: 'Value must be five.',
        showErrorMessage: true,
      },
    });
    const stopValidation = vi.fn(() => {
      throw new Error('stop observer failed');
    });
    const stopHarness = attachFormulaBarHarness(
      sheet,
      stopValidation,
      stopStore,
      () => sheet.workbook,
    );
    const stopOutcomes: string[] = [];
    const stopDraft = asExternalDraftController(stopHarness.controller).beginExternalDraft(
      addr,
      '=',
      { onFinish: (outcome) => stopOutcomes.push(outcome) },
    );
    if (!stopDraft) throw new Error('Stop draft did not start.');
    stopDraft.setValue('4');

    expect(() => stopDraft.commit()).not.toThrow();
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 1 });
    expect(stopValidation).toHaveBeenCalledTimes(1);
    expect(stopOutcomes).toEqual([]);
    expect(stopHarness.controller.isEditing()).toBe(true);
    stopDraft.setValue('5');
    expect(stopDraft.commit()).toBe(true);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 5 });
    expect(stopOutcomes).toEqual(['committed']);
    stopHarness.detach();
    warn.mockRestore();
  });

  it('contains cancellation finish throws and keeps terminal ownership until cleanup', () => {
    const harness = attachFormulaBarHarness(
      sheet,
      vi.fn(),
      sheet.instance.store,
      () => sheet.instance.workbook,
    );
    const addr = { sheet: 0, row: 0, col: 0 };
    let draft: ExternalDraftTestHandle;
    let reentrant: ExternalDraftTestHandle | null = null;
    const finish = vi.fn(() => {
      reentrant = asExternalDraftController(harness.controller).beginExternalDraft(
        addr,
        '=ACOS(1)',
        { onFinish: vi.fn() },
      );
      draft.cancel();
      throw new Error('finish observer failed');
    });
    const started = asExternalDraftController(harness.controller).beginExternalDraft(addr, '=', {
      onFinish: finish,
    });
    if (!started) throw new Error('External draft did not start.');
    draft = started;

    expect(() => draft.cancel()).not.toThrow();
    expect(finish).toHaveBeenCalledTimes(1);
    expect(reentrant).toBeNull();
    expect(draft.commit()).toBe(false);
    const next = asExternalDraftController(harness.controller).beginExternalDraft(addr, '=', {
      onFinish: vi.fn(),
    });
    expect(next).not.toBeNull();
    next?.cancel();
    harness.detach();
  });

  it('detaches open drafts write-free and rejects stale handles and reopening', () => {
    const harness = attachFormulaBarHarness(
      sheet,
      vi.fn(),
      sheet.instance.store,
      () => sheet.instance.workbook,
    );
    const addr = { sheet: 0, row: 0, col: 0 };
    sheet.workbook.setNumber(addr, 1);
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, addr);
    sheet.instance.history.clear();
    const outcomes: string[] = [];
    const draft = asExternalDraftController(harness.controller).beginExternalDraft(addr, '=', {
      onFinish: (outcome) => outcomes.push(outcome),
    });
    if (!draft) throw new Error('External draft did not start.');
    draft.setValue('=ACOS(1)');

    harness.controller.detach();
    expect(outcomes).toEqual(['cancelled']);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 1 });
    expect(sheet.instance.history.canUndo()).toBe(false);
    draft.setValue('stale');
    expect(draft.value()).toBe('');
    expect(draft.commit()).toBe(false);
    expect(
      asExternalDraftController(harness.controller).beginExternalDraft(addr, '=', {
        onFinish: vi.fn(),
      }),
    ).toBeNull();
  });

  it('cancels external drafts without writing and fails closed after workbook identity changes', async () => {
    const harness = attachFormulaBarHarness(
      sheet,
      vi.fn(),
      sheet.instance.store,
      () => sheet.instance.workbook,
    );
    const addr = { sheet: 0, row: 0, col: 0 };
    sheet.workbook.setNumber(addr, 1);
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, addr);
    sheet.instance.history.clear();
    harness.fxInput.value = '1';
    const cancelled: string[] = [];
    const draft = asExternalDraftController(harness.controller).beginExternalDraft(addr, '=', {
      onFinish: (outcome) => cancelled.push(outcome),
    });
    if (!draft) throw new Error('External draft did not start.');
    draft.setValue('9');
    harness.fxInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 1 });
    expect(harness.fxInput.value).toBe('1');
    expect(sheet.instance.history.canUndo()).toBe(false);
    expect(cancelled).toEqual(['cancelled']);
    expect(draft.commit()).toBe(false);
    draft.cancel();

    const second = await WorkbookHandle.createDefault();
    expect(second.isStub).toBe(false);
    harness.fxInput.value = '1';
    const swapped = asExternalDraftController(harness.controller).beginExternalDraft(addr, '=', {
      onFinish: (outcome) => cancelled.push(outcome),
    });
    if (!swapped) throw new Error('External draft did not restart.');
    swapped.setValue('77');
    await sheet.instance.setWorkbook(second);
    expect(swapped.commit()).toBe(false);
    expect(second.getValue(addr)).toEqual({ kind: 'blank' });
    expect(sheet.instance.history.canUndo()).toBe(false);
    swapped.cancel();
    expect(second.getValue(addr)).toEqual({ kind: 'blank' });
    harness.detach();
  });

  it('keeps rejected external drafts open with raw input and pending format intact', () => {
    const harness = attachFormulaBarHarness(
      sheet,
      vi.fn(),
      sheet.instance.store,
      () => sheet.instance.workbook,
    );
    const addr = { sheet: 0, row: 0, col: 0 };
    sheet.workbook.setNumber(addr, 1);
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, addr);
    sheet.instance.history.clear();
    sheet.instance.setPolicy({ readOnly: true });
    const draft = asExternalDraftController(harness.controller).beginExternalDraft(addr, '=', {
      onFinish: vi.fn(),
    });
    if (!draft) throw new Error('External draft did not start.');
    draft.setValue('denied');
    mutators.setPendingFormat(sheet.instance.store, { addr, format: { italic: true } });

    expect(draft.commit()).toBe(false);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 1 });
    expect(draft.value()).toBe('denied');
    expect(harness.controller.isEditing()).toBe(true);
    expect(sheet.instance.store.getState().ui.pendingFormat).toEqual({
      addr,
      format: { italic: true },
    });
    expect(sheet.instance.history.canUndo()).toBe(false);
    draft.cancel();
    sheet.instance.setPolicy();
    harness.detach();
  });

  it('preserves external drafts on controller throws and fallback stop validation', () => {
    const harness = attachFormulaBarHarness(
      sheet,
      vi.fn(),
      sheet.instance.store,
      () => sheet.instance.workbook,
    );
    const addr = { sheet: 0, row: 0, col: 0 };
    sheet.workbook.setNumber(addr, 1);
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, addr);
    mutators.setPendingFormat(sheet.instance.store, { addr, format: { underline: true } });
    const execute = vi.spyOn(sheet.instance.commands, 'execute').mockImplementation(() => {
      throw new Error('controller failed');
    });
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const draft = asExternalDraftController(harness.controller).beginExternalDraft(addr, '=', {
      onFinish: vi.fn(),
    });
    if (!draft) throw new Error('External draft did not start.');
    draft.setValue('throw');
    expect(draft.commit()).toBe(false);
    expect(draft.value()).toBe('throw');
    expect(sheet.instance.store.getState().ui.pendingFormat).toEqual({
      addr,
      format: { underline: true },
    });
    expect(harness.controller.isEditing()).toBe(true);
    draft.cancel();
    execute.mockRestore();
    warn.mockRestore();
    harness.detach();

    const store = createSpreadsheetStore();
    const onValidation = vi.fn();
    const fallback = attachFormulaBarHarness(sheet, onValidation, store, () => sheet.workbook);
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
    mutators.setPendingFormat(store, { addr, format: { italic: true } });
    fallback.fxInput.value = '1';
    const stopDraft = asExternalDraftController(fallback.controller).beginExternalDraft(addr, '=', {
      onFinish: vi.fn(),
    });
    if (!stopDraft) throw new Error('Fallback external draft did not start.');
    stopDraft.setValue('4');
    expect(stopDraft.commit()).toBe(false);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 1 });
    expect(stopDraft.value()).toBe('4');
    expect(store.getState().ui.pendingFormat).toEqual({ addr, format: { italic: true } });
    expect(fallback.controller.isEditing()).toBe(true);
    expect(onValidation).toHaveBeenCalledWith(
      expect.objectContaining({ severity: 'stop', message: 'Value must be five.' }),
    );
    stopDraft.cancel();
    fallback.detach();
  });

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

  it('records a formula-bar value and pending nested format as one undo step', () => {
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(sheet.instance.store, addr);
    sheet.instance.history.clear();
    mutators.setPendingFormat(sheet.instance.store, {
      addr,
      format: { bold: true, borders: { bottom: { style: 'thin' } } },
    });
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.value = '42';
    harness.fxInput.dispatchEvent(new Event('input'));

    expect(harness.controller.commitFx('none')).toBe(true);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 42 });
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')).toMatchObject({
      bold: true,
      borders: { bottom: { style: 'thin' } },
    });
    expect(sheet.instance.store.getState().ui.pendingFormat).toBeNull();
    expect(sheet.instance.history.canUndo()).toBe(true);
    expect(sheet.instance.history.undo()).toBe(true);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'blank' });
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')).toBeUndefined();
    expect(sheet.instance.history.canUndo()).toBe(false);
    expect(sheet.instance.history.redo()).toBe(true);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 42 });
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')).toMatchObject({
      bold: true,
      borders: { bottom: { style: 'thin' } },
    });
    harness.detach();
  });

  it('records a pending style-only edit when the committed input is unchanged', () => {
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    const addr = { sheet: 0, row: 0, col: 0 };
    sheet.workbook.setNumber(addr, 42);
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, addr);
    sheet.instance.history.clear();
    mutators.setPendingFormat(sheet.instance.store, { addr, format: { italic: true } });
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.value = '42';
    harness.fxInput.dispatchEvent(new Event('input'));

    expect(harness.controller.commitFx('none')).toBe(true);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 42 });
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')?.italic).toBe(true);
    expect(sheet.instance.history.canUndo()).toBe(true);
    sheet.instance.setPolicy({ operations: { valueEdit: true, format: false } });
    expect(sheet.instance.history.undo()).toBe(false);
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')?.italic).toBe(true);
    sheet.instance.setPolicy({ operations: { valueEdit: true, format: true } });
    expect(sheet.instance.history.undo()).toBe(true);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'number', value: 42 });
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')).toBeUndefined();
    sheet.instance.setPolicy({ operations: { valueEdit: true, format: false } });
    expect(sheet.instance.history.redo()).toBe(false);
    sheet.instance.setPolicy({ operations: { valueEdit: true, format: true } });
    expect(sheet.instance.history.redo()).toBe(true);
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')?.italic).toBe(true);
    harness.detach();
  });

  it('records an external draft value and pending style as one undo step', () => {
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    const addr = { sheet: 0, row: 0, col: 0 };
    const finished = vi.fn();
    mutators.setActive(sheet.instance.store, addr);
    sheet.instance.history.clear();
    mutators.setPendingFormat(sheet.instance.store, {
      addr,
      format: { bold: true, numFmt: { kind: 'text' } },
    });
    const draft = asExternalDraftController(harness.controller).beginExternalDraft(addr, '=', {
      onFinish: finished,
    });
    if (!draft) throw new Error('External draft did not start.');
    draft.setValue('42');

    expect(draft.commit()).toBe(true);
    expect(finished).toHaveBeenCalledWith('committed');
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'text', value: '42' });
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')?.bold).toBe(true);
    expect(sheet.instance.history.canUndo()).toBe(true);
    expect(sheet.instance.history.undo()).toBe(true);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'blank' });
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')).toBeUndefined();
    expect(sheet.instance.history.canUndo()).toBe(false);
    expect(sheet.instance.history.redo()).toBe(true);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'text', value: '42' });
    harness.detach();
  });

  it('rejects a pending format denied by policy without consuming the draft', () => {
    const onValidation = vi.fn();
    const harness = attachFormulaBarHarness(sheet, onValidation);
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(sheet.instance.store, addr);
    sheet.instance.history.clear();
    sheet.instance.setPolicy({ operations: { valueEdit: true, format: false } });
    mutators.setPendingFormat(sheet.instance.store, { addr, format: { bold: true } });
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.value = '42';
    harness.fxInput.dispatchEvent(new Event('input'));

    expect(harness.controller.commitFx('none')).toBe(false);
    expect(sheet.workbook.getValue(addr)).toEqual({ kind: 'blank' });
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')).toBeUndefined();
    expect(sheet.instance.store.getState().ui.pendingFormat).toEqual({
      addr,
      format: { bold: true },
    });
    expect(sheet.instance.history.canUndo()).toBe(false);
    expect(harness.controller.isEditing()).toBe(true);
    expect(onValidation).toHaveBeenCalledWith(expect.objectContaining({ severity: 'stop' }));
    sheet.instance.setPolicy();
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
