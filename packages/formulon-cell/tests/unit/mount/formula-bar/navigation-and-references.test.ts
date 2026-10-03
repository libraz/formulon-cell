import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { defaultStrings } from '../../../../src/i18n/strings.js';
import { attachFormulaBarController } from '../../../../src/mount/formula-bar.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/index.js';
import { attachFormulaBarHarness, selectRange } from './fixtures.js';

describe('mount/formula-bar — edit lifecycle', () => {
  let sheet: MountedStubSheet;
  let fxInput: HTMLTextAreaElement;
  let formulabar: HTMLDivElement;

  beforeEach(async () => {
    sheet = await mountStubSheet();
    fxInput = sheet.host.querySelector('.fc-host__formulabar-input') as HTMLTextAreaElement;
    formulabar = sheet.host.querySelector('.fc-host__formulabar') as HTMLDivElement;
  });

  afterEach(() => sheet.dispose());

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

  it('Control+Return fills the selection directly when no controller is registered', () => {
    const store = createSpreadsheetStore();
    const onValidation = vi.fn();
    const harness = attachFormulaBarHarness(sheet, onValidation, store);
    const active = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(store, active);
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.value = 'both';
    harness.fxInput.dispatchEvent(new Event('input'));
    harness.fxInput.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, cancelable: true }),
    );

    expect(onValidation).not.toHaveBeenCalled();
    expect(sheet.workbook.getValue(active)).toEqual({ kind: 'text', value: 'both' });
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
      kind: 'text',
      value: 'both',
    });
    expect(harness.controller.isEditing()).toBe(false);
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
