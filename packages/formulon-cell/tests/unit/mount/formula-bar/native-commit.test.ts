import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { setProtectedSheet } from '../../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/index.js';
import { asExternalDraftController, attachFormulaBarHarness } from './fixtures.js';

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
