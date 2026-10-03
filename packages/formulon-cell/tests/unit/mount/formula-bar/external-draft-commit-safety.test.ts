import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/index.js';
import {
  asExternalDraftController,
  attachFormulaBarHarness,
  type ExternalDraftTestHandle,
} from './fixtures.js';

describe('mount/formula-bar — native commit result', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    const workbook = await WorkbookHandle.createDefault();
    sheet = await mountStubSheet({ workbook });
    expect(sheet.workbook.isStub).toBe(false);
  });

  afterEach(() => sheet.dispose());

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
    expect(outcomes).toEqual(['discarded']);
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
    expect(outcomes).toEqual(['discarded']);
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

  it('contains cancellation finish throws after clearing terminal ownership for reentry', () => {
    const harness = attachFormulaBarHarness(
      sheet,
      vi.fn(),
      sheet.instance.store,
      () => sheet.instance.workbook,
    );
    const addr = { sheet: 0, row: 0, col: 0 };
    let draft: ExternalDraftTestHandle;
    const reentrant: { handle: ExternalDraftTestHandle | null } = { handle: null };
    const finish = vi.fn(() => {
      reentrant.handle = asExternalDraftController(harness.controller).beginExternalDraft(
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
    expect(reentrant.handle).not.toBeNull();
    expect(draft.commit()).toBe(false);
    reentrant.handle?.cancel();
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
    expect(outcomes).toEqual(['discarded']);
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
    // This standalone controller is not the mounted formula-bar owner, so
    // probe the stale handle to trigger its own guarded discard first.
    expect(swapped.snapshot()).toBeNull();
    expect(cancelled).toEqual(['cancelled', 'discarded']);
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
});
