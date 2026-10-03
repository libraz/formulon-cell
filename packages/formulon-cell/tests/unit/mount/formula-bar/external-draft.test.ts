import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { mutators } from '../../../../src/store/store.js';
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
});
