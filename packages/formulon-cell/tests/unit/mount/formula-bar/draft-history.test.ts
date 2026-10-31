import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/index.js';
import { asExternalDraftController, attachFormulaBarHarness } from './fixtures.js';

describe('mount/formula-bar — active draft history', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
    sheet.instance.history.clear();
  });

  afterEach(() => sheet.dispose());

  it('groups pointer picks and point continuation and records F4 as one command', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 3, col: 3 });
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    const { controller, fxInput: input } = harness;
    input.value = '=SUM(';
    input.focus();
    input.dispatchEvent(new FocusEvent('focus'));
    input.setSelectionRange(input.value.length, input.value.length);
    controller.insertRefAtCaret('B5', { phase: 'start' });
    const arrow = new KeyboardEvent('keydown', { key: 'ArrowRight', cancelable: true });
    input.dispatchEvent(arrow);
    expect(arrow.defaultPrevented).toBe(true);
    expect(input.value).toBe('=SUM(C5');

    input.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, cancelable: true }),
    );
    expect(input.value).toBe('=SUM(');
    input.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, shiftKey: true, cancelable: true }),
    );
    expect(input.value).toBe('=SUM(C5');

    input.value = '=SUM(A1)';
    input.setSelectionRange(8, 8);
    input.dispatchEvent(new Event('input'));
    input.setSelectionRange(7, 7);
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'F4', cancelable: true }));
    expect(input.value).toBe('=SUM($A$1)');
    input.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, cancelable: true }),
    );
    expect(input.value).toBe('=SUM(A1)');
    input.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, shiftKey: true, cancelable: true }),
    );
    expect(input.value).toBe('=SUM($A$1)');
    expect(sheet.instance.history.canUndo()).toBe(false);
    harness.detach();
  });

  it('records autocomplete and F9 changes as atomic commands', () => {
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    const { controller, fxInput: input } = harness;
    input.value = '=SU';
    input.focus();
    input.dispatchEvent(new FocusEvent('focus'));
    controller.beforeAutocompleteInsert?.();
    input.value = '=SUM(';
    input.setSelectionRange(5, 5);
    controller.notifyAutocompleteInsert?.({ kind: 'function' });
    input.dispatchEvent(new Event('input'));
    input.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, cancelable: true }),
    );
    expect(input.value).toBe('=SU');
    input.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, shiftKey: true, cancelable: true }),
    );
    expect(input.value).toBe('=SUM(');

    const source = { sheet: 0, row: 0, col: 0 };
    sheet.workbook.setNumber(source, 7);
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    input.value = '=A1+1';
    input.setSelectionRange(1, 3);
    input.dispatchEvent(new Event('input'));
    input.setSelectionRange(1, 3);
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'F9', cancelable: true }));
    expect(input.value).toBe('=7+1');
    input.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, cancelable: true }),
    );
    expect(input.value).toBe('=A1+1');
    harness.controller.cancelFx();
    harness.detach();
  });

  it('leaves a read-only formula display untouched on F9', () => {
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    const { fxInput: input } = harness;
    sheet.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 7);
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    input.readOnly = true;
    input.value = '=A1+1';
    input.focus();
    input.dispatchEvent(new FocusEvent('focus'));
    input.setSelectionRange(1, 3);
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'F9', cancelable: true }));
    expect(input.value).toBe('=A1+1');
    harness.detach();
  });

  it('leaves a read-only formula display untouched on Cmd+T', () => {
    sheet.host.dataset.fcPlatform = 'mac';
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    const { fxInput: input } = harness;
    input.readOnly = true;
    input.value = '=A1+1';
    input.focus();
    input.dispatchEvent(new FocusEvent('focus'));
    input.setSelectionRange(3, 3);
    input.dispatchEvent(
      new KeyboardEvent('keydown', { key: 't', metaKey: true, cancelable: true }),
    );
    expect(input.value).toBe('=A1+1');
    harness.detach();
  });

  it('undoes ordinary typing and prevents a cancelled history event from undoing twice', () => {
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    const { controller, fxInput: input } = harness;
    input.value = '=A1';
    input.setSelectionRange(3, 3);
    input.focus();
    input.dispatchEvent(new FocusEvent('focus'));

    const type = (value: string): void => {
      input.dispatchEvent(
        new InputEvent('beforeinput', {
          inputType: 'insertText',
          bubbles: true,
          cancelable: true,
        }),
      );
      input.value = value;
      input.setSelectionRange(value.length, value.length);
      input.dispatchEvent(new InputEvent('input', { inputType: 'insertText', bubbles: true }));
    };
    type('=A1+2');
    type('=A1+23');

    input.dispatchEvent(
      new InputEvent('beforeinput', {
        inputType: 'historyUndo',
        bubbles: true,
        cancelable: true,
      }),
    );
    expect(input.value).toBe('=A1+2');
    input.value = '=A1';
    input.dispatchEvent(new InputEvent('input', { inputType: 'historyUndo', bubbles: true }));
    expect(input.value).toBe('=A1+2');

    input.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, cancelable: true }),
    );
    expect(input.value).toBe('=A1');
    expect(sheet.instance.history.canUndo()).toBe(false);
    controller.cancelFx();
    harness.detach();
  });

  it('overrides input-only native history with the current snapshot when local history is empty', () => {
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    const { controller, fxInput: input } = harness;
    input.value = '=5';
    input.setSelectionRange(2, 2);
    input.focus();
    input.dispatchEvent(new FocusEvent('focus'));
    input.dispatchEvent(
      new InputEvent('beforeinput', {
        inputType: 'historyUndo',
        bubbles: true,
        cancelable: false,
      }),
    );
    input.value = 'stale native value';
    input.dispatchEvent(new InputEvent('input', { inputType: 'historyUndo', bubbles: true }));

    expect(input.value).toBe('=5');
    expect(sheet.instance.history.canUndo()).toBe(false);
    controller.cancelFx();
    harness.detach();
  });

  it('keeps a committing external draft unchanged when terminal history input reenters', () => {
    const address = { sheet: 0, row: 0, col: 0 };
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    const { controller, fxInput: input } = harness;
    const draft = asExternalDraftController(controller).beginExternalDraft(address, '=1', {
      onFinish: vi.fn(),
    });
    if (!draft) throw new Error('External draft did not start.');
    draft.setValue('=2');

    const originalApplyCellPatch = sheet.workbook.applyCellPatchAtomic.bind(sheet.workbook);
    const terminalUndo = new InputEvent('beforeinput', {
      inputType: 'historyUndo',
      bubbles: true,
      cancelable: true,
    });
    const applyCellPatch = vi
      .spyOn(sheet.workbook, 'applyCellPatchAtomic')
      .mockImplementation((patches) => {
        input.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
        input.dispatchEvent(terminalUndo);
        input.value = '=1';
        input.dispatchEvent(
          new InputEvent('input', {
            inputType: 'insertCompositionText',
            isComposing: true,
            bubbles: true,
          }),
        );
        input.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true }));
        return originalApplyCellPatch(patches);
      });

    expect(draft.commit()).toBe(true);

    expect(applyCellPatch).toHaveBeenCalledOnce();
    expect(terminalUndo.defaultPrevented).toBe(true);
    expect(input.value).toBe('=2');
    expect(sheet.workbook.cellFormula(address)).toBe('=2');
    expect(draft.snapshot()).toBeNull();
    expect(sheet.instance.history.canUndo()).toBe(true);
    harness.detach();
  });

  it('records external-draft F4 and F9 before store observers can synchronously undo', () => {
    const address = { sheet: 0, row: 0, col: 0 };
    sheet.workbook.setNumber(address, 7);
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    const { controller, fxInput: input } = harness;
    const draft = asExternalDraftController(controller).beginExternalDraft(address, '=A1', {
      onFinish: vi.fn(),
    });
    if (!draft) throw new Error('External draft did not start.');
    draft.setValue('=A1+1');
    const undo = (): void => {
      input.dispatchEvent(
        new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, bubbles: true, cancelable: true }),
      );
    };
    const originalSetSelectionRange = input.setSelectionRange.bind(input);
    vi.spyOn(input, 'setSelectionRange').mockImplementation((start, end, direction) => {
      originalSetSelectionRange(start, end, direction);
      input.dispatchEvent(new Event('select'));
    });
    let reentered = false;
    const unsubscribe = draft.subscribe((view) => {
      if (reentered || view.raw !== '=$A$1+1') return;
      reentered = true;
      undo();
    });
    input.setSelectionRange(3, 3);
    input.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'F4', bubbles: true, cancelable: true }),
    );
    unsubscribe();

    expect(reentered).toBe(true);
    expect(input.value).toBe('=A1+1');
    undo();
    expect(input.value).toBe('=A1');

    draft.setValue('=A1+1');
    reentered = false;
    const unsubscribeF9 = draft.subscribe((view) => {
      if (reentered || view.raw !== '=7+1') return;
      reentered = true;
      undo();
    });
    input.setSelectionRange(1, 3);
    input.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'F9', bubbles: true, cancelable: true }),
    );
    unsubscribeF9();

    expect(reentered).toBe(true);
    expect(input.value).toBe('=A1+1');
    undo();
    expect(input.value).toBe('=A1');
    draft.cancel();
    harness.detach();
  });

  it('records an unphased formula-bar reference pick before observers can undo', () => {
    const address = { sheet: 0, row: 3, col: 3 };
    mutators.setActive(sheet.instance.store, address);
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    const { controller, fxInput: input } = harness;
    const draft = asExternalDraftController(controller).beginExternalDraft(address, '=SUM', {
      onFinish: vi.fn(),
    });
    if (!draft) throw new Error('External draft did not start.');
    draft.setValue('=SUM(');

    let reentered = false;
    const unsubscribe = sheet.instance.store.subscribe(() => {
      if (reentered || input.value !== '=SUM(C4') return;
      reentered = true;
      input.dispatchEvent(
        new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, bubbles: true, cancelable: true }),
      );
    });
    controller.insertRefAtCaret('C4');
    unsubscribe();

    expect(reentered).toBe(true);
    expect(input.value).toBe('=SUM(');
    input.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, bubbles: true, cancelable: true }),
    );
    expect(input.value).toBe('=SUM');
    draft.cancel();
    harness.detach();
  });

  it('keeps the external palette history separate and restores the formula-bar owner journal', () => {
    const address = { sheet: 0, row: 3, col: 3 };
    mutators.setActive(sheet.instance.store, address);
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    const { controller, fxInput: input } = harness;
    input.value = '=SUM(';
    input.focus();
    input.dispatchEvent(new FocusEvent('focus'));
    controller.insertRefAtCaret('B5', { phase: 'start' });
    const lease = controller.suspendForFormulaPalette({
      getLocale: () => 'en-US',
      contextCurrent: () => true,
    });
    expect(lease).not.toBeNull();
    const draft = asExternalDraftController(controller).beginExternalDraft(
      address,
      '=',
      { onFinish: vi.fn() },
      { lease: lease ?? undefined },
    );
    if (!draft) throw new Error('External draft did not start.');
    draft.setValue('=SUM(9)');
    input.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, cancelable: true }),
    );
    expect(input.value).toBe('=SUM(B5');
    draft.cancel();
    expect(input.value).toBe('=SUM(B5');
    input.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, cancelable: true }),
    );
    expect(input.value).toBe('=SUM(');
    controller.cancelFx();
    harness.detach();
  });

  it('resolves R1C1 reference highlights and F4 rotations against the edited cell', () => {
    const address = { sheet: 0, row: 3, col: 3 };
    mutators.setR1C1(sheet.instance.store, true);
    mutators.setActive(sheet.instance.store, address);
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    const { controller, fxInput: input } = harness;
    input.value = '=SUM(R[-2]C[-2])';
    input.focus();
    input.dispatchEvent(new FocusEvent('focus'));
    expect(sheet.instance.store.getState().ui.editorRefs).toEqual([
      { r0: 1, c0: 1, r1: 1, c1: 1, colorIndex: 0 },
    ]);
    const rotate = (): void => {
      const caret = input.value.indexOf(')');
      input.setSelectionRange(caret, caret);
      input.dispatchEvent(new KeyboardEvent('keydown', { key: 'F4', cancelable: true }));
    };
    rotate();
    expect(input.value).toBe('=SUM(R2C2)');
    rotate();
    expect(input.value).toBe('=SUM(R2C[-2])');
    rotate();
    expect(input.value).toBe('=SUM(R[-2]C2)');
    rotate();
    expect(input.value).toBe('=SUM(R[-2]C[-2])');
    controller.cancelFx();
    harness.detach();
  });
});
