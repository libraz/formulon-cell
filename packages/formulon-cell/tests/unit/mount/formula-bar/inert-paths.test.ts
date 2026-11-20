import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/index.js';
import { attachFormulaBarHarness } from './fixtures.js';

describe('mount/formula-bar — paths that leave the bar inert', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
    sheet.instance.history.clear();
  });

  afterEach(() => sheet.dispose());

  it('releasing a suspended lease leaves the bar inert with a fresh baseline', () => {
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    const { controller, fxInput: input } = harness;
    const anchor = { sheet: 0, row: 0, col: 0 };
    input.value = '=1';
    input.focus();
    input.dispatchEvent(new FocusEvent('focus'));
    input.dispatchEvent(new Event('input'));
    expect(controller.isEditing()).toBe(true);
    const lease = controller.suspendForFormulaPalette({
      getLocale: () => 'en',
      contextCurrent: () => true,
    });
    if (!lease) throw new Error('Lease was not granted.');

    input.value = '=2';
    lease.discard();

    expect(controller.isEditing()).toBe(false);
    // The changed text is the new baseline: a blur must not commit it.
    input.dispatchEvent(new FocusEvent('blur'));
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', cancelable: true }));
    expect(sheet.workbook.cellFormula(anchor)).toBeNull();
    // The draft history was dropped: undo has nothing to step back to.
    input.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, cancelable: true }),
    );
    expect(input.value).toBe('=2');
    // A fresh focus starts a new edit session.
    input.dispatchEvent(new FocusEvent('focus'));
    expect(controller.isEditing()).toBe(true);
    controller.cancelFx();
    harness.detach();
  });

  it('a clean blur ends the edit and clears history and reference highlights', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 3, col: 3 });
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    const { controller, fxInput: input } = harness;
    input.value = '=A1+B2';
    input.focus();
    input.dispatchEvent(new FocusEvent('focus'));
    input.dispatchEvent(new Event('input'));
    expect(controller.isEditing()).toBe(true);
    expect(sheet.instance.store.getState().ui.editorRefs.length).toBeGreaterThan(0);

    input.dispatchEvent(new FocusEvent('blur'));

    expect(controller.isEditing()).toBe(false);
    expect(sheet.instance.store.getState().ui.editorRefs).toEqual([]);
    input.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, cancelable: true }),
    );
    expect(input.value).toBe('=A1+B2');
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 3, col: 3 })).toBeNull();
    harness.detach();
  });

  it('resolves R1C1 highlights against the active cell after a clean blur on another cell', () => {
    const store = sheet.instance.store;
    store.setState((s) => ({ ...s, ui: { ...s.ui, r1c1: true } }));
    mutators.setActive(store, { sheet: 0, row: 3, col: 3 });
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    const { controller, fxInput: input } = harness;
    input.value = '=R[1]C';
    input.focus();
    input.dispatchEvent(new FocusEvent('focus'));
    input.dispatchEvent(new FocusEvent('blur'));
    expect(controller.isEditing()).toBe(false);

    mutators.setActive(store, { sheet: 0, row: 10, col: 5 });
    controller.syncFxRefs();

    const refs = store.getState().ui.editorRefs;
    expect(refs).toHaveLength(1);
    expect(refs[0]).toMatchObject({ r0: 11, c0: 5, r1: 11, c1: 5 });
    harness.detach();
  });
});
