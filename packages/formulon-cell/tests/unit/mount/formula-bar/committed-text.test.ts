import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/index.js';
import { asExternalDraftController, attachFormulaBarHarness } from './fixtures.js';

// The engine stores a formula in its own spelling (`=sum(1.50,2)` is stored as `=SUM(1.5,2)`),
// so the bar must show what was stored.
const TYPED = '=sum(1.50,2)';

describe('mount/formula-bar — the committed text is the stored formula', () => {
  let sheet: MountedStubSheet;
  const addr = { sheet: 0, row: 1, col: 1 };

  beforeEach(async () => {
    sheet = await mountStubSheet({ workbook: await WorkbookHandle.createDefault() });
    expect(sheet.workbook.isStub).toBe(false);
  });

  afterEach(() => sheet.dispose());

  it('shows the stored formula after the accept button commits in place, and keeps it on Escape', () => {
    const fxInput = sheet.host.querySelector('.fc-host__formulabar-input') as HTMLTextAreaElement;
    const fxAccept = sheet.host.querySelector(
      '.fc-host__formulabar-action--accept',
    ) as HTMLButtonElement;
    mutators.setActive(sheet.instance.store, addr);
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.value = TYPED;
    fxInput.dispatchEvent(new Event('input'));
    fxAccept.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const stored = sheet.workbook.cellFormula(addr);
    expect(stored).not.toBeNull();
    expect(fxInput.value).toBe(stored);
    fxInput.focus();
    fxInput.dispatchEvent(new FocusEvent('focus'));
    fxInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(fxInput.value).toBe(stored);
  });

  it('takes the stored formula as the baseline after a commit in place', () => {
    const harness = attachFormulaBarHarness(
      sheet,
      vi.fn(),
      sheet.instance.store,
      () => sheet.workbook,
      () => {},
      (fxInput) => {
        fxInput.value = sheet.workbook.cellFormula(addr) ?? '';
      },
    );
    mutators.setActive(sheet.instance.store, addr);
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.value = TYPED;
    harness.fxInput.dispatchEvent(new Event('input'));
    expect(harness.controller.commitFx('none')).toBe(true);
    expect(harness.fxInput.value).toBe(sheet.workbook.cellFormula(addr));
    harness.detach();
  });

  it('shows the stored formula after an external draft commits', () => {
    const harness = attachFormulaBarHarness(
      sheet,
      vi.fn(),
      sheet.instance.store,
      () => sheet.workbook,
      () => {},
      (fxInput) => {
        fxInput.value = sheet.workbook.cellFormula(addr) ?? '';
      },
    );
    mutators.setActive(sheet.instance.store, addr);
    const draft = asExternalDraftController(harness.controller).beginExternalDraft(addr, '=', {
      onFinish: () => {},
    });
    if (!draft) throw new Error('External draft did not start.');
    draft.setValue(TYPED, TYPED.length);
    expect(draft.commit()).toBe(true);
    expect(harness.fxInput.value).toBe(sheet.workbook.cellFormula(addr));
    harness.detach();
  });
});
