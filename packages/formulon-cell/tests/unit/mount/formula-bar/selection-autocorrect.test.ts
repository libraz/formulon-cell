import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/index.js';
import { attachFormulaBarHarness, selectRange } from './fixtures.js';

describe('mount/formula-bar - Control+Enter fill', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
  });
  afterEach(() => sheet.dispose());

  it('applies AutoCorrect to the text it fills', () => {
    const harness = attachFormulaBarHarness(sheet, vi.fn());
    selectRange(sheet, { sheet: 0, row: 0, col: 0 }, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.value = 'THe';
    harness.fxInput.dispatchEvent(new Event('input'));
    harness.fxInput.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, cancelable: true }),
    );
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'The',
    });
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
      kind: 'text',
      value: 'The',
    });
    harness.detach();
  });
});
