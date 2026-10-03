import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/index.js';
import { type PaletteSetupArgs, paletteRoot, setupPalette } from './fixtures.js';

describe('interact/mac-formula-palette argument fields', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet({ workbook: await WorkbookHandle.createDefault() });
    expect(sheet.workbook.isStub).toBe(false);
  });

  afterEach(() => sheet.dispose());

  const setup = (...args: PaletteSetupArgs) => setupPalette(sheet, ...args);

  it('round-trips nested and quoted arguments, preserves unsynchronized raw input, and exposes the focused range target', () => {
    const { anchor, formulaBar, palette } = setup();
    palette.open('IF');
    const root = paletteRoot(palette);
    expect(root.dataset.state).toBe('arguments-editing');
    const first = root.querySelector<HTMLInputElement>('[data-argument-index="0"]');
    expect(first).not.toBeNull();
    const liveTarget = palette.rangeInsertTarget();
    expect(liveTarget?.isFormulaEdit()).toBe(true);
    liveTarget?.insertRefAtCaret('B3');
    expect(formulaBar.input.value).toBe('=IF(B3)');
    formulaBar.input.value = '=IF(SUM(A1,A2)>1,"a,b",ACOS(1))';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="0"]')?.value).toBe(
      'SUM(A1,A2)>1',
    );
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="1"]')?.value).toBe('"a,b"');
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="2"]')?.value).toBe(
      'ACOS(1)',
    );

    formulaBar.input.value = '=ACOS(';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    expect(formulaBar.input.value).toBe('=ACOS(');
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="0"]')?.value).toBe(
      'SUM(A1,A2)>1',
    );

    const target = palette.rangeInsertTarget();
    expect(target?.isFormulaEdit()).toBe(false);
    expect(target).not.toBeNull();
    target?.insertRefAtCaret('B3');
    expect(formulaBar.input.value).toBe('=ACOS(');
    expect(anchor).toEqual({ sheet: 0, row: 0, col: 0 });
  });

  it('preserves explicit trailing blank argument slots when reverse-projected raw is edited', () => {
    const { formulaBar, palette } = setup();
    palette.open('IF');
    const root = paletteRoot(palette);
    formulaBar.input.value = '=IF(FALSE,1,)';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="2"]')?.value).toBe('');

    const valueIfTrue = root.querySelector<HTMLInputElement>('[data-argument-index="1"]');
    expect(valueIfTrue).not.toBeNull();
    if (valueIfTrue) {
      valueIfTrue.value = '2';
      valueIfTrue.dispatchEvent(new Event('input', { bubbles: true }));
    }
    expect(formulaBar.input.value).toBe('=IF(FALSE,2,)');

    formulaBar.input.value = '=SUM({1,2},3)';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    expect(root.dataset.rawSynchronized).toBe('false');
    expect(formulaBar.input.value).toBe('=SUM({1,2},3)');
    palette.close();
  });

  it('keeps palette-owned multiargument edits and range insertion in the assembled formula', () => {
    const { formulaBar, palette } = setup();
    palette.open('IF');
    const root = paletteRoot(palette);
    const valueIfTrue = root.querySelector<HTMLInputElement>('[data-argument-index="1"]');
    const valueIfFalse = root.querySelector<HTMLInputElement>('[data-argument-index="2"]');
    expect(valueIfTrue).not.toBeNull();
    expect(valueIfFalse).not.toBeNull();
    if (valueIfTrue) {
      valueIfTrue.value = '1';
      valueIfTrue.dispatchEvent(new Event('input', { bubbles: true }));
    }
    if (valueIfFalse) {
      valueIfFalse.value = '2';
      valueIfFalse.dispatchEvent(new Event('input', { bubbles: true }));
    }
    expect(formulaBar.input.value).toBe('=IF(,1,2)');
    palette.rangeInsertTarget()?.insertRefAtCaret('B3');
    expect(formulaBar.input.value).toBe('=IF(,1,B3)');
  });

  it('grows externally projected IF arguments through fields and range insertion', () => {
    const { formulaBar, palette } = setup();
    palette.open('IF');
    const root = paletteRoot(palette);
    formulaBar.input.value = '=IF(TRUE)';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    const valueIfTrue = root.querySelector<HTMLInputElement>('[data-argument-index="1"]');
    const valueIfFalse = root.querySelector<HTMLInputElement>('[data-argument-index="2"]');
    expect(valueIfTrue).not.toBeNull();
    expect(valueIfFalse).not.toBeNull();
    if (valueIfTrue) {
      valueIfTrue.focus();
      valueIfTrue.value = '1';
      valueIfTrue.dispatchEvent(new Event('input', { bubbles: true }));
      expect(root.querySelector('[data-argument-index="1"]')).toBe(valueIfTrue);
    }
    expect(formulaBar.input.value).toBe('=IF(TRUE,1)');
    if (valueIfFalse) {
      valueIfFalse.focus();
      valueIfFalse.value = '';
      valueIfFalse.dispatchEvent(new Event('input', { bubbles: true }));
    }
    expect(formulaBar.input.value).toBe('=IF(TRUE,1,)');
    palette.rangeInsertTarget()?.insertRefAtCaret('B3');
    expect(formulaBar.input.value).toBe('=IF(TRUE,1,B3)');
  });

  it('reverse-projects array and structured-reference fields while preserving edits and blanks', () => {
    const { formulaBar, palette } = setup();
    palette.open('SUM');
    const root = paletteRoot(palette);
    formulaBar.input.value = '=SUM({1,2;3,4},Table1[[Last, First],[Amount]],)';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    expect(root.dataset.rawSynchronized).toBe('true');
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="0"]')?.value).toBe(
      '{1,2;3,4}',
    );
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="1"]')?.value).toBe(
      'Table1[[Last, First],[Amount]]',
    );
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="2"]')?.value).toBe('');

    const array = root.querySelector<HTMLInputElement>('[data-argument-index="0"]');
    expect(array).not.toBeNull();
    if (array) {
      array.value = '{5,6;7,8}';
      array.dispatchEvent(new Event('input', { bubbles: true }));
    }
    expect(formulaBar.input.value).toBe('=SUM({5,6;7,8},Table1[[Last, First],[Amount]],)');

    const structuredReference = root.querySelector<HTMLInputElement>('[data-argument-index="1"]');
    expect(structuredReference).not.toBeNull();
    structuredReference?.focus();
    palette.rangeInsertTarget()?.insertRefAtCaret('B3');
    expect(formulaBar.input.value).toBe('=SUM({5,6;7,8},B3,)');
  });

  it('fails closed for crossed, unclosed, or different-name raw formulas', () => {
    const { formulaBar, palette } = setup();
    palette.open('SUM');
    const root = paletteRoot(palette);

    formulaBar.input.value = '=SUM({1,2],3)';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    expect(root.dataset.rawSynchronized).toBe('false');
    expect(formulaBar.input.value).toBe('=SUM({1,2],3)');
    root.querySelector<HTMLButtonElement>('[data-action="add-argument"]')?.click();
    expect(root.dataset.rawSynchronized).toBe('false');
    expect(formulaBar.input.value).toBe('=SUM({1,2],3)');

    formulaBar.input.value = '=SUM(Table1[[Last, First],3)';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    expect(root.dataset.rawSynchronized).toBe('false');
    expect(formulaBar.input.value).toBe('=SUM(Table1[[Last, First],3)');

    formulaBar.input.value = '=AVERAGE({1,2},3)';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    expect(root.dataset.rawSynchronized).toBe('false');
    expect(formulaBar.input.value).toBe('=AVERAGE({1,2},3)');
  });

  it('grows variadic arguments without dropping an existing trailing blank slot', () => {
    const { formulaBar, palette } = setup();
    palette.open('SUM');
    const root = paletteRoot(palette);
    formulaBar.input.value = '=SUM(1,)';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    root.querySelector<HTMLButtonElement>('[data-action="add-argument"]')?.click();
    expect(formulaBar.input.value).toBe('=SUM(1,,)');
    const second = paletteRoot(palette).querySelector<HTMLInputElement>(
      '[data-argument-index="1"]',
    );
    expect(second).not.toBeNull();
    if (second) {
      second.value = '2';
      second.dispatchEvent(new Event('input', { bubbles: true }));
    }
    expect(formulaBar.input.value).toBe('=SUM(1,2,)');
  });
});
