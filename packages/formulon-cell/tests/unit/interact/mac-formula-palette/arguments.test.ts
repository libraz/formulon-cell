import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { defaultStrings } from '../../../../src/i18n/strings.js';
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

  it('round-trips nested and quoted arguments and exposes the focused range target', () => {
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
    expect(root.dataset.rawSynchronized).toBe('true');
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="0"]')?.value).toBe('');
    const incompleteDone = root.querySelector<HTMLButtonElement>('[data-action="done"]');
    expect(incompleteDone?.disabled).toBe(true);
    expect(incompleteDone?.getAttribute('aria-description')).toBe(
      defaultStrings.fxDialog.macPalette.draftConflict,
    );

    const target = palette.rangeInsertTarget();
    expect(target?.isFormulaEdit()).toBe(true);
    expect(target).not.toBeNull();
    target?.insertRefAtCaret('B3');
    expect(formulaBar.input.value).toBe('=ACOS(B3)');
    expect(anchor).toEqual({ sheet: 0, row: 0, col: 0 });
  });

  it('adopts a caret-only move between nested calls without reopening the draft', () => {
    const { formulaBar, palette } = setup();
    palette.open('SUM');
    const raw = '=SUM(1,MAX(2,3))';
    formulaBar.input.value = raw;
    formulaBar.input.setSelectionRange(raw.indexOf('2') + 1, raw.indexOf('2') + 1);
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    let root = paletteRoot(palette);
    expect(root.querySelector('.fc-mac-formula-palette__args-name')?.textContent).toBe('MAX');
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="0"]')?.value).toBe('2');

    const outerArg = raw.indexOf('1') + 1;
    formulaBar.input.setSelectionRange(outerArg, outerArg);
    palette.open();
    root = paletteRoot(palette);
    expect(root.querySelector('.fc-mac-formula-palette__args-name')?.textContent).toBe('SUM');
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="0"]')?.value).toBe('1');
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="1"]')?.value).toBe(
      'MAX(2,3)',
    );
  });

  it('reconciles passive formula-bar caret moves before field and range writes', () => {
    const { formulaBar, palette } = setup();
    palette.open('SUM');
    const raw = '=SUM(1,MAX(2,3))';
    formulaBar.input.value = raw;
    formulaBar.input.setSelectionRange(raw.indexOf('2') + 1, raw.indexOf('2') + 1);
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    const maxField = paletteRoot(palette).querySelector<HTMLInputElement>(
      '[data-argument-index="0"]',
    );
    expect(maxField?.value).toBe('2');

    const outerCaret = raw.indexOf('1') + 1;
    formulaBar.input.setSelectionRange(outerCaret, outerCaret);
    formulaBar.input.dispatchEvent(new KeyboardEvent('keyup', { key: 'ArrowLeft', bubbles: true }));
    formulaBar.input.dispatchEvent(new Event('select', { bubbles: true }));

    const root = paletteRoot(palette);
    expect(root.querySelector('.fc-mac-formula-palette__args-name')?.textContent).toBe('SUM');
    expect(root.querySelector<HTMLInputElement>('[data-argument-index="0"]')?.value).toBe('1');
    const first = root.querySelector<HTMLInputElement>('[data-argument-index="0"]');
    expect(first).not.toBeNull();
    if (first) {
      first.focus();
      first.value = '5';
      first.dispatchEvent(new Event('input', { bubbles: true }));
    }
    expect(formulaBar.input.value).toBe('=SUM(5,MAX(2,3))');
    palette.rangeInsertTarget()?.insertRefAtCaret('A1');
    expect(formulaBar.input.value).toBe('=SUM(A1,MAX(2,3))');
  });

  it('rejects a stale field event when selection notification is missing', () => {
    const { formulaBar, palette } = setup();
    palette.open('SUM');
    const raw = '=SUM(1,MAX(2,3))';
    formulaBar.input.value = raw;
    formulaBar.input.setSelectionRange(raw.indexOf('2') + 1, raw.indexOf('2') + 1);
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    const staleField = paletteRoot(palette).querySelector<HTMLInputElement>(
      '[data-argument-index="0"]',
    );
    expect(staleField?.value).toBe('2');

    const outerCaret = raw.indexOf('1') + 1;
    formulaBar.input.setSelectionRange(outerCaret, outerCaret);
    if (staleField) {
      staleField.value = '9';
      staleField.dispatchEvent(new Event('input', { bubbles: true }));
    }

    expect(formulaBar.input.value).toBe(raw);
    expect(
      paletteRoot(palette).querySelector('.fc-mac-formula-palette__args-name')?.textContent,
    ).toBe('SUM');
    expect(
      paletteRoot(palette).querySelector<HTMLInputElement>('[data-argument-index="0"]')?.value,
    ).toBe('1');
    palette.rangeInsertTarget()?.insertRefAtCaret('A1');
    expect(formulaBar.input.value).toBe('=SUM(A1,MAX(2,3))');
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
    expect(root.dataset.rawSynchronized).toBe('true');
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

  it('keeps the edited first field as the range insertion target after the call caret moves', () => {
    const { formulaBar, palette } = setup();
    palette.open('SUM');
    formulaBar.input.value = '=SUM(1,2)';
    formulaBar.input.setSelectionRange(6, 6);
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    const first = paletteRoot(palette).querySelector<HTMLInputElement>('[data-argument-index="0"]');
    expect(first).not.toBeNull();
    if (first) {
      first.focus();
      first.value = '5';
      first.dispatchEvent(new Event('input', { bubbles: true }));
    }
    expect(formulaBar.input.value).toBe('=SUM(5,2)');
    palette.rangeInsertTarget()?.insertRefAtCaret('A1');
    expect(formulaBar.input.value).toBe('=SUM(A1,2)');
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

  it('fails closed for crossed, unclosed, or unknown raw formulas', () => {
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

    for (const raw of [
      '=SUM({1,2',
      '=SUM((1+2',
      '=SUM("unterminated',
      "=SUM('unterminated",
      '=SUM(UNKNOWN(1,2',
    ]) {
      formulaBar.input.value = raw;
      formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
      expect(root.dataset.rawSynchronized).toBe('false');
      expect(formulaBar.input.value).toBe(raw);
    }

    formulaBar.input.value = '=MISSING_FUNCTION({1,2},3)';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    expect(root.dataset.rawSynchronized).toBe('false');
    expect(formulaBar.input.value).toBe('=MISSING_FUNCTION({1,2},3)');

    formulaBar.input.value = '=SUM(1,2';
    formulaBar.input.dispatchEvent(new Event('input', { bubbles: true }));
    expect(root.dataset.rawSynchronized).toBe('true');
    const second = root.querySelector<HTMLInputElement>('[data-argument-index="1"]');
    expect(second).not.toBeNull();
    if (second) {
      second.value = '3';
      second.dispatchEvent(new Event('input', { bubbles: true }));
    }
    expect(formulaBar.input.value).toBe('=SUM(1,3)');
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
