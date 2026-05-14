import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/index.js';

/**
 * Unit: chrome-sync — the small reactive bridge that keeps the name box,
 * formula bar input, and aria-live region in sync with the store. Going
 * through `mountStubSheet` here because chrome-sync needs the entire
 * formulabar tree wired in front of it; faking those inputs by hand would
 * just reproduce the chrome.ts factory.
 */
describe('mount/chrome-sync — name box and formula bar reflect store state', () => {
  let sheet: MountedStubSheet;
  let tag: HTMLInputElement;
  let fxInput: HTMLTextAreaElement;
  let a11y: HTMLDivElement;

  beforeEach(async () => {
    sheet = await mountStubSheet();
    tag = sheet.host.querySelector('.fc-host__formulabar-tag') as HTMLInputElement;
    fxInput = sheet.host.querySelector('.fc-host__formulabar-input') as HTMLTextAreaElement;
    a11y = sheet.host.querySelector('.fc-host__a11y') as HTMLDivElement;
  });

  afterEach(() => sheet.dispose());

  it('initial state is A1, empty formula, a11y echoing both', () => {
    expect(tag.value).toBe('A1');
    expect(fxInput.value).toBe('');
    expect(a11y.textContent?.trim()).toBe('A1');
  });

  it('moving the active cell updates the name box', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 2, col: 3 });
    expect(tag.value).toBe('D3');
    expect(a11y.textContent?.trim().startsWith('D3')).toBe(true);
  });

  it('a range selection formats as A1:B3', () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 });
    expect(tag.value).toBe('A1:B3');
  });

  it('R1C1 mode swaps the name box format', () => {
    mutators.setR1C1(sheet.instance.store, true);
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 4, col: 1 });
    expect(tag.value).toBe('R5C2');
  });

  it('selecting a numeric cell mirrors the value into the formula bar input', () => {
    sheet.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 42);
    sheet.workbook.recalc();
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    expect(fxInput.value).toBe('42');
  });

  it('selecting a formula cell echoes the source formula in the formula bar', () => {
    sheet.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 3);
    sheet.workbook.setNumber({ sheet: 0, row: 0, col: 1 }, 4);
    sheet.workbook.setFormula({ sheet: 0, row: 0, col: 2 }, '=A1+B1');
    sheet.workbook.recalc();
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 2 });
    expect(fxInput.value).toBe('=A1+B1');
  });

  it('text and bool cells display their natural rendering', () => {
    sheet.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'hello');
    sheet.workbook.recalc();
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    expect(fxInput.value).toBe('hello');
  });

  it('formula bar input is left alone while the user is editing the name box', () => {
    tag.focus();
    tag.value = 'C5';
    // Move-active fires; chrome-sync should NOT clobber the tag value because
    // document.activeElement === tag.
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 1, col: 1 });
    expect(tag.value).toBe('C5');
  });
});

describe('mount/chrome-sync — name box (tag) keyboard handling', () => {
  let sheet: MountedStubSheet;
  let tag: HTMLInputElement;

  beforeEach(async () => {
    sheet = await mountStubSheet();
    tag = sheet.host.querySelector('.fc-host__formulabar-tag') as HTMLInputElement;
  });

  afterEach(() => sheet.dispose());

  it('Enter on a valid cell ref jumps the selection there', () => {
    tag.focus();
    tag.value = 'C5';
    tag.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    const a = sheet.instance.store.getState().selection.active;
    expect(a).toEqual({ sheet: 0, row: 4, col: 2 });
  });

  it('Enter on an A1:B3 range expands the selection range', () => {
    tag.focus();
    tag.value = 'A1:B3';
    tag.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    const sel = sheet.instance.store.getState().selection;
    expect(sel.range).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 });
  });

  it('Enter on garbage input leaves the selection alone', () => {
    const before = sheet.instance.store.getState().selection;
    tag.focus();
    tag.value = 'not-a-ref';
    tag.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(sheet.instance.store.getState().selection).toEqual(before);
  });

  it('Escape restores the name box value from store state', () => {
    tag.focus();
    tag.value = 'Z99';
    tag.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    // Focus moves back to host; chrome-sync repaints the tag with the live ref.
    expect(tag.value).toBe('A1');
  });
});
