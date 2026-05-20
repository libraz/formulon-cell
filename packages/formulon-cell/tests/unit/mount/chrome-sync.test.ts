import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/index.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

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

  it('R1C1 mode formats formula bar source as relative R1C1 references', () => {
    sheet.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 3);
    sheet.workbook.setNumber({ sheet: 0, row: 0, col: 1 }, 4);
    sheet.workbook.setFormula({ sheet: 0, row: 0, col: 2 }, '=A1+B1');
    sheet.workbook.recalc();
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setR1C1(sheet.instance.store, true);
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 2 });
    expect(fxInput.value).toBe('=RC[-2]+RC[-1]');
  });

  it('hides formula bar source when the cell is Hidden and the sheet is protected', () => {
    sheet.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 3);
    sheet.workbook.setNumber({ sheet: 0, row: 0, col: 1 }, 4);
    sheet.workbook.setFormula({ sheet: 0, row: 0, col: 2 }, '=A1+B1');
    sheet.workbook.recalc();
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setCellFormat(
      sheet.instance.store,
      { sheet: 0, row: 0, col: 2 },
      {
        formulaHidden: true,
      },
    );
    mutators.setSheetProtected(sheet.instance.store, 0, true);
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 2 });

    expect(fxInput.value).toBe('');

    mutators.setSheetProtected(sheet.instance.store, 0, false);
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

  it('Alt+ArrowDown opens the Excel-like defined-name list and selecting one jumps there', () => {
    const wb = sheet.workbook as unknown as {
      definedNames: () => Generator<{ name: string; formula: string }>;
    };
    wb.definedNames = function* () {
      yield { name: 'Sales', formula: 'Sheet1!$B$2:$C$4' };
    };

    tag.focus();
    tag.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'ArrowDown', altKey: true, bubbles: true }),
    );

    const menu = document.querySelector<HTMLElement>('.fc-namebox-menu');
    expect(menu).not.toBeNull();
    const item = menu?.querySelector<HTMLButtonElement>('[role="option"]');
    expect(item?.textContent).toBe('Sales');
    item?.click();

    const sel = sheet.instance.store.getState().selection;
    expect(sel.range).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 3, c1: 2 });
    expect(document.querySelector('.fc-namebox-menu')).toBeNull();
  });

  it('Alt+ArrowDown shows an empty defined-name menu when no names exist', () => {
    tag.focus();
    tag.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'ArrowDown', altKey: true, bubbles: true }),
    );

    const menu = document.querySelector<HTMLElement>('.fc-namebox-menu');
    expect(menu).not.toBeNull();
    expect(menu?.querySelector('.fc-namebox-menu__empty')?.textContent).toBeTruthy();
    expect(menu?.querySelector('[role="option"]')).toBeNull();
  });

  it('uses the localized name box label for the defined-name list', async () => {
    sheet.dispose();
    sheet = await mountStubSheet({ locale: 'ja' });
    tag = sheet.host.querySelector('.fc-host__formulabar-tag') as HTMLInputElement;

    tag.focus();
    tag.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'ArrowDown', altKey: true, bubbles: true }),
    );

    const menu = document.querySelector<HTMLElement>('.fc-namebox-menu');
    expect(menu?.getAttribute('aria-label')).toBe('名前ボックス');
  });

  it('Ctrl-click in the name box list adds a defined range to the multi-selection', () => {
    const wb = sheet.workbook as unknown as {
      definedNames: () => Generator<{ name: string; formula: string }>;
    };
    wb.definedNames = function* () {
      yield { name: 'North', formula: 'Sheet1!$B$2:$B$4' };
      yield { name: 'South', formula: 'Sheet1!$D$2:$D$4' };
    };

    tag.focus();
    tag.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'ArrowDown', altKey: true, bubbles: true }),
    );
    const items = document.querySelectorAll<HTMLButtonElement>('.fc-namebox-menu [role="option"]');
    expect(items.length).toBe(2);
    items[0]?.dispatchEvent(new MouseEvent('click', { bubbles: true, ctrlKey: true }));
    items[1]?.dispatchEvent(new MouseEvent('click', { bubbles: true, ctrlKey: true }));

    const sel = sheet.instance.store.getState().selection;
    expect(sel.range).toEqual({ sheet: 0, r0: 1, c0: 3, r1: 3, c1: 3 });
    expect(sel.extraRanges).toEqual([
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 1 },
    ]);
    expect(document.querySelector('.fc-namebox-menu')).not.toBeNull();
  });

  it('defined-name list supports Arrow/Home/End navigation and Enter activation', () => {
    const wb = sheet.workbook as unknown as {
      definedNames: () => Generator<{ name: string; formula: string }>;
    };
    wb.definedNames = function* () {
      yield { name: 'North', formula: 'Sheet1!$B$2:$B$4' };
      yield { name: 'South', formula: 'Sheet1!$D$2:$D$4' };
      yield { name: 'West', formula: 'Sheet1!$F$2:$F$4' };
    };

    tag.focus();
    tag.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'ArrowDown', altKey: true, bubbles: true }),
    );
    const items = Array.from(
      document.querySelectorAll<HTMLButtonElement>('.fc-namebox-menu [role="option"]'),
    );
    expect(items.map((item) => item.textContent)).toEqual(['North', 'South', 'West']);
    expect(document.activeElement).toBe(items[0]);

    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true }));
    expect(document.activeElement).toBe(items[2]);
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    expect(document.activeElement).toBe(items[0]);
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', bubbles: true }));
    expect(document.activeElement).toBe(items[0]);
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowUp', bubbles: true }));
    expect(document.activeElement).toBe(items[2]);
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));

    const sel = sheet.instance.store.getState().selection;
    expect(sel.range).toEqual({ sheet: 0, r0: 1, c0: 5, r1: 3, c1: 5 });
    expect(document.querySelector('.fc-namebox-menu')).toBeNull();
  });

  it('Enter on an Excel-style name defines the current selection as a workbook name', () => {
    const calls: { name: string; formula: string }[] = [];
    Object.defineProperty(sheet.workbook, 'capabilities', {
      configurable: true,
      value: { ...sheet.workbook.capabilities, definedNameMutate: true },
    });
    const wb = sheet.workbook as unknown as {
      setDefinedNameEntry(name: string, formula: string): boolean;
    };
    wb.setDefinedNameEntry = (name, formula) => {
      calls.push({ name, formula });
      return true;
    };

    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 });
    tag.focus();
    tag.value = 'Sales_Total';
    tag.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));

    expect(calls).toEqual([{ name: 'Sales_Total', formula: '=Sheet1!$A$1:$B$3' }]);
  });

  it('quotes sheet names when defining a name from the name box', () => {
    const calls: { name: string; formula: string }[] = [];
    Object.defineProperty(sheet.workbook, 'capabilities', {
      configurable: true,
      value: { ...sheet.workbook.capabilities, definedNameMutate: true },
    });
    const wb = sheet.workbook as unknown as {
      setDefinedNameEntry(name: string, formula: string): boolean;
      sheetName(idx: number): string;
    };
    wb.sheetName = () => "FY 26's Plan";
    wb.setDefinedNameEntry = (name, formula) => {
      calls.push({ name, formula });
      return true;
    };

    tag.focus();
    tag.value = 'PlanCell';
    tag.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));

    expect(calls).toEqual([{ name: 'PlanCell', formula: "='FY 26''s Plan'!$A$1" }]);
  });

  it('Enter does not define invalid names or cell references from the name box', () => {
    const calls: { name: string; formula: string }[] = [];
    Object.defineProperty(sheet.workbook, 'capabilities', {
      configurable: true,
      value: { ...sheet.workbook.capabilities, definedNameMutate: true },
    });
    const wb = sheet.workbook as unknown as {
      setDefinedNameEntry(name: string, formula: string): boolean;
    };
    wb.setDefinedNameEntry = (name, formula) => {
      calls.push({ name, formula });
      return true;
    };

    for (const value of ['not a name', 'A1', 'C', 'R']) {
      tag.focus();
      tag.value = value;
      tag.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    }

    expect(calls).toEqual([]);
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

  it('keeps name-box dropdown items on the shared host button helper', () => {
    const source = readFileSync(join(root, 'src/mount/chrome-sync.ts'), 'utf8');
    expect(source).toContain('createHostButton({');
    expect(source).not.toContain("document.createElement('button')");
  });
});
