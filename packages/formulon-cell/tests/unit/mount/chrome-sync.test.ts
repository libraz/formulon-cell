import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { SpreadsheetEmitter } from '../../../src/events.js';
import { defaultStrings } from '../../../src/i18n/strings.js';
import { attachChromeSync } from '../../../src/mount/chrome-sync.js';
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
  let a11yLive: HTMLDivElement;
  let grid: HTMLDivElement;

  beforeEach(async () => {
    sheet = await mountStubSheet();
    tag = sheet.host.querySelector('.fc-host__formulabar-tag') as HTMLInputElement;
    fxInput = sheet.host.querySelector('.fc-host__formulabar-input') as HTMLTextAreaElement;
    a11y = sheet.host.querySelector('.fc-host__a11y') as HTMLDivElement;
    a11yLive = sheet.host.querySelector('.fc-host__a11y-live') as HTMLDivElement;
    grid = sheet.host.querySelector('.fc-host__grid') as HTMLDivElement;
  });

  afterEach(() => sheet.dispose());

  it('initial state is A1, empty formula, a11y echoing both', () => {
    expect(tag.value).toBe('A1');
    expect(fxInput.value).toBe('');
    const activeCell = a11y.querySelector<HTMLElement>('[aria-selected="true"]');
    expect(activeCell?.textContent?.trim()).toBe('A1');
    expect(grid.getAttribute('aria-activedescendant')).toBe(activeCell?.id);
    expect(grid.getAttribute('aria-rowcount')).toBe('1048576');
    expect(grid.getAttribute('aria-colcount')).toBe('16384');
    expect(grid.getAttribute('aria-multiselectable')).toBe('true');
    expect(activeCell?.getAttribute('role')).toBe('gridcell');
    expect(activeCell?.getAttribute('aria-rowindex')).toBe('1');
    expect(activeCell?.getAttribute('aria-colindex')).toBe('1');
    expect(activeCell?.getAttribute('aria-label')).toBe('A1');
    expect(a11yLive.textContent?.trim()).toBe('A1');
  });

  it('mirrors cells as grid > rowgroup > row > gridcell', () => {
    expect(a11y.getAttribute('role')).toBe('rowgroup');
    expect(a11y.parentElement).toBe(grid);
    for (const rowMirror of a11y.querySelectorAll<HTMLElement>(':scope > *')) {
      expect(rowMirror.getAttribute('role')).toBe('row');
      expect(rowMirror.getAttribute('aria-rowindex')).toBeTruthy();
      for (const cell of rowMirror.querySelectorAll<HTMLElement>(':scope > *')) {
        expect(cell.getAttribute('role')).toBe('gridcell');
      }
    }
    // Every gridcell is a real descendant of the grid, so nothing is owned
    // across the tree — `aria-owns` would re-parent them under the grid.
    expect(grid.getAttribute('aria-owns')).toBeNull();
    expect(a11y.querySelectorAll('[role="gridcell"]').length).toBeGreaterThan(0);
  });

  it('moving the active cell updates the name box', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 2, col: 3 });
    expect(tag.value).toBe('D3');
    expect(a11yLive.textContent?.trim().startsWith('D3')).toBe(true);
    const activeCell = a11y.querySelector<HTMLElement>('[aria-selected="true"]');
    expect(activeCell?.getAttribute('aria-rowindex')).toBe('3');
    expect(activeCell?.getAttribute('aria-colindex')).toBe('4');
    expect(activeCell?.parentElement?.getAttribute('aria-rowindex')).toBe('3');
  });

  it('mirrors a bounded visible viewport for screen readers', () => {
    sheet.workbook.setText({ sheet: 0, row: 1, col: 1 }, 'visible');
    sheet.workbook.recalc();
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    sheet.instance.store.setState((s) => ({
      ...s,
      viewport: { ...s.viewport, rowStart: 0, rowCount: 100, colStart: 0, colCount: 100 },
    }));

    const rows = a11y.querySelectorAll<HTMLElement>('[role="row"]');
    const mirrors = a11y.querySelectorAll<HTMLElement>('[role="gridcell"]');
    expect(rows.length).toBe(20);
    expect(mirrors.length).toBe(200);
    const b2 = a11y.querySelector<HTMLElement>('[aria-label="B2 visible"]');
    expect(b2?.getAttribute('aria-rowindex')).toBe('2');
    expect(b2?.getAttribute('aria-colindex')).toBe('2');
    expect(b2?.getAttribute('aria-selected')).toBe('false');
    expect(b2?.parentElement?.getAttribute('aria-rowindex')).toBe('2');
  });

  it('keeps the active cell in its own row when it is outside the mirrored viewport', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 200, col: 40 });
    sheet.instance.store.setState((s) => ({
      ...s,
      viewport: { ...s.viewport, rowStart: 0, rowCount: 20, colStart: 0, colCount: 10 },
    }));

    const activeCell = a11y.querySelector<HTMLElement>('[aria-selected="true"]');
    expect(grid.getAttribute('aria-activedescendant')).toBe(activeCell?.id);
    const activeRow = activeCell?.parentElement;
    expect(activeRow?.getAttribute('role')).toBe('row');
    expect(activeRow?.getAttribute('aria-rowindex')).toBe('201');
    expect(activeRow?.parentElement).toBe(a11y);
    expect(a11y.querySelectorAll('[role="row"]').length).toBe(21);
  });

  it('a range selection formats as A1:B3', () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 });
    expect(tag.value).toBe('A1:B3');
  });

  it('shows a merged logical cell by its anchor while retaining the complete selection', () => {
    const store = sheet.instance.store;
    const range = { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 3 };
    mutators.mergeRange(store, range);
    mutators.setActive(store, { sheet: 0, row: 2, col: 3 });
    expect(tag.value).toBe('B2');
    expect(store.getState().selection.range).toEqual(range);
    mutators.setR1C1(store, true);
    expect(tag.value).toBe('R2C2');
    mutators.setR1C1(store, false);
    mutators.setRange(store, { ...range, r1: 3 });
    expect(tag.value).toBe('B2:D4');
  });

  it('a selected chart or picture names itself and empties the formula bar', () => {
    const store = sheet.instance.store;
    sheet.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 42);
    sheet.workbook.recalc();
    mutators.replaceCells(store, sheet.workbook.cells(0));
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    expect(fxInput.value).toBe('42');
    mutators.setSelectedObject(store, { id: 'chart-1', name: 'Chart 1' });
    expect(tag.value).toBe('Chart 1');
    expect(fxInput.value).toBe('');
    mutators.setSelectedObject(store, null);
    expect(tag.value).toBe('A1');
    expect(fxInput.value).toBe('42');
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
      definedNames: () => Generator<{ name: string; formula: string; localSheetId: number }>;
    };
    wb.definedNames = function* () {
      yield { name: 'Sales', formula: 'Sheet1!$B$2:$C$4', localSheetId: -1 };
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

  it('Enter on a defined name prefers the active sheet scoped name', () => {
    const wb = sheet.workbook as unknown as {
      definedNames: () => Generator<{ name: string; formula: string; localSheetId: number }>;
    };
    wb.definedNames = function* () {
      yield { name: 'Sales', formula: 'Sheet1!$A$1:$A$3', localSheetId: -1 };
      yield { name: 'Sales', formula: 'Sheet1!$B$2:$B$4', localSheetId: 0 };
    };

    tag.focus();
    tag.value = 'Sales';
    tag.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));

    const sel = sheet.instance.store.getState().selection;
    expect(sel.range).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 3, c1: 1 });
  });

  it('defined-name list shadows workbook names with active sheet scoped names', () => {
    const wb = sheet.workbook as unknown as {
      definedNames: () => Generator<{ name: string; formula: string; localSheetId: number }>;
    };
    wb.definedNames = function* () {
      yield { name: 'Sales', formula: 'Sheet1!$A$1:$A$3', localSheetId: -1 };
      yield { name: 'Sales', formula: 'Sheet1!$B$2:$B$4', localSheetId: 0 };
      yield { name: 'OtherSheetOnly', formula: 'Sheet2!$C$1:$C$3', localSheetId: 1 };
    };

    tag.focus();
    tag.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'ArrowDown', altKey: true, bubbles: true }),
    );

    const items = Array.from(
      document.querySelectorAll<HTMLButtonElement>('.fc-namebox-menu [role="option"]'),
    );
    expect(items.map((item) => item.textContent)).toEqual(['Sales']);
    items[0]?.click();

    const sel = sheet.instance.store.getState().selection;
    expect(sel.range).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 3, c1: 1 });
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
      definedNames: () => Generator<{ name: string; formula: string; localSheetId: number }>;
    };
    wb.definedNames = function* () {
      yield { name: 'North', formula: 'Sheet1!$B$2:$B$4', localSheetId: -1 };
      yield { name: 'South', formula: 'Sheet1!$D$2:$D$4', localSheetId: -1 };
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
      definedNames: () => Generator<{ name: string; formula: string; localSheetId: number }>;
    };
    wb.definedNames = function* () {
      yield { name: 'North', formula: 'Sheet1!$B$2:$B$4', localSheetId: -1 };
      yield { name: 'South', formula: 'Sheet1!$D$2:$D$4', localSheetId: -1 };
      yield { name: 'West', formula: 'Sheet1!$F$2:$F$4', localSheetId: -1 };
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

  const stubNames = (entries: { name: string; formula: string; localSheetId: number }[]): void => {
    const wb = sheet.workbook as unknown as { definedNames: () => Iterable<unknown> };
    wb.definedNames = function* () {
      yield* entries;
    };
  };
  /** A second chrome-sync bound to its own name box, so the mounted one stays out of the way. */
  const standaloneChrome = (
    extra: Partial<Parameters<typeof attachChromeSync>[0]>,
  ): { controller: ReturnType<typeof attachChromeSync>; ownTag: HTMLInputElement } => {
    const ownTag = document.createElement('input');
    document.body.appendChild(ownTag);
    const controller = attachChromeSync({
      a11y: document.createElement('div'),
      a11yLive: document.createElement('div'),
      emitter: new SpreadsheetEmitter(),
      fxInput: document.createElement('textarea'),
      getFormulaEditing: () => false,
      getSheetTabs: () => null,
      getStrings: () => defaultStrings,
      getWb: () => sheet.workbook,
      grid: document.createElement('div'),
      host: sheet.host,
      invalidate: () => {},
      store: sheet.instance.store,
      tag: ownTag,
      ...extra,
    });
    return { controller, ownTag };
  };
  const pressEnter = (value: string, target: HTMLInputElement = tag): void => {
    target.focus();
    target.value = value;
    target.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
  };

  it('Enter on a cross-sheet name switches sheet and selects its range', () => {
    sheet.workbook.addSheet('My Sheet');
    stubNames([{ name: 'Remote', formula: "'My Sheet'!$A$1:$B$3", localSheetId: -1 }]);
    pressEnter('remote');
    const state = sheet.instance.store.getState();
    expect(state.data.sheetIndex).toBe(1);
    expect(state.selection.range).toEqual({ sheet: 1, r0: 0, c0: 0, r1: 2, c1: 1 });
  });

  it('Enter on a sheet-qualified address jumps to that sheet', () => {
    sheet.workbook.addSheet('Sheet2');
    pressEnter('sheet2!B4');
    const state = sheet.instance.store.getState();
    expect(state.data.sheetIndex).toBe(1);
    expect(state.selection.active).toEqual({ sheet: 1, row: 3, col: 1 });
  });

  it('Enter on a multi-area entry fills extraRanges and activates the first cell', () => {
    pressEnter('A1,C3:D4');
    const sel = sheet.instance.store.getState().selection;
    expect(sel.active).toEqual({ sheet: 0, row: 0, col: 0 });
    expect(sel.range).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    expect(sel.extraRanges).toEqual([{ sheet: 0, r0: 2, c0: 2, r1: 3, c1: 3 }]);
  });

  it('shows the name when the selection exactly matches a named range', () => {
    stubNames([
      { name: 'Zeta', formula: 'Sheet1!$A$1:$B$2', localSheetId: -1 },
      { name: 'alpha', formula: 'Sheet1!$A$1:$B$2', localSheetId: -1 },
      { name: 'Scoped', formula: 'Sheet1!$A$1:$B$2', localSheetId: 0 },
    ]);
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    expect(tag.value).toBe('Scoped');
    stubNames([
      { name: 'Zeta', formula: 'Sheet1!$A$1:$B$2', localSheetId: -1 },
      { name: 'alpha', formula: 'Sheet1!$A$1:$B$2', localSheetId: -1 },
    ]);
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 });
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    expect(tag.value).toBe('alpha');
  });

  it('the dropdown lists only range names, sorted case-insensitively', () => {
    stubNames([
      { name: 'beta', formula: 'Sheet1!$A$1', localSheetId: -1 },
      { name: 'Tax', formula: '0.08', localSheetId: -1 },
      { name: 'Calc', formula: 'SUM(Sheet1!$A$1:$A$3)', localSheetId: -1 },
      { name: 'Alpha', formula: 'Sheet1!$B$1', localSheetId: -1 },
    ]);
    tag.focus();
    tag.dispatchEvent(new KeyboardEvent('keydown', { key: 'F4', bubbles: true }));
    const items = [...document.querySelectorAll('.fc-namebox-menu [role="option"]')];
    expect(items.map((el) => el.textContent)).toEqual(['Alpha', 'beta']);
  });

  it('Enter with an unusable entry alerts, restores the box and keeps focus', () => {
    const alerts: string[] = [];
    const { controller, ownTag } = standaloneChrome({
      getStrings: () =>
        ({ a11y: { nameBox: 'Name box', nameBoxInvalidReference: 'bad ref' } }) as never,
      onAlert: (o) => alerts.push(o.message),
    });
    pressEnter('Sheet9!A1', ownTag);
    expect(alerts).toEqual(['bad ref']);
    expect(ownTag.value).toBe('A1');
    expect(document.activeElement).toBe(ownTag);
    controller.detach();
    ownTag.remove();
  });

  it('defining a name from the box is undoable', () => {
    Object.defineProperty(sheet.workbook, 'capabilities', {
      configurable: true,
      value: { ...sheet.workbook.capabilities, definedNameMutate: true },
    });
    const store = new Map<string, string>();
    const wb = sheet.workbook as unknown as {
      setDefinedNameEntry(name: string, formula: string): boolean;
      definedNames: () => Iterable<unknown>;
    };
    wb.setDefinedNameEntry = (name, formula) => {
      if (formula) store.set(name, formula);
      else store.delete(name);
      return true;
    };
    wb.definedNames = function* () {
      for (const [name, formula] of store) yield { name, formula, localSheetId: -1 };
    };
    const history = new History();
    const { controller, ownTag } = standaloneChrome({ getHistory: () => history });
    pressEnter('Fresh', ownTag);
    expect(store.get('Fresh')).toBe('=Sheet1!$A$1');
    history.undo();
    expect(store.has('Fresh')).toBe(false);
    controller.detach();
    ownTag.remove();
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
