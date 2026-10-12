import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { attachEvaluateFormulaDialog } from '../../../src/interact/evaluate-formula-dialog.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const overlay = (): HTMLElement | null => document.querySelector<HTMLElement>('.fc-evaldlg');
const boxes = (): HTMLElement[] =>
  Array.from(document.querySelectorAll<HTMLElement>('.fc-evaldlg__box'));
const evalButton = (): HTMLButtonElement | null =>
  Array.from(document.querySelectorAll<HTMLButtonElement>('.fc-evaldlg__btn')).find(
    (button) => !button.classList.contains('fc-evaldlg__btn--primary'),
  ) ?? null;

describe('attachEvaluateFormulaDialog', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    while (document.body.firstChild) document.body.removeChild(document.body.firstChild);
  });

  it('shows the active formula, step evaluation, and current result', () => {
    const store = createSpreadsheetStore();
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(store, addr);
    mutators.setCell(store, addr, { kind: 'number', value: 3 }, '=A2+B2');
    const wb = {
      cellFormula: vi.fn(() => '=A2+B2'),
      getValue: vi.fn((a) => {
        if (a.row === 1 && a.col === 0) return { kind: 'number', value: 1 };
        if (a.row === 1 && a.col === 1) return { kind: 'number', value: 2 };
        return { kind: 'number', value: 3 };
      }),
    } as unknown as WorkbookHandle;

    const handle = attachEvaluateFormulaDialog({ host, store, getWb: () => wb, strings: en });
    handle.open();

    expect(overlay()?.hidden).toBe(false);
    expect(boxes()[0]?.textContent).toBe('=A2+B2');
    expect(boxes()[1]?.textContent).toBe('=A2+B2');
    expect(boxes()[2]?.textContent).toBe('3');
    expect(evalButton()?.disabled).toBe(false);

    evalButton()?.click();
    expect(boxes()[1]?.textContent).toBe('=1+B2');
    expect(evalButton()?.disabled).toBe(false);

    evalButton()?.click();
    expect(boxes()[1]?.textContent).toBe('=1+2');
    expect(evalButton()?.disabled).toBe(true);
    expect(evalButton()?.dataset.disabledReason).toBe(
      'The formula has already been fully evaluated.',
    );
    handle.detach();
  });

  it('shows an empty-state when the active cell has no formula', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    const wb = {
      cellFormula: vi.fn(() => null),
      getValue: vi.fn(() => ({ kind: 'blank' })),
    } as unknown as WorkbookHandle;

    const handle = attachEvaluateFormulaDialog({ host, store, getWb: () => wb, strings: en });
    handle.open();

    expect(boxes()[0]?.textContent).toBe('The active cell does not contain a formula.');
    expect(boxes()[1]?.textContent).toBe('');
    expect(boxes()[2]?.textContent).toBe('');
    expect(evalButton()?.disabled).toBe(true);
    expect(evalButton()?.getAttribute('aria-description')).toBe(
      'Select a cell that contains a formula.',
    );
    handle.detach();
  });

  it('localizes formula displays after canonical reference substitution', () => {
    const store = createSpreadsheetStore();
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(store, addr);
    mutators.setCell(store, addr, { kind: 'number', value: 3.5 }, '=SUM(A2,B2)');
    const wb = {
      cellFormula: vi.fn(() => '=SUM(A2,B2)'),
      getValue: vi.fn((a) => {
        if (a.row === 1 && a.col === 0) return { kind: 'number', value: 1.5 };
        if (a.row === 1 && a.col === 1) return { kind: 'number', value: 2 };
        return { kind: 'number', value: 3.5 };
      }),
      getFormulaForDisplay: vi.fn((formula: string) =>
        formula.replace('SUM', 'SUMME').replaceAll(',', ';').replaceAll('.', ','),
      ),
      getDisplayText: vi.fn(() => '3,5'),
    } as unknown as WorkbookHandle;

    const handle = attachEvaluateFormulaDialog({
      host,
      store,
      getWb: () => wb,
      getLocale: () => 'de-DE',
      strings: en,
    });
    handle.open();

    expect(boxes()[0]?.textContent).toBe('=SUMME(A2;B2)');
    expect(boxes()[1]?.textContent).toBe('=SUMME(A2;B2)');
    expect(boxes()[2]?.textContent).toBe('3,5');
    evalButton()?.click();
    expect(boxes()[1]?.textContent).toBe('=SUMME(1,5;B2)');
    expect(wb.cellFormula(addr)).toBe('=SUM(A2,B2)');
    handle.detach();
  });

  it('reads qualified references from their own sheets with canonical literals', () => {
    const store = createSpreadsheetStore();
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(store, addr);
    mutators.setCell(store, addr, { kind: 'number', value: 1235 }, "='Data Sheet'!A1+Sheet3!A1");
    const getValue = vi.fn((a: { sheet: number; row: number; col: number }) => {
      if (a.sheet === 1) return { kind: 'number', value: 1234 };
      if (a.sheet === 2) return { kind: 'bool', value: true };
      return { kind: 'number', value: 1235 };
    });
    const wb = {
      cellFormula: vi.fn(() => "='Data Sheet'!A1+Sheet3!A1"),
      getValue,
      sheetCount: 3,
      sheetName: (sheet: number) => ['Sheet1', 'Data Sheet', 'Sheet3'][sheet] ?? '',
      getFormulaForDisplay: (formula: string) => formula,
      getDisplayText: () => '1235',
    } as unknown as WorkbookHandle;

    const handle = attachEvaluateFormulaDialog({ host, store, getWb: () => wb, strings: en });
    handle.open();
    evalButton()?.click();
    expect(boxes()[1]?.textContent).toBe('=1234+Sheet3!A1');
    evalButton()?.click();
    expect(boxes()[1]?.textContent).toBe('=1234+TRUE');
    expect(getValue).toHaveBeenCalledWith({ sheet: 1, row: 0, col: 0 });
    expect(getValue).toHaveBeenCalledWith({ sheet: 2, row: 0, col: 0 });
    handle.detach();
  });

  it('substitutes every occurrence of a stepped repeated reference', () => {
    const store = createSpreadsheetStore();
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(store, addr);
    mutators.setCell(store, addr, { kind: 'number', value: 2468 }, '=A1+A1+B1');
    const wb = {
      cellFormula: vi.fn(() => '=A1+A1+B1'),
      getValue: vi.fn((a) =>
        a.col === 0 ? { kind: 'number', value: 1234 } : { kind: 'number', value: 5 },
      ),
      getFormulaForDisplay: (formula: string) => formula,
      getDisplayText: () => '2468',
    } as unknown as WorkbookHandle;

    const handle = attachEvaluateFormulaDialog({ host, store, getWb: () => wb, strings: en });
    handle.open();
    evalButton()?.click();
    expect(boxes()[1]?.textContent).toBe('=1234+1234+B1');
    expect(evalButton()?.disabled).toBe(false);
    evalButton()?.click();
    expect(boxes()[1]?.textContent).toBe('=1234+1234+5');
    expect(evalButton()?.disabled).toBe(true);
    handle.detach();
  });

  it('substitutes repeated qualified references from a non-active sheet together', () => {
    const store = createSpreadsheetStore();
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(store, addr);
    mutators.setCell(store, addr, { kind: 'number', value: 20 }, '=Data!A1+Data!A1');
    const getValue = vi.fn((a: { sheet: number; row: number; col: number }) =>
      a.sheet === 1 ? { kind: 'number', value: 10 } : { kind: 'number', value: 20 },
    );
    const wb = {
      cellFormula: vi.fn(() => '=Data!A1+Data!A1'),
      getValue,
      sheetCount: 2,
      sheetName: (sheet: number) => ['Sheet1', 'Data'][sheet] ?? '',
      getFormulaForDisplay: (formula: string) => formula,
      getDisplayText: () => '20',
    } as unknown as WorkbookHandle;

    const handle = attachEvaluateFormulaDialog({ host, store, getWb: () => wb, strings: en });
    handle.open();
    evalButton()?.click();
    expect(boxes()[1]?.textContent).toBe('=10+10');
    expect(evalButton()?.disabled).toBe(true);
    expect(getValue).toHaveBeenCalledTimes(1);
    handle.detach();
  });

  it('disables Evaluate with a reason when the formula has no references', () => {
    const store = createSpreadsheetStore();
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setActive(store, addr);
    mutators.setCell(store, addr, { kind: 'number', value: 3 }, '=1+2');
    const wb = {
      cellFormula: vi.fn(() => '=1+2'),
      getValue: vi.fn(() => ({ kind: 'number', value: 3 })),
    } as unknown as WorkbookHandle;

    const handle = attachEvaluateFormulaDialog({ host, store, getWb: () => wb, strings: en });
    handle.open();

    expect(evalButton()?.disabled).toBe(true);
    expect(evalButton()?.dataset.disabledReason).toBe(
      'The formula has no cell references to evaluate.',
    );
    handle.detach();
  });

  it('keeps Evaluate Formula on compact desktop dialog geometry', () => {
    const css = readFileSync(
      join(root, 'src/styles/core/app/dialogs/evaluate-formula.css'),
      'utf8',
    );

    expect(css).toMatch(/\.fc-evaldlg__panel\s*\{[\s\S]*?border-radius: 2px;/);
    expect(css).toMatch(
      /\.fc-evaldlg__header\s*\{[\s\S]*?min-height: 34px;[\s\S]*?padding: 8px 14px;/,
    );
    expect(css).toMatch(/\.fc-evaldlg__body\s*\{[\s\S]*?gap: 6px;[\s\S]*?padding: 12px 14px;/);
    expect(css).toMatch(
      /\.fc-evaldlg__box\s*\{[\s\S]*?padding: 6px 8px;[\s\S]*?border-radius: 2px;[\s\S]*?background: var\(--fc-bg, Canvas\);/,
    );
    expect(css).toMatch(
      /\.fc-evaldlg__footer\s*\{[\s\S]*?gap: 6px;[\s\S]*?padding: 10px 14px 12px;/,
    );
    expect(css).toMatch(/\.fc-evaldlg__btn\s*\{[\s\S]*?height: 28px;[\s\S]*?border-radius: 2px;/);
    expect(css).not.toContain('box-shadow: var(--fc-shadow-16)');
  });
});
