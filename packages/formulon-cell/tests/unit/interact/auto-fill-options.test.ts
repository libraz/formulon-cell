import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings/en.js';
import { attachAutoFillOptions } from '../../../src/interact/auto-fill-options.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const seedAndMirror = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  cells: Array<{ row: number; col: number; value: number | string }>,
): void => {
  store.setState((s) => {
    const map = new Map(s.data.cells);
    for (const c of cells) {
      const key = `${0}:${c.row}:${c.col}`;
      if (typeof c.value === 'number') {
        wb.setNumber({ sheet: 0, row: c.row, col: c.col }, c.value);
        map.set(key, { value: { kind: 'number', value: c.value }, formula: null });
      } else {
        wb.setText({ sheet: 0, row: c.row, col: c.col }, c.value);
        map.set(key, { value: { kind: 'text', value: c.value }, formula: null });
      }
    }
    return { ...s, data: { ...s.data, cells: map } };
  });
  wb.recalc();
};

const num = (wb: WorkbookHandle, row: number, col: number): number => {
  const value = wb.getValue({ sheet: 0, row, col });
  return value.kind === 'number' ? value.value : Number.NaN;
};

const dateSerial = (year: number, month: number, day: number): number =>
  Date.UTC(year, month - 1, day) / 86_400_000 + 25569;

const visibleLabels = (): Array<string | null> =>
  Array.from(document.querySelectorAll<HTMLButtonElement>('.fc-autofill-options__item'))
    .filter((item) => item.style.display !== 'none')
    .map((item) => item.textContent);

describe('auto fill options', () => {
  let host: HTMLDivElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let detach: () => void;

  beforeEach(async () => {
    document.body.innerHTML = '';
    host = document.createElement('div');
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await newWb();
    detach = () => {};
  });

  afterEach(() => {
    detach();
    document.body.innerHTML = '';
  });

  it('opens the Excel-style smart button and can switch a fill series to Copy Cells', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 1, col: 0, value: 2 },
    ]);
    const handle = attachAutoFillOptions({
      host,
      store,
      wb,
      strings: en,
      onAfterCommit: () => {
        mutators.replaceCells(store, wb.cells(0));
      },
    });
    detach = () => handle.detach();

    host.dispatchEvent(
      new CustomEvent('fc:autofilloptions', {
        detail: {
          src: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 },
          dest: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 },
          mode: 'series',
          clientX: 120,
          clientY: 90,
        },
      }),
    );

    const button = document.querySelector<HTMLButtonElement>('.fc-autofill-options__button');
    expect(button?.style.display).toBe('block');
    button?.click();
    expect(visibleLabels()).toEqual([
      'Copy Cells',
      'Fill Series',
      'Fill Formatting Only',
      'Fill Without Formatting',
    ]);
    const copy = Array.from(
      document.querySelectorAll<HTMLButtonElement>('.fc-autofill-options__item'),
    ).find((item) => item.textContent === 'Copy Cells');
    copy?.click();
    wb.recalc();

    expect(num(wb, 2, 0)).toBe(1);
    expect(num(wb, 3, 0)).toBe(2);
    expect(store.getState().selection.range).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 });
  });

  it('shows date fill choices for date-formatted cells and can fill months', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: dateSerial(2024, 1, 31) }]);
    store.setState((s) => {
      const formats = new Map(s.format.formats);
      formats.set('0:0:0', { numFmt: { kind: 'date', pattern: 'yyyy-mm-dd' } });
      return { ...s, format: { ...s.format, formats } };
    });
    const handle = attachAutoFillOptions({
      host,
      store,
      wb,
      strings: en,
      onAfterCommit: () => {
        mutators.replaceCells(store, wb.cells(0));
      },
    });
    detach = () => handle.detach();

    host.dispatchEvent(
      new CustomEvent('fc:autofilloptions', {
        detail: {
          src: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
          dest: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 },
          mode: 'series',
          clientX: 120,
          clientY: 90,
        },
      }),
    );

    document.querySelector<HTMLButtonElement>('.fc-autofill-options__button')?.click();
    expect(visibleLabels()).toEqual([
      'Copy Cells',
      'Fill Series',
      'Fill Formatting Only',
      'Fill Without Formatting',
      'Fill Days',
      'Fill Weekdays',
      'Fill Months',
      'Fill Years',
    ]);
    const months = Array.from(
      document.querySelectorAll<HTMLButtonElement>('.fc-autofill-options__item'),
    ).find((item) => item.textContent === 'Fill Months');
    months?.click();
    wb.recalc();

    expect(num(wb, 1, 0)).toBe(dateSerial(2024, 2, 29));
    expect(num(wb, 2, 0)).toBe(dateSerial(2024, 3, 31));
  });

  it('detects date fill choices in huge sources from materialized formats only', () => {
    seedAndMirror(store, wb, [{ row: 100_000, col: 0, value: dateSerial(2024, 1, 31) }]);
    store.setState((s) => {
      const formats = new Map(s.format.formats);
      formats.set('0:100000:0', { numFmt: { kind: 'date', pattern: 'yyyy-mm-dd' } });
      return { ...s, format: { ...s.format, formats } };
    });
    const handle = attachAutoFillOptions({
      host,
      store,
      wb,
      strings: en,
    });
    detach = () => handle.detach();

    host.dispatchEvent(
      new CustomEvent('fc:autofilloptions', {
        detail: {
          src: { sheet: 0, r0: 0, c0: 0, r1: 100_000, c1: 0 },
          dest: { sheet: 0, r0: 0, c0: 0, r1: 100_001, c1: 0 },
          mode: 'series',
          clientX: 120,
          clientY: 90,
        },
      }),
    );

    document.querySelector<HTMLButtonElement>('.fc-autofill-options__button')?.click();
    expect(visibleLabels()).toContain('Fill Months');
  });

  it('keeps the smart button and menu within the viewport', () => {
    Object.defineProperty(window, 'innerWidth', { configurable: true, value: 320 });
    Object.defineProperty(window, 'innerHeight', { configurable: true, value: 180 });
    const handle = attachAutoFillOptions({
      host,
      store,
      wb,
      strings: en,
    });
    detach = () => handle.detach();
    const button = document.querySelector<HTMLButtonElement>('.fc-autofill-options__button');
    const menu = document.querySelector<HTMLDivElement>('.fc-autofill-options__menu');
    expect(button).not.toBeNull();
    expect(menu).not.toBeNull();
    if (!button || !menu) throw new Error('missing auto fill options elements');
    Object.defineProperty(button, 'offsetWidth', { configurable: true, value: 28 });
    Object.defineProperty(button, 'offsetHeight', { configurable: true, value: 28 });
    Object.defineProperty(menu, 'offsetWidth', { configurable: true, value: 180 });
    Object.defineProperty(menu, 'offsetHeight', { configurable: true, value: 80 });

    host.dispatchEvent(
      new CustomEvent('fc:autofilloptions', {
        detail: {
          src: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
          dest: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 },
          mode: 'series',
          clientX: 310,
          clientY: 170,
        },
      }),
    );

    expect(button.style.left).toBe('288px');
    expect(button.style.top).toBe('148px');
    expect(menu.style.left).toBe('136px');
    expect(menu.style.top).toBe('96px');
  });

  it('keeps smart button and menu item DOM on the shared floating options helper', () => {
    const source = readFileSync(join(root, 'src/interact/auto-fill-options.ts'), 'utf8');
    const helperSource = readFileSync(join(root, 'src/interact/floating-options-menu.ts'), 'utf8');

    expect(source).toContain('createFloatingOptionsButton({');
    expect(source).toContain('createFloatingOptionsMenuItem({');
    expect(source).not.toContain("const button = document.createElement('button')");
    expect(source).not.toContain("const copyItem = document.createElement('button')");
    expect(source).not.toContain("const item = document.createElement('button')");
    expect(helperSource).toContain("import { createInteractionButton } from './chip-button.js'");
    expect(helperSource).not.toContain("document.createElement('button')");
  });
});
