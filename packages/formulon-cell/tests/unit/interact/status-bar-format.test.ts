import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import type { CellValue } from '../../../src/engine/types.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { attachStatusBar } from '../../../src/interact/status-bar.js';
import { createSpreadsheetStore, type NumFmt } from '../../../src/store/store.js';

const show = (values: number[], numFmt?: NumFmt): string => {
  const bar = document.createElement('div');
  document.body.appendChild(bar);
  const store = createSpreadsheetStore();
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    const formats = new Map(s.format.formats);
    values.forEach((value, row) => {
      const key = addrKey({ sheet: 0, row, col: 0 });
      cells.set(key, { value: { kind: 'number', value } as CellValue, formula: null });
      if (numFmt) formats.set(key, { numFmt });
    });
    return {
      ...s,
      data: { ...s.data, cells },
      format: { ...s.format, formats },
      selection: {
        active: { sheet: 0, row: 0, col: 0 },
        anchor: { sheet: 0, row: 0, col: 0 },
        range: { sheet: 0, r0: 0, c0: 0, r1: values.length - 1, c1: 0 },
      },
    };
  });
  const handle = attachStatusBar({ statusbar: bar, store, strings: en, getEngineLabel: () => '' });
  handle.refresh();
  const text = bar.querySelector('.fc-host__statusbar-aggs')?.textContent ?? '';
  handle.detach();
  return text;
};

// Excel for Mac: General results fit 11 characters and never group; a formatted
// first cell lends its number format to the results.
describe('status bar aggregate number format', () => {
  beforeEach(() => {
    document.body.replaceChildren();
  });
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('shows General results in at most 11 characters without grouping', () => {
    expect(show([1 / 3, 2 / 3, 1])).toBe(`Average: 0.666666667 · Count: 3 · Sum: 2`);
    expect(show([123456789, 987654321])).toBe('Average: 555555555 · Count: 2 · Sum: 1111111110');
    expect(show([1.23456789012345, 2])).toBe('Average: 1.617283945 · Count: 2 · Sum: 3.23456789');
    expect(show([10, 20, 40])).toBe('Average: 23.33333333 · Count: 3 · Sum: 70');
  });

  it('applies the first numeric cell number format to the results', () => {
    const fmt: NumFmt = { kind: 'currency', decimals: 0, symbol: '$' };
    expect(show([1234.5, 2], fmt)).toBe('Average: $618 · Count: 2 · Sum: $1,237');
  });
});
