import { describe, expect, it } from 'vitest';

import { encodeHtml } from '../../../../src/commands/clipboard/html.js';
import { addrKey, WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';

const seed = async (
  cells: { row: number; col: number; value: string | number; fmt?: Record<string, unknown> }[],
) => {
  const wb = await WorkbookHandle.createDefault({ preferStub: true });
  const store = createSpreadsheetStore();
  store.setState((s) => {
    const map = new Map(s.data.cells);
    const fmtMap = new Map(s.format.formats);
    for (const c of cells) {
      const addr = { sheet: 0, row: c.row, col: c.col };
      const k = addrKey(addr);
      if (typeof c.value === 'number') {
        wb.setNumber(addr, c.value);
        map.set(k, { value: { kind: 'number', value: c.value }, formula: null });
      } else {
        wb.setText(addr, c.value);
        map.set(k, { value: { kind: 'text', value: c.value }, formula: null });
      }
      if (c.fmt) fmtMap.set(k, c.fmt);
    }
    return {
      ...s,
      data: { ...s.data, cells: map },
      format: { ...s.format, formats: fmtMap },
    };
  });
  return store;
};

describe('commands/clipboard/encodeHtml', () => {
  it('wraps the range in <table><tr><td>…', async () => {
    const store = await seed([{ row: 0, col: 0, value: 'A' }]);
    const html = encodeHtml(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    expect(html).toBe('<table><tr><td>A</td></tr></table>');
  });

  it('escapes HTML special characters in text values', async () => {
    const store = await seed([{ row: 0, col: 0, value: '<a&b>"c"' }]);
    const html = encodeHtml(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    expect(html).toContain('&lt;a&amp;b&gt;&quot;c&quot;');
  });

  it('emits inline styles for bold + italic + underline + strike', async () => {
    const store = await seed([
      {
        row: 0,
        col: 0,
        value: 'rich',
        fmt: { bold: true, italic: true, underline: true, strike: true },
      },
    ]);
    const html = encodeHtml(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    expect(html).toContain('font-weight:bold');
    expect(html).toContain('font-style:italic');
    expect(html).toContain('text-decoration:underline line-through');
  });

  it('emits color + fill + align styles', async () => {
    const store = await seed([
      {
        row: 0,
        col: 0,
        value: 'styled',
        fmt: { color: '#abc', fill: '#fed', align: 'center' },
      },
    ]);
    const html = encodeHtml(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    expect(html).toContain('color:#abc');
    expect(html).toContain('background-color:#fed');
    expect(html).toContain('text-align:center');
  });

  it('renders multi-row, multi-column ranges as <tr> rows of <td> cells', async () => {
    const store = await seed([
      { row: 0, col: 0, value: 'a' },
      { row: 0, col: 1, value: 'b' },
      { row: 1, col: 0, value: 'c' },
      { row: 1, col: 1, value: 'd' },
    ]);
    const html = encodeHtml(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    expect(html).toBe('<table><tr><td>a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr></table>');
  });

  it('renders blank cells as empty <td>', async () => {
    const store = await seed([{ row: 0, col: 0, value: 'a' }]); // (0,1) blank
    const html = encodeHtml(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    expect(html).toBe('<table><tr><td>a</td><td></td></tr></table>');
  });

  it('wraps cells with hyperlinks in <a href>', async () => {
    const store = await seed([
      {
        row: 0,
        col: 0,
        value: 'Click',
        fmt: { hyperlink: 'https://example.com' },
      },
    ]);
    const html = encodeHtml(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    expect(html).toContain('<a href="https://example.com">Click</a>');
  });

  it('escapes the hyperlink href + text', async () => {
    const store = await seed([
      {
        row: 0,
        col: 0,
        value: 'a&b',
        fmt: { hyperlink: 'https://x.test/?a&b' },
      },
    ]);
    const html = encodeHtml(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    expect(html).toContain('href="https://x.test/?a&amp;b"');
    expect(html).toContain('>a&amp;b<');
  });

  it('applies numFmt before stringification (e.g. currency)', async () => {
    const store = await seed([
      {
        row: 0,
        col: 0,
        value: 1234.5,
        fmt: { numFmt: { kind: 'currency', decimals: 2, symbol: '$' } },
      },
    ]);
    const html = encodeHtml(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    // currency formatter should include the symbol; exact spacing varies but '$' must appear
    expect(html).toContain('$');
    expect(html).toContain('1,234.50');
  });
});
