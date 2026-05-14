import { describe, expect, it } from 'vitest';

import { encodeHtml } from '../../../../src/commands/clipboard/html.js';
import { addrKey, WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { type CellFormat, createSpreadsheetStore } from '../../../../src/store/store.js';

/** Seed a tiny workbook + store with values and per-cell formats. Returns the
 *  store so encodeHtml can be called directly. */
const seed = async (
  cells: { row: number; col: number; value: string | number; fmt?: CellFormat }[],
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

/** Parse the encoder's HTML and turn it back into a comparable
 *  cells×styles shape. Uses happy-dom's DOMParser so the test exercises
 *  the same browser-side parse the paste path would. */
function parseHtmlTable(html: string): {
  text: string[][];
  styles: string[][];
  hrefs: (string | null)[][];
} {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const trs = Array.from(doc.querySelectorAll('table tr'));
  const text: string[][] = [];
  const styles: string[][] = [];
  const hrefs: (string | null)[][] = [];
  for (const tr of trs) {
    const cellsT: string[] = [];
    const cellsS: string[] = [];
    const cellsH: (string | null)[] = [];
    for (const td of Array.from(tr.children)) {
      cellsT.push((td.textContent ?? '').trim());
      cellsS.push(td.getAttribute('style') ?? '');
      const a = td.querySelector('a');
      cellsH.push(a?.getAttribute('href') ?? null);
    }
    text.push(cellsT);
    styles.push(cellsS);
    hrefs.push(cellsH);
  }
  return { text, styles, hrefs };
}

describe('commands/clipboard — HTML encode roundtrip via DOMParser', () => {
  it('values survive encode → parse for plain text + numbers', async () => {
    const store = await seed([
      { row: 0, col: 0, value: 'Header A' },
      { row: 0, col: 1, value: 'Header B' },
      { row: 1, col: 0, value: 42 },
      { row: 1, col: 1, value: 'hello' },
    ]);
    const html = encodeHtml(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    const parsed = parseHtmlTable(html);
    expect(parsed.text).toEqual([
      ['Header A', 'Header B'],
      ['42', 'hello'],
    ]);
  });

  it('styles survive encode → parse for bold + color + fill', async () => {
    const store = await seed([
      {
        row: 0,
        col: 0,
        value: 'styled',
        fmt: { bold: true, color: '#ff0000', fill: '#00ff00' },
      },
    ]);
    const html = encodeHtml(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const parsed = parseHtmlTable(html);
    expect(parsed.styles[0]?.[0]).toContain('font-weight:bold');
    expect(parsed.styles[0]?.[0]).toContain('color:#ff0000');
    expect(parsed.styles[0]?.[0]).toContain('background-color:#00ff00');
  });

  it('hyperlinks survive encode → parse with href intact', async () => {
    const store = await seed([
      {
        row: 0,
        col: 0,
        value: 'Click me',
        fmt: { hyperlink: 'https://example.com/path?q=1&r=2' },
      },
    ]);
    const html = encodeHtml(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const parsed = parseHtmlTable(html);
    // The encoder escapes `&` to `&amp;` in the href; DOMParser unescapes it
    // back when reading getAttribute, so the round-trip matches the original.
    expect(parsed.hrefs[0]?.[0]).toBe('https://example.com/path?q=1&r=2');
    expect(parsed.text[0]?.[0]).toBe('Click me');
  });

  it('HTML special characters in cell text survive both legs of the trip', async () => {
    const store = await seed([
      { row: 0, col: 0, value: '<b>not bold</b>' },
      { row: 0, col: 1, value: 'A & B "quoted"' },
    ]);
    const html = encodeHtml(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    const parsed = parseHtmlTable(html);
    expect(parsed.text[0]).toEqual(['<b>not bold</b>', 'A & B "quoted"']);
  });

  it('idempotent encoding: re-encoding parsed values produces identical structure', async () => {
    const store = await seed([
      { row: 0, col: 0, value: 'a' },
      { row: 0, col: 1, value: 'b' },
    ]);
    const html1 = encodeHtml(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    const html2 = encodeHtml(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    expect(html2).toBe(html1);

    // Re-seed a fresh store with the parsed values and verify the
    // re-encoded HTML matches.
    const parsed = parseHtmlTable(html1);
    const replay = await seed(
      parsed.text.flatMap((row, r) => row.map((value, c) => ({ row: r, col: c, value }))),
    );
    const html3 = encodeHtml(replay.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    expect(html3).toBe(html1);
  });

  it('multi-row range row/column counts survive the roundtrip', async () => {
    const store = await seed([
      { row: 0, col: 0, value: 'r0c0' },
      { row: 0, col: 1, value: 'r0c1' },
      { row: 0, col: 2, value: 'r0c2' },
      { row: 1, col: 0, value: 'r1c0' },
      { row: 1, col: 1, value: 'r1c1' },
      { row: 1, col: 2, value: 'r1c2' },
      { row: 2, col: 0, value: 'r2c0' },
      { row: 2, col: 1, value: 'r2c1' },
      { row: 2, col: 2, value: 'r2c2' },
    ]);
    const html = encodeHtml(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 });
    const parsed = parseHtmlTable(html);
    expect(parsed.text).toHaveLength(3);
    expect(parsed.text.every((row) => row.length === 3)).toBe(true);
  });

  it('blank cells round-trip as empty strings (preserving column count)', async () => {
    const store = await seed([
      { row: 0, col: 0, value: 'a' },
      // (0,1) blank
      { row: 0, col: 2, value: 'c' },
    ]);
    const html = encodeHtml(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
    const parsed = parseHtmlTable(html);
    expect(parsed.text[0]).toEqual(['a', '', 'c']);
  });

  it('numeric values with currency numFmt keep their formatted string after parse', async () => {
    const store = await seed([
      {
        row: 0,
        col: 0,
        value: 1234.5,
        fmt: { numFmt: { kind: 'currency', decimals: 2, symbol: '$' } },
      },
    ]);
    const html = encodeHtml(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const parsed = parseHtmlTable(html);
    const text = parsed.text[0]?.[0] ?? '';
    expect(text).toContain('$');
    expect(text).toContain('1,234.50');
  });

  it('text alignment + font family + size styles survive parsing', async () => {
    const store = await seed([
      {
        row: 0,
        col: 0,
        value: 'big',
        fmt: { align: 'right', fontFamily: 'Arial', fontSize: 18 },
      },
    ]);
    const html = encodeHtml(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const parsed = parseHtmlTable(html);
    const style = parsed.styles[0]?.[0] ?? '';
    expect(style).toContain('text-align:right');
    expect(style).toContain('font-family:Arial');
    expect(style).toContain('font-size:18px');
  });
});
