import { describe, expect, it } from 'vitest';
import { encodeCSV, parseCSV } from '../../../../src/commands/clipboard/csv.js';

describe('encodeCSV', () => {
  it('emits commas between cells and CRLF between rows by default', () => {
    expect(
      encodeCSV([
        ['a', 'b'],
        ['c', 'd'],
      ]),
    ).toBe('a,b\r\nc,d');
  });

  it('quotes cells containing commas, quotes, or line breaks', () => {
    expect(encodeCSV([['hello, world', 'a"b', 'line1\nline2']])).toBe(
      '"hello, world","a""b","line1\nline2"',
    );
  });

  it('quotes leading/trailing whitespace so trim-on-read is non-destructive', () => {
    expect(encodeCSV([[' hello', 'world ', 'middle space']])).toBe(
      '" hello","world ",middle space',
    );
  });

  it('honours eol option', () => {
    expect(
      encodeCSV(
        [
          ['a', 'b'],
          ['c', 'd'],
        ],
        { eol: '\n' },
      ),
    ).toBe('a,b\nc,d');
  });

  it('prepends a BOM when requested', () => {
    expect(encodeCSV([['a']], { bom: true }).charCodeAt(0)).toBe(0xfeff);
  });
});

describe('parseCSV', () => {
  it('splits a simple comma-separated grid', () => {
    expect(parseCSV('a,b,c\n1,2,3')).toEqual([
      ['a', 'b', 'c'],
      ['1', '2', '3'],
    ]);
  });

  it('honours \\r\\n, \\n, and \\r as row terminators', () => {
    expect(parseCSV('a,b\r\nc,d')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
    expect(parseCSV('a,b\rc,d')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
    expect(parseCSV('a,b\nc,d')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
  });

  it('drops a single trailing terminator (no phantom empty row)', () => {
    expect(parseCSV('a\n')).toEqual([['a']]);
    expect(parseCSV('a\r\n')).toEqual([['a']]);
  });

  it('preserves a fully empty trailing row when caller injects two breaks', () => {
    expect(parseCSV('a\n\n')).toEqual([['a'], ['']]);
  });

  it('handles quoted fields with commas inside', () => {
    expect(parseCSV('"hello, world",x')).toEqual([['hello, world', 'x']]);
  });

  it('handles doubled quotes inside quoted fields', () => {
    expect(parseCSV('"a""b",c')).toEqual([['a"b', 'c']]);
  });

  it('handles embedded newlines inside quoted fields', () => {
    expect(parseCSV('"line1\nline2",b')).toEqual([['line1\nline2', 'b']]);
  });

  it('strips a leading UTF-8 BOM', () => {
    expect(parseCSV('\u{FEFF}a,b')).toEqual([['a', 'b']]);
  });

  it('round-trips the encode → parse pipeline', () => {
    const grid = [
      ['a', 'b,c', 'd"e'],
      ['', 'line1\nline2', '   space'],
    ];
    expect(parseCSV(encodeCSV(grid))).toEqual(grid);
  });
});
