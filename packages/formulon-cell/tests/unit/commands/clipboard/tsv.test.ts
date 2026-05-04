import { describe, expect, it } from 'vitest';
import { encodeTSV, parseTSV } from '../../../../src/commands/clipboard/tsv.js';

describe('encodeTSV', () => {
  it('joins rows with CRLF and cells with tab', () => {
    expect(encodeTSV([['a', 'b'], ['c']])).toBe('a\tb\r\nc');
  });

  it('quotes cells containing tab / newline / quote', () => {
    expect(encodeTSV([['a\tb']])).toBe('"a\tb"');
    expect(encodeTSV([['a\nb']])).toBe('"a\nb"');
    expect(encodeTSV([['a"b']])).toBe('"a""b"');
  });

  it('leaves plain cells unquoted', () => {
    expect(encodeTSV([['hello', 'world']])).toBe('hello\tworld');
  });
});

describe('parseTSV', () => {
  it('parses a single row', () => {
    expect(parseTSV('a\tb\tc')).toEqual([['a', 'b', 'c']]);
  });

  it('parses multiple rows separated by LF', () => {
    expect(parseTSV('a\tb\nc\td')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
  });

  it('treats CRLF as a single row break', () => {
    expect(parseTSV('a\r\nb')).toEqual([['a'], ['b']]);
  });

  it('strips a single trailing terminator', () => {
    expect(parseTSV('a\nb\n')).toEqual([['a'], ['b']]);
    expect(parseTSV('a\r\nb\r\n')).toEqual([['a'], ['b']]);
  });

  it('handles empty cells (consecutive tabs)', () => {
    expect(parseTSV('a\t\tc')).toEqual([['a', '', 'c']]);
  });

  it('unquotes cells and decodes doubled quotes', () => {
    expect(parseTSV('"a""b"')).toEqual([['a"b']]);
  });

  it('preserves embedded tabs and newlines inside quoted cells', () => {
    expect(parseTSV('"a\tb"\tc')).toEqual([['a\tb', 'c']]);
    expect(parseTSV('"a\nb"\tc')).toEqual([['a\nb', 'c']]);
  });

  it('round-trips encodeTSV → parseTSV for tricky inputs', () => {
    const rows = [
      ['plain', 'with tab\there', 'with "quote"'],
      ['', 'two\nlines', 'mixed\t"and"\n'],
    ];
    expect(parseTSV(encodeTSV(rows))).toEqual(rows);
  });
});
