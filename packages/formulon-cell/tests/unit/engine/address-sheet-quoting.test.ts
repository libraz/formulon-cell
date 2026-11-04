import { describe, expect, it } from 'vitest';
import { quoteSheetName } from '../../../src/engine/address.js';

describe('engine/address quoteSheetName', () => {
  it('leaves plain identifiers bare', () => {
    expect(quoteSheetName('Sheet1')).toBe('Sheet1');
    expect(quoteSheetName('Data_2')).toBe('Data_2');
  });

  it('quotes names that read as A1 references', () => {
    expect(quoteSheetName('Q1')).toBe("'Q1'");
    expect(quoteSheetName('abc12')).toBe("'abc12'");
  });

  it('quotes R1C1-shaped and all-digit names', () => {
    expect(quoteSheetName('R1C1')).toBe("'R1C1'");
    expect(quoteSheetName('2024')).toBe("'2024'");
  });

  it('quotes names with spaces and doubles embedded quotes', () => {
    expect(quoteSheetName('My Sheet')).toBe("'My Sheet'");
    expect(quoteSheetName("O'Brien")).toBe("'O''Brien'");
  });
});
