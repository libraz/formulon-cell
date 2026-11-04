import { describe, expect, it } from 'vitest';
import { caretIndexAt, textStartOffset } from '../../../src/interact/editor-caret.js';

// A fixed 8 px advance per character keeps the arithmetic readable.
const measure = (text: string): number => text.length * 8;

describe('caretIndexAt', () => {
  it('puts the caret at the boundary nearest the click', () => {
    expect(caretIndexAt('200', -5, measure)).toBe(0);
    expect(caretIndexAt('200', 3, measure)).toBe(0);
    expect(caretIndexAt('200', 5, measure)).toBe(1);
    expect(caretIndexAt('200', 13, measure)).toBe(2);
    expect(caretIndexAt('200', 100, measure)).toBe(3);
  });

  it('places the caret by code point, never inside a surrogate pair', () => {
    expect(caretIndexAt('a😀b', 12, measure)).toBe(1);
    expect(caretIndexAt('a😀b', 21, measure)).toBe(3);
  });
});

describe('textStartOffset', () => {
  const box = { left: 100, width: 104, paddingLeft: 3, paddingRight: 3 };

  it('starts left-aligned text at the left padding', () => {
    expect(textStartOffset({ ...box, align: 'left' }, 24)).toBe(103);
  });

  it('ends right-aligned text at the right padding', () => {
    // A right-aligned "200" (24 px) in a 104 px box: 100 + 104 - 3 - 24.
    expect(textStartOffset({ ...box, align: 'right' }, 24)).toBe(177);
  });

  it('centres centred text between the paddings', () => {
    expect(textStartOffset({ ...box, align: 'center' }, 24)).toBe(140);
  });

  it('maps a click left of right-aligned text to the start, as the desktop does', () => {
    const start = textStartOffset({ ...box, align: 'right' }, measure('200'));
    expect(caretIndexAt('200', 150 - start, measure)).toBe(0);
  });
});
