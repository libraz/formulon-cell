import { describe, expect, it } from 'vitest';
import {
  insertFormulaReference,
  resolveFormulaReferenceInsertionSpan,
} from '../../../src/commands/formula-reference-input.js';

describe('insertFormulaReference', () => {
  it.each([
    { text: '=A1+B2', start: 0, end: 6, ref: 'C4', want: '=C4', caret: 3 },
    { text: '=A1+B2', start: 2, end: 2, ref: 'C4', want: '=C4+B2', caret: 3 },
    { text: '=SUM(A1:B2)', start: 8, end: 8, ref: 'C4', want: '=SUM(C4)', caret: 7 },
    { text: '=SUM(A1:B2)', start: 6, end: 6, ref: 'C4', want: '=SUM(C4)', caret: 7 },
    { text: '=SUM(A:C)', start: 7, end: 7, ref: 'B2', want: '=SUM(B2)', caret: 7 },
    { text: '=SUM(1:3)', start: 7, end: 7, ref: 'B2', want: '=SUM(B2)', caret: 7 },
    { text: '=SUM(A', start: 6, end: 6, ref: 'B2', want: '=SUM(B2', caret: 7 },
    { text: '=A1+B2', start: 1, end: 3, ref: 'C4', want: '=C4+B2', caret: 3 },
    { text: "='Other Sheet'!A1+B2", start: 16, end: 16, ref: 'C4', want: '=C4+B2', caret: 3 },
    {
      text: '=SUM(R[-2]C[-2]:R[-1]C[-1])',
      start: 20,
      end: 20,
      ref: 'R[1]C[1]',
      r1c1: true,
      want: '=SUM(R[1]C[1])',
      caret: 13,
    },
    {
      text: '=SUM(R[-2]C[-2]',
      start: 15,
      end: 15,
      ref: 'R[-2]C[-2]:R[-1]C[-1]',
      r1c1: true,
      want: '=SUM(R[-2]C[-2]:R[-1]C[-1]',
      caret: 26,
    },
  ])(
    'replaces the reference at $start in $text',
    ({ text, start, end, ref, r1c1, want, caret }) => {
      const result = insertFormulaReference(text, start, end, ref, { r1c1 });
      expect(result.text).toBe(want);
      expect(result.caret).toBe(caret);
    },
  );

  it.each(['="A1"', '=SUM("A1")', '=SUM(Table1[Q1])'])(
    'ignores a collapsed caret inside a quoted or bracketed expression: %s',
    (text) => {
      const caret = text.indexOf('A1') >= 0 ? text.indexOf('A1') + 1 : text.indexOf('Q1') + 1;
      expect(insertFormulaReference(text, caret, caret, 'C4').text).toBe(text);
    },
  );

  it('keeps the equals sign when a selection includes it and preserves trailing expression text', () => {
    expect(insertFormulaReference('=A1+B2', 0, 3, 'C4')).toEqual({ text: '=C4+B2', caret: 3 });
  });

  it('does not consume a function name as a partial cell reference', () => {
    expect(insertFormulaReference('=SUM(A1)', 2, 2, 'C4').text).toBe('=SUM(A1)');
    expect(insertFormulaReference('=SUM(A1)', 4, 4, 'C4').text).toBe('=SUM(A1)');
  });

  it.each([
    { text: '=PriceName+1', caret: 4 },
    { text: '=MyName+1', caret: 7 },
    { text: '=Foo.A1+1', caret: 6 },
    { text: '=XFE1+1', caret: 5 },
    { text: '=A1048577+1', caret: 9 },
    { text: '=1+2', caret: 4 },
    { text: '=12+3', caret: 2 },
  ])('does not splice a reference into an identifier or number in $text', ({ text, caret }) => {
    expect(insertFormulaReference(text, caret, caret, 'C4')).toEqual({ text, caret });
  });

  it('still inserts at a formula operand boundary', () => {
    expect(insertFormulaReference('=SUM(', 5, 5, 'C4')).toEqual({ text: '=SUM(C4', caret: 7 });
  });

  it.each(['=A1+', '=SUM(A1,', '=SUM(A1;', '=SUM(A1:', '=(A1+'])(
    'inserts at an empty operand after an operator or separator in %s',
    (text) => {
      const result = insertFormulaReference(text, text.length, text.length, 'C4');
      expect(result.text).toBe(`${text}C4`);
    },
  );

  it.each(['=SUM(A1)', '=SUM("x")', '=A1%', '=A1#', '="x"', '=1!'])(
    'does not append a reference after a completed operand in %s',
    (text) => {
      expect(insertFormulaReference(text, text.length, text.length, 'C4')).toEqual({
        text,
        caret: text.length,
      });
    },
  );

  it('continues to replace an explicit selection inside a quoted operand', () => {
    expect(insertFormulaReference('="SUM"', 2, 5, 'C4')).toEqual({ text: '="C4"', caret: 4 });
  });

  it('rejects a selection that is not entirely in expression context when asked', () => {
    const options = { requireExpressionSelection: true };
    expect(resolveFormulaReferenceInsertionSpan('="abc', 2, 5, options)).toBeNull();
    expect(resolveFormulaReferenceInsertionSpan('="SUM"', 2, 5, options)).toBeNull();
    expect(resolveFormulaReferenceInsertionSpan('=1+SUM(A1)+2', 3, 10, options)).toEqual({
      start: 3,
      end: 10,
    });
  });

  it('does not create an operand before an existing string literal', () => {
    const text = '=A1+"x"';
    expect(insertFormulaReference(text, 4, 4, 'C4')).toEqual({ text, caret: 4 });
  });

  it.each(['=@', '=Sheet1!', "='Other Sheet'!", '=@Sheet1!', "=@'Other Sheet'!"])(
    'keeps an operand after a valid prefix or sheet qualifier in %s',
    (text) => {
      expect(insertFormulaReference(text, text.length, text.length, 'C4')).toEqual({
        text: `${text}C4`,
        caret: text.length + 2,
      });
    },
  );

  it('replaces a complete R1C1 reference at its closing bracket boundary', () => {
    const text = '=SUM(R[-1]C[-1])';
    const caret = text.indexOf(')');
    expect(insertFormulaReference(text, caret, caret, 'R[0]C[0]', { r1c1: true })).toEqual({
      text: '=SUM(R[0]C[0])',
      caret: 13,
    });
  });

  it.each(['B1', 'C4', 'D5'])(
    'replaces the complete A1 colon chain when the collapsed caret is inside %s',
    (atom) => {
      const formula = '=SUM(B1:C4:D5)+E1';
      const caret = formula.indexOf(atom) + 1;

      expect(insertFormulaReference(formula, caret, caret, 'Z9')).toEqual({
        text: '=SUM(Z9)+E1',
        caret: 7,
      });
    },
  );

  it('replaces the complete R1C1 colon chain when the caret is inside a middle endpoint', () => {
    const formula = '=SUM(R[-2]C[-3]:R[-1]C[-2]:RC)+E1';
    const caret = formula.indexOf('R[-1]C[-2]') + 4;

    expect(insertFormulaReference(formula, caret, caret, 'R[2]C[2]', { r1c1: true })).toEqual({
      text: '=SUM(R[2]C[2])+E1',
      caret: 13,
    });
  });
});
