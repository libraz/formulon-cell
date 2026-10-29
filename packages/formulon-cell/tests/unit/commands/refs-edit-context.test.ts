import { describe, expect, it } from 'vitest';
import { extractRefs, rotateRefAt, suggestFunctions } from '../../../src/commands/refs.js';

describe('rotateRefAt lexical boundaries', () => {
  it.each([
    { text: '="A1"', target: 'A1', r1c1: false },
    { text: '=LOG10(A1)', target: 'LOG10', r1c1: false },
    { text: '=Table1[A1]', target: 'A1', r1c1: false },
    { text: '=R1C1', target: 'R1C1', r1c1: true },
    { text: "='A1'!B2", target: 'A1', r1c1: false },
  ])('does not treat non-reference text as A1 refs in $text', ({ text, target, r1c1 }) => {
    const caret = text.indexOf(target) + Math.floor(target.length / 2);
    expect(rotateRefAt(text, caret, { r1c1 })).toEqual({ text, caret });
  });

  it('rotates both endpoints of a valid A1 range independently', () => {
    expect(rotateRefAt('=SUM(A1:B2)', 7)).toEqual({ text: '=SUM($A$1:B2)', caret: 9 });
    expect(rotateRefAt('=SUM(A1:B2)', 10)).toEqual({ text: '=SUM(A1:$B$2)', caret: 12 });
  });

  it('rotates the real cell ref in a formula whose function name has digits', () => {
    expect(rotateRefAt('=LOG10(A1)', 9)).toEqual({ text: '=LOG10($A$1)', caret: 11 });
  });

  it('extracts multi-letter columns but skips function names, strings, and table headers', () => {
    const refs = extractRefs('=AA1+LOG10(A1)+"A1"+Table1[A1]');
    expect(refs.map(({ r0, c0, start }) => ({ r0, c0, start }))).toEqual([
      { r0: 0, c0: 26, start: 1 },
      { r0: 0, c0: 0, start: 11 },
    ]);
  });

  it('treats dotted defined names as one identifier when extracting and rotating refs', () => {
    expect(extractRefs('=Foo.A1+1')).toEqual([]);
    expect(rotateRefAt('=Foo.A1+1', 6)).toEqual({ text: '=Foo.A1+1', caret: 6 });
  });

  it('extracts colon chains as one bounding ref and assigns colors after merging', () => {
    const formula = '=E1+SUM(B1:C4:D5)+SUM(B1:D5)';
    const refs = extractRefs(formula);

    expect(refs).toEqual([
      {
        r0: 0,
        c0: 4,
        r1: 0,
        c1: 4,
        start: formula.indexOf('E1'),
        end: formula.indexOf('E1') + 2,
        colorIndex: 0,
      },
      {
        r0: 0,
        c0: 1,
        r1: 4,
        c1: 3,
        start: formula.indexOf('B1'),
        end: formula.indexOf('D5') + 2,
        colorIndex: 1,
      },
      {
        r0: 0,
        c0: 1,
        r1: 4,
        c1: 3,
        start: formula.lastIndexOf('B1'),
        end: formula.lastIndexOf('D5') + 2,
        colorIndex: 1,
      },
    ]);
  });

  it('does not offer function suggestions inside string literals', () => {
    expect(suggestFunctions('="SUM"', 5)).toBeNull();
  });

  it('does not rotate an R1C1 reference in R1C1 mode', () => {
    expect(rotateRefAt('=R[-1]C[2]', 5, { r1c1: true })).toEqual({
      text: '=R[-1]C[2]',
      caret: 5,
    });
  });
});
