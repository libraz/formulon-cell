import { describe, expect, it } from 'vitest';
import {
  assembledFormula,
  formulaWithArgumentCount,
  parseOuterCall,
  projectFormulaCallAtCaret,
  replaceProjectedFormulaCall,
} from '../../../src/interact/mac-formula-call.js';

const resolveKnownName = (rawName: string): string | null => {
  const name = rawName.toUpperCase();
  return new Set(['AVERAGE', 'MAX', 'SUM']).has(name) ? name : null;
};

describe('interact/mac-formula-call', () => {
  it('parses the outer call case-insensitively and keeps empty arguments', () => {
    expect(parseOuterCall('= sum(A1,,B2)', 'SUM')).toEqual({ name: 'SUM', args: ['A1', '', 'B2'] });
    expect(parseOuterCall('=AVERAGE(A1)', 'SUM')).toBeNull();
  });

  it('drops trailing blank arguments when assembling', () => {
    expect(assembledFormula('SUM', ['A1', '', ' '])).toBe('=SUM(A1)');
    expect(assembledFormula('TODAY', ['', ''])).toBe('=TODAY()');
  });

  it('trims trailing blanks when no explicit count is kept', () => {
    expect(formulaWithArgumentCount('SUM', ['A1', '', ''], null)).toBe('=SUM(A1)');
  });

  it('keeps exactly the explicit argument count, blanks included', () => {
    expect(formulaWithArgumentCount('SUM', ['A1', '', ''], 2)).toBe('=SUM(A1,)');
    expect(formulaWithArgumentCount('SUM', ['A1', 'B1', 'C1'], 1)).toBe('=SUM(A1)');
    expect(formulaWithArgumentCount('SUM', ['A1'], 0)).toBe('=SUM()');
  });

  it('projects the innermost known nested call at the caret', () => {
    const source = '=1+SUM(2,MAX(3,4))*5';
    const nested = projectFormulaCallAtCaret(source, source.indexOf('4') + 1, resolveKnownName);
    const outer = projectFormulaCallAtCaret(source, source.indexOf('2') + 1, resolveKnownName);

    expect(nested).toMatchObject({ source, args: ['3', '4'] });
    expect(nested?.span.canonicalName).toBe('MAX');
    expect(outer).toMatchObject({ source, args: ['2', 'MAX(3,4)'] });
    expect(outer?.span.canonicalName).toBe('SUM');
  });

  it('keeps blank fields, arrays, structured references, quotes, and R1C1 text as fields', () => {
    const source = '=SUM("a,b",Table1[[#Headers],[Amount]],R[1]C[-1],{1,2},  ,)';
    const projection = projectFormulaCallAtCaret(
      source,
      source.indexOf('Amount') + 2,
      resolveKnownName,
    );

    expect(projection?.args).toEqual([
      '"a,b"',
      'Table1[[#Headers],[Amount]]',
      'R[1]C[-1]',
      '{1,2}',
      '  ',
      '',
    ]);
    expect(projection?.span.complete).toBe(true);
  });

  it('keeps untouched argument slices verbatim when one field changes', () => {
    const source = '=SUM( 1 ,  2  )';
    const projection = projectFormulaCallAtCaret(source, source.indexOf('1') + 1, resolveKnownName);
    if (!projection) throw new Error('expected a SUM projection');
    expect(projection.args).toEqual([' 1 ', '  2  ']);
    const untouchedArgument = projection.args[1];
    if (untouchedArgument === undefined) throw new Error('expected the second SUM argument');

    const replaced = replaceProjectedFormulaCall(
      source,
      projection,
      'SUM',
      ['5', untouchedArgument],
      2,
    );

    expect(replaced?.raw).toBe('=SUM(5,  2  )');
  });

  it('replaces only the projected call and preserves prefix and suffix', () => {
    const source = '=1+SUM(2,3)*4';
    const projection = projectFormulaCallAtCaret(source, source.indexOf('3') + 1, resolveKnownName);
    if (!projection) throw new Error('expected a SUM projection');

    const replaced = replaceProjectedFormulaCall(source, projection, 'AVERAGE', ['2', '3'], 2);

    expect(replaced).toEqual({ raw: '=1+AVERAGE(2,3)*4', caret: '=1+AVERAGE(2,3)'.length });
  });

  it('keeps explicit trailing blanks when rebuilding a projected call', () => {
    const source = '=SUM(1,  ,)';
    const projection = projectFormulaCallAtCaret(source, source.indexOf('1') + 1, resolveKnownName);
    if (!projection) throw new Error('expected a SUM projection');
    expect(projection.args).toEqual(['1', '  ', '']);

    const replaced = replaceProjectedFormulaCall(source, projection, 'SUM', projection.args, 3);

    expect(replaced?.raw).toBe('=SUM(1,  ,)');
  });

  it('rejects a stale projection rather than overwriting a newer source', () => {
    const source = '=1+SUM(2,3)*4';
    const projection = projectFormulaCallAtCaret(source, source.indexOf('2') + 1, resolveKnownName);
    if (!projection) throw new Error('expected a SUM projection');

    expect(
      replaceProjectedFormulaCall('=1+SUM(9,3)*4', projection, 'SUM', ['9', '3'], 2),
    ).toBeNull();
  });

  it('projects an incomplete call and can complete it at the close boundary', () => {
    const source = '=1+SUM(2,';
    const projection = projectFormulaCallAtCaret(source, source.length, resolveKnownName);
    if (!projection) throw new Error('expected an incomplete SUM projection');
    expect(projection.args).toEqual(['2', '']);
    expect(projection.span.complete).toBe(false);

    const completed = replaceProjectedFormulaCall(source, projection, 'SUM', ['2', '3'], 2);
    if (!completed) throw new Error('expected the incomplete SUM replacement');
    expect(completed).toEqual({ raw: '=1+SUM(2,3)', caret: '=1+SUM(2,3)'.length });

    const rescanned = projectFormulaCallAtCaret(completed.raw, completed.caret, resolveKnownName);
    expect(rescanned?.span.complete).toBe(true);
    expect(rescanned?.args).toEqual(['2', '3']);
  });

  it('rejects unsafe incomplete calls while preserving the raw source', () => {
    const cases = [
      '=SUM({1,2',
      '=SUM(Table1[[#Headers],[Amount]',
      '=SUM((1+2',
      '=SUM("unterminated',
      "=SUM('unterminated",
      '=SUM(UNKNOWN(1,2',
    ];

    for (const source of cases) {
      expect(projectFormulaCallAtCaret(source, source.length, resolveKnownName)).toBeNull();
    }
  });
});
