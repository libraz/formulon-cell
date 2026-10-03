import { describe, expect, it } from 'vitest';
import {
  dblClickRange,
  FUNCTION_NAMES,
  FUNCTION_SIGNATURES,
  findActiveSignature,
  findFunctionCallAtCaret,
  formatA1FormulaAsR1C1,
  normalizeR1C1Formula,
  shiftFormulaRefs,
  suggestFunctions,
} from '../../../src/commands/refs.js';

describe('normalizeR1C1Formula', () => {
  const base = { row: 3, col: 3 }; // D4

  it('converts relative, absolute, and same-axis R1C1 refs to A1 refs', () => {
    expect(normalizeR1C1Formula('=R[-2]C[-2]+R1C1+RC[-1]', base)).toBe('=B2+A1+C4');
  });

  it('converts R1C1 ranges atom-by-atom', () => {
    expect(normalizeR1C1Formula('=SUM(R[-2]C[-2]:R[-1]C[-1])', base)).toBe('=SUM(B2:C3)');
  });

  it('does not rewrite R1C1-looking text inside strings or identifiers', () => {
    expect(normalizeR1C1Formula('="R[-2]C[-2]"+FOOR1C1+R[1]C', base)).toBe(
      '="R[-2]C[-2]"+FOOR1C1+D5',
    );
  });
});

describe('formatA1FormulaAsR1C1', () => {
  const base = { row: 3, col: 3 }; // D4

  it('formats relative, absolute, and mixed A1 refs as R1C1', () => {
    expect(formatA1FormulaAsR1C1('=A1+$A$1+A$1+$A1+D4', base)).toBe(
      '=R[-3]C[-3]+R1C1+R1C[-3]+R[-3]C1+RC',
    );
  });

  it('formats A1 ranges atom-by-atom', () => {
    expect(formatA1FormulaAsR1C1('=SUM(B2:C3)', base)).toBe('=SUM(R[-2]C[-2]:R[-1]C[-1])');
  });

  it('does not rewrite A1-looking text inside strings or identifiers', () => {
    expect(formatA1FormulaAsR1C1('="A1"+FOOA1+B2', base)).toBe('="A1"+FOOA1+R[-2]C[-2]');
  });
});

describe('shiftFormulaRefs', () => {
  it('returns the input verbatim when delta is zero', () => {
    expect(shiftFormulaRefs('=A1+B2', 0, 0)).toBe('=A1+B2');
  });

  it('returns the input verbatim when not a formula', () => {
    expect(shiftFormulaRefs('A1+B2', 1, 1)).toBe('A1+B2');
  });

  it('shifts a single relative ref', () => {
    expect(shiftFormulaRefs('=A1', 2, 3)).toBe('=D3');
  });

  it('shifts a sum of relative refs', () => {
    expect(shiftFormulaRefs('=A1+B2', 1, 1)).toBe('=B2+C3');
  });

  it('respects $-locked column ($A1 — column anchored)', () => {
    expect(shiftFormulaRefs('=$A1', 2, 3)).toBe('=$A3');
  });

  it('respects $-locked row (A$1 — row anchored)', () => {
    expect(shiftFormulaRefs('=A$1', 2, 3)).toBe('=D$1');
  });

  it('does not shift a fully absolute ref ($A$1)', () => {
    expect(shiftFormulaRefs('=$A$1', 5, 5)).toBe('=$A$1');
  });

  it('handles ranges atom-by-atom', () => {
    expect(shiftFormulaRefs('=SUM(A1:B5)', 1, 1)).toBe('=SUM(B2:C6)');
  });

  it('handles a range with mixed locks', () => {
    expect(shiftFormulaRefs('=SUM($A1:B$5)', 2, 3)).toBe('=SUM($A3:E$5)');
  });

  it('does not touch refs inside string literals', () => {
    expect(shiftFormulaRefs('="A1 is here"+B2', 1, 1)).toBe('="A1 is here"+C3');
  });

  it('does not treat a function call as a ref (SIN(A1) — only A1 shifts)', () => {
    expect(shiftFormulaRefs('=SIN(A1)', 1, 1)).toBe('=SIN(B2)');
  });

  it('shifts every ref in a SUM range plus an addend', () => {
    expect(shiftFormulaRefs('=SUM(A1:B2)+C3', 1, 1)).toBe('=SUM(B2:C3)+D4');
  });

  it('shifts multi-letter columns (AA1 → AB1 with +1 col)', () => {
    expect(shiftFormulaRefs('=AA1', 0, 1)).toBe('=AB1');
  });

  it('handles negative shifts', () => {
    expect(shiftFormulaRefs('=C3', -1, -1)).toBe('=B2');
  });

  it('turns out-of-range shifted refs into #REF!', () => {
    expect(shiftFormulaRefs('=A1', -1, 0)).toBe('=#REF!');
  });

  it('handles a sheet-qualified ref (Sheet1!A1)', () => {
    // Atom regex doesn't match the sheet token, so only the cell part shifts.
    expect(shiftFormulaRefs('=Sheet1!A1', 1, 1)).toBe('=Sheet1!B2');
  });
});

describe('findActiveSignature', () => {
  it('includes the ACOS argument label in the static help signature', () => {
    expect(FUNCTION_SIGNATURES.ACOS).toEqual(['number']);
  });

  it('returns null when caret is outside a function call', () => {
    expect(findActiveSignature('=A1+B2', 5)).toBeNull();
  });

  it('returns null when text is not a formula', () => {
    expect(findActiveSignature('SUM(1, 2)', 5)).toBeNull();
  });

  it('returns the matching signature for a known function', () => {
    const sig = findActiveSignature('=SUM(', 5);
    expect(sig?.name).toBe('SUM');
    expect(sig?.activeArgIndex).toBe(0);
  });

  it('keeps the legacy signature for an unsafe incomplete call', () => {
    const text = '=SUM({1,2';
    expect(findActiveSignature(text, text.length)).toEqual({
      name: 'SUM',
      args: FUNCTION_SIGNATURES.SUM,
      activeArgIndex: 0,
    });
  });

  it('bumps activeArgIndex once per top-level comma', () => {
    const text = '=IF(A1>0, B1,';
    const sig = findActiveSignature(text, text.length);
    expect(sig?.name).toBe('IF');
    expect(sig?.activeArgIndex).toBe(2);
  });

  it('ignores commas inside nested calls', () => {
    const text = '=SUMIF(A1:A10, ">5", SUM(B1, B2),';
    const sig = findActiveSignature(text, text.length);
    expect(sig?.name).toBe('SUMIF');
    expect(sig?.activeArgIndex).toBe(3);
  });

  it('ignores commas inside string literals', () => {
    const text = '=CONCAT("a, b", ';
    const sig = findActiveSignature(text, text.length);
    expect(sig?.name).toBe('CONCAT');
    expect(sig?.activeArgIndex).toBe(1);
  });

  it('returns null for unknown function names', () => {
    expect(findActiveSignature('=NOTAFUNC(', 10)).toBeNull();
  });

  it('resolves the signature when the caret sits on the function name', () => {
    // caret inside "SUM" of "=SUM(F4:F8)"
    const sig = findActiveSignature('=SUM(F4:F8)', 2);
    expect(sig?.name).toBe('SUM');
    expect(sig?.activeArgIndex).toBe(0);
  });

  it('does not resolve a bare name with no following paren', () => {
    expect(findActiveSignature('=SUM', 2)).toBeNull();
  });

  it('keeps the legacy null result just after a closed top-level call', () => {
    const text = '=SUM(1)';
    expect(findActiveSignature(text, text.length)).toBeNull();
    const nested = '=SUM(IF(1))';
    const innerClose = nested.indexOf(')', nested.indexOf('IF'));
    expect(findActiveSignature(nested, innerClose + 1)?.name).toBe('SUM');
  });

  it('keeps an unknown innermost call from being masked by a known outer call', () => {
    const text = '=SUM(UNKNOWN(1),2)';
    const innerStart = text.indexOf('UNKNOWN');
    expect(findActiveSignature(text, innerStart + 2)).toBeNull();
    const innerClose = text.indexOf(')', innerStart);
    expect(findActiveSignature(text, innerClose + 1)).toEqual({
      name: 'SUM',
      args: FUNCTION_SIGNATURES.SUM,
      activeArgIndex: 0,
    });
  });
});

describe('findFunctionCallAtCaret', () => {
  const resolve = (rawName: string): string | null => {
    const canonical = rawName.toUpperCase();
    return FUNCTION_SIGNATURES[canonical] ? canonical : null;
  };

  it('returns raw UTF-16 spans for an inner call in a compound formula', () => {
    const text = '=1+SUM( A1 , IF(B1, "x,y", {2,3}) )*2';
    const caret = text.indexOf('IF') + 4;
    const span = findFunctionCallAtCaret(text, caret, resolve);
    expect(span?.rawName).toBe('IF');
    expect(span?.name).toEqual({ start: text.indexOf('IF'), end: text.indexOf('IF') + 2 });
    expect(text.slice(span?.call.start, span?.call.end)).toBe('IF(B1, "x,y", {2,3})');
    expect(span?.argumentSpans.map(({ start, end }) => text.slice(start, end))).toEqual([
      'B1',
      ' "x,y"',
      ' {2,3}',
    ]);
    expect(span?.complete).toBe(true);
    expect(span?.activeArgumentIndex).toBe(0);
  });

  it('uses the enclosing call exactly one position after an inner close', () => {
    const text = '=SUM(IF(1,2),3)';
    const innerClose = text.indexOf(')', text.indexOf('IF'));
    const inner = findFunctionCallAtCaret(text, innerClose, resolve);
    const outer = findFunctionCallAtCaret(text, innerClose + 1, resolve);
    expect(inner?.canonicalName).toBe('IF');
    expect(inner?.activeArgumentIndex).toBe(1);
    expect(outer?.canonicalName).toBe('SUM');
    expect(text.slice(outer?.call.start, outer?.call.end)).toBe(text.slice(1));
    expect(outer?.activeArgumentIndex).toBe(0);
    expect(findFunctionCallAtCaret(text, text.length, resolve)?.canonicalName).toBe('SUM');
  });

  it('falls through an unknown inner call to a live-resolved enclosing call', () => {
    const text = '=LOCAL(UNKNOWN(1),2)';
    const span = findFunctionCallAtCaret(text, text.indexOf('UNKNOWN') + 9, (rawName) =>
      rawName === 'LOCAL' ? 'LOCAL.CANONICAL' : null,
    );
    expect(span?.rawName).toBe('LOCAL');
    expect(span?.canonicalName).toBe('LOCAL.CANONICAL');
    expect(span?.argumentSpans.map(({ start, end }) => text.slice(start, end))).toEqual([
      'UNKNOWN(1)',
      '2',
    ]);
  });

  it('passes dotted and _xlfn names to the live resolver without normalizing raw spans', () => {
    const text = '=_xlfn.BETA.DIST (1)';
    const span = findFunctionCallAtCaret(text, text.indexOf('1'), (rawName) =>
      rawName === '_xlfn.BETA.DIST' ? 'BETA.DIST' : null,
    );
    expect(span?.rawName).toBe('_xlfn.BETA.DIST');
    expect(span?.canonicalName).toBe('BETA.DIST');
    expect(text.slice(span?.name.start, span?.name.end)).toBe('_xlfn.BETA.DIST');
    expect(text.slice(span?.call.start, span?.call.end)).toBe('=_xlfn.BETA.DIST (1)'.slice(1));
  });

  it('keeps emoji prefixes in UTF-16 offsets and preserves the source text', () => {
    const text = '="😀"+SUM(A1)';
    const span = findFunctionCallAtCaret(text, text.indexOf('A1') + 1, resolve);
    expect(span?.name).toEqual({ start: text.indexOf('SUM'), end: text.indexOf('SUM') + 3 });
    expect(span?.call).toEqual({ start: text.indexOf('SUM'), end: text.length });
    expect(text.slice(span?.argumentSpans[0]?.start, span?.argumentSpans[0]?.end)).toBe('A1');
  });

  it('ignores commas in doubled quotes, arrays, and escaped structured references', () => {
    const text = '=SUM("a,b","a""b",{1,2},Table1[[#This Row],[Col\'[x\']#\'@\'\'x]])';
    const span = findFunctionCallAtCaret(text, text.indexOf('Col'), resolve);
    expect(span?.argumentSpans.map(({ start, end }) => text.slice(start, end))).toEqual([
      '"a,b"',
      '"a""b"',
      '{1,2}',
      "Table1[[#This Row],[Col'[x']#'@''x]]",
    ]);
    expect(span?.activeArgumentIndex).toBe(3);
  });

  it('ignores commas in doubled apostrophe sheet quotes', () => {
    const text = "=SUM('O''Brien, West'!A1,2)";
    const span = findFunctionCallAtCaret(text, text.indexOf('Brien'), resolve);
    expect(span?.argumentSpans.map(({ start, end }) => text.slice(start, end))).toEqual([
      "'O''Brien, West'!A1",
      '2',
    ]);
  });

  it('keeps explicit and trailing blank argument spans zero-length', () => {
    const text = '=SUM(1,,)';
    const firstComma = text.indexOf(',');
    const secondComma = text.indexOf(',', firstComma + 1);
    const span = findFunctionCallAtCaret(text, secondComma, resolve);
    expect(span?.argumentSpans.map(({ start, end }) => [start, end])).toEqual([
      [5, 6],
      [secondComma, secondComma],
      [secondComma + 1, secondComma + 1],
    ]);
    expect(span?.activeArgumentIndex).toBe(1);
    expect(findFunctionCallAtCaret(text, secondComma + 1, resolve)?.activeArgumentIndex).toBe(2);
    const empty = '=SUM()';
    expect(findFunctionCallAtCaret(empty, empty.length, resolve)?.argumentSpans).toEqual([
      { start: 5, end: 5 },
    ]);
  });

  it('preserves raw whitespace in nonblank argument ranges', () => {
    const text = '=SUM(  A1  ,B1 )';
    const span = findFunctionCallAtCaret(text, text.indexOf('B1'), resolve);
    expect(span?.argumentSpans.map(({ start, end }) => text.slice(start, end))).toEqual([
      '  A1  ',
      'B1 ',
    ]);
  });

  it('reports EOF-incomplete calls and rejects crossed delimiters', () => {
    const incomplete = '=SUM(IF(1,2),';
    const span = findFunctionCallAtCaret(incomplete, incomplete.length, resolve);
    expect(span?.canonicalName).toBe('SUM');
    expect(span?.closeParen).toBeNull();
    expect(span?.complete).toBe(false);
    expect(span?.safeToComplete).toBe(true);
    expect(span?.argumentSpans.map(({ start, end }) => incomplete.slice(start, end))).toEqual([
      'IF(1,2)',
      '',
    ]);
    expect(findFunctionCallAtCaret('=SUM({1,2],3)', 8, resolve)).toBeNull();
    expect(findFunctionCallAtCaret('=SUM("unterminated)', 8, resolve)).toMatchObject({
      complete: false,
      safeToComplete: false,
    });
  });

  it('marks only structurally closable EOF calls as safe to complete', () => {
    const cases: readonly [string, boolean][] = [
      ['=SUM(1,2', true],
      ['=SUM({1,2', false],
      ['=SUM(Table1[[#Headers],[Amount]', false],
      ['=SUM((1+2', false],
      ['=SUM("unterminated', false],
      ["=SUM('unterminated", false],
      ['=SUM(UNKNOWN(1,2', false],
    ];

    for (const [text, expected] of cases) {
      expect(findFunctionCallAtCaret(text, text.length, resolve)?.safeToComplete).toBe(expected);
    }
  });

  it('rejects nonformulas and invalid caret positions', () => {
    const text = '=SUM(1)';
    expect(findFunctionCallAtCaret('SUM(1)', 4, resolve)).toBeNull();
    expect(findFunctionCallAtCaret(text, -1, resolve)).toBeNull();
    expect(findFunctionCallAtCaret(text, text.length + 1, resolve)).toBeNull();
    expect(findFunctionCallAtCaret(text, 1.5, resolve)).toBeNull();
  });
});

describe('dblClickRange', () => {
  it('selects the function name when the probe is on it', () => {
    expect(dblClickRange('=SUM(F4:F8)', 2)).toEqual({ start: 1, end: 4 });
  });

  it('selects a whole range argument rather than splitting at the colon', () => {
    const text = '=SUM(F4:F8)';
    expect(dblClickRange(text, 7)).toEqual({ start: 5, end: 10 });
    expect(text.slice(5, 10)).toBe('F4:F8');
  });

  it('selects the argument between commas', () => {
    const text = '=IF(A1>0, B7, C9)';
    const r = dblClickRange(text, 11); // probe on "B7"
    expect(text.slice(r?.start, r?.end)).toBe('B7');
  });

  it('trims surrounding whitespace from the argument', () => {
    const text = '=SUM( F4:F8 )';
    const r = dblClickRange(text, 8);
    expect(text.slice(r?.start, r?.end)).toBe('F4:F8');
  });

  it('selects the nested function name inside a call', () => {
    const text = '=SUM(AVERAGE(A1:A5), B1)';
    expect(dblClickRange(text, 8)).toEqual({ start: 5, end: 12 });
  });

  it('does not split a quoted string argument at its comma', () => {
    const text = '=CONCAT("a, b", C1)';
    const r = dblClickRange(text, 10); // probe inside "a, b"
    expect(text.slice(r?.start, r?.end)).toBe('"a, b"');
  });

  it('returns null at the top level so native word selection wins', () => {
    expect(dblClickRange('=A1+B2', 2)).toBeNull();
  });

  it('extends to the end when the enclosing paren is unclosed', () => {
    const text = '=SUM(F4:F8';
    const r = dblClickRange(text, 7);
    expect(text.slice(r?.start, r?.end)).toBe('F4:F8');
  });
});

describe('suggestFunctions', () => {
  it('returns built-in matches by default', () => {
    const r = suggestFunctions('=SU', 3);
    expect(r?.token).toBe('SU');
    expect(r?.matches).toContain('SUM');
  });

  it('returns null when not in a formula', () => {
    expect(suggestFunctions('SUM', 3)).toBeNull();
  });

  it('uses opts.names when supplied (engine catalog override)', () => {
    const engineNames = ['CUSTOM_FN', 'CUSTOM_OTHER', 'OTHER'];
    const r = suggestFunctions('=CUS', 4, 8, { names: engineNames });
    expect(r?.matches).toEqual(['CUSTOM_FN', 'CUSTOM_OTHER']);
  });

  it('does not fall back to built-ins when opts.names is empty', () => {
    expect(suggestFunctions('=SU', 3, 8, { names: [] })).toBeNull();
  });

  it('keeps the dotted function prefix and replacement start from the engine catalog', () => {
    const names = ['BETA.DIST', 'BETA.INV', 'NORM.DIST', 'NORM.INV'];
    expect(suggestFunctions('=BETA.D', 7, 8, { names })).toEqual({
      token: 'BETA.D',
      tokenStart: 1,
      matches: ['BETA.DIST'],
    });
    expect(suggestFunctions('=SUM(A1,norm.', 13, 8, { names })).toEqual({
      token: 'norm.',
      tokenStart: 8,
      matches: ['NORM.DIST', 'NORM.INV'],
    });
    expect(suggestFunctions('=1.5', 4, 8, { names })).toBeNull();
  });
});

describe('365 dynamic-array function catalog', () => {
  it('FUNCTION_NAMES includes the marquee 365 array functions', () => {
    const expected = [
      'GROUPBY',
      'PIVOTBY',
      'TEXTSPLIT',
      'VSTACK',
      'HSTACK',
      'TOROW',
      'TOCOL',
      'CHOOSEROWS',
      'CHOOSECOLS',
      'TAKE',
      'DROP',
      'EXPAND',
      'LAMBDA',
      'LET',
      'MAP',
      'REDUCE',
      'SCAN',
      'BYROW',
      'BYCOL',
      'MAKEARRAY',
      'XMATCH',
      'SORTBY',
      'SEQUENCE',
      'RANDARRAY',
      'IMAGE',
    ];
    for (const n of expected) expect(FUNCTION_NAMES).toContain(n);
  });

  it('FUNCTION_SIGNATURES exposes the same names with parameter lists', () => {
    expect(FUNCTION_SIGNATURES.GROUPBY?.[0]).toBe('row_fields');
    expect(FUNCTION_SIGNATURES.TEXTSPLIT?.[0]).toBe('text');
    expect(FUNCTION_SIGNATURES.LAMBDA?.[0]).toBe('parameter');
    expect(FUNCTION_SIGNATURES.LET?.length).toBeGreaterThanOrEqual(3);
    // Signatures end with the calculation slot for LAMBDA / LET / MAP.
    expect(FUNCTION_SIGNATURES.MAP?.at(-1)).toBe('lambda');
  });

  it('suggestFunctions surfaces a 365 function from a partial token', () => {
    const r = suggestFunctions('=GRO', 4);
    expect(r?.matches).toContain('GROUPBY');
  });

  it('findActiveSignature resolves TEXTSPLIT inside an open call', () => {
    const text = '=TEXTSPLIT(';
    const sig = findActiveSignature(text, text.length);
    expect(sig?.name).toBe('TEXTSPLIT');
    expect(sig?.activeArgIndex).toBe(0);
  });
});
