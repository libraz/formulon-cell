import { afterEach, describe, expect, it } from 'vitest';
import { evaluateConditional } from '../../../src/render/conditional.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import { formulaRule, seedCell, withConditionalRules } from './conditional-fixtures.js';

/** True when the one-cell formula rule over `text` in A1 fires. */
const fires = (text: string, formula: string): boolean => {
  const store = createSpreadsheetStore();
  let s = store.getState();
  s = seedCell(s, 0, 0, { kind: 'text', value: text });
  s = withConditionalRules(s, [
    formulaRule({ sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }, formula, '#hit'),
  ]);
  return evaluateConditional(s).get('0:0:1')?.fill === '#hit';
};

describe('SEARCH wildcards', () => {
  it.each([
    ['abcdef', '"c?e"', '', 3],
    ['abcdef', '"?a"', '', null],
    ['abcdef', '"?b"', '', 1],
    ['abcdef', '"c*f"', '', 3],
    ['abcabc', '"b?"', ',3', 5],
    ['abcabc', '"b?"', ',6', null],
    ['abcabc', '"b?"', ',7', null],
    ['ABCdef', '"c?e"', '', 3],
    ['a*c', '"~*"', '', 2],
    ['a?c', '"a~?c"', '', 1],
    ['abc', '"a~?c"', '', null],
    ['a~c', '"a~~c"', '', 1],
    ['abc', '"x*"', '', null],
    ['abc', '"*"', '', 1],
    ['abc', '"*"', ',4', 4],
  ])('SEARCH(%j, %s%s) -> %j', (text, pattern, start, expected) => {
    const formula = `SEARCH(${pattern},A1${start})`;
    if (expected === null) {
      expect(fires(text, `=ISERROR(${formula})`)).toBe(true);
    } else {
      expect(fires(text, `=${formula}=${expected}`)).toBe(true);
    }
  });

  it('start beyond the text is #VALUE!', () => {
    expect(fires('abc', '=ISERROR(SEARCH("a*",A1,5))')).toBe(true);
  });

  it('FIND keeps wildcards literal', () => {
    expect(fires('a*c', '=FIND("*",A1)=2')).toBe(true);
    expect(fires('abc', '=ISERROR(FIND("a*c",A1))')).toBe(true);
  });
});

describe('COUNTIF wildcards', () => {
  const count = (text: string, criteria: string): boolean =>
    fires(text, `=COUNTIF($A$1:$A$1,${criteria})=1`);

  it.each([
    ['North', '"N*"', true],
    ['north', '"N*H"', true],
    ['North', '"*x"', false],
    ['abc', '"a?c"', true],
    ['abbc', '"a?c"', false],
    ['abc', '"?bc"', true],
    ['a*c', '"a~*c"', true],
    ['abc', '"a~*c"', false],
    ['a?c', '"a~?c"', true],
    ['abc', '"a~?c"', false],
    ['a~c', '"a~~c"', true],
    ['xabc', '"abc*"', false],
    ['abcx', '"*abc"', false],
  ])('COUNTIF over %j with %s -> %s', (text, criteria, expected) => {
    expect(count(text, criteria)).toBe(expected);
  });

  it('<> negates the wildcard match', () => {
    expect(fires('South', '=COUNTIF($A$1:$A$1,"<>S*")=0')).toBe(true);
    expect(fires('North', '=COUNTIF($A$1:$A$1,"<>S*")=1')).toBe(true);
  });
});

describe('wildcard pattern compilation', () => {
  const NativeRegExp = RegExp;
  afterEach(() => {
    globalThis.RegExp = NativeRegExp;
  });

  it('compiles a COUNTIF wildcard criterion once for a whole range', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    for (let row = 0; row < 40; row += 1) {
      s = seedCell(s, row, 0, { kind: 'text', value: row % 2 ? 'North' : 'South' });
    }
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=COUNTIF($A$1:$A$40,"No*")=20',
        '#hit',
      ),
    ]);
    let compiled = 0;
    globalThis.RegExp = new Proxy(NativeRegExp, {
      construct(target, args, newTarget) {
        if (typeof args[0] === 'string' && args[0].includes('.*')) compiled += 1;
        return Reflect.construct(target, args, newTarget);
      },
    });

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#hit');
    expect(compiled).toBeLessThanOrEqual(1);
  });
});
