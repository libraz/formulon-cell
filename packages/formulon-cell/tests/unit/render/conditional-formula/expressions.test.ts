import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  _resetConditionalCache,
  evaluateConditional,
  parseFormulaPredicate,
} from '../../../../src/render/conditional.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';
import {
  formulaRule,
  seedCell,
  seedNumber,
  withConditionalRules,
} from '../conditional-fixtures.js';

describe('evaluateConditional', () => {
  afterEach(() => {
    _resetConditionalCache();
    vi.useRealTimers();
  });

  it('formula rule fires for comparator-prefix predicates and skips unparseable forms', () => {
    expect(parseFormulaPredicate('>10')?.test({ kind: 'number', value: 11 })).toBe(true);
    expect(parseFormulaPredicate('>10')?.test({ kind: 'number', value: 5 })).toBe(false);
    expect(parseFormulaPredicate('<>"foo"')?.test({ kind: 'text', value: 'bar' })).toBe(true);
    expect(parseFormulaPredicate('<>"foo"')?.test({ kind: 'text', value: 'foo' })).toBe(false);
    // `=A1>0` is reserved for engine evaluator (not implemented in v1) — null.
    expect(parseFormulaPredicate('=A1>0')).toBeNull();
    // Bare `=42` after stripping the leading `=` has no comparator — null.
    expect(parseFormulaPredicate('=42')).toBeNull();
    // `==42` after stripping the leading `=` becomes `=42` which matches.
    expect(parseFormulaPredicate('==42')?.test({ kind: 'number', value: 42 })).toBe(true);
  });

  it('formula rules evaluate A1 references relative to the rule range anchor', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 1);
    s = seedNumber(s, 1, 0, -1);
    s = withConditionalRules(s, [
      formulaRule({ sheet: 0, r0: 0, c0: 1, r1: 1, c1: 1 }, '=A1>0', '#ref'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#ref');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
  });

  it('formula rules honor absolute rows and columns in A1 references', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 1);
    s = withConditionalRules(s, [
      formulaRule({ sheet: 0, r0: 0, c0: 1, r1: 1, c1: 1 }, '=$A$1>0', '#abs'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#abs');
    expect(overlay.get('0:1:1')?.fill).toBe('#abs');
  });

  it('formula rules combine simple comparisons with AND/OR/NOT', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 7);
    s = seedNumber(s, 0, 1, 1);
    s = seedNumber(s, 1, 0, 3);
    s = seedNumber(s, 1, 1, 1);
    s = seedNumber(s, 2, 0, 7);
    s = seedNumber(s, 2, 1, 9);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 2, c1: 2 },
        '=AND(A1>5,OR(B1=1,NOT(A1<5)))',
        '#logic',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#logic');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
    expect(overlay.get('0:2:2')?.fill).toBe('#logic');
  });

  it('formula rules combine simple comparisons with XOR', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 7);
    s = seedNumber(s, 0, 1, 1);
    s = seedNumber(s, 1, 0, 7);
    s = seedNumber(s, 1, 1, 9);
    s = seedNumber(s, 2, 0, 3);
    s = seedNumber(s, 2, 1, 9);
    s = withConditionalRules(s, [
      formulaRule({ sheet: 0, r0: 0, c0: 2, r1: 2, c1: 2 }, '=XOR(A1>5,B1>5,FALSE)', '#xor'),
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=AND(A1>5,B1>0)=TRUE()',
        '#logical-comparison',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=N(XOR(A1>5,B1>5,FALSE))=1',
        '#logical-coerce',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#xor');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
    expect(overlay.get('0:2:2')?.fill).toBe('#xor');
    expect(overlay.get('0:0:3')?.fill).toBe('#logical-comparison');
    expect(overlay.get('0:0:4')?.fill).toBe('#logical-coerce');
  });

  it('formula rules evaluate simple arithmetic expressions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 4);
    s = seedNumber(s, 0, 1, 3);
    s = seedNumber(s, 1, 0, 2);
    s = seedNumber(s, 1, 1, 3);
    s = withConditionalRules(s, [
      formulaRule({ sheet: 0, r0: 0, c0: 2, r1: 1, c1: 2 }, '=A1+B1*2>8', '#math'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#math');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate arithmetic over aggregate operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 1);
    s = seedNumber(s, 1, 0, 2);
    s = seedNumber(s, 2, 0, 3);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=SUM($A$1:$A$3)/COUNT($A$1:$A$3)=2',
        '#math-agg',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#math-agg');
  });

  it('formula arithmetic keeps negative numeric literals attached to the operand', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 4);
    s = withConditionalRules(s, [
      formulaRule({ sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }, '=A1*-1=-4', '#negative'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#negative');
  });

  it('formula rules evaluate exponent arithmetic', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 3);
    s = withConditionalRules(s, [
      formulaRule({ sheet: 0, r0: 0, c0: 1, r1: 0, c1: 2 }, '=$A$1^2=9', '#pow'),
      formulaRule({ sheet: 0, r0: 1, c0: 1, r1: 1, c1: 2 }, '=2^3^2=512', '#pow-right'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#pow');
    expect(overlay.get('0:0:2')?.fill).toBe('#pow');
    expect(overlay.get('0:1:1')?.fill).toBe('#pow-right');
    expect(overlay.get('0:1:2')?.fill).toBe('#pow-right');
  });

  it('formula rules accept explicit current-sheet A1 references', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 1);
    s = seedNumber(s, 1, 0, -1);
    s = seedNumber(s, 2, 0, 3);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 1, c1: 2 },
        "=AND(Sheet1!A1>0,SUM('Sheet1'!$A$1:$A$3)>2)",
        '#sheet',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#sheet');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate IF with boolean branches', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 7);
    s = seedNumber(s, 0, 1, 3);
    s = seedNumber(s, 1, 0, 7);
    s = seedNumber(s, 1, 1, 1);
    s = seedNumber(s, 2, 0, 2);
    s = seedNumber(s, 2, 1, 9);
    s = withConditionalRules(s, [
      formulaRule({ sheet: 0, r0: 0, c0: 2, r1: 2, c1: 2 }, '=IF(A1>5,B1>2,FALSE)', '#if'),
      formulaRule({ sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 }, '=TRUE()', '#true-function'),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=IF(A1>5,,TRUE)',
        '#if-omitted-boolean',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#if');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
    expect(overlay.get('0:2:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBe('#true-function');
    expect(overlay.get('0:0:4')?.fill).toBeUndefined();
  });

  it('formula rules evaluate ISBLANK/ISERROR/ISNUMBER/ISTEXT predicates', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'label' });
    s = seedNumber(s, 1, 0, 42);
    s = seedCell(s, 2, 0, { kind: 'error', code: 1, text: '#DIV/0!' });
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 3, c1: 1 },
        '=OR(ISTEXT(A1),ISNUMBER(A1),ISERROR(A1),ISBLANK(A1))',
        '#is',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#is');
    expect(overlay.get('0:1:1')?.fill).toBe('#is');
    expect(overlay.get('0:2:1')?.fill).toBe('#is');
    expect(overlay.get('0:3:1')?.fill).toBe('#is');
  });

  it('formula rules leave unsupported sheet-qualified references unapplied', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 1);
    s = withConditionalRules(s, [
      formulaRule({ sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }, '=Sheet2!A1>0', '#sheet2'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
  });
});
