import { afterEach, describe, expect, it, vi } from 'vitest';
import { _resetConditionalCache, evaluateConditional } from '../../../../src/render/conditional.js';
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

  it('formula rules evaluate ISEVEN/ISODD boolean operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = withConditionalRules(s, [
      formulaRule({ sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 }, '=ISEVEN(ROW())', '#is-even-row'),
      formulaRule({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 }, '=ISODD(COLUMN())', '#is-odd-column'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#is-odd-column');
    expect(overlay.get('0:1:0')?.fill).toBe('#is-even-row');
    expect(overlay.get('0:2:0')?.fill).toBeUndefined();
    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBe('#is-odd-column');
  });

  it('formula rules evaluate ABS/MOD/ROUND numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, -3.4);
    s = seedNumber(s, 1, 0, 125);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=AND(ABS(A1)>3,MOD(A1,2)>0.5,ROUND(A1,0)=-3)',
        '#numeric-fns',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
        '=ROUND(A2,-1)=130',
        '#round-negative-digits',
      ),
      formulaRule({ sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 }, '=MOD(A1,0)=0', '#mod-zero'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#numeric-fns');
    expect(overlay.get('0:1:1')?.fill).toBe('#round-negative-digits');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate INT/TRUNC/SQRT/POWER/SIGN numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, -3.75);
    s = seedNumber(s, 1, 0, 125.987);
    s = seedNumber(s, 2, 0, 16);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=AND(INT(A1)=-4,TRUNC(A1)=-3,SIGN(A1)=-1)',
        '#int-trunc-sign',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
        '=AND(TRUNC(A2,1)=125.9,TRUNC(A2,-1)=120)',
        '#trunc-digits',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 3, r1: 1, c1: 3 },
        '=TRUNC(A2,)=125',
        '#trunc-omitted-digits',
      ),
      formulaRule(
        { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
        '=AND(SQRT(A3)=4,POWER(A3,2)=256)',
        '#sqrt-power',
      ),
      formulaRule({ sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 }, '=SQRT(A1)=0', '#sqrt-negative'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#int-trunc-sign');
    expect(overlay.get('0:1:1')?.fill).toBe('#trunc-digits');
    expect(overlay.get('0:1:3')?.fill).toBe('#trunc-omitted-digits');
    expect(overlay.get('0:2:1')?.fill).toBe('#sqrt-power');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate ROUNDUP/ROUNDDOWN numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, -3.275);
    s = seedNumber(s, 1, 0, 125.987);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=AND(ROUNDUP(A1,2)=-3.28,ROUNDDOWN(A1,2)=-3.27)',
        '#roundup-rounddown',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
        '=AND(ROUNDUP(A2,-1)=130,ROUNDDOWN(A2,-1)=120)',
        '#roundup-rounddown-negative-digits',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#roundup-rounddown');
    expect(overlay.get('0:1:1')?.fill).toBe('#roundup-rounddown-negative-digits');
  });

  it('formula rules evaluate EVEN/ODD numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 3.2);
    s = seedNumber(s, 1, 0, -3.2);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=AND(EVEN(A1)=4,ODD(A1)=5)',
        '#even-odd-positive',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
        '=AND(EVEN(A2)=-4,ODD(A2)=-5)',
        '#even-odd-negative',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#even-odd-positive');
    expect(overlay.get('0:1:1')?.fill).toBe('#even-odd-negative');
  });

  it('formula rules evaluate CEILING.MATH/FLOOR.MATH numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 4.3);
    s = seedNumber(s, 1, 0, -4.3);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=AND(CEILING.MATH(A1,2)=6,FLOOR.MATH(A1,2)=4)',
        '#math-round-positive',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
        '=AND(CEILING.MATH(A2,2)=-4,FLOOR.MATH(A2,2)=-6)',
        '#math-round-negative-default',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
        '=AND(CEILING.MATH(A2,2,1)=-6,FLOOR.MATH(A2,2,1)=-4)',
        '#math-round-negative-mode',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=AND(CEILING.MATH(A1,)=5,FLOOR.MATH(A1,)=4)',
        '#math-round-omitted-significance',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 3, r1: 1, c1: 3 },
        '=AND(CEILING.MATH(A2,,1)=-5,FLOOR.MATH(A2,,1)=-4)',
        '#math-round-omitted-significance-mode',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#math-round-positive');
    expect(overlay.get('0:1:1')?.fill).toBe('#math-round-negative-default');
    expect(overlay.get('0:1:2')?.fill).toBe('#math-round-negative-mode');
    expect(overlay.get('0:0:3')?.fill).toBe('#math-round-omitted-significance');
    expect(overlay.get('0:1:3')?.fill).toBe('#math-round-omitted-significance-mode');
  });

  it('formula rules evaluate CEILING.PRECISE/FLOOR.PRECISE numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 4.3);
    s = seedNumber(s, 1, 0, -4.3);
    s = seedCell(s, 2, 0, { kind: 'text', value: 'north' });
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=AND(CEILING.PRECISE(A1,2)=6,FLOOR.PRECISE(A1,2)=4)',
        '#precise-round-positive',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
        '=AND(CEILING.PRECISE(A2,2)=-4,FLOOR.PRECISE(A2,2)=-6)',
        '#precise-round-negative',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=AND(CEILING.PRECISE(A1,)=5,FLOOR.PRECISE(A1,)=4,ISO.CEILING(A1)=5)',
        '#precise-round-omitted-significance',
      ),
      formulaRule(
        { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
        '=CEILING.PRECISE(A3)=0',
        '#precise-round-text',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#precise-round-positive');
    expect(overlay.get('0:1:1')?.fill).toBe('#precise-round-negative');
    expect(overlay.get('0:0:2')?.fill).toBe('#precise-round-omitted-significance');
    expect(overlay.get('0:2:1')?.fill).toBeUndefined();
  });

  it('formula rules evaluate legacy CEILING/FLOOR numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 4.3);
    s = seedNumber(s, 1, 0, -4.3);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=AND(CEILING(A1,2)=6,FLOOR(A1,2)=4)',
        '#legacy-round-positive',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
        '=AND(CEILING(A2,-2)=-6,FLOOR(A2,-2)=-4)',
        '#legacy-round-negative',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
        '=CEILING(A2,2)=0',
        '#legacy-round-sign-mismatch',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#legacy-round-positive');
    expect(overlay.get('0:1:1')?.fill).toBe('#legacy-round-negative');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate MROUND numeric function', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 10);
    s = seedNumber(s, 1, 0, -10);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=AND(MROUND(A1,3)=9,MROUND(10.5,3)=12)',
        '#mround-positive',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
        '=MROUND(A2,-3)=-9',
        '#mround-negative',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
        '=MROUND(A2,3)=0',
        '#mround-sign-mismatch',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#mround-positive');
    expect(overlay.get('0:1:1')?.fill).toBe('#mround-negative');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate GCD/LCM numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 24);
    s = seedNumber(s, 0, 1, 36);
    s = seedNumber(s, 0, 2, 54);
    s = seedNumber(s, 1, 0, -12);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=AND(GCD(A1,B1,C1)=6,LCM(4,6,8)=24)',
        '#gcd-lcm',
      ),
      formulaRule({ sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 }, '=LCM(A1,0)=0', '#lcm-zero'),
      formulaRule({ sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 }, '=GCD(A2,6)=6', '#gcd-negative'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#gcd-lcm');
    expect(overlay.get('0:0:4')?.fill).toBe('#lcm-zero');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
  });

  it('formula rules evaluate factorial, gamma, and combinatoric numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        '=AND(FACT(5)=120,FACTDOUBLE(6)=48,FACT(5.9)=120)',
        '#factorial',
      ),
      {
        kind: 'formula',
        range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        formula:
          '=AND(COMBIN(10,3)=120,COMBINA(3,2)=6,PERMUT(10,3)=720,PERMUTATIONA(4,3)=64,MULTINOMIAL(2,3,4)=1260)',
        apply: { fill: '#combinatoric' },
      },
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=AND(ROUND(GAMMALN(4),6)=1.791759,ROUND(GAMMALN.PRECISE(4),6)=1.791759)',
        '#gammaln',
      ),
      {
        kind: 'formula',
        range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
        formula:
          '=AND(ROUND(GAMMA(5),6)=24,ROUND(GAMMA(0.5),6)=1.772454,ROUND(GAMMA(-0.5),6)=-3.544908)',
        apply: { fill: '#gamma' },
      },
      formulaRule({ sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 }, '=FACT(-1)=0', '#fact-negative'),
      formulaRule({ sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 }, '=COMBIN(2,3)=0', '#combin-invalid'),
      formulaRule({ sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 }, '=GAMMALN(0)=0', '#gammaln-invalid'),
      formulaRule(
        { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
        '=MULTINOMIAL(2,-1)=0',
        '#multinomial-invalid',
      ),
      formulaRule({ sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 }, '=GAMMA(0)=0', '#gamma-invalid'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#factorial');
    expect(overlay.get('0:0:1')?.fill).toBe('#combinatoric');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
    expect(overlay.get('0:0:4')?.fill).toBe('#gammaln');
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
    expect(overlay.get('0:0:6')?.fill).toBeUndefined();
    expect(overlay.get('0:0:7')?.fill).toBe('#gamma');
    expect(overlay.get('0:0:8')?.fill).toBeUndefined();
  });

  it('formula rules evaluate QUOTIENT numeric function', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 17);
    s = seedNumber(s, 1, 0, -17);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=QUOTIENT(A1,5)=3',
        '#quotient-positive',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
        '=QUOTIENT(A2,5)=-3',
        '#quotient-negative',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=QUOTIENT(A1,0)=0',
        '#quotient-zero-divisor',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#quotient-positive');
    expect(overlay.get('0:1:1')?.fill).toBe('#quotient-negative');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate trigonometric numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 30);
    s = seedNumber(s, 1, 0, 45);
    s = seedCell(s, 2, 0, { kind: 'text', value: 'north' });
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=AND(ROUND(SIN(RADIANS(A1)),6)=0.5,ROUND(COS(RADIANS(60)),6)=0.5)',
        '#trig-sin-cos',
      ),
      {
        kind: 'formula',
        range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
        formula:
          '=AND(ROUND(TAN(RADIANS(A2)),6)=1,ROUND(DEGREES(PI()),6)=180,ROUND(DEGREES(ATAN2(0,-1)),6)=-90)',
        apply: { fill: '#trig-tan-pi' },
      },
      formulaRule({ sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 }, '=ATAN2(0,0)=0', '#atan2-zero-origin'),
      {
        kind: 'formula',
        range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        formula:
          '=AND(ROUND(SEC(RADIANS(60)),6)=2,ROUND(CSC(RADIANS(30)),6)=2,ROUND(COT(RADIANS(45)),6)=1)',
        apply: { fill: '#reciprocal-trig' },
      },
      formulaRule({ sheet: 0, r0: 1, c0: 3, r1: 1, c1: 3 }, '=COT(0)=0', '#cot-zero'),
      formulaRule({ sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 }, '=SIN(A3)=0', '#trig-text'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#trig-sin-cos');
    expect(overlay.get('0:1:1')?.fill).toBe('#trig-tan-pi');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBe('#reciprocal-trig');
    expect(overlay.get('0:1:3')?.fill).toBeUndefined();
    expect(overlay.get('0:2:1')?.fill).toBeUndefined();
  });

  it('formula rules evaluate inverse and hyperbolic numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 0.5);
    s = seedNumber(s, 1, 0, 2);
    s = seedNumber(s, 2, 0, 4);
    s = withConditionalRules(s, [
      {
        kind: 'formula',
        range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        formula:
          '=AND(ROUND(DEGREES(ASIN(A1)),6)=30,ROUND(DEGREES(ACOS(A1)),6)=60,ROUND(DEGREES(ATAN(1)),6)=45,ROUND(DEGREES(ACOT(1)),6)=45)',
        apply: { fill: '#inverse-trig' },
      },
      {
        kind: 'formula',
        range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
        formula:
          '=AND(ROUND(SINH(1),6)=1.175201,ROUND(COSH(1),6)=1.543081,ROUND(TANH(1),6)=0.761594,ROUND(COTH(1),6)=1.313035,ROUND(SECH(1),6)=0.648054,ROUND(CSCH(1),6)=0.850918,ROUND(ASINH(1),6)=0.881374,ROUND(ACOSH(2),6)=1.316958,ROUND(ATANH(0.5),6)=0.549306,ROUND(ACOTH(2),6)=0.549306)',
        apply: { fill: '#hyperbolic' },
      },
      formulaRule(
        { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
        '=ROUND(SQRTPI(A3),6)=3.544908',
        '#sqrtpi',
      ),
      formulaRule({ sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 }, '=ASIN(A2)=0', '#asin-out-of-range'),
      formulaRule({ sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 }, '=ACOSH(0)=0', '#acosh-out-of-range'),
      formulaRule({ sheet: 0, r0: 2, c0: 3, r1: 2, c1: 3 }, '=COTH(0)=0', '#coth-zero'),
      formulaRule({ sheet: 0, r0: 2, c0: 4, r1: 2, c1: 4 }, '=ACOTH(0.5)=0', '#acoth-out-of-range'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#inverse-trig');
    expect(overlay.get('0:1:1')?.fill).toBe('#hyperbolic');
    expect(overlay.get('0:2:1')?.fill).toBe('#sqrtpi');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
    expect(overlay.get('0:2:2')?.fill).toBeUndefined();
    expect(overlay.get('0:2:3')?.fill).toBeUndefined();
    expect(overlay.get('0:2:4')?.fill).toBeUndefined();
  });

  it('formula rules evaluate SUMSQ and SERIESSUM numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 3);
    s = seedNumber(s, 0, 1, 4);
    s = seedCell(s, 1, 0, { kind: 'text', value: 'north' });
    s = seedNumber(s, 2, 0, 1);
    s = seedNumber(s, 2, 1, 2);
    s = seedNumber(s, 2, 2, 3);
    s = withConditionalRules(s, [
      formulaRule({ sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 }, '=SUMSQ(A1,B1,12)=169', '#sumsq'),
      formulaRule({ sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 }, '=SUMSQ(A2,1)=1', '#sumsq-text'),
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=SERIESSUM(2,1,1,$A$3:$C$3)=34',
        '#series-sum',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=SERIESSUM(2,0,2,3)=3',
        '#series-sum-scalar',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=SERIESSUM(2,1,1,A2)=0',
        '#series-sum-text',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#sumsq');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBe('#series-sum');
    expect(overlay.get('0:0:4')?.fill).toBe('#series-sum-scalar');
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
  });

  it('formula rules evaluate logarithmic numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 7);
    s = seedNumber(s, 1, 0, 8);
    s = seedNumber(s, 2, 0, -1);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=AND(ROUND(EXP(LN(A1)),6)=7,ROUND(LOG10(1000),6)=3)',
        '#exp-ln-log10',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
        '=AND(ROUND(LOG(A2,2),6)=3,ROUND(LOG(100),6)=2)',
        '#log',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=ROUND(LOG(100,),6)=2',
        '#log-omitted-base',
      ),
      formulaRule({ sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 }, '=LN(A3)=0', '#ln-negative'),
      formulaRule({ sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 }, '=LOG(A2,1)=0', '#log-base-one'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#exp-ln-log10');
    expect(overlay.get('0:1:1')?.fill).toBe('#log');
    expect(overlay.get('0:0:3')?.fill).toBe('#log-omitted-base');
    expect(overlay.get('0:2:1')?.fill).toBeUndefined();
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate error function numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'north' });
    s = withConditionalRules(s, [
      {
        kind: 'formula',
        range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        formula:
          '=AND(ROUND(ERF(1),6)=0.842701,ROUND(ERF(0,1),6)=0.842701,ROUND(ERF.PRECISE(1),6)=0.842701)',
        apply: { fill: '#erf' },
      },
      {
        kind: 'formula',
        range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        formula:
          '=AND(ROUND(ERFC(1),6)=0.157299,ROUND(ERFC.PRECISE(1),6)=0.157299,ROUND(GAUSS(1),6)=0.341345)',
        apply: { fill: '#erfc-gauss' },
      },
      formulaRule({ sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 }, '=ERF(A1)=0', '#erf-text'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#erf');
    expect(overlay.get('0:0:2')?.fill).toBe('#erfc-gauss');
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
  });

  it('formula rules evaluate base conversion numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = withConditionalRules(s, [
      {
        kind: 'formula',
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        formula:
          '=AND(BASE(255,16)="FF",BASE(5,2,8)="00000101",DECIMAL("FF",16)=255,BIN2DEC("1111111111")=-1,DEC2BIN(5,8)="00000101",DEC2BIN(-1)="1111111111",HEX2DEC("FFFFFFFFFF")=-1,DEC2HEX(255,4)="00FF",DEC2HEX(-1)="FFFFFFFFFF",OCT2DEC("7777777777")=-1,DEC2OCT(64,4)="0100",DEC2OCT(-1)="7777777777",BIN2HEX("1111",4)="000F",HEX2BIN("F",8)="00001111",BIN2OCT("1111",4)="0017",OCT2BIN("17",8)="00001111",HEX2OCT("F",4)="0017",OCT2HEX("17",4)="000F")',
        apply: { fill: '#base-decimal' },
      },
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=BASE(10,1)="10"',
        '#base-invalid-radix',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=DECIMAL("2",2)=2',
        '#decimal-invalid-digit',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=BIN2DEC("10000000000")=0',
        '#bin2dec-too-long',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=DEC2BIN(512)="1000000000"',
        '#dec2bin-out-of-range',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=HEX2DEC("G")=16',
        '#hex2dec-invalid-digit',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
        '=DEC2HEX(549755813888)="8000000000"',
        '#dec2hex-out-of-range',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
        '=OCT2DEC("8")=8',
        '#oct2dec-invalid-digit',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
        '=DEC2OCT(536870912)="4000000000"',
        '#dec2oct-out-of-range',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
        '=HEX2BIN("200")="1000000000"',
        '#hex2bin-out-of-range',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 10, r1: 0, c1: 10 },
        '=BIN2HEX("10000000000")="400"',
        '#bin2hex-too-long',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 11, r1: 0, c1: 11 },
        '=OCT2HEX("8")="8"',
        '#oct2hex-invalid-digit',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#base-decimal');
    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
    expect(overlay.get('0:0:4')?.fill).toBeUndefined();
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
    expect(overlay.get('0:0:6')?.fill).toBeUndefined();
    expect(overlay.get('0:0:7')?.fill).toBeUndefined();
    expect(overlay.get('0:0:8')?.fill).toBeUndefined();
    expect(overlay.get('0:0:9')?.fill).toBeUndefined();
    expect(overlay.get('0:0:10')?.fill).toBeUndefined();
    expect(overlay.get('0:0:11')?.fill).toBeUndefined();
  });

  it('formula rules evaluate roman numeral numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        '=AND(ROMAN(1999)="MCMXCIX",ROMAN(944,0)="CMXLIV",ARABIC("MCMXCIX")=1999)',
        '#roman-arabic',
      ),
      formulaRule({ sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }, '=ROMAN(0)="N"', '#roman-zero'),
      formulaRule({ sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 }, '=ARABIC("IIII")=4', '#arabic-invalid'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#roman-arabic');
    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate threshold numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'north' });
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=AND(DELTA(4,4)=1,DELTA(4,5)=0,DELTA(0)=1,DELTA(0,)=1)',
        '#delta',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=AND(GESTEP(5,4)=1,GESTEP(4,5)=0,GESTEP(0)=1,GESTEP(0,)=1)',
        '#gestep',
      ),
      formulaRule({ sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 }, '=DELTA(A1,0)=0', '#delta-text'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#delta');
    expect(overlay.get('0:0:2')?.fill).toBe('#gestep');
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
  });

  it('formula rules evaluate bitwise numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        '=AND(BITAND(13,7)=5,BITOR(9,6)=15,BITXOR(9,6)=15)',
        '#bitwise',
      ),
      {
        kind: 'formula',
        range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        formula: '=AND(BITLSHIFT(3,4)=48,BITRSHIFT(48,4)=3,BITLSHIFT(48,-4)=3,BITRSHIFT(3,-4)=48)',
        apply: { fill: '#bitshift' },
      },
      formulaRule({ sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 }, '=BITAND(-1,1)=1', '#bit-negative'),
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=BITLSHIFT(1,54)=0',
        '#bit-shift-too-large',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#bitwise');
    expect(overlay.get('0:0:1')?.fill).toBe('#bitshift');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
  });
});
