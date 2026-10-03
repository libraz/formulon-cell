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

  it('formula rules evaluate aggregate ranges relative to the rule range anchor', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 1);
    s = seedNumber(s, 1, 0, 2);
    s = seedNumber(s, 2, 0, 3);
    s = withConditionalRules(s, [
      formulaRule({ sheet: 0, r0: 0, c0: 2, r1: 1, c1: 2 }, '=SUM(A1:A3)>5', '#sum'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#sum');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
  });

  it('formula rules support absolute aggregate ranges and aggregate operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 1);
    s = seedNumber(s, 1, 0, 2);
    s = seedNumber(s, 2, 0, 3);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 1, c1: 2 },
        '=AND(AVERAGE($A$1:$A$3)=2,MAX($A$1:$A$3)=COUNT($A$1:$A$3))',
        '#agg',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#agg');
    expect(overlay.get('0:1:2')?.fill).toBe('#agg');
  });

  it('formula rules evaluate COUNTA/COUNTBLANK/PRODUCT aggregate ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 2);
    s = seedCell(s, 1, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 2, 0, { kind: 'bool', value: true });
    s = seedCell(s, 4, 0, { kind: 'error', code: 1, text: '#DIV/0!' });
    s = seedNumber(s, 0, 1, 2);
    s = seedNumber(s, 1, 1, 3);
    s = seedNumber(s, 2, 1, 4);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=AND(COUNTA($A$1:$A$5)=4,COUNTBLANK($A$1:$A$5)=1)',
        '#counta-countblank',
      ),
      formulaRule({ sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 }, '=PRODUCT($B$1:$B$3)=24', '#product'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#counta-countblank');
    expect(overlay.get('0:0:3')?.fill).toBe('#product');
  });

  it('formula rules evaluate MEDIAN aggregate ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 12);
    s = seedNumber(s, 1, 0, 8);
    s = seedNumber(s, 2, 0, 15);
    s = seedCell(s, 3, 0, { kind: 'text', value: 'ignored' });
    s = seedNumber(s, 4, 0, 4);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=MEDIAN($A$1:$A$5)=10',
        '#median-even',
      ),
      formulaRule({ sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 }, '=MEDIAN($A$1:$A$3)=12', '#median-odd'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#median-even');
    expect(overlay.get('0:0:2')?.fill).toBe('#median-odd');
  });

  it('formula rules evaluate MIN over ranges and mixed arguments', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 12);
    s = seedNumber(s, 1, 0, 8);
    s = seedCell(s, 2, 0, { kind: 'text', value: 'ignored' });
    s = seedNumber(s, 4, 0, 4);
    s = withConditionalRules(s, [
      formulaRule({ sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }, '=MIN($A$1:$A$5)=4', '#min-range'),
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=MIN($A$1:$A$2,10,-3)=-3',
        '#min-args',
      ),
      formulaRule({ sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 }, '=MIN($A$1:$A$2)=12', '#min-miss'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#min-range');
    expect(overlay.get('0:0:2')?.fill).toBe('#min-args');
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
  });

  it('formula rules evaluate multi-argument aggregate operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 2);
    s = seedNumber(s, 0, 1, 4);
    s = seedNumber(s, 0, 2, 6);
    s = seedNumber(s, 1, 0, 8);
    s = seedCell(s, 1, 1, { kind: 'text', value: 'ignored' });
    s = seedCell(s, 1, 2, { kind: 'bool', value: true });
    s = seedNumber(s, 3, 0, 1);
    s = seedNumber(s, 3, 1, 2);
    s = seedNumber(s, 3, 2, 3);
    s = seedNumber(s, 3, 3, 2);
    s = seedNumber(s, 3, 4, 4);
    s = seedNumber(s, 3, 5, 7);
    s = seedNumber(s, 4, 0, 1);
    s = seedNumber(s, 4, 1, 2);
    s = seedNumber(s, 4, 2, 3);
    s = seedNumber(s, 4, 3, 4);
    s = seedNumber(s, 5, 0, 0.1);
    s = seedNumber(s, 5, 1, 0.2);
    s = seedNumber(s, 5, 2, 0.3);
    s = seedNumber(s, 5, 3, 0.4);
    s = seedNumber(s, 6, 0, 10);
    s = seedNumber(s, 6, 1, 20);
    s = seedNumber(s, 7, 0, 20);
    s = seedNumber(s, 7, 1, 40);
    s = seedNumber(s, 6, 3, 15);
    s = seedNumber(s, 6, 4, 15);
    s = seedNumber(s, 7, 3, 15);
    s = seedNumber(s, 7, 4, 45);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=AND(SUM(A1,B1,$C$1:$C$1)=12,PRODUCT(A1,B1,3)=24)',
        '#multi-sum-product',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=AND(AVERAGE($A$1:$C$1,A2)=5,MEDIAN($A$1:$C$1,A2)=5)',
        '#multi-average-median',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=AND(AVERAGEA($A$1:$C$2)=3.5,MINA($A$1:$C$2)=0,MAXA($A$1:$C$2)=8)',
        '#aggregate-a',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
        '=AND(COUNT(A1,B1,B2)=2,COUNTA(A1,B2,C3)=2,COUNTBLANK(B2,C3)=1)',
        '#multi-counts',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 3, r1: 1, c1: 3 },
        '=AND(ROUND(STDEV.S($A$1:$C$1),6)=2,ROUND(STDEV.P($A$1:$C$1),6)=1.632993)',
        '#stdev-range',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 4, r1: 1, c1: 4 },
        '=AND(VAR.S(A1,B1,C1)=4,ROUND(VAR.P(A1,B1,C1),6)=2.666667)',
        '#var-args',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
        '=AND(ROUND(STDEV($A$1:$C$1),6)=2,VAR(A1,B1,C1)=4)',
        '#legacy-stdev-var',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
        '=AND(ROUND(STDEVP($A$1:$C$1),6)=1.632993,ROUND(VARP(A1,B1,C1),6)=2.666667)',
        '#legacy-stdevp-varp',
      ),
      formulaRule({ sheet: 0, r0: 1, c0: 5, r1: 1, c1: 5 }, '=STDEV.S(A1)=0', '#stdev-s-single'),
      {
        kind: 'formula',
        range: { sheet: 0, r0: 1, c0: 6, r1: 1, c1: 6 },
        formula: '=AND(ROUND(GEOMEAN($A$1:$C$1),6)=3.634241,ROUND(HARMEAN(A1,B1,C1),6)=3.272727)',
        apply: { fill: '#geo-harmonic' },
      },
      formulaRule(
        { sheet: 0, r0: 1, c0: 7, r1: 1, c1: 7 },
        '=GEOMEAN(A1,-1)=0',
        '#geomean-negative',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 8, r1: 1, c1: 8 },
        '=AND(DEVSQ(A1,B1,C1)=8,DEVSQ($A$1:$C$1,A2)=20)',
        '#devsq',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 9, r1: 1, c1: 9 },
        '=AND(ROUND(AVEDEV(A1,B1,C1),6)=1.333333,AVEDEV($A$1:$C$1,A2)=2)',
        '#avedev',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 10, r1: 1, c1: 10 },
        '=ROUND(SKEW(A1,B1,C1,20),6)=1.763633',
        '#skew',
      ),
      formulaRule({ sheet: 0, r0: 1, c0: 11, r1: 1, c1: 11 }, '=SKEW(A1,B1)=0', '#skew-too-few'),
      {
        kind: 'formula',
        range: { sheet: 0, r0: 1, c0: 32, r1: 1, c1: 32 },
        formula:
          '=AND(ROUND(SKEW.P(A1,B1,C1,20),6)=1.018234,ROUND(SKEW.P($A$1:$C$1,20),6)=1.018234)',
        apply: { fill: '#skew-p' },
      },
      formulaRule(
        { sheet: 0, r0: 1, c0: 33, r1: 1, c1: 33 },
        '=SKEW.P(A1,B1)=0',
        '#skew-p-too-few',
      ),
      {
        kind: 'formula',
        range: { sheet: 0, r0: 1, c0: 34, r1: 1, c1: 34 },
        formula:
          '=AND(ROUND(Z.TEST($A$1:$C$1,3,2),6)=0.193238,ROUND(Z.TEST($A$1:$C$1,3),6)=0.193238,ROUND(ZTEST($A$1:$C$1,3,2),6)=0.193238)',
        apply: { fill: '#z-test' },
      },
      formulaRule(
        { sheet: 0, r0: 1, c0: 35, r1: 1, c1: 35 },
        '=Z.TEST($A$1:$C$1,3,0)=0',
        '#z-test-invalid',
      ),
      {
        kind: 'formula',
        range: { sheet: 0, r0: 1, c0: 36, r1: 1, c1: 36 },
        formula:
          '=AND(ROUND(F.TEST($A$4:$C$4,$D$4:$F$4),6)=0.272727,ROUND(FTEST($A$4:$C$4,$D$4:$F$4),6)=0.272727)',
        apply: { fill: '#f-test' },
      },
      formulaRule(
        { sheet: 0, r0: 1, c0: 37, r1: 1, c1: 37 },
        '=F.TEST($A$1:$A$1,$D$4:$F$4)=0',
        '#f-test-invalid',
      ),
      {
        kind: 'formula',
        range: { sheet: 0, r0: 1, c0: 38, r1: 1, c1: 38 },
        formula:
          '=AND(ROUND(T.TEST($A$4:$C$4,$D$4:$F$4,1,1),6)=0.059041,ROUND(T.TEST($A$4:$C$4,$D$4:$F$4,2,1),6)=0.118083)',
        apply: { fill: '#t-test-paired' },
      },
      {
        kind: 'formula',
        range: { sheet: 0, r0: 1, c0: 39, r1: 1, c1: 39 },
        formula:
          '=AND(ROUND(T.TEST($A$4:$C$4,$D$4:$F$4,2,2),6)=0.209875,ROUND(TTEST($A$4:$C$4,$D$4:$F$4,2,3),6)=0.245113)',
        apply: { fill: '#t-test-independent' },
      },
      formulaRule(
        { sheet: 0, r0: 1, c0: 40, r1: 1, c1: 40 },
        '=T.TEST($A$4:$C$4,$D$4:$F$4,3,2)=0',
        '#t-test-invalid',
      ),
      {
        kind: 'formula',
        range: { sheet: 0, r0: 1, c0: 41, r1: 1, c1: 41 },
        formula:
          '=AND(ROUND(CHISQ.TEST($A$7:$B$8,$D$7:$E$8),6)=ROUND(CHISQ.DIST.RT(5.555555555555556,1),6),ROUND(CHITEST($A$7:$B$8,$D$7:$E$8),6)=ROUND(CHISQ.DIST.RT(5.555555555555556,1),6))',
        apply: { fill: '#chisq-test' },
      },
      formulaRule(
        { sheet: 0, r0: 1, c0: 42, r1: 1, c1: 42 },
        '=CHISQ.TEST($A$7:$B$8,$D$7:$F$8)=0',
        '#chisq-test-invalid',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 12, r1: 1, c1: 12 },
        '=ROUND(KURT(A1,B1,C1,20),6)=3.228',
        '#kurt',
      ),
      formulaRule({ sheet: 0, r0: 1, c0: 13, r1: 1, c1: 13 }, '=KURT(A1,B1,C1)=0', '#kurt-too-few'),
      formulaRule(
        { sheet: 0, r0: 1, c0: 14, r1: 1, c1: 14 },
        '=AND(MODE.SNGL(A1,B1,C1,B1)=4,MODE.SNGL($A$1:$C$1,4)=4)',
        '#mode-single',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 14, r1: 0, c1: 14 },
        '=AND(MODE(A1,B1,C1,B1)=4,MODE($A$1:$C$1,4)=4)',
        '#legacy-mode',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 15, r1: 1, c1: 15 },
        '=MODE.SNGL(A1,B1,C1)=2',
        '#mode-no-duplicate',
      ),
      {
        kind: 'formula',
        range: { sheet: 0, r0: 1, c0: 16, r1: 1, c1: 16 },
        formula:
          '=AND(ROUND(CORREL($A$4:$C$4,$D$4:$F$4),6)=0.993399,ROUND(COVARIANCE.P($A$4:$C$4,$D$4:$F$4),6)=1.666667,COVARIANCE.S($A$4:$C$4,$D$4:$F$4)=2.5)',
        apply: { fill: '#paired-stats' },
      },
      {
        kind: 'formula',
        range: { sheet: 0, r0: 0, c0: 16, r1: 0, c1: 16 },
        formula:
          '=AND(ROUND(PEARSON($A$4:$C$4,$D$4:$F$4),6)=0.993399,ROUND(COVAR($A$4:$C$4,$D$4:$F$4),6)=1.666667)',
        apply: { fill: '#legacy-paired-stats' },
      },
      {
        kind: 'formula',
        range: { sheet: 0, r0: 1, c0: 19, r1: 1, c1: 19 },
        formula:
          '=AND(SLOPE($D$4:$F$4,$A$4:$C$4)=2.5,ROUND(INTERCEPT($D$4:$F$4,$A$4:$C$4),6)=-0.666667,ROUND(RSQ($D$4:$F$4,$A$4:$C$4),6)=0.986842)',
        apply: { fill: '#regression-stats' },
      },
      {
        kind: 'formula',
        range: { sheet: 0, r0: 1, c0: 21, r1: 1, c1: 21 },
        formula:
          '=AND(ROUND(STEYX($D$4:$F$4,$A$4:$C$4),6)=0.408248,ROUND(FORECAST.LINEAR(4,$D$4:$F$4,$A$4:$C$4),6)=9.333333,ROUND(FORECAST(4,$D$4:$F$4,$A$4:$C$4),6)=9.333333)',
        apply: { fill: '#forecast-stats' },
      },
      formulaRule(
        { sheet: 0, r0: 1, c0: 17, r1: 1, c1: 17 },
        '=CORREL($A$4:$B$4,$D$4:$F$4)=1',
        '#paired-stats-mismatch',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 18, r1: 1, c1: 18 },
        '=CORREL($A$1:$C$1,$B$1:$B$3)=1',
        '#paired-stats-zero-variance',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 20, r1: 1, c1: 20 },
        '=SLOPE($A$1:$C$1,$B$1:$B$3)=1',
        '#regression-zero-x-variance',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 22, r1: 1, c1: 22 },
        '=FORECAST("x",$D$4:$F$4,$A$4:$C$4)=1',
        '#forecast-nonnumeric-x',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 23, r1: 1, c1: 23 },
        '=AND(PROB($A$5:$D$5,$A$6:$D$6,2,3)=0.5,PROB($A$5:$D$5,$A$6:$D$6,4)=0.4)',
        '#probability-range',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 24, r1: 1, c1: 24 },
        '=PROB($A$5:$D$5,$A$6:$C$6,2)=0.2',
        '#probability-mismatch',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 25, r1: 1, c1: 25 },
        '=PROB($A$5:$C$5,$A$6:$C$6,2)=0.2',
        '#probability-total-invalid',
      ),
      {
        kind: 'formula',
        range: { sheet: 0, r0: 1, c0: 26, r1: 1, c1: 26 },
        formula:
          '=AND(SUMX2MY2($A$4:$C$4,$D$4:$F$4)=-55,SUMX2PY2($A$4:$C$4,$D$4:$F$4)=83,SUMXMY2($A$4:$C$4,$D$4:$F$4)=21)',
        apply: { fill: '#sumx-paired' },
      },
      formulaRule(
        { sheet: 0, r0: 1, c0: 27, r1: 1, c1: 27 },
        '=SUMXMY2($A$4:$B$4,$D$4:$F$4)=1',
        '#sumx-mismatch',
      ),
      {
        kind: 'formula',
        range: { sheet: 0, r0: 1, c0: 28, r1: 1, c1: 28 },
        formula:
          '=AND(SUBTOTAL(9,$A$1:$C$1)=12,SUBTOTAL(101,$A$1:$C$1)=4,SUBTOTAL(103,$A$1:$C$2)=6,ROUND(SUBTOTAL(107,$A$1:$C$1),6)=2)',
        apply: { fill: '#subtotal' },
      },
      formulaRule(
        { sheet: 0, r0: 1, c0: 29, r1: 1, c1: 29 },
        '=SUBTOTAL(12,$A$1:$C$1)=0',
        '#subtotal-invalid',
      ),
      {
        kind: 'formula',
        range: { sheet: 0, r0: 1, c0: 30, r1: 1, c1: 30 },
        formula:
          '=AND(AGGREGATE(9,0,$A$1:$C$1)=12,AGGREGATE(1,0,$A$1:$C$1)=4,ROUND(AGGREGATE(7,0,$A$1:$C$1),6)=2)',
        apply: { fill: '#aggregate-function' },
      },
      formulaRule(
        { sheet: 0, r0: 1, c0: 31, r1: 1, c1: 31 },
        '=AGGREGATE(20,0,$A$1:$C$1)=0',
        '#aggregate-function-invalid',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#multi-sum-product');
    expect(overlay.get('0:0:4')?.fill).toBe('#multi-average-median');
    expect(overlay.get('0:0:5')?.fill).toBe('#aggregate-a');
    expect(overlay.get('0:1:2')?.fill).toBe('#multi-counts');
    expect(overlay.get('0:1:3')?.fill).toBe('#stdev-range');
    expect(overlay.get('0:1:4')?.fill).toBe('#var-args');
    expect(overlay.get('0:0:6')?.fill).toBe('#legacy-stdev-var');
    expect(overlay.get('0:0:7')?.fill).toBe('#legacy-stdevp-varp');
    expect(overlay.get('0:1:5')?.fill).toBeUndefined();
    expect(overlay.get('0:1:6')?.fill).toBe('#geo-harmonic');
    expect(overlay.get('0:1:7')?.fill).toBeUndefined();
    expect(overlay.get('0:1:8')?.fill).toBe('#devsq');
    expect(overlay.get('0:1:9')?.fill).toBe('#avedev');
    expect(overlay.get('0:1:10')?.fill).toBe('#skew');
    expect(overlay.get('0:1:11')?.fill).toBeUndefined();
    expect(overlay.get('0:1:12')?.fill).toBe('#kurt');
    expect(overlay.get('0:1:13')?.fill).toBeUndefined();
    expect(overlay.get('0:0:14')?.fill).toBe('#legacy-mode');
    expect(overlay.get('0:1:14')?.fill).toBe('#mode-single');
    expect(overlay.get('0:1:15')?.fill).toBeUndefined();
    expect(overlay.get('0:0:16')?.fill).toBe('#legacy-paired-stats');
    expect(overlay.get('0:1:16')?.fill).toBe('#paired-stats');
    expect(overlay.get('0:1:17')?.fill).toBeUndefined();
    expect(overlay.get('0:1:18')?.fill).toBeUndefined();
    expect(overlay.get('0:1:19')?.fill).toBe('#regression-stats');
    expect(overlay.get('0:1:20')?.fill).toBeUndefined();
    expect(overlay.get('0:1:21')?.fill).toBe('#forecast-stats');
    expect(overlay.get('0:1:22')?.fill).toBeUndefined();
    expect(overlay.get('0:1:23')?.fill).toBe('#probability-range');
    expect(overlay.get('0:1:24')?.fill).toBeUndefined();
    expect(overlay.get('0:1:25')?.fill).toBeUndefined();
    expect(overlay.get('0:1:26')?.fill).toBe('#sumx-paired');
    expect(overlay.get('0:1:27')?.fill).toBeUndefined();
    expect(overlay.get('0:1:28')?.fill).toBe('#subtotal');
    expect(overlay.get('0:1:29')?.fill).toBeUndefined();
    expect(overlay.get('0:1:30')?.fill).toBe('#aggregate-function');
    expect(overlay.get('0:1:31')?.fill).toBeUndefined();
    expect(overlay.get('0:1:32')?.fill).toBe('#skew-p');
    expect(overlay.get('0:1:33')?.fill).toBeUndefined();
    expect(overlay.get('0:1:34')?.fill).toBe('#z-test');
    expect(overlay.get('0:1:35')?.fill).toBeUndefined();
    expect(overlay.get('0:1:36')?.fill).toBe('#f-test');
    expect(overlay.get('0:1:37')?.fill).toBeUndefined();
    expect(overlay.get('0:1:38')?.fill).toBe('#t-test-paired');
    expect(overlay.get('0:1:39')?.fill).toBe('#t-test-independent');
    expect(overlay.get('0:1:40')?.fill).toBeUndefined();
    expect(overlay.get('0:1:41')?.fill).toBe('#chisq-test');
    expect(overlay.get('0:1:42')?.fill).toBeUndefined();
  });

  it('formula rules evaluate statistical functions over dynamic ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 2);
    s = seedNumber(s, 0, 1, 4);
    s = seedNumber(s, 0, 2, 6);
    s = seedNumber(s, 3, 0, 1);
    s = seedNumber(s, 3, 1, 2);
    s = seedNumber(s, 3, 2, 3);
    s = seedNumber(s, 3, 3, 2);
    s = seedNumber(s, 3, 4, 4);
    s = seedNumber(s, 3, 5, 7);
    s = seedNumber(s, 4, 0, 1);
    s = seedNumber(s, 4, 1, 2);
    s = seedNumber(s, 4, 2, 3);
    s = seedNumber(s, 4, 3, 4);
    s = seedNumber(s, 5, 0, 0.1);
    s = seedNumber(s, 5, 1, 0.2);
    s = seedNumber(s, 5, 2, 0.3);
    s = seedNumber(s, 5, 3, 0.4);
    s = seedNumber(s, 6, 0, 10);
    s = seedNumber(s, 6, 1, 20);
    s = seedNumber(s, 7, 0, 20);
    s = seedNumber(s, 7, 1, 40);
    s = seedNumber(s, 6, 3, 15);
    s = seedNumber(s, 6, 4, 15);
    s = seedNumber(s, 7, 3, 15);
    s = seedNumber(s, 7, 4, 45);
    s = withConditionalRules(s, [
      {
        kind: 'formula',
        range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
        formula:
          '=AND(ROUND(CORREL(OFFSET(A4,0,0,1,3),INDIRECT("D4:F4")),6)=0.993399,ROUND(F.TEST(OFFSET(A4,0,0,1,3),INDIRECT("D4:F4")),6)=0.272727,ROUND(FORECAST.LINEAR(4,OFFSET(D4,0,0,1,3),INDIRECT("A4:C4")),6)=9.333333)',
        apply: { fill: '#stats-dynamic-paired' },
      },
      {
        kind: 'formula',
        range: { sheet: 0, r0: 1, c0: 6, r1: 1, c1: 6 },
        formula:
          '=AND(ROUND(Z.TEST(OFFSET(A1,0,0,1,3),3,2),6)=0.193238,ROUND(T.TEST(INDIRECT("A4:C4"),OFFSET(D4,0,0,1,3),2,2),6)=0.209875,PROB(OFFSET(A5,0,0,1,4),INDIRECT("A6:D6"),2,3)=0.5)',
        apply: { fill: '#stats-dynamic-tests' },
      },
      {
        kind: 'formula',
        range: { sheet: 0, r0: 2, c0: 6, r1: 2, c1: 6 },
        formula:
          '=ROUND(CHISQ.TEST(OFFSET(A7,0,0,2,2),INDIRECT("D7:E8")),6)=ROUND(CHISQ.DIST.RT(5.555555555555556,1),6)',
        apply: { fill: '#stats-dynamic-chisq' },
      },
      formulaRule(
        { sheet: 0, r0: 3, c0: 6, r1: 3, c1: 6 },
        '=CHISQ.TEST(OFFSET(A7,0,0,2,2),INDIRECT("D7:F8"))=0',
        '#stats-dynamic-mismatch',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:6')?.fill).toBe('#stats-dynamic-paired');
    expect(overlay.get('0:1:6')?.fill).toBe('#stats-dynamic-tests');
    expect(overlay.get('0:2:6')?.fill).toBe('#stats-dynamic-chisq');
    expect(overlay.get('0:3:6')?.fill).toBeUndefined();
  });

  it('formula rules evaluate LARGE and SMALL ranked range operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 12);
    s = seedNumber(s, 1, 0, 8);
    s = seedNumber(s, 2, 0, 15);
    s = seedCell(s, 3, 0, { kind: 'text', value: 'ignored' });
    s = seedNumber(s, 4, 0, 4);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=AND(LARGE($A$1:$A$5,2)=12,SMALL($A$1:$A$5,2)=8)',
        '#ranked-range',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=LARGE($A$1:$A$5,5)=0',
        '#ranked-out-of-range',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=AND(PERCENTILE.INC($A$1:$A$5,0.5)=10,QUARTILE.INC($A$1:$A$5,3)=12.75)',
        '#percentile-quartile',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 3, r1: 1, c1: 3 },
        '=AND(PERCENTILE($A$1:$A$5,0.5)=10,QUARTILE($A$1:$A$5,3)=12.75)',
        '#legacy-percentile-quartile',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=PERCENTILE.INC($A$1:$A$5,1.2)=15',
        '#percentile-out-of-range',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=QUARTILE.INC($A$1:$A$5,5)=15',
        '#quartile-out-of-range',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
        '=AND(PERCENTRANK.INC($A$1:$A$5,10)=0.5,PERCENTRANK.INC($A$1:$A$5,12,2)=0.66)',
        '#percentrank',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 6, r1: 1, c1: 6 },
        '=AND(PERCENTRANK($A$1:$A$5,10)=0.5,PERCENTRANK($A$1:$A$5,12,2)=0.66)',
        '#legacy-percentrank',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
        '=PERCENTRANK.INC($A$1:$A$5,10,)=0.5',
        '#percentrank-omitted-significance',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
        '=PERCENTRANK.INC($A$1:$A$5,20)=1',
        '#percentrank-out-of-range',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
        '=AND(PERCENTILE.EXC($A$1:$A$5,0.5)=10,QUARTILE.EXC($A$1:$A$5,1)=5)',
        '#percentile-exc',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 10, r1: 0, c1: 10 },
        '=AND(PERCENTRANK.EXC($A$1:$A$5,10)=0.5,PERCENTRANK.EXC($A$1:$A$5,12,2)=0.6)',
        '#percentrank-exc',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 11, r1: 0, c1: 11 },
        '=PERCENTILE.EXC($A$1:$A$5,0.1)=4',
        '#percentile-exc-out-of-range',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 12, r1: 0, c1: 12 },
        '=QUARTILE.EXC($A$1:$A$5,0)=4',
        '#quartile-exc-out-of-range',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 13, r1: 0, c1: 13 },
        '=AND(AGGREGATE(14,0,$A$1:$A$5,2)=12,AGGREGATE(15,0,$A$1:$A$5,2)=8)',
        '#aggregate-ranked',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 14, r1: 0, c1: 14 },
        '=AND(AGGREGATE(16,0,$A$1:$A$5,0.5)=10,AGGREGATE(17,0,$A$1:$A$5,3)=12.75)',
        '#aggregate-percentile-inc',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 15, r1: 0, c1: 15 },
        '=AND(AGGREGATE(18,0,$A$1:$A$5,0.5)=10,AGGREGATE(19,0,$A$1:$A$5,1)=5)',
        '#aggregate-percentile-exc',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 16, r1: 0, c1: 16 },
        '=AGGREGATE(14,8,$A$1:$A$5,2)=12',
        '#aggregate-invalid-option',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#ranked-range');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBe('#percentile-quartile');
    expect(overlay.get('0:1:3')?.fill).toBe('#legacy-percentile-quartile');
    expect(overlay.get('0:0:4')?.fill).toBeUndefined();
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
    expect(overlay.get('0:0:6')?.fill).toBe('#percentrank');
    expect(overlay.get('0:1:6')?.fill).toBe('#legacy-percentrank');
    expect(overlay.get('0:0:7')?.fill).toBe('#percentrank-omitted-significance');
    expect(overlay.get('0:0:8')?.fill).toBeUndefined();
    expect(overlay.get('0:0:9')?.fill).toBe('#percentile-exc');
    expect(overlay.get('0:0:10')?.fill).toBe('#percentrank-exc');
    expect(overlay.get('0:0:11')?.fill).toBeUndefined();
    expect(overlay.get('0:0:12')?.fill).toBeUndefined();
    expect(overlay.get('0:0:13')?.fill).toBe('#aggregate-ranked');
    expect(overlay.get('0:0:14')?.fill).toBe('#aggregate-percentile-inc');
    expect(overlay.get('0:0:15')?.fill).toBe('#aggregate-percentile-exc');
    expect(overlay.get('0:0:16')?.fill).toBeUndefined();
  });

  it('formula rules evaluate RANK variants over numeric ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 12);
    s = seedNumber(s, 1, 0, 8);
    s = seedNumber(s, 2, 0, 12);
    s = seedNumber(s, 3, 0, 15);
    s = seedCell(s, 4, 0, { kind: 'text', value: 'ignored' });
    s = seedNumber(s, 5, 0, 4);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=AND(RANK.EQ(12,$A$1:$A$6)=2,RANK(12,$A$1:$A$6)=2)',
        '#rank-eq',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=RANK.AVG(12,$A$1:$A$6)=2.5',
        '#rank-avg',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=RANK.EQ(12,$A$1:$A$6,1)=3',
        '#rank-ascending',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=RANK.EQ(12,$A$1:$A$6,)=2',
        '#rank-omitted-order',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=RANK.EQ(10,$A$1:$A$6)=3',
        '#rank-not-found',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#rank-eq');
    expect(overlay.get('0:0:2')?.fill).toBe('#rank-avg');
    expect(overlay.get('0:0:3')?.fill).toBe('#rank-ascending');
    expect(overlay.get('0:0:4')?.fill).toBe('#rank-omitted-order');
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
  });

  it('formula rules evaluate COUNTIF with relative criteria references', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 7);
    s = seedNumber(s, 1, 0, 3);
    s = seedNumber(s, 2, 0, 7);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 2, c1: 1 },
        '=COUNTIF($A$1:$A$3,A1)>1',
        '#countif-ref',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#countif-ref');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
    expect(overlay.get('0:2:1')?.fill).toBe('#countif-ref');
  });

  it('formula rules evaluate COUNTIF with comparator criteria', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 1);
    s = seedNumber(s, 1, 0, 2);
    s = seedNumber(s, 2, 0, 3);
    s = seedCell(s, 0, 1, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 1, { kind: 'text', value: 'north' });
    s = seedCell(s, 2, 1, { kind: 'text', value: 'South' });
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=AND(COUNTIF($A$1:$A$3,">1")=2,COUNTIF($B$1:$B$3,"north")=2)',
        '#countif-criteria',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#countif-criteria');
  });

  it('formula rules evaluate COUNTIF wildcard criteria with tilde escapes', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'Northeast' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'N* literal' });
    s = seedCell(s, 3, 0, { kind: 'text', value: 'South' });
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=AND(COUNTIF($A$1:$A$4,"Nor*")=2,COUNTIF($A$1:$A$4,"N~* literal")=1)',
        '#countif-wild',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
        '=COUNTIF($A$1:$A$4,"<>S*")=3',
        '#countif-not-wild',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#countif-wild');
    expect(overlay.get('0:1:1')?.fill).toBe('#countif-not-wild');
  });

  it('formula rules evaluate COUNTIFS with multiple criteria ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'Northwest' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'South' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = seedNumber(s, 2, 1, 15);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=COUNTIFS($A$1:$A$3,"North*",$B$1:$B$3,">10")=1',
        '#countifs',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#countifs');
  });

  it('formula rules leave mismatched COUNTIFS ranges unapplied', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedNumber(s, 0, 1, 12);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=COUNTIFS($A$1:$A$2,"North*",$B$1:$B$3,">10")=1',
        '#countifs-mismatch',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate SUMIF with optional sum ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'Northwest' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'South' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = seedNumber(s, 2, 1, 15);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=SUMIF($A$1:$A$3,"North*",$B$1:$B$3)=20',
        '#sumif',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
        '=SUMIF($B$1:$B$3,">10")=27',
        '#sumif-self',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#sumif');
    expect(overlay.get('0:1:2')?.fill).toBe('#sumif-self');
  });

  it('formula rules evaluate SUMIFS with multiple criteria ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'Northwest' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'South' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = seedNumber(s, 2, 1, 15);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=SUMIFS($B$1:$B$3,$A$1:$A$3,"North*",$B$1:$B$3,">10")=12',
        '#sumifs',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#sumifs');
  });

  it('formula rules evaluate SUMPRODUCT over aligned ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 2);
    s = seedNumber(s, 1, 0, 3);
    s = seedNumber(s, 2, 0, 4);
    s = seedNumber(s, 0, 1, 5);
    s = seedCell(s, 1, 1, { kind: 'text', value: 'skip' });
    s = seedNumber(s, 2, 1, 7);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=SUMPRODUCT($A$1:$A$3,$B$1:$B$3)=38',
        '#sumproduct',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
        '=SUMPRODUCT($A$1:$A$2,$B$1:$B$3)=10',
        '#sumproduct-mismatch',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#sumproduct');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate SUMPRODUCT over dynamic ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 2);
    s = seedNumber(s, 1, 0, 3);
    s = seedNumber(s, 2, 0, 4);
    s = seedNumber(s, 0, 1, 5);
    s = seedCell(s, 1, 1, { kind: 'text', value: 'skip' });
    s = seedNumber(s, 2, 1, 7);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=SUMPRODUCT(OFFSET(A1,0,0,3,1),INDIRECT("B1:B3"))=38',
        '#sumproduct-dynamic',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
        '=SUMPRODUCT(OFFSET(A1,0,0,2,1),INDIRECT("B1:B3"))=10',
        '#sumproduct-dynamic-mismatch',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#sumproduct-dynamic');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
  });

  it('formula rules leave mismatched SUMIFS ranges unapplied', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedNumber(s, 0, 1, 12);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=SUMIFS($B$1:$B$2,$A$1:$A$3,"North*")=12',
        '#sumifs-mismatch',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate AVERAGEIF with optional average ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'Northwest' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'South' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = seedNumber(s, 2, 1, 15);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=AVERAGEIF($A$1:$A$3,"North*",$B$1:$B$3)=10',
        '#averageif',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
        '=AVERAGEIF($B$1:$B$3,">10")=13.5',
        '#averageif-self',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#averageif');
    expect(overlay.get('0:1:2')?.fill).toBe('#averageif-self');
  });

  it('formula rules evaluate AVERAGEIFS and fail closed on empty numeric matches', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'Northwest' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'South' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = seedNumber(s, 2, 1, 15);
    s = seedCell(s, 0, 3, { kind: 'text', value: 'not numeric' });
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=AVERAGEIFS($B$1:$B$3,$A$1:$A$3,"North*",$B$1:$B$3,">10")=12',
        '#averageifs',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
        '=AVERAGEIF($A$1:$A$1,"North",$D$1:$D$1)=0',
        '#averageif-empty',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#averageifs');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate MINIFS and MAXIFS with multiple criteria ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'Northwest' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'South' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = seedNumber(s, 2, 1, 15);
    s = withConditionalRules(s, [
      {
        kind: 'formula',
        range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        formula:
          '=AND(MINIFS($B$1:$B$3,$A$1:$A$3,"North*")=8,MAXIFS($B$1:$B$3,$A$1:$A$3,"North*")=12)',
        apply: { fill: '#minmaxifs' },
      },
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#minmaxifs');
  });

  it('formula rules leave MINIFS/MAXIFS without numeric matches unapplied', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 0, 1, { kind: 'text', value: 'not numeric' });
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=MINIFS($B$1:$B$1,$A$1:$A$1,"North")=0',
        '#minifs-empty',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate criteria functions over dynamic ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'Northwest' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'South' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = seedNumber(s, 2, 1, 15);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=COUNTIF(OFFSET(A1,0,0,3,1),"North*")=2',
        '#countif-offset',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=COUNTIFS(INDIRECT("A1:A3"),"North*",OFFSET(B1,0,0,3,1),">10")=1',
        '#countifs-dynamic',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=SUMIF(INDIRECT("A1:A3"),"North*",OFFSET(B1,0,0,3,1))=20',
        '#sumif-dynamic',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=AVERAGEIF(OFFSET(A1,0,0,3,1),"North*",INDIRECT("B1:B3"))=10',
        '#averageif-dynamic',
      ),
      {
        kind: 'formula',
        range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
        formula:
          '=SUMIFS(OFFSET(B1,0,0,3,1),INDIRECT("A1:A3"),"North*",INDIRECT("B1:B3"),">10")=12',
        apply: { fill: '#sumifs-dynamic' },
      },
      {
        kind: 'formula',
        range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
        formula:
          '=AND(MINIFS(OFFSET(B1,0,0,3,1),INDIRECT("A1:A3"),"North*")=8,MAXIFS(INDIRECT("B1:B3"),OFFSET(A1,0,0,3,1),"North*")=12)',
        apply: { fill: '#minmaxifs-dynamic' },
      },
      formulaRule(
        { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
        '=COUNTIF(OFFSET(A1,0,0,10001,1),"North*")=2',
        '#countif-dynamic-too-large',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
        '=SUMIF(INDIRECT("Sheet2!A1:A3"),"North*",B1:B3)=20',
        '#sumif-dynamic-other-sheet',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#countif-offset');
    expect(overlay.get('0:0:3')?.fill).toBe('#countifs-dynamic');
    expect(overlay.get('0:0:4')?.fill).toBe('#sumif-dynamic');
    expect(overlay.get('0:0:5')?.fill).toBe('#averageif-dynamic');
    expect(overlay.get('0:0:6')?.fill).toBe('#sumifs-dynamic');
    expect(overlay.get('0:0:7')?.fill).toBe('#minmaxifs-dynamic');
    expect(overlay.get('0:0:8')?.fill).toBeUndefined();
    expect(overlay.get('0:0:9')?.fill).toBeUndefined();
  });
});
