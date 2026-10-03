import { afterEach, describe, expect, it, vi } from 'vitest';
import { _resetConditionalCache, evaluateConditional } from '../../../../src/render/conditional.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';
import { seedCell, seedNumber } from '../conditional-fixtures.js';

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 1, c1: 2 },
            formula: '=SUM(A1:A3)>5',
            apply: { fill: '#sum' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 1, c1: 2 },
            formula: '=AND(AVERAGE($A$1:$A$3)=2,MAX($A$1:$A$3)=COUNT($A$1:$A$3))',
            apply: { fill: '#agg' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=AND(COUNTA($A$1:$A$5)=4,COUNTBLANK($A$1:$A$5)=1)',
            apply: { fill: '#counta-countblank' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=PRODUCT($B$1:$B$3)=24',
            apply: { fill: '#product' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=MEDIAN($A$1:$A$5)=10',
            apply: { fill: '#median-even' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=MEDIAN($A$1:$A$3)=12',
            apply: { fill: '#median-odd' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=MIN($A$1:$A$5)=4',
            apply: { fill: '#min-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=MIN($A$1:$A$2,10,-3)=-3',
            apply: { fill: '#min-args' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=MIN($A$1:$A$2)=12',
            apply: { fill: '#min-miss' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=AND(SUM(A1,B1,$C$1:$C$1)=12,PRODUCT(A1,B1,3)=24)',
            apply: { fill: '#multi-sum-product' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=AND(AVERAGE($A$1:$C$1,A2)=5,MEDIAN($A$1:$C$1,A2)=5)',
            apply: { fill: '#multi-average-median' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=AND(AVERAGEA($A$1:$C$2)=3.5,MINA($A$1:$C$2)=0,MAXA($A$1:$C$2)=8)',
            apply: { fill: '#aggregate-a' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=AND(COUNT(A1,B1,B2)=2,COUNTA(A1,B2,C3)=2,COUNTBLANK(B2,C3)=1)',
            apply: { fill: '#multi-counts' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 3, r1: 1, c1: 3 },
            formula: '=AND(ROUND(STDEV.S($A$1:$C$1),6)=2,ROUND(STDEV.P($A$1:$C$1),6)=1.632993)',
            apply: { fill: '#stdev-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 4, r1: 1, c1: 4 },
            formula: '=AND(VAR.S(A1,B1,C1)=4,ROUND(VAR.P(A1,B1,C1),6)=2.666667)',
            apply: { fill: '#var-args' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=AND(ROUND(STDEV($A$1:$C$1),6)=2,VAR(A1,B1,C1)=4)',
            apply: { fill: '#legacy-stdev-var' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=AND(ROUND(STDEVP($A$1:$C$1),6)=1.632993,ROUND(VARP(A1,B1,C1),6)=2.666667)',
            apply: { fill: '#legacy-stdevp-varp' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 5, r1: 1, c1: 5 },
            formula: '=STDEV.S(A1)=0',
            apply: { fill: '#stdev-s-single' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 6, r1: 1, c1: 6 },
            formula:
              '=AND(ROUND(GEOMEAN($A$1:$C$1),6)=3.634241,ROUND(HARMEAN(A1,B1,C1),6)=3.272727)',
            apply: { fill: '#geo-harmonic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 7, r1: 1, c1: 7 },
            formula: '=GEOMEAN(A1,-1)=0',
            apply: { fill: '#geomean-negative' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 8, r1: 1, c1: 8 },
            formula: '=AND(DEVSQ(A1,B1,C1)=8,DEVSQ($A$1:$C$1,A2)=20)',
            apply: { fill: '#devsq' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 9, r1: 1, c1: 9 },
            formula: '=AND(ROUND(AVEDEV(A1,B1,C1),6)=1.333333,AVEDEV($A$1:$C$1,A2)=2)',
            apply: { fill: '#avedev' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 10, r1: 1, c1: 10 },
            formula: '=ROUND(SKEW(A1,B1,C1,20),6)=1.763633',
            apply: { fill: '#skew' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 11, r1: 1, c1: 11 },
            formula: '=SKEW(A1,B1)=0',
            apply: { fill: '#skew-too-few' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 32, r1: 1, c1: 32 },
            formula:
              '=AND(ROUND(SKEW.P(A1,B1,C1,20),6)=1.018234,ROUND(SKEW.P($A$1:$C$1,20),6)=1.018234)',
            apply: { fill: '#skew-p' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 33, r1: 1, c1: 33 },
            formula: '=SKEW.P(A1,B1)=0',
            apply: { fill: '#skew-p-too-few' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 34, r1: 1, c1: 34 },
            formula:
              '=AND(ROUND(Z.TEST($A$1:$C$1,3,2),6)=0.193238,ROUND(Z.TEST($A$1:$C$1,3),6)=0.193238,ROUND(ZTEST($A$1:$C$1,3,2),6)=0.193238)',
            apply: { fill: '#z-test' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 35, r1: 1, c1: 35 },
            formula: '=Z.TEST($A$1:$C$1,3,0)=0',
            apply: { fill: '#z-test-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 36, r1: 1, c1: 36 },
            formula:
              '=AND(ROUND(F.TEST($A$4:$C$4,$D$4:$F$4),6)=0.272727,ROUND(FTEST($A$4:$C$4,$D$4:$F$4),6)=0.272727)',
            apply: { fill: '#f-test' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 37, r1: 1, c1: 37 },
            formula: '=F.TEST($A$1:$A$1,$D$4:$F$4)=0',
            apply: { fill: '#f-test-invalid' },
          },
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
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 40, r1: 1, c1: 40 },
            formula: '=T.TEST($A$4:$C$4,$D$4:$F$4,3,2)=0',
            apply: { fill: '#t-test-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 41, r1: 1, c1: 41 },
            formula:
              '=AND(ROUND(CHISQ.TEST($A$7:$B$8,$D$7:$E$8),6)=ROUND(CHISQ.DIST.RT(5.555555555555556,1),6),ROUND(CHITEST($A$7:$B$8,$D$7:$E$8),6)=ROUND(CHISQ.DIST.RT(5.555555555555556,1),6))',
            apply: { fill: '#chisq-test' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 42, r1: 1, c1: 42 },
            formula: '=CHISQ.TEST($A$7:$B$8,$D$7:$F$8)=0',
            apply: { fill: '#chisq-test-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 12, r1: 1, c1: 12 },
            formula: '=ROUND(KURT(A1,B1,C1,20),6)=3.228',
            apply: { fill: '#kurt' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 13, r1: 1, c1: 13 },
            formula: '=KURT(A1,B1,C1)=0',
            apply: { fill: '#kurt-too-few' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 14, r1: 1, c1: 14 },
            formula: '=AND(MODE.SNGL(A1,B1,C1,B1)=4,MODE.SNGL($A$1:$C$1,4)=4)',
            apply: { fill: '#mode-single' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 14, r1: 0, c1: 14 },
            formula: '=AND(MODE(A1,B1,C1,B1)=4,MODE($A$1:$C$1,4)=4)',
            apply: { fill: '#legacy-mode' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 15, r1: 1, c1: 15 },
            formula: '=MODE.SNGL(A1,B1,C1)=2',
            apply: { fill: '#mode-no-duplicate' },
          },
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
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 17, r1: 1, c1: 17 },
            formula: '=CORREL($A$4:$B$4,$D$4:$F$4)=1',
            apply: { fill: '#paired-stats-mismatch' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 18, r1: 1, c1: 18 },
            formula: '=CORREL($A$1:$C$1,$B$1:$B$3)=1',
            apply: { fill: '#paired-stats-zero-variance' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 20, r1: 1, c1: 20 },
            formula: '=SLOPE($A$1:$C$1,$B$1:$B$3)=1',
            apply: { fill: '#regression-zero-x-variance' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 22, r1: 1, c1: 22 },
            formula: '=FORECAST("x",$D$4:$F$4,$A$4:$C$4)=1',
            apply: { fill: '#forecast-nonnumeric-x' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 23, r1: 1, c1: 23 },
            formula: '=AND(PROB($A$5:$D$5,$A$6:$D$6,2,3)=0.5,PROB($A$5:$D$5,$A$6:$D$6,4)=0.4)',
            apply: { fill: '#probability-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 24, r1: 1, c1: 24 },
            formula: '=PROB($A$5:$D$5,$A$6:$C$6,2)=0.2',
            apply: { fill: '#probability-mismatch' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 25, r1: 1, c1: 25 },
            formula: '=PROB($A$5:$C$5,$A$6:$C$6,2)=0.2',
            apply: { fill: '#probability-total-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 26, r1: 1, c1: 26 },
            formula:
              '=AND(SUMX2MY2($A$4:$C$4,$D$4:$F$4)=-55,SUMX2PY2($A$4:$C$4,$D$4:$F$4)=83,SUMXMY2($A$4:$C$4,$D$4:$F$4)=21)',
            apply: { fill: '#sumx-paired' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 27, r1: 1, c1: 27 },
            formula: '=SUMXMY2($A$4:$B$4,$D$4:$F$4)=1',
            apply: { fill: '#sumx-mismatch' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 28, r1: 1, c1: 28 },
            formula:
              '=AND(SUBTOTAL(9,$A$1:$C$1)=12,SUBTOTAL(101,$A$1:$C$1)=4,SUBTOTAL(103,$A$1:$C$2)=6,ROUND(SUBTOTAL(107,$A$1:$C$1),6)=2)',
            apply: { fill: '#subtotal' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 29, r1: 1, c1: 29 },
            formula: '=SUBTOTAL(12,$A$1:$C$1)=0',
            apply: { fill: '#subtotal-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 30, r1: 1, c1: 30 },
            formula:
              '=AND(AGGREGATE(9,0,$A$1:$C$1)=12,AGGREGATE(1,0,$A$1:$C$1)=4,ROUND(AGGREGATE(7,0,$A$1:$C$1),6)=2)',
            apply: { fill: '#aggregate-function' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 31, r1: 1, c1: 31 },
            formula: '=AGGREGATE(20,0,$A$1:$C$1)=0',
            apply: { fill: '#aggregate-function-invalid' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
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
          {
            kind: 'formula',
            range: { sheet: 0, r0: 3, c0: 6, r1: 3, c1: 6 },
            formula: '=CHISQ.TEST(OFFSET(A7,0,0,2,2),INDIRECT("D7:F8"))=0',
            apply: { fill: '#stats-dynamic-mismatch' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(LARGE($A$1:$A$5,2)=12,SMALL($A$1:$A$5,2)=8)',
            apply: { fill: '#ranked-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=LARGE($A$1:$A$5,5)=0',
            apply: { fill: '#ranked-out-of-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=AND(PERCENTILE.INC($A$1:$A$5,0.5)=10,QUARTILE.INC($A$1:$A$5,3)=12.75)',
            apply: { fill: '#percentile-quartile' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 3, r1: 1, c1: 3 },
            formula: '=AND(PERCENTILE($A$1:$A$5,0.5)=10,QUARTILE($A$1:$A$5,3)=12.75)',
            apply: { fill: '#legacy-percentile-quartile' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=PERCENTILE.INC($A$1:$A$5,1.2)=15',
            apply: { fill: '#percentile-out-of-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=QUARTILE.INC($A$1:$A$5,5)=15',
            apply: { fill: '#quartile-out-of-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=AND(PERCENTRANK.INC($A$1:$A$5,10)=0.5,PERCENTRANK.INC($A$1:$A$5,12,2)=0.66)',
            apply: { fill: '#percentrank' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 6, r1: 1, c1: 6 },
            formula: '=AND(PERCENTRANK($A$1:$A$5,10)=0.5,PERCENTRANK($A$1:$A$5,12,2)=0.66)',
            apply: { fill: '#legacy-percentrank' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=PERCENTRANK.INC($A$1:$A$5,10,)=0.5',
            apply: { fill: '#percentrank-omitted-significance' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=PERCENTRANK.INC($A$1:$A$5,20)=1',
            apply: { fill: '#percentrank-out-of-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=AND(PERCENTILE.EXC($A$1:$A$5,0.5)=10,QUARTILE.EXC($A$1:$A$5,1)=5)',
            apply: { fill: '#percentile-exc' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 10, r1: 0, c1: 10 },
            formula: '=AND(PERCENTRANK.EXC($A$1:$A$5,10)=0.5,PERCENTRANK.EXC($A$1:$A$5,12,2)=0.6)',
            apply: { fill: '#percentrank-exc' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 11, r1: 0, c1: 11 },
            formula: '=PERCENTILE.EXC($A$1:$A$5,0.1)=4',
            apply: { fill: '#percentile-exc-out-of-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 12, r1: 0, c1: 12 },
            formula: '=QUARTILE.EXC($A$1:$A$5,0)=4',
            apply: { fill: '#quartile-exc-out-of-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 13, r1: 0, c1: 13 },
            formula: '=AND(AGGREGATE(14,0,$A$1:$A$5,2)=12,AGGREGATE(15,0,$A$1:$A$5,2)=8)',
            apply: { fill: '#aggregate-ranked' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 14, r1: 0, c1: 14 },
            formula: '=AND(AGGREGATE(16,0,$A$1:$A$5,0.5)=10,AGGREGATE(17,0,$A$1:$A$5,3)=12.75)',
            apply: { fill: '#aggregate-percentile-inc' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 15, r1: 0, c1: 15 },
            formula: '=AND(AGGREGATE(18,0,$A$1:$A$5,0.5)=10,AGGREGATE(19,0,$A$1:$A$5,1)=5)',
            apply: { fill: '#aggregate-percentile-exc' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 16, r1: 0, c1: 16 },
            formula: '=AGGREGATE(14,8,$A$1:$A$5,2)=12',
            apply: { fill: '#aggregate-invalid-option' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(RANK.EQ(12,$A$1:$A$6)=2,RANK(12,$A$1:$A$6)=2)',
            apply: { fill: '#rank-eq' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=RANK.AVG(12,$A$1:$A$6)=2.5',
            apply: { fill: '#rank-avg' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=RANK.EQ(12,$A$1:$A$6,1)=3',
            apply: { fill: '#rank-ascending' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=RANK.EQ(12,$A$1:$A$6,)=2',
            apply: { fill: '#rank-omitted-order' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=RANK.EQ(10,$A$1:$A$6)=3',
            apply: { fill: '#rank-not-found' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 2, c1: 1 },
            formula: '=COUNTIF($A$1:$A$3,A1)>1',
            apply: { fill: '#countif-ref' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=AND(COUNTIF($A$1:$A$3,">1")=2,COUNTIF($B$1:$B$3,"north")=2)',
            apply: { fill: '#countif-criteria' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(COUNTIF($A$1:$A$4,"Nor*")=2,COUNTIF($A$1:$A$4,"N~* literal")=1)',
            apply: { fill: '#countif-wild' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=COUNTIF($A$1:$A$4,"<>S*")=3',
            apply: { fill: '#countif-not-wild' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=COUNTIFS($A$1:$A$3,"North*",$B$1:$B$3,">10")=1',
            apply: { fill: '#countifs' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#countifs');
  });

  it('formula rules leave mismatched COUNTIFS ranges unapplied', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedNumber(s, 0, 1, 12);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=COUNTIFS($A$1:$A$2,"North*",$B$1:$B$3,">10")=1',
            apply: { fill: '#countifs-mismatch' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=SUMIF($A$1:$A$3,"North*",$B$1:$B$3)=20',
            apply: { fill: '#sumif' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=SUMIF($B$1:$B$3,">10")=27',
            apply: { fill: '#sumif-self' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=SUMIFS($B$1:$B$3,$A$1:$A$3,"North*",$B$1:$B$3,">10")=12',
            apply: { fill: '#sumifs' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=SUMPRODUCT($A$1:$A$3,$B$1:$B$3)=38',
            apply: { fill: '#sumproduct' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=SUMPRODUCT($A$1:$A$2,$B$1:$B$3)=10',
            apply: { fill: '#sumproduct-mismatch' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=SUMPRODUCT(OFFSET(A1,0,0,3,1),INDIRECT("B1:B3"))=38',
            apply: { fill: '#sumproduct-dynamic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=SUMPRODUCT(OFFSET(A1,0,0,2,1),INDIRECT("B1:B3"))=10',
            apply: { fill: '#sumproduct-dynamic-mismatch' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#sumproduct-dynamic');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
  });

  it('formula rules leave mismatched SUMIFS ranges unapplied', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedNumber(s, 0, 1, 12);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=SUMIFS($B$1:$B$2,$A$1:$A$3,"North*")=12',
            apply: { fill: '#sumifs-mismatch' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=AVERAGEIF($A$1:$A$3,"North*",$B$1:$B$3)=10',
            apply: { fill: '#averageif' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=AVERAGEIF($B$1:$B$3,">10")=13.5',
            apply: { fill: '#averageif-self' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=AVERAGEIFS($B$1:$B$3,$A$1:$A$3,"North*",$B$1:$B$3,">10")=12',
            apply: { fill: '#averageifs' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=AVERAGEIF($A$1:$A$1,"North",$D$1:$D$1)=0',
            apply: { fill: '#averageif-empty' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula:
              '=AND(MINIFS($B$1:$B$3,$A$1:$A$3,"North*")=8,MAXIFS($B$1:$B$3,$A$1:$A$3,"North*")=12)',
            apply: { fill: '#minmaxifs' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#minmaxifs');
  });

  it('formula rules leave MINIFS/MAXIFS without numeric matches unapplied', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 0, 1, { kind: 'text', value: 'not numeric' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=MINIFS($B$1:$B$1,$A$1:$A$1,"North")=0',
            apply: { fill: '#minifs-empty' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=COUNTIF(OFFSET(A1,0,0,3,1),"North*")=2',
            apply: { fill: '#countif-offset' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=COUNTIFS(INDIRECT("A1:A3"),"North*",OFFSET(B1,0,0,3,1),">10")=1',
            apply: { fill: '#countifs-dynamic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=SUMIF(INDIRECT("A1:A3"),"North*",OFFSET(B1,0,0,3,1))=20',
            apply: { fill: '#sumif-dynamic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=AVERAGEIF(OFFSET(A1,0,0,3,1),"North*",INDIRECT("B1:B3"))=10',
            apply: { fill: '#averageif-dynamic' },
          },
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
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=COUNTIF(OFFSET(A1,0,0,10001,1),"North*")=2',
            apply: { fill: '#countif-dynamic-too-large' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=SUMIF(INDIRECT("Sheet2!A1:A3"),"North*",B1:B3)=20',
            apply: { fill: '#sumif-dynamic-other-sheet' },
          },
        ],
      },
    };

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
