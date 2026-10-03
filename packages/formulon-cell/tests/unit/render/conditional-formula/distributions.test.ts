import { afterEach, describe, expect, it, vi } from 'vitest';
import { _resetConditionalCache, evaluateConditional } from '../../../../src/render/conditional.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';

describe('evaluateConditional', () => {
  afterEach(() => {
    _resetConditionalCache();
    vi.useRealTimers();
  });

  it('formula rules evaluate normal distribution numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            formula:
              '=AND(STANDARDIZE(42,40,2)=1,ROUND(NORM.S.DIST(1,TRUE()),6)=0.841345,ROUND(NORM.S.DIST(0,FALSE()),6)=0.398942,ROUND(PHI(0),6)=0.398942)',
            apply: { fill: '#standard-normal' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula:
              '=AND(ROUND(NORM.DIST(42,40,2,TRUE()),6)=0.841345,ROUND(NORM.DIST(42,40,2,FALSE()),6)=0.120985,ROUND(CONFIDENCE.NORM(0.05,2.5,50),6)=0.692952)',
            apply: { fill: '#normal' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 11, r1: 0, c1: 11 },
            formula: '=ROUND(CONFIDENCE.T(0.05,2.5,50),6)=ROUND(T.INV.2T(0.05,49)*2.5/SQRT(50),6)',
            apply: { fill: '#confidence-t' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula:
              '=AND(ROUND(NORMSDIST(1),6)=0.841345,ROUND(NORMDIST(42,40,2,FALSE()),6)=0.120985,ROUND(CONFIDENCE(0.05,2.5,50),6)=0.692952)',
            apply: { fill: '#legacy-normal' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula:
              '=AND(ROUND(FISHER(0.75),6)=0.972955,ROUND(FISHERINV(0.9729550745276566),6)=0.75)',
            apply: { fill: '#fisher' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=FISHER(1)=0',
            apply: { fill: '#fisher-out-of-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula:
              '=AND(ROUND(NORM.S.INV(0.841344746068543),6)=1,ROUND(NORM.INV(0.841344746068543,40,2),6)=42,ROUND(LOGNORM.INV(0.841344746068543,1,0.5),6)=4.481689)',
            apply: { fill: '#normal-inv' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula:
              '=AND(ROUND(NORMSINV(0.841344746068543),6)=1,ROUND(NORMINV(0.841344746068543,40,2),6)=42,ROUND(LOGINV(0.841344746068543,1,0.5),6)=4.481689)',
            apply: { fill: '#legacy-normal-inv' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=NORM.S.INV(0)=0',
            apply: { fill: '#normal-inv-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 10, r1: 0, c1: 10 },
            formula: '=CONFIDENCE.NORM(0,2.5,50)=0',
            apply: { fill: '#confidence-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 12, r1: 0, c1: 12 },
            formula: '=CONFIDENCE.T(0.05,2.5,1)=0',
            apply: { fill: '#confidence-t-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=STANDARDIZE(42,40,0)=0',
            apply: { fill: '#standardize-zero-deviation' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=NORM.DIST(42,40,0,TRUE())=0',
            apply: { fill: '#normal-zero-deviation' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#standard-normal');
    expect(overlay.get('0:0:1')?.fill).toBe('#normal');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
    expect(overlay.get('0:0:4')?.fill).toBe('#legacy-normal');
    expect(overlay.get('0:0:5')?.fill).toBe('#fisher');
    expect(overlay.get('0:0:6')?.fill).toBeUndefined();
    expect(overlay.get('0:0:7')?.fill).toBe('#normal-inv');
    expect(overlay.get('0:0:8')?.fill).toBe('#legacy-normal-inv');
    expect(overlay.get('0:0:9')?.fill).toBeUndefined();
    expect(overlay.get('0:0:10')?.fill).toBeUndefined();
  });

  it('formula rules evaluate probability distribution numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            formula:
              '=AND(ROUND(BINOM.DIST(6,10,0.5,FALSE()),6)=0.205078,ROUND(BINOM.DIST(6,10,0.5,TRUE()),6)=0.828125,BINOM.INV(10,0.5,0.8)=6)',
            apply: { fill: '#binom-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula:
              '=AND(ROUND(POISSON.DIST(2,5,FALSE()),6)=0.084224,ROUND(POISSON.DIST(2,5,TRUE()),6)=0.124652)',
            apply: { fill: '#poisson-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula:
              '=AND(ROUND(EXPON.DIST(0.2,10,TRUE()),6)=0.864665,ROUND(EXPON.DIST(0.2,10,FALSE()),6)=1.353353)',
            apply: { fill: '#expon-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula:
              '=AND(ROUND(BINOMDIST(6,10,0.5,FALSE()),6)=0.205078,CRITBINOM(10,0.5,0.8)=6,ROUND(POISSON(2,5,TRUE()),6)=0.124652,ROUND(EXPONDIST(0.2,10,FALSE()),6)=1.353353)',
            apply: { fill: '#legacy-distributions' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula:
              '=AND(ROUND(LOGNORM.DIST(4,1,0.5,TRUE()),6)=0.780117,ROUND(LOGNORM.DIST(4,1,0.5,FALSE()),6)=0.148002)',
            apply: { fill: '#lognorm-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula:
              '=AND(ROUND(WEIBULL.DIST(3,2,4,TRUE()),6)=0.430217,ROUND(WEIBULL.DIST(3,2,4,FALSE()),6)=0.213669)',
            apply: { fill: '#weibull-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 10, r1: 0, c1: 10 },
            formula:
              '=AND(ROUND(LOGNORMDIST(4,1,0.5),6)=0.780117,ROUND(WEIBULL(3,2,4,FALSE()),6)=0.213669)',
            apply: { fill: '#legacy-lognorm-weibull' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 17, r1: 0, c1: 17 },
            formula:
              '=AND(ROUND(GAMMA.DIST(4,2,3,TRUE()),6)=0.38494,ROUND(GAMMA.DIST(4,2,3,FALSE()),6)=0.117154)',
            apply: { fill: '#gamma-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 18, r1: 0, c1: 18 },
            formula:
              '=AND(ROUND(GAMMADIST(2,3,2,TRUE()),6)=0.080301,ROUND(GAMMADIST(2,3,2,FALSE()),6)=0.09197)',
            apply: { fill: '#legacy-gamma-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 26, r1: 0, c1: 26 },
            formula:
              '=AND(ROUND(GAMMA.INV(0.384940011063304,2,3),6)=4,ROUND(GAMMAINV(0.080301397071394,3,2),6)=2)',
            apply: { fill: '#gamma-inv' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 28, r1: 0, c1: 28 },
            formula:
              '=AND(ROUND(BETA.DIST(0.4,2,3,TRUE()),6)=0.5248,ROUND(BETA.DIST(0.4,2,3,FALSE()),6)=1.728)',
            apply: { fill: '#beta-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 29, r1: 0, c1: 29 },
            formula:
              '=AND(ROUND(BETA.DIST(4,2,3,TRUE(),0,10),6)=0.5248,ROUND(BETA.DIST(4,2,3,FALSE(),0,10),6)=0.1728)',
            apply: { fill: '#beta-dist-scaled' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 30, r1: 0, c1: 30 },
            formula:
              '=AND(ROUND(BETADIST(0.4,2,3),6)=0.5248,ROUND(BETA.INV(0.5248,2,3,0,10),6)=4,ROUND(BETAINV(0.5248,2,3),6)=0.4)',
            apply: { fill: '#beta-inv-legacy' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 33, r1: 0, c1: 33 },
            formula:
              '=AND(ROUND(F.DIST(3,2,10,TRUE()),6)=0.904633,ROUND(F.DIST(3,2,10,FALSE()),6)=0.059605)',
            apply: { fill: '#f-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 34, r1: 0, c1: 34 },
            formula: '=AND(ROUND(F.DIST.RT(3,2,10),6)=0.095367,ROUND(FDIST(3,2,10),6)=0.095367)',
            apply: { fill: '#f-dist-rt' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 35, r1: 0, c1: 35 },
            formula:
              '=AND(ROUND(F.INV(0.904632568359375,2,10),6)=3,ROUND(F.INV.RT(0.095367431640625,2,10),6)=3,ROUND(FINV(0.095367431640625,2,10),6)=3)',
            apply: { fill: '#f-inv' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 38, r1: 0, c1: 38 },
            formula:
              '=AND(ROUND(T.DIST(1,10,TRUE()),6)=0.829553,ROUND(T.DIST(1,10,FALSE()),6)=0.230362)',
            apply: { fill: '#t-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 39, r1: 0, c1: 39 },
            formula:
              '=AND(ROUND(T.DIST.RT(1,10),6)=0.170447,ROUND(T.DIST.2T(1,10),6)=0.340893,ROUND(TDIST(1,10,2),6)=0.340893)',
            apply: { fill: '#t-dist-tails' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 40, r1: 0, c1: 40 },
            formula:
              '=AND(ROUND(T.INV(0.82955343384897,10),6)=1,ROUND(T.INV.2T(0.34089313230206,10),6)=1,ROUND(TINV(0.34089313230206,10),6)=1)',
            apply: { fill: '#t-inv' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 21, r1: 0, c1: 21 },
            formula:
              '=AND(ROUND(CHISQ.DIST(3,2,TRUE()),6)=0.77687,ROUND(CHISQ.DIST(3,2,FALSE()),6)=0.111565)',
            apply: { fill: '#chisq-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 22, r1: 0, c1: 22 },
            formula: '=AND(ROUND(CHISQ.DIST.RT(3,2),6)=0.22313,ROUND(CHIDIST(3,2),6)=0.22313)',
            apply: { fill: '#chisq-dist-rt' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 24, r1: 0, c1: 24 },
            formula:
              '=AND(ROUND(CHISQ.INV(0.77686983985157,2),6)=3,ROUND(CHISQ.INV.RT(0.22313016014843,2),6)=3,ROUND(CHIINV(0.22313016014843,2),6)=3)',
            apply: { fill: '#chisq-inv' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 11, r1: 0, c1: 11 },
            formula:
              '=AND(ROUND(NEGBINOM.DIST(10,5,0.25,FALSE()),6)=0.055049,ROUND(NEGBINOM.DIST(10,5,0.25,TRUE()),6)=0.313514)',
            apply: { fill: '#negbinom-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 12, r1: 0, c1: 12 },
            formula:
              '=AND(ROUND(HYPGEOM.DIST(1,4,8,20,FALSE()),6)=0.363261,ROUND(HYPGEOM.DIST(1,4,8,20,TRUE()),6)=0.465428)',
            apply: { fill: '#hypgeom-dist' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 15, r1: 0, c1: 15 },
            formula:
              '=AND(ROUND(NEGBINOMDIST(10,5,0.25),6)=0.055049,ROUND(HYPGEOMDIST(1,4,8,20),6)=0.363261)',
            apply: { fill: '#legacy-discrete-distributions' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=BINOM.DIST(11,10,0.5,FALSE())=0',
            apply: { fill: '#binom-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 16, r1: 0, c1: 16 },
            formula: '=BINOM.INV(10,0.5,0)=0',
            apply: { fill: '#binom-inv-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 19, r1: 0, c1: 19 },
            formula: '=GAMMA.DIST(1,0,2,TRUE())=0',
            apply: { fill: '#gamma-dist-invalid-alpha' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 20, r1: 0, c1: 20 },
            formula: '=GAMMA.DIST(1,2,0,TRUE())=0',
            apply: { fill: '#gamma-dist-invalid-beta' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 27, r1: 0, c1: 27 },
            formula: '=GAMMA.INV(0,2,3)=0',
            apply: { fill: '#gamma-inv-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 31, r1: 0, c1: 31 },
            formula: '=BETA.DIST(0.4,0,3,TRUE())=0',
            apply: { fill: '#beta-dist-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 32, r1: 0, c1: 32 },
            formula: '=BETA.INV(0,2,3)=0',
            apply: { fill: '#beta-inv-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 36, r1: 0, c1: 36 },
            formula: '=F.DIST(1,0,10,TRUE())=0',
            apply: { fill: '#f-dist-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 37, r1: 0, c1: 37 },
            formula: '=F.INV(0,2,10)=0',
            apply: { fill: '#f-inv-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 41, r1: 0, c1: 41 },
            formula: '=T.DIST.RT(-1,10)=0',
            apply: { fill: '#t-dist-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 42, r1: 0, c1: 42 },
            formula: '=T.INV(0,10)=0',
            apply: { fill: '#t-inv-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 23, r1: 0, c1: 23 },
            formula: '=CHISQ.DIST(1,0,TRUE())=0',
            apply: { fill: '#chisq-dist-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 25, r1: 0, c1: 25 },
            formula: '=CHISQ.INV(0,2)=0',
            apply: { fill: '#chisq-inv-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=EXPON.DIST(1,0,TRUE())=0',
            apply: { fill: '#expon-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=LOGNORM.DIST(0,1,0.5,TRUE())=0',
            apply: { fill: '#lognorm-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=WEIBULL.DIST(3,0,4,TRUE())=0',
            apply: { fill: '#weibull-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 13, r1: 0, c1: 13 },
            formula: '=NEGBINOM.DIST(10,0,0.25,FALSE())=0',
            apply: { fill: '#negbinom-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 14, r1: 0, c1: 14 },
            formula: '=HYPGEOM.DIST(5,4,8,20,FALSE())=0',
            apply: { fill: '#hypgeom-invalid' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#binom-dist');
    expect(overlay.get('0:0:1')?.fill).toBe('#poisson-dist');
    expect(overlay.get('0:0:2')?.fill).toBe('#expon-dist');
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
    expect(overlay.get('0:0:4')?.fill).toBeUndefined();
    expect(overlay.get('0:0:5')?.fill).toBe('#legacy-distributions');
    expect(overlay.get('0:0:6')?.fill).toBe('#lognorm-dist');
    expect(overlay.get('0:0:7')?.fill).toBe('#weibull-dist');
    expect(overlay.get('0:0:8')?.fill).toBeUndefined();
    expect(overlay.get('0:0:9')?.fill).toBeUndefined();
    expect(overlay.get('0:0:10')?.fill).toBe('#legacy-lognorm-weibull');
    expect(overlay.get('0:0:11')?.fill).toBe('#negbinom-dist');
    expect(overlay.get('0:0:12')?.fill).toBe('#hypgeom-dist');
    expect(overlay.get('0:0:13')?.fill).toBeUndefined();
    expect(overlay.get('0:0:14')?.fill).toBeUndefined();
    expect(overlay.get('0:0:15')?.fill).toBe('#legacy-discrete-distributions');
    expect(overlay.get('0:0:16')?.fill).toBeUndefined();
    expect(overlay.get('0:0:17')?.fill).toBe('#gamma-dist');
    expect(overlay.get('0:0:18')?.fill).toBe('#legacy-gamma-dist');
    expect(overlay.get('0:0:19')?.fill).toBeUndefined();
    expect(overlay.get('0:0:20')?.fill).toBeUndefined();
    expect(overlay.get('0:0:21')?.fill).toBe('#chisq-dist');
    expect(overlay.get('0:0:22')?.fill).toBe('#chisq-dist-rt');
    expect(overlay.get('0:0:23')?.fill).toBeUndefined();
    expect(overlay.get('0:0:24')?.fill).toBe('#chisq-inv');
    expect(overlay.get('0:0:25')?.fill).toBeUndefined();
    expect(overlay.get('0:0:26')?.fill).toBe('#gamma-inv');
    expect(overlay.get('0:0:27')?.fill).toBeUndefined();
    expect(overlay.get('0:0:28')?.fill).toBe('#beta-dist');
    expect(overlay.get('0:0:29')?.fill).toBe('#beta-dist-scaled');
    expect(overlay.get('0:0:30')?.fill).toBe('#beta-inv-legacy');
    expect(overlay.get('0:0:31')?.fill).toBeUndefined();
    expect(overlay.get('0:0:32')?.fill).toBeUndefined();
    expect(overlay.get('0:0:33')?.fill).toBe('#f-dist');
    expect(overlay.get('0:0:34')?.fill).toBe('#f-dist-rt');
    expect(overlay.get('0:0:35')?.fill).toBe('#f-inv');
    expect(overlay.get('0:0:36')?.fill).toBeUndefined();
    expect(overlay.get('0:0:37')?.fill).toBeUndefined();
  });
});
