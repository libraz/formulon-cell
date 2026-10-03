import { afterEach, describe, expect, it, vi } from 'vitest';
import { _resetConditionalCache, evaluateConditional } from '../../../../src/render/conditional.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';
import { dateSerial, seedNumber } from '../conditional-fixtures.js';

describe('evaluateConditional', () => {
  afterEach(() => {
    _resetConditionalCache();
    vi.useRealTimers();
  });

  it('formula rules evaluate financial numeric functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 1, 0, 0.09);
    s = seedNumber(s, 1, 1, 0.11);
    s = seedNumber(s, 1, 2, 0.1);
    s = seedNumber(s, 2, 0, 1000);
    s = seedNumber(s, 2, 1, 2000);
    s = seedNumber(s, 2, 2, 3000);
    s = seedNumber(s, 2, 3, 4000);
    s = seedNumber(s, 3, 0, -120000);
    s = seedNumber(s, 3, 1, 39000);
    s = seedNumber(s, 3, 2, 30000);
    s = seedNumber(s, 3, 3, 21000);
    s = seedNumber(s, 3, 4, 37000);
    s = seedNumber(s, 3, 5, 46000);
    s = seedNumber(s, 4, 0, -10000);
    s = seedNumber(s, 4, 1, 2750);
    s = seedNumber(s, 4, 2, 4250);
    s = seedNumber(s, 4, 3, 3250);
    s = seedNumber(s, 4, 4, 2750);
    s = seedNumber(s, 5, 0, dateSerial(2024, 1, 1));
    s = seedNumber(s, 5, 1, dateSerial(2024, 3, 1));
    s = seedNumber(s, 5, 2, dateSerial(2024, 10, 30));
    s = seedNumber(s, 5, 3, dateSerial(2025, 2, 15));
    s = seedNumber(s, 5, 4, dateSerial(2025, 4, 1));
    s = seedNumber(s, 5, 5, dateSerial(2024, 7, 1));
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            formula:
              '=AND(ROUND(PMT(0.05/12,60,10000),6)=-188.712336,ROUND(PMT(0.05/12,60,10000,,1),6)=-187.929298)',
            apply: { fill: '#pmt' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(ROUND(PV(0.05/12,60,-188.712336),6)=9999.999977,PV(0,10,-100,0,0)=1000)',
            apply: { fill: '#pv' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula:
              '=AND(ROUND(FV(0.05/12,60,-100),6)=6800.608284,ROUND(FV(0.05/12,60,-100,,1),6)=6828.944152,FV(0,10,-100)=1000)',
            apply: { fill: '#fv' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula:
              '=AND(ROUND(NPER(0.05/12,-188.712336,10000),6)=60,NPER(0,-100,1000)=10,ROUND(NPER(0.05/12,-187.929298,10000,,1),6)=60)',
            apply: { fill: '#nper' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula:
              '=AND(ROUND(RATE(60,-188.712336,10000),9)=0.004166667,ROUND(RATE(60,-187.929298,10000,,1),9)=0.004166667,ROUND(RATE(10,-100,1000),9)=0)',
            apply: { fill: '#rate' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula:
              '=AND(ROUND(IPMT(0.05/12,1,60,10000),6)=-41.666667,ROUND(PPMT(0.05/12,1,60,10000),6)=-147.04567,ROUND(IPMT(0.05/12,2,60,10000),6)=-41.053976)',
            apply: { fill: '#ipmt-ppmt' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula:
              '=AND(IPMT(0.05/12,1,60,10000,,1)=0,ROUND(PPMT(0.05/12,1,60,10000,,1),6)=-187.929298,ROUND(IPMT(0.05/12,2,60,10000,,1),6)=-40.100589,PPMT(0,1,10,1000)=-100)',
            apply: { fill: '#ipmt-ppmt-type' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=PMT(0.05/12,0,10000)=0',
            apply: { fill: '#pmt-zero-periods' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=FV(0.05/12,60,-100,0,2)=0',
            apply: { fill: '#fv-invalid-type' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=NPER(0.05/12,0,10000)=0',
            apply: { fill: '#nper-zero-payment' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 10, r1: 0, c1: 10 },
            formula: '=RATE(0,-100,1000)=0',
            apply: { fill: '#rate-zero-periods' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 11, r1: 0, c1: 11 },
            formula: '=IPMT(0.05/12,0,60,10000)=0',
            apply: { fill: '#ipmt-invalid-period' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 12, r1: 0, c1: 12 },
            formula:
              '=AND(SLN(30000,7500,10)=2250,ROUND(SYD(30000,7500,10,1),6)=4090.909091,ROUND(SYD(30000,7500,10,10),6)=409.090909)',
            apply: { fill: '#depreciation-linear-syd' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 13, r1: 0, c1: 13 },
            formula:
              '=AND(DDB(2400,300,10,1)=480,DDB(2400,300,10,2)=384,DDB(2400,300,10,1,1.5)=360)',
            apply: { fill: '#depreciation-ddb' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 14, r1: 0, c1: 14 },
            formula: '=DDB(2400,300,10,0)=0',
            apply: { fill: '#ddb-invalid-period' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 15, r1: 0, c1: 15 },
            formula:
              '=AND(ROUND(DB(1000000,100000,6,1,7),2)=186083.33,ROUND(DB(1000000,100000,6,2,7),2)=259639.42,ROUND(DB(1000000,100000,6,7,7),2)=15845.1)',
            apply: { fill: '#depreciation-db' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 16, r1: 0, c1: 16 },
            formula: '=DB(1000000,100000,6,8,7)=0',
            apply: { fill: '#db-invalid-period' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 17, r1: 0, c1: 17 },
            formula: '=DB(1000000,100000,6,1,13)=0',
            apply: { fill: '#db-invalid-month' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 18, r1: 0, c1: 18 },
            formula:
              '=AND(ROUND(CUMIPMT(0.05/12,60,10000,1,12,0),6)=-458.995507,ROUND(CUMPRINC(0.05/12,60,10000,1,12,0),6)=-1805.55253)',
            apply: { fill: '#cumipmt-cumprinc' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 19, r1: 0, c1: 19 },
            formula:
              '=AND(ROUND(CUMIPMT(0.05/12,60,10000,1,12,1),6)=-406.802051,ROUND(CUMPRINC(0.05/12,60,10000,1,12,1),6)=-1848.349521)',
            apply: { fill: '#cumipmt-cumprinc-type' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 20, r1: 0, c1: 20 },
            formula: '=CUMIPMT(0,60,10000,1,12,0)=0',
            apply: { fill: '#cumipmt-invalid-rate' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 21, r1: 0, c1: 21 },
            formula: '=CUMPRINC(0.05/12,60,10000,12,1,0)=0',
            apply: { fill: '#cumprinc-invalid-periods' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 22, r1: 0, c1: 22 },
            formula:
              '=AND(ROUND(ISPMT(0.1/12,1,36,8000),6)=-64.814815,ROUND(ISPMT(0.1/12,0,36,8000),6)=-66.666667)',
            apply: { fill: '#ispmt' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 23, r1: 0, c1: 23 },
            formula:
              '=AND(ROUND(EFFECT(0.0525,4),9)=0.053542667,ROUND(NOMINAL(0.05354266737075822,4),9)=0.0525)',
            apply: { fill: '#effect-nominal' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 24, r1: 0, c1: 24 },
            formula: '=ISPMT(0.1/12,1,0,8000)=0',
            apply: { fill: '#ispmt-zero-periods' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 25, r1: 0, c1: 25 },
            formula: '=EFFECT(0.0525,0)=0',
            apply: { fill: '#effect-invalid-periods' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 26, r1: 0, c1: 26 },
            formula: '=NOMINAL(0,4)=0',
            apply: { fill: '#nominal-invalid-rate' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 27, r1: 0, c1: 27 },
            formula:
              '=AND(FVSCHEDULE(1000,$A$2:$C$2)=1330.89,ROUND(RRI(10,1000,10000),6)=0.258925,ROUND(PDURATION(0.1,1000,10000),6)=24.158858)',
            apply: { fill: '#fvschedule-rri-pduration' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 28, r1: 0, c1: 28 },
            formula: '=FVSCHEDULE("principal",$A$2:$C$2)=0',
            apply: { fill: '#fvschedule-invalid-principal' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 29, r1: 0, c1: 29 },
            formula: '=RRI(0,1000,10000)=0',
            apply: { fill: '#rri-invalid-periods' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 30, r1: 0, c1: 30 },
            formula: '=PDURATION(0,1000,10000)=0',
            apply: { fill: '#pduration-invalid-rate' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 31, r1: 0, c1: 31 },
            formula:
              '=AND(ROUND(NPV(0.1,-10000,3000,4200,6800),6)=1188.443412,ROUND(NPV(0.08,$A$3:$C$3,D3),6)=7962.219701)',
            apply: { fill: '#npv' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 32, r1: 0, c1: 32 },
            formula: '=NPV(-1,$A$3:$C$3)=0',
            apply: { fill: '#npv-invalid-rate' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 33, r1: 0, c1: 33 },
            formula: '=NPV(0.1,"bad")=0',
            apply: { fill: '#npv-invalid-value' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 34, r1: 0, c1: 34 },
            formula: '=ROUND(MIRR($A$4:$F$4,0.1,0.12),6)=0.126094',
            apply: { fill: '#mirr' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 35, r1: 0, c1: 35 },
            formula: '=MIRR($A$3:$D$3,0.1,0.12)=0',
            apply: { fill: '#mirr-no-negative' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 36, r1: 0, c1: 36 },
            formula: '=MIRR($A$4:$F$4,-1,0.12)=0',
            apply: { fill: '#mirr-invalid-rate' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 37, r1: 0, c1: 37 },
            formula: '=ROUND(XNPV(0.09,$A$5:$E$5,$A$6:$E$6),6)=2086.647602',
            apply: { fill: '#xnpv' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 38, r1: 0, c1: 38 },
            formula: '=XNPV(0.09,$A$5:$E$5,$A$6:$D$6)=0',
            apply: { fill: '#xnpv-mismatched-ranges' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 39, r1: 0, c1: 39 },
            formula: '=XNPV(0.09,$A$5:$E$5,$B$6:$F$6)=0',
            apply: { fill: '#xnpv-invalid-date' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 40, r1: 0, c1: 40 },
            formula:
              '=AND(ROUND(XIRR($A$5:$E$5,$A$6:$E$6),9)=0.373362534,ROUND(XIRR($A$5:$E$5,$A$6:$E$6,0.2),9)=0.373362534)',
            apply: { fill: '#xirr' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 41, r1: 0, c1: 41 },
            formula: '=XIRR($A$5:$E$5,$A$6:$D$6)=0',
            apply: { fill: '#xirr-mismatched-ranges' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 42, r1: 0, c1: 42 },
            formula: '=XIRR($A$3:$D$3,$A$6:$D$6)=0',
            apply: { fill: '#xirr-no-negative' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 43, r1: 0, c1: 43 },
            formula:
              '=AND(ROUND(IRR($A$4:$F$4),9)=0.130735539,ROUND(IRR($A$4:$F$4,0.2),9)=0.130735539)',
            apply: { fill: '#irr' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 44, r1: 0, c1: 44 },
            formula: '=IRR($A$3:$D$3)=0',
            apply: { fill: '#irr-no-negative' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 45, r1: 0, c1: 45 },
            formula: '=IRR($A$4:$F$4,-1)=0',
            apply: { fill: '#irr-invalid-guess' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 46, r1: 0, c1: 46 },
            formula:
              '=AND(DOLLARDE(1.02,16)=1.125,DOLLARFR(1.125,16)=1.02,DOLLARDE(-1.02,16)=-1.125)',
            apply: { fill: '#dollarde-dollarfr' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 47, r1: 0, c1: 47 },
            formula: '=DOLLARDE(1.02,0)=0',
            apply: { fill: '#dollarde-invalid-fraction' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 48, r1: 0, c1: 48 },
            formula: '=DOLLARFR(1.125,0)=0',
            apply: { fill: '#dollarfr-invalid-fraction' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 49, r1: 0, c1: 49 },
            formula:
              '=AND(ROUND(DISC($A$6,$F$6,97.5,100,3),9)=0.050137363,ROUND(INTRATE($A$6,$F$6,9700,10000,3),9)=0.062025603)',
            apply: { fill: '#disc-intrate' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 50, r1: 0, c1: 50 },
            formula: '=DISC($F$6,$A$6,97.5,100,3)=0',
            apply: { fill: '#disc-invalid-dates' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 51, r1: 0, c1: 51 },
            formula: '=INTRATE($A$6,$F$6,9700,10000,5)=0',
            apply: { fill: '#intrate-invalid-basis' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 52, r1: 0, c1: 52 },
            formula:
              '=AND(ROUND(PRICEDISC($A$6,$F$6,0.05,100,3),9)=97.506849315,ROUND(RECEIVED($A$6,$F$6,9700,0.05,3),6)=9948.019106)',
            apply: { fill: '#pricedisc-received' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 53, r1: 0, c1: 53 },
            formula: '=PRICEDISC($A$6,$F$6,0,100,3)=0',
            apply: { fill: '#pricedisc-invalid-discount' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 54, r1: 0, c1: 54 },
            formula: '=RECEIVED($A$6,$F$6,9700,0,3)=0',
            apply: { fill: '#received-invalid-discount' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 55, r1: 0, c1: 55 },
            formula:
              '=AND(ROUND(TBILLPRICE($A$6,$F$6,0.05),9)=97.472222222,ROUND(TBILLYIELD($A$6,$F$6,97.5),9)=0.050718512,ROUND(TBILLEQ($A$6,$F$6,0.05),9)=0.052009119)',
            apply: { fill: '#tbill' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 56, r1: 0, c1: 56 },
            formula: '=TBILLPRICE($F$6,$A$6,0.05)=0',
            apply: { fill: '#tbill-invalid-dates' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 57, r1: 0, c1: 57 },
            formula: '=TBILLYIELD($A$6,$F$6,0)=0',
            apply: { fill: '#tbill-invalid-price' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 58, r1: 0, c1: 58 },
            formula: '=TBILLEQ($A$6,$E$6,0.05)=0',
            apply: { fill: '#tbill-too-long' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 59, r1: 0, c1: 59 },
            formula:
              '=AND(ACCRINTM($A$6,$F$6,0.05)=25,ROUND(ACCRINTM($A$6,$F$6,0.05,1000,3),9)=24.931506849,ACCRINTM($A$6,$F$6,0.05,,)=25)',
            apply: { fill: '#accrintm' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 60, r1: 0, c1: 60 },
            formula: '=ACCRINTM($F$6,$A$6,0.05)=0',
            apply: { fill: '#accrintm-invalid-dates' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 61, r1: 0, c1: 61 },
            formula: '=ACCRINTM($A$6,$F$6,0)=0',
            apply: { fill: '#accrintm-invalid-rate' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#pmt');
    expect(overlay.get('0:0:1')?.fill).toBe('#pv');
    expect(overlay.get('0:0:2')?.fill).toBe('#fv');
    expect(overlay.get('0:0:3')?.fill).toBe('#nper');
    expect(overlay.get('0:0:4')?.fill).toBe('#rate');
    expect(overlay.get('0:0:5')?.fill).toBe('#ipmt-ppmt');
    expect(overlay.get('0:0:6')?.fill).toBe('#ipmt-ppmt-type');
    expect(overlay.get('0:0:7')?.fill).toBeUndefined();
    expect(overlay.get('0:0:8')?.fill).toBeUndefined();
    expect(overlay.get('0:0:9')?.fill).toBeUndefined();
    expect(overlay.get('0:0:10')?.fill).toBeUndefined();
    expect(overlay.get('0:0:11')?.fill).toBeUndefined();
    expect(overlay.get('0:0:12')?.fill).toBe('#depreciation-linear-syd');
    expect(overlay.get('0:0:13')?.fill).toBe('#depreciation-ddb');
    expect(overlay.get('0:0:14')?.fill).toBeUndefined();
    expect(overlay.get('0:0:15')?.fill).toBe('#depreciation-db');
    expect(overlay.get('0:0:16')?.fill).toBeUndefined();
    expect(overlay.get('0:0:17')?.fill).toBeUndefined();
    expect(overlay.get('0:0:18')?.fill).toBe('#cumipmt-cumprinc');
    expect(overlay.get('0:0:19')?.fill).toBe('#cumipmt-cumprinc-type');
    expect(overlay.get('0:0:20')?.fill).toBeUndefined();
    expect(overlay.get('0:0:21')?.fill).toBeUndefined();
    expect(overlay.get('0:0:22')?.fill).toBe('#ispmt');
    expect(overlay.get('0:0:23')?.fill).toBe('#effect-nominal');
    expect(overlay.get('0:0:24')?.fill).toBeUndefined();
    expect(overlay.get('0:0:25')?.fill).toBeUndefined();
    expect(overlay.get('0:0:26')?.fill).toBeUndefined();
    expect(overlay.get('0:0:27')?.fill).toBe('#fvschedule-rri-pduration');
    expect(overlay.get('0:0:28')?.fill).toBeUndefined();
    expect(overlay.get('0:0:29')?.fill).toBeUndefined();
    expect(overlay.get('0:0:30')?.fill).toBeUndefined();
    expect(overlay.get('0:0:31')?.fill).toBe('#npv');
    expect(overlay.get('0:0:32')?.fill).toBeUndefined();
    expect(overlay.get('0:0:33')?.fill).toBeUndefined();
    expect(overlay.get('0:0:34')?.fill).toBe('#mirr');
    expect(overlay.get('0:0:35')?.fill).toBeUndefined();
    expect(overlay.get('0:0:36')?.fill).toBeUndefined();
    expect(overlay.get('0:0:37')?.fill).toBe('#xnpv');
    expect(overlay.get('0:0:38')?.fill).toBeUndefined();
    expect(overlay.get('0:0:39')?.fill).toBeUndefined();
    expect(overlay.get('0:0:40')?.fill).toBe('#xirr');
    expect(overlay.get('0:0:41')?.fill).toBeUndefined();
    expect(overlay.get('0:0:42')?.fill).toBeUndefined();
    expect(overlay.get('0:0:43')?.fill).toBe('#irr');
    expect(overlay.get('0:0:44')?.fill).toBeUndefined();
    expect(overlay.get('0:0:45')?.fill).toBeUndefined();
    expect(overlay.get('0:0:46')?.fill).toBe('#dollarde-dollarfr');
    expect(overlay.get('0:0:47')?.fill).toBeUndefined();
    expect(overlay.get('0:0:48')?.fill).toBeUndefined();
    expect(overlay.get('0:0:49')?.fill).toBe('#disc-intrate');
    expect(overlay.get('0:0:50')?.fill).toBeUndefined();
    expect(overlay.get('0:0:51')?.fill).toBeUndefined();
    expect(overlay.get('0:0:52')?.fill).toBe('#pricedisc-received');
    expect(overlay.get('0:0:53')?.fill).toBeUndefined();
    expect(overlay.get('0:0:54')?.fill).toBeUndefined();
    expect(overlay.get('0:0:55')?.fill).toBe('#tbill');
    expect(overlay.get('0:0:56')?.fill).toBeUndefined();
    expect(overlay.get('0:0:57')?.fill).toBeUndefined();
    expect(overlay.get('0:0:58')?.fill).toBeUndefined();
    expect(overlay.get('0:0:59')?.fill).toBe('#accrintm');
    expect(overlay.get('0:0:60')?.fill).toBeUndefined();
    expect(overlay.get('0:0:61')?.fill).toBeUndefined();
  });

  it('formula rules evaluate financial functions over dynamic ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 0.09);
    s = seedNumber(s, 0, 1, 0.11);
    s = seedNumber(s, 0, 2, 0.1);
    s = seedNumber(s, 1, 0, -120000);
    s = seedNumber(s, 1, 1, 39000);
    s = seedNumber(s, 1, 2, 30000);
    s = seedNumber(s, 1, 3, 21000);
    s = seedNumber(s, 1, 4, 37000);
    s = seedNumber(s, 1, 5, 46000);
    s = seedNumber(s, 2, 0, -10000);
    s = seedNumber(s, 2, 1, 2750);
    s = seedNumber(s, 2, 2, 4250);
    s = seedNumber(s, 2, 3, 3250);
    s = seedNumber(s, 2, 4, 2750);
    s = seedNumber(s, 3, 0, dateSerial(2024, 1, 1));
    s = seedNumber(s, 3, 1, dateSerial(2024, 3, 1));
    s = seedNumber(s, 3, 2, dateSerial(2024, 10, 30));
    s = seedNumber(s, 3, 3, dateSerial(2025, 2, 15));
    s = seedNumber(s, 3, 4, dateSerial(2025, 4, 1));
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula:
              '=AND(FVSCHEDULE(1000,OFFSET(A1,0,0,1,3))=1330.89,ROUND(IRR(OFFSET(A2,0,0,1,6)),9)=0.130735539,ROUND(MIRR(INDIRECT("A2:F2"),0.1,0.12),6)=0.126094)',
            apply: { fill: '#financial-dynamic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 6, r1: 1, c1: 6 },
            formula:
              '=AND(ROUND(XNPV(0.09,OFFSET(A3,0,0,1,5),INDIRECT("A4:E4")),6)=2086.647602,ROUND(XIRR(INDIRECT("A3:E3"),OFFSET(A4,0,0,1,5)),9)=0.373362534)',
            apply: { fill: '#financial-dynamic-x' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 6, r1: 2, c1: 6 },
            formula: '=XNPV(0.09,OFFSET(A3,0,0,1,5),INDIRECT("A4:D4"))=0',
            apply: { fill: '#financial-dynamic-mismatch' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:6')?.fill).toBe('#financial-dynamic');
    expect(overlay.get('0:1:6')?.fill).toBe('#financial-dynamic-x');
    expect(overlay.get('0:2:6')?.fill).toBeUndefined();
  });
});
