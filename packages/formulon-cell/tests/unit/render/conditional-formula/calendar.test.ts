import { afterEach, describe, expect, it, vi } from 'vitest';
import { _resetConditionalCache, evaluateConditional } from '../../../../src/render/conditional.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';
import { dateSerial, seedCell, seedNumber } from '../conditional-fixtures.js';

describe('evaluateConditional', () => {
  afterEach(() => {
    _resetConditionalCache();
    vi.useRealTimers();
  });

  it('formula rules evaluate DATE/YEAR/MONTH/DAY operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, dateSerial(2026, 7, 12));
    s = seedNumber(s, 1, 0, dateSerial(2025, 2, 1));
    s = seedNumber(s, 2, 0, dateSerial(2024, 12, 25));
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(YEAR(A1)=2026,MONTH(A1)=7,DAY(A1)=12,DATE(2026,7,12)=A1)',
            apply: { fill: '#date-parts' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=DATE(2024,14,1)=A2',
            apply: { fill: '#date-overflow' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
            formula: '=AND(DATEVALUE("2024-12-25")=A3,DATEVALUE("12/25/2024")=A3)',
            apply: { fill: '#datevalue' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
            formula: '=DATEVALUE("2024-02-31")=DATE(2024,3,2)',
            apply: { fill: '#datevalue-invalid' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#date-parts');
    expect(overlay.get('0:1:1')?.fill).toBe('#date-overflow');
    expect(overlay.get('0:2:1')?.fill).toBe('#datevalue');
    expect(overlay.get('0:2:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate TIME/HOUR/MINUTE/SECOND operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, dateSerial(2026, 7, 12) + (13 * 3600 + 5 * 60 + 9) / 86_400);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula:
              '=AND(HOUR(A1)=13,MINUTE(A1)=5,SECOND(A1)=9,ROUND(TIME(13,5,9)*86400,0)=47109)',
            apply: { fill: '#time-parts' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=TIME(25,0,0)=1/24',
            apply: { fill: '#time-overflow' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=TIME(-1,0,0)=0',
            apply: { fill: '#time-negative' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula:
              '=AND(ROUND(TIMEVALUE("13:05:09")*86400,0)=47109,TIMEVALUE("1:05 PM")=TIME(13,5,0))',
            apply: { fill: '#timevalue' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=TIMEVALUE("25:00")=1/24',
            apply: { fill: '#timevalue-invalid' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#time-parts');
    expect(overlay.get('0:0:2')?.fill).toBe('#time-overflow');
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
    expect(overlay.get('0:0:4')?.fill).toBe('#timevalue');
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
  });

  it('formula rules evaluate EDATE/EOMONTH date operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, dateSerial(2024, 1, 31));
    s = seedNumber(s, 1, 0, dateSerial(2026, 7, 12));
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(EDATE(A1,1)=DATE(2024,2,29),EOMONTH(A1,1)=DATE(2024,2,29))',
            apply: { fill: '#edate-eomonth-leap' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=AND(EDATE(A2,-1)=DATE(2026,6,12),EOMONTH(A2,0)=DATE(2026,7,31))',
            apply: { fill: '#edate-eomonth' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#edate-eomonth-leap');
    expect(overlay.get('0:1:1')?.fill).toBe('#edate-eomonth');
  });

  it('formula rules evaluate DAYS date operand', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, dateSerial(2026, 7, 12));
    s = seedNumber(s, 0, 1, dateSerial(2026, 7, 1));
    s = seedCell(s, 1, 0, { kind: 'text', value: 'not-a-date' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=DAYS(A1,B1)=11',
            apply: { fill: '#days' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=DAYS(A2,B1)=0',
            apply: { fill: '#days-text' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#days');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
  });

  it('formula rules evaluate DAYS360 date operand', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, dateSerial(2024, 1, 1));
    s = seedNumber(s, 0, 1, dateSerial(2024, 2, 1));
    s = seedNumber(s, 1, 0, dateSerial(2024, 2, 29));
    s = seedNumber(s, 1, 1, dateSerial(2024, 3, 31));
    s = seedCell(s, 2, 0, { kind: 'text', value: 'not-a-date' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=AND(DAYS360(A1,B1)=30,DAYS360(A1,B1,)=30)',
            apply: { fill: '#days360-default' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=AND(DAYS360(A2,B2)=30,DAYS360(A2,B2,TRUE())=31)',
            apply: { fill: '#days360-method' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
            formula: '=DAYS360(A3,B1)=0',
            apply: { fill: '#days360-text' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#days360-default');
    expect(overlay.get('0:1:2')?.fill).toBe('#days360-method');
    expect(overlay.get('0:2:1')?.fill).toBeUndefined();
  });

  it('formula rules evaluate DATEDIF date operand', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, dateSerial(2020, 1, 15));
    s = seedNumber(s, 0, 1, dateSerial(2023, 3, 20));
    s = seedNumber(s, 1, 0, dateSerial(2020, 3, 31));
    s = seedNumber(s, 1, 1, dateSerial(2020, 5, 2));
    s = seedCell(s, 2, 0, { kind: 'text', value: 'not-a-date' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=AND(DATEDIF(A1,B1,"Y")=3,DATEDIF(A1,B1,"M")=38,DATEDIF(A1,B1,"D")=1160)',
            apply: { fill: '#datedif-primary' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=AND(DATEDIF(A1,B1,"YM")=2,DATEDIF(A1,B1,"YD")=64,DATEDIF(A2,B2,"MD")=2)',
            apply: { fill: '#datedif-units' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
            formula: '=DATEDIF(A3,B1,"D")=0',
            apply: { fill: '#datedif-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
            formula: '=DATEDIF(B1,A1,"D")=0',
            apply: { fill: '#datedif-reversed' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 3, r1: 2, c1: 3 },
            formula: '=DATEDIF(A1,B1,"BAD")=0',
            apply: { fill: '#datedif-invalid-unit' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#datedif-primary');
    expect(overlay.get('0:1:2')?.fill).toBe('#datedif-units');
    expect(overlay.get('0:2:1')?.fill).toBeUndefined();
    expect(overlay.get('0:2:2')?.fill).toBeUndefined();
    expect(overlay.get('0:2:3')?.fill).toBeUndefined();
  });

  it('formula rules evaluate YEARFRAC date operand', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, dateSerial(2024, 1, 1));
    s = seedNumber(s, 0, 1, dateSerial(2024, 7, 1));
    s = seedNumber(s, 1, 0, dateSerial(2024, 2, 29));
    s = seedNumber(s, 1, 1, dateSerial(2024, 3, 31));
    s = seedCell(s, 2, 0, { kind: 'text', value: 'not-a-date' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula:
              '=AND(YEARFRAC(A1,B1)=0.5,YEARFRAC(A1,B1,)=0.5,ROUND(YEARFRAC(A1,B1,1),6)=0.497268)',
            apply: { fill: '#yearfrac-primary' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula:
              '=AND(ROUND(YEARFRAC(A1,B1,2),6)=0.505556,ROUND(YEARFRAC(A1,B1,3),6)=0.49863,ROUND(YEARFRAC(A2,B2,4),6)=0.086111)',
            apply: { fill: '#yearfrac-basis' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
            formula: '=YEARFRAC(A3,B1)=0',
            apply: { fill: '#yearfrac-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
            formula: '=YEARFRAC(A1,B1,5)=0',
            apply: { fill: '#yearfrac-invalid-basis' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#yearfrac-primary');
    expect(overlay.get('0:1:2')?.fill).toBe('#yearfrac-basis');
    expect(overlay.get('0:2:1')?.fill).toBeUndefined();
    expect(overlay.get('0:2:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate NETWORKDAYS and WORKDAY date operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, dateSerial(2024, 12, 23));
    s = seedNumber(s, 0, 1, dateSerial(2024, 12, 27));
    s = seedCell(s, 0, 2, { kind: 'text', value: 'north' });
    s = seedNumber(s, 0, 8, dateSerial(2024, 12, 25));
    s = seedNumber(s, 0, 9, dateSerial(2024, 12, 26));
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=AND(NETWORKDAYS(A1,B1)=5,NETWORKDAYS(B1,A1)=-5)',
            apply: { fill: '#networkdays' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=AND(WORKDAY(A1,5)=DATE(2024,12,30),WORKDAY(B1,-5)=DATE(2024,12,20))',
            apply: { fill: '#workday' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=NETWORKDAYS(A1,C1)=0',
            apply: { fill: '#networkdays-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=NETWORKDAYS(A1,B1,$I$1:$J$1)=3',
            apply: { fill: '#networkdays-holidays' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=WORKDAY(A1,2,$I$1:$J$1)=DATE(2024,12,27)',
            apply: { fill: '#workday-holidays' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 10, r1: 0, c1: 10 },
            formula: '=NETWORKDAYS.INTL(A1,B1,7)=4',
            apply: { fill: '#networkdays-intl-code' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 11, r1: 0, c1: 11 },
            formula: '=NETWORKDAYS.INTL(A1,B1,"0011000")=3',
            apply: { fill: '#networkdays-intl-mask' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 12, r1: 0, c1: 12 },
            formula: '=NETWORKDAYS.INTL(A1,B1,,$I$1:$J$1)=3',
            apply: { fill: '#networkdays-intl-omitted-weekend' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 13, r1: 0, c1: 13 },
            formula: '=WORKDAY.INTL(A1,3,7,$I$1:$J$1)=DATE(2024,12,30)',
            apply: { fill: '#workday-intl' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 14, r1: 0, c1: 14 },
            formula: '=NETWORKDAYS.INTL(A1,B1,"1111111")=0',
            apply: { fill: '#networkdays-intl-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 15, r1: 0, c1: 15 },
            formula: '=NETWORKDAYS(A1,B1,OFFSET(I1,0,0,1,2))=3',
            apply: { fill: '#networkdays-dynamic-holidays' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 16, r1: 0, c1: 16 },
            formula: '=WORKDAY(A1,2,INDIRECT("I1:J1"))=DATE(2024,12,27)',
            apply: { fill: '#workday-dynamic-holidays' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 17, r1: 0, c1: 17 },
            formula: '=NETWORKDAYS.INTL(A1,B1,,OFFSET(I1,0,0,1,2))=3',
            apply: { fill: '#networkdays-intl-dynamic-holidays' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#networkdays');
    expect(overlay.get('0:0:4')?.fill).toBe('#workday');
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
    expect(overlay.get('0:0:6')?.fill).toBe('#networkdays-holidays');
    expect(overlay.get('0:0:7')?.fill).toBe('#workday-holidays');
    expect(overlay.get('0:0:10')?.fill).toBe('#networkdays-intl-code');
    expect(overlay.get('0:0:11')?.fill).toBe('#networkdays-intl-mask');
    expect(overlay.get('0:0:12')?.fill).toBe('#networkdays-intl-omitted-weekend');
    expect(overlay.get('0:0:13')?.fill).toBe('#workday-intl');
    expect(overlay.get('0:0:14')?.fill).toBeUndefined();
    expect(overlay.get('0:0:15')?.fill).toBe('#networkdays-dynamic-holidays');
    expect(overlay.get('0:0:16')?.fill).toBe('#workday-dynamic-holidays');
    expect(overlay.get('0:0:17')?.fill).toBe('#networkdays-intl-dynamic-holidays');
  });

  it('formula rules evaluate WEEKDAY return types', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, dateSerial(2026, 7, 12)); // Sunday.
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(WEEKDAY(A1)=1,WEEKDAY(A1,2)=7,WEEKDAY(A1,3)=6,WEEKDAY(A1,17)=1)',
            apply: { fill: '#weekday' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=WEEKDAY(A1,99)=1',
            apply: { fill: '#weekday-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=WEEKDAY(A1,)=1',
            apply: { fill: '#weekday-omitted-return-type' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#weekday');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBe('#weekday-omitted-return-type');
  });

  it('formula rules evaluate WEEKNUM and ISOWEEKNUM date operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, dateSerial(2026, 1, 1));
    s = seedNumber(s, 1, 0, dateSerial(2026, 1, 4));
    s = seedNumber(s, 2, 0, dateSerial(2026, 1, 5));
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(WEEKNUM(A1)=1,WEEKNUM(A2)=2,WEEKNUM(A3,2)=2)',
            apply: { fill: '#weeknum' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=AND(ISOWEEKNUM(A1)=1,ISOWEEKNUM(A3)=2,WEEKNUM(A3,21)=2)',
            apply: { fill: '#isoweeknum' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
            formula: '=WEEKNUM(A1,99)=1',
            apply: { fill: '#weeknum-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=WEEKNUM(A1,)=1',
            apply: { fill: '#weeknum-omitted-return-type' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#weeknum');
    expect(overlay.get('0:1:1')?.fill).toBe('#isoweeknum');
    expect(overlay.get('0:2:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBe('#weeknum-omitted-return-type');
  });

  it('formula rules evaluate TODAY/NOW volatile date operands', () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date(Date.UTC(2026, 6, 12, 12, 0, 0)));
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, dateSerial(2026, 7, 11));
    s = seedNumber(s, 1, 0, dateSerial(2026, 7, 12));
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 1, c1: 1 },
            formula: '=A1<TODAY()',
            apply: { fill: '#before-today' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=AND(NOW()>TODAY(),NOW()<TODAY()+1)',
            apply: { fill: '#now-within-today' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#before-today');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBe('#now-within-today');
  });
});
