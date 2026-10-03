import { afterEach, describe, expect, it, vi } from 'vitest';
import { _resetConditionalCache, evaluateConditional } from '../../../../src/render/conditional.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';
import { seedCell, seedNumber } from '../conditional-fixtures.js';

describe('evaluateConditional', () => {
  afterEach(() => {
    _resetConditionalCache();
    vi.useRealTimers();
  });

  it('formula rules evaluate exact MATCH over row and column ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'East' });
    s = seedCell(s, 3, 0, { kind: 'text', value: 'A*B' });
    s = seedNumber(s, 0, 3, 10);
    s = seedNumber(s, 0, 4, 20);
    s = seedNumber(s, 0, 5, 30);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 2, c1: 1 },
            formula: '=ISNUMBER(MATCH(A1,$A$1:$A$3,0))',
            apply: { fill: '#match-col' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=MATCH(20,$D$1:$F$1,0)=2',
            apply: { fill: '#match-row' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=ISNA(MATCH("West",$A$1:$A$3,0))',
            apply: { fill: '#match-na' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=ISNA(MATCH("North",$A$1:$A$3,1))',
            apply: { fill: '#match-unsupported' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=MATCH("No*",$A$1:$A$3,0)=1',
            apply: { fill: '#match-wildcard' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 10, r1: 0, c1: 10 },
            formula: '=MATCH("A~*B",$A$4:$A$4,0)=1',
            apply: { fill: '#match-escaped-wildcard' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#match-col');
    expect(overlay.get('0:1:1')?.fill).toBe('#match-col');
    expect(overlay.get('0:2:1')?.fill).toBe('#match-col');
    expect(overlay.get('0:0:6')?.fill).toBe('#match-row');
    expect(overlay.get('0:0:7')?.fill).toBe('#match-na');
    expect(overlay.get('0:0:8')?.fill).toBe('#match-unsupported');
    expect(overlay.get('0:0:9')?.fill).toBe('#match-wildcard');
    expect(overlay.get('0:0:10')?.fill).toBe('#match-escaped-wildcard');
  });

  it('formula rules evaluate approximate MATCH over monotonic ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 5);
    s = seedNumber(s, 1, 0, 10);
    s = seedNumber(s, 2, 0, 20);
    s = seedNumber(s, 3, 0, 40);
    s = seedNumber(s, 0, 1, 40);
    s = seedNumber(s, 1, 1, 20);
    s = seedNumber(s, 2, 1, 10);
    s = seedNumber(s, 3, 1, 5);
    s = seedCell(s, 0, 3, { kind: 'text', value: 'East' });
    s = seedCell(s, 0, 4, { kind: 'text', value: 'North' });
    s = seedCell(s, 0, 5, { kind: 'text', value: 'South' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=MATCH(17,$A$1:$A$4,1)=2',
            apply: { fill: '#match-ascending' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=MATCH(17,$B$1:$B$4,-1)=2',
            apply: { fill: '#match-descending' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=MATCH("Nor",$D$1:$F$1,1)=1',
            apply: { fill: '#match-text-ascending' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 6, r1: 1, c1: 6 },
            formula: '=MATCH(17,$A$1:$A$4)=2',
            apply: { fill: '#match-omitted-approx' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 3, c0: 2, r1: 3, c1: 2 },
            formula: '=ISNA(MATCH(17,$A$1:$B$2,1))',
            apply: { fill: '#match-not-one-dimensional' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#match-ascending');
    expect(overlay.get('0:1:2')?.fill).toBe('#match-descending');
    expect(overlay.get('0:0:6')?.fill).toBe('#match-text-ascending');
    expect(overlay.get('0:1:6')?.fill).toBe('#match-omitted-approx');
    expect(overlay.get('0:3:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate XMATCH exact, wildcard, reverse search, and INDEX composition', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'East' });
    s = seedCell(s, 3, 0, { kind: 'text', value: 'North' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = seedNumber(s, 2, 1, 4);
    s = seedNumber(s, 3, 1, 14);
    s = seedCell(s, 5, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 5, 1, { kind: 'text', value: 'South' });
    s = seedCell(s, 5, 2, { kind: 'text', value: 'East' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=XMATCH("South",$A$1:$A$4)=2',
            apply: { fill: '#xmatch-exact' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=XMATCH("No*",$A$1:$A$4,2)=1',
            apply: { fill: '#xmatch-wildcard' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=XMATCH("North",$A$1:$A$4,0,-1)=4',
            apply: { fill: '#xmatch-reverse' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=INDEX($B$1:$B$4,XMATCH("North",$A$1:$A$4,0,-1))=14',
            apply: { fill: '#xmatch-index' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=XMATCH("East",$A$6:$C$6)=3',
            apply: { fill: '#xmatch-horizontal' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=ISNA(XMATCH("South",$A$1:$A$4,-1))',
            apply: { fill: '#xmatch-unsupported' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=XMATCH("North",$A$1:$A$4,,-1)=4',
            apply: { fill: '#xmatch-omitted-match-mode' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#xmatch-exact');
    expect(overlay.get('0:0:4')?.fill).toBe('#xmatch-wildcard');
    expect(overlay.get('0:0:5')?.fill).toBe('#xmatch-reverse');
    expect(overlay.get('0:0:6')?.fill).toBe('#xmatch-index');
    expect(overlay.get('0:0:7')?.fill).toBe('#xmatch-horizontal');
    expect(overlay.get('0:0:8')?.fill).toBe('#xmatch-unsupported');
    expect(overlay.get('0:0:9')?.fill).toBe('#xmatch-omitted-match-mode');
  });

  it('formula rules evaluate approximate XMATCH over monotonic ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 5);
    s = seedNumber(s, 1, 0, 10);
    s = seedNumber(s, 2, 0, 20);
    s = seedNumber(s, 3, 0, 40);
    s = seedCell(s, 0, 1, { kind: 'text', value: 'East' });
    s = seedCell(s, 1, 1, { kind: 'text', value: 'North' });
    s = seedCell(s, 2, 1, { kind: 'text', value: 'South' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=XMATCH(17,$A$1:$A$4,-1)=2',
            apply: { fill: '#xmatch-next-smaller' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=XMATCH(17,$A$1:$A$4,1)=3',
            apply: { fill: '#xmatch-next-larger' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
            formula: '=XMATCH("Nor",$B$1:$B$3,-1)=1',
            apply: { fill: '#xmatch-text-next-smaller' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 3, c0: 2, r1: 3, c1: 2 },
            formula: '=XMATCH("Nor",$B$1:$B$3,1)=2',
            apply: { fill: '#xmatch-text-next-larger' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#xmatch-next-smaller');
    expect(overlay.get('0:1:2')?.fill).toBe('#xmatch-next-larger');
    expect(overlay.get('0:2:2')?.fill).toBe('#xmatch-text-next-smaller');
    expect(overlay.get('0:3:2')?.fill).toBe('#xmatch-text-next-larger');
  });

  it('formula rules evaluate lookup and rank functions over dynamic ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'East' });
    s = seedCell(s, 3, 0, { kind: 'text', value: 'West' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = seedNumber(s, 2, 1, 4);
    s = seedNumber(s, 3, 1, 14);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=MATCH("South",OFFSET(A1,0,0,4,1),0)=2',
            apply: { fill: '#match-offset' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=XMATCH("West",INDIRECT("A1:A4"))=4',
            apply: { fill: '#xmatch-indirect' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=INDEX(OFFSET(B1,0,0,4,1),XMATCH("West",A1:A4))=14',
            apply: { fill: '#index-offset' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=VLOOKUP("South",OFFSET(A1,0,0,4,2),2,FALSE)=8',
            apply: { fill: '#vlookup-offset' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=LARGE(INDIRECT("B1:B4"),2)=12',
            apply: { fill: '#large-indirect' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=RANK(8,OFFSET(B1,0,0,4,1),0)=3',
            apply: { fill: '#rank-offset' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=MATCH("South",OFFSET(A1,0,0,2,2),0)=2',
            apply: { fill: '#match-dynamic-not-one-dimensional' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#match-offset');
    expect(overlay.get('0:0:3')?.fill).toBe('#xmatch-indirect');
    expect(overlay.get('0:0:4')?.fill).toBe('#index-offset');
    expect(overlay.get('0:0:5')?.fill).toBe('#vlookup-offset');
    expect(overlay.get('0:0:6')?.fill).toBe('#large-indirect');
    expect(overlay.get('0:0:7')?.fill).toBe('#rank-offset');
    expect(overlay.get('0:0:8')?.fill).toBeUndefined();
  });

  it('formula rules evaluate scalar CHOOSE operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'East' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = seedNumber(s, 2, 1, 4);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=CHOOSE(2,"North","South","East")="South"',
            apply: { fill: '#choose-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=CHOOSE(XMATCH("South",$A$1:$A$3),$B$1,$B$2,$B$3)=8',
            apply: { fill: '#choose-xmatch' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=CHOOSE(3,1+1,2+2,3+3)=6',
            apply: { fill: '#choose-expression' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=ISERROR(CHOOSE(4,"North","South","East"))',
            apply: { fill: '#choose-error' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#choose-text');
    expect(overlay.get('0:0:4')?.fill).toBe('#choose-xmatch');
    expect(overlay.get('0:0:5')?.fill).toBe('#choose-expression');
    expect(overlay.get('0:0:6')?.fill).toBe('#choose-error');
  });

  it('formula rules evaluate scalar SWITCH operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'East' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = seedNumber(s, 2, 1, 4);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=SWITCH(A1,"North","Region N","South","Region S","Other")="Region N"',
            apply: { fill: '#switch-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=SWITCH(XMATCH("South",$A$1:$A$3),1,$B$1,2,$B$2,3,$B$3)=8',
            apply: { fill: '#switch-xmatch' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=SWITCH("West","North",1,"South",2,99)=99',
            apply: { fill: '#switch-default' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=ISNA(SWITCH("West","North",1,"South",2))',
            apply: { fill: '#switch-error' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#switch-text');
    expect(overlay.get('0:0:4')?.fill).toBe('#switch-xmatch');
    expect(overlay.get('0:0:5')?.fill).toBe('#switch-default');
    expect(overlay.get('0:0:6')?.fill).toBe('#switch-error');
  });

  it('formula rules evaluate scalar IFS operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'East' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = seedNumber(s, 2, 1, 4);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=IFS(A1="North","Region N",A1="South","Region S",TRUE,"Other")="Region N"',
            apply: { fill: '#ifs-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=IFS(AND(A1="North",B1>10),B1,TRUE,0)=12',
            apply: { fill: '#ifs-and' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 3, r1: 2, c1: 3 },
            formula: '=IFS(B3>10,"High",TRUE,"Low")="Low"',
            apply: { fill: '#ifs-fallback' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=ISNA(IFS(A1="West","Missing",A1="Central","Missing"))',
            apply: { fill: '#ifs-error' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#ifs-text');
    expect(overlay.get('0:0:4')?.fill).toBe('#ifs-and');
    expect(overlay.get('0:2:3')?.fill).toBe('#ifs-fallback');
    expect(overlay.get('0:0:5')?.fill).toBe('#ifs-error');
  });

  it('formula rules evaluate scalar IF result operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 1, c1: 2 },
            formula: '=IF(A1="North","Region N","Other")="Region N"',
            apply: { fill: '#if-text-result' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 1, c1: 3 },
            formula: '=IF(AND(A1="North",B1>10),B1,0)=12',
            apply: { fill: '#if-number-result' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=IF(TRUE(),42,1/0)=42',
            apply: { fill: '#if-short-circuit-true' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=IF(FALSE(),1/0,"fallback")="fallback"',
            apply: { fill: '#if-short-circuit-false' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=IF(A1="North",,99)=0',
            apply: { fill: '#if-omitted-true' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 6, r1: 1, c1: 6 },
            formula: '=IF(A2="North",99)=FALSE',
            apply: { fill: '#if-omitted-false' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#if-text-result');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBe('#if-number-result');
    expect(overlay.get('0:1:3')?.fill).toBeUndefined();
    expect(overlay.get('0:0:4')?.fill).toBe('#if-short-circuit-true');
    expect(overlay.get('0:0:5')?.fill).toBe('#if-short-circuit-false');
    expect(overlay.get('0:0:6')?.fill).toBe('#if-omitted-true');
    expect(overlay.get('0:1:6')?.fill).toBe('#if-omitted-false');
  });

  it('formula rules evaluate INDEX over vector and rectangular ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'East' });
    s = seedNumber(s, 0, 3, 10);
    s = seedNumber(s, 0, 4, 20);
    s = seedNumber(s, 0, 5, 30);
    s = seedCell(s, 3, 3, { kind: 'text', value: 'Red' });
    s = seedCell(s, 3, 4, { kind: 'text', value: 'Blue' });
    s = seedCell(s, 4, 3, { kind: 'text', value: 'Green' });
    s = seedCell(s, 4, 4, { kind: 'text', value: 'Gold' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=INDEX($A$1:$A$3,MATCH("South",$A$1:$A$3,0))="South"',
            apply: { fill: '#index-column' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=INDEX($D$1:$F$1,2)=20',
            apply: { fill: '#index-row' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=INDEX($D$4:$E$5,2,2)="Gold"',
            apply: { fill: '#index-rect' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=ISERROR(INDEX($D$4:$E$5,3,1))',
            apply: { fill: '#index-error' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#index-column');
    expect(overlay.get('0:0:6')?.fill).toBe('#index-row');
    expect(overlay.get('0:0:7')?.fill).toBe('#index-rect');
    expect(overlay.get('0:0:8')?.fill).toBe('#index-error');
  });

  it('formula rules evaluate scalar OFFSET references', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 18);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=OFFSET(A1,1,0)="South"',
            apply: { fill: '#offset-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 1, c1: 3 },
            formula: '=OFFSET(A1,0,1)=12',
            apply: { fill: '#offset-relative' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=OFFSET(A1,0,0,1,1)="North"',
            apply: { fill: '#offset-sized' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=OFFSET(A1,0,0,2,1)="North"',
            apply: { fill: '#offset-range' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=OFFSET(OFFSET(A1,1,0),0,0)="South"',
            apply: { fill: '#offset-nested-offset' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=OFFSET(INDIRECT("A1"),1,1)=18',
            apply: { fill: '#offset-nested-indirect' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#offset-text');
    expect(overlay.get('0:0:3')?.fill).toBe('#offset-relative');
    expect(overlay.get('0:1:3')?.fill).toBeUndefined();
    expect(overlay.get('0:0:4')?.fill).toBe('#offset-sized');
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
    expect(overlay.get('0:0:6')?.fill).toBe('#offset-nested-offset');
    expect(overlay.get('0:0:7')?.fill).toBe('#offset-nested-indirect');
  });

  it('formula rules evaluate aggregate dynamic ranges from OFFSET and INDIRECT', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 10);
    s = seedNumber(s, 1, 0, 20);
    s = seedCell(s, 2, 0, { kind: 'text', value: 'not numeric' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=SUM(OFFSET(A1,0,0,2,1))=30',
            apply: { fill: '#offset-range-sum' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=COUNT(OFFSET(A1,0,0,3,1))=2',
            apply: { fill: '#offset-range-count' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=SUM(INDIRECT("A1:A2"))=30',
            apply: { fill: '#indirect-range-a1' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=SUM(INDIRECT("R1C1:R2C1",FALSE))=30',
            apply: { fill: '#indirect-range-r1c1' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=SUM(INDIRECT("RC[-5]:R[1]C[-5]",FALSE))=30',
            apply: { fill: '#indirect-range-relative-r1c1' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=SUM(OFFSET(A1,0,0,10001,1))=30',
            apply: { fill: '#offset-range-too-large' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=SUM(INDIRECT("Sheet2!A1:A2"))=30',
            apply: { fill: '#indirect-range-other-sheet' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=SUM(OFFSET(OFFSET(A1,0,0,2,1),0,0,2,1))=30',
            apply: { fill: '#offset-range-nested-offset' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=SUM(OFFSET(INDIRECT("A1:A2"),0,0,2,1))=30',
            apply: { fill: '#offset-range-nested-indirect' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#offset-range-sum');
    expect(overlay.get('0:0:2')?.fill).toBe('#offset-range-count');
    expect(overlay.get('0:0:3')?.fill).toBe('#indirect-range-a1');
    expect(overlay.get('0:0:4')?.fill).toBe('#indirect-range-r1c1');
    expect(overlay.get('0:0:5')?.fill).toBe('#indirect-range-relative-r1c1');
    expect(overlay.get('0:0:6')?.fill).toBeUndefined();
    expect(overlay.get('0:0:7')?.fill).toBeUndefined();
    expect(overlay.get('0:0:8')?.fill).toBe('#offset-range-nested-offset');
    expect(overlay.get('0:0:9')?.fill).toBe('#offset-range-nested-indirect');
  });

  it('formula rules evaluate scalar INDIRECT A1 references', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedCell(s, 0, 1, { kind: 'text', value: 'A1' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=INDIRECT("A1")="North"',
            apply: { fill: '#indirect-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=INDIRECT(B1,TRUE)="North"',
            apply: { fill: '#indirect-ref-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 1, c1: 4 },
            formula: '=INDIRECT("A1")="North"',
            apply: { fill: '#indirect-fixed' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=INDIRECT("R1C1",FALSE)="North"',
            apply: { fill: '#indirect-r1c1' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=INDIRECT("Sheet2!A1")="North"',
            apply: { fill: '#indirect-other-sheet' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=INDIRECT("$A$1")="North"',
            apply: { fill: '#indirect-absolute' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=INDIRECT("Sheet1!A1")="North"',
            apply: { fill: '#indirect-current-sheet' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=INDIRECT("\'Sheet1\'!A1")="North"',
            apply: { fill: '#indirect-quoted-current-sheet' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 10, r1: 0, c1: 10 },
            formula: '=INDIRECT("Sheet1!R1C1",FALSE)="North"',
            apply: { fill: '#indirect-r1c1-current-sheet' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 11, r1: 0, c1: 11 },
            formula: '=INDIRECT("RC[-11]",FALSE)="North"',
            apply: { fill: '#indirect-r1c1-relative' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            formula: '=INDIRECT("RC",FALSE)="North"',
            apply: { fill: '#indirect-r1c1-current-cell' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 12, r1: 0, c1: 12 },
            formula: '=INDIRECT("R[-1]C",FALSE)="North"',
            apply: { fill: '#indirect-r1c1-out-of-bounds' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 13, r1: 0, c1: 13 },
            formula: '=INDIRECT("Sheet2!R1C1",FALSE)="North"',
            apply: { fill: '#indirect-r1c1-other-sheet' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#indirect-text');
    expect(overlay.get('0:0:3')?.fill).toBe('#indirect-ref-text');
    expect(overlay.get('0:0:4')?.fill).toBe('#indirect-fixed');
    expect(overlay.get('0:1:4')?.fill).toBe('#indirect-fixed');
    expect(overlay.get('0:0:5')?.fill).toBe('#indirect-r1c1');
    expect(overlay.get('0:0:6')?.fill).toBeUndefined();
    expect(overlay.get('0:0:7')?.fill).toBe('#indirect-absolute');
    expect(overlay.get('0:0:8')?.fill).toBe('#indirect-current-sheet');
    expect(overlay.get('0:0:9')?.fill).toBe('#indirect-quoted-current-sheet');
    expect(overlay.get('0:0:10')?.fill).toBe('#indirect-r1c1-current-sheet');
    expect(overlay.get('0:0:11')?.fill).toBe('#indirect-r1c1-relative');
    expect(overlay.get('0:0:0')?.fill).toBe('#indirect-r1c1-current-cell');
    expect(overlay.get('0:0:12')?.fill).toBeUndefined();
    expect(overlay.get('0:0:13')?.fill).toBeUndefined();
  });

  it('formula rules evaluate exact VLOOKUP and HLOOKUP table lookups', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedNumber(s, 0, 1, 12);
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedNumber(s, 1, 1, 8);
    s = seedCell(s, 2, 0, { kind: 'text', value: 'East' });
    s = seedNumber(s, 2, 1, 4);
    s = seedCell(s, 4, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 4, 1, { kind: 'text', value: 'South' });
    s = seedCell(s, 4, 2, { kind: 'text', value: 'East' });
    s = seedNumber(s, 5, 0, 12);
    s = seedNumber(s, 5, 1, 8);
    s = seedNumber(s, 5, 2, 4);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=VLOOKUP("South",$A$1:$B$3,2,FALSE)=8',
            apply: { fill: '#vlookup-exact' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=VLOOKUP("Nor*",$A$1:$B$3,2,0)=12',
            apply: { fill: '#vlookup-wildcard' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=HLOOKUP("East",$A$5:$C$6,2,FALSE)=4',
            apply: { fill: '#hlookup-exact' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=ISNA(VLOOKUP("South",$A$1:$B$3,2,TRUE))',
            apply: { fill: '#vlookup-unsupported' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#vlookup-exact');
    expect(overlay.get('0:0:4')?.fill).toBe('#vlookup-wildcard');
    expect(overlay.get('0:0:5')?.fill).toBe('#hlookup-exact');
    expect(overlay.get('0:0:6')?.fill).toBe('#vlookup-unsupported');
  });

  it('formula rules evaluate approximate VLOOKUP and HLOOKUP over sorted lookup axes', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 5);
    s = seedNumber(s, 0, 1, 50);
    s = seedNumber(s, 1, 0, 10);
    s = seedNumber(s, 1, 1, 100);
    s = seedNumber(s, 2, 0, 20);
    s = seedNumber(s, 2, 1, 200);
    s = seedNumber(s, 3, 0, 40);
    s = seedNumber(s, 3, 1, 400);
    s = seedCell(s, 5, 0, { kind: 'text', value: 'East' });
    s = seedCell(s, 5, 1, { kind: 'text', value: 'North' });
    s = seedCell(s, 5, 2, { kind: 'text', value: 'South' });
    s = seedNumber(s, 6, 0, 4);
    s = seedNumber(s, 6, 1, 12);
    s = seedNumber(s, 6, 2, 8);
    s = seedNumber(s, 0, 5, 10);
    s = seedNumber(s, 0, 6, 100);
    s = seedNumber(s, 1, 5, 5);
    s = seedNumber(s, 1, 6, 50);
    s = seedNumber(s, 2, 5, 20);
    s = seedNumber(s, 2, 6, 200);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=VLOOKUP(17,$A$1:$B$4,2,TRUE)=100',
            apply: { fill: '#vlookup-approx' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 3, r1: 1, c1: 3 },
            formula: '=HLOOKUP("Nor",$A$6:$C$7,2,TRUE)=4',
            apply: { fill: '#hlookup-approx-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 3, r1: 2, c1: 3 },
            formula: '=ISNA(VLOOKUP(17,$F$1:$G$3,2,TRUE))',
            apply: { fill: '#vlookup-approx-unsorted' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=VLOOKUP(17,$A$1:$B$4,2)=100',
            apply: { fill: '#vlookup-omitted-approx' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 4, r1: 1, c1: 4 },
            formula: '=HLOOKUP("Nor",$A$6:$C$7,2)=4',
            apply: { fill: '#hlookup-omitted-approx' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#vlookup-approx');
    expect(overlay.get('0:1:3')?.fill).toBe('#hlookup-approx-text');
    expect(overlay.get('0:2:3')?.fill).toBe('#vlookup-approx-unsorted');
    expect(overlay.get('0:0:4')?.fill).toBe('#vlookup-omitted-approx');
    expect(overlay.get('0:1:4')?.fill).toBe('#hlookup-omitted-approx');
  });

  it('formula rules evaluate XLOOKUP exact, wildcard, fallback, and reverse search', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'East' });
    s = seedCell(s, 3, 0, { kind: 'text', value: 'North' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = seedNumber(s, 2, 1, 4);
    s = seedNumber(s, 3, 1, 14);
    s = seedCell(s, 5, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 5, 1, { kind: 'text', value: 'South' });
    s = seedCell(s, 5, 2, { kind: 'text', value: 'East' });
    s = seedNumber(s, 6, 0, 12);
    s = seedNumber(s, 6, 1, 8);
    s = seedNumber(s, 6, 2, 4);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=XLOOKUP("South",$A$1:$A$4,$B$1:$B$4)=8',
            apply: { fill: '#xlookup-exact' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=XLOOKUP("No*",$A$1:$A$4,$B$1:$B$4,0,2)=12',
            apply: { fill: '#xlookup-wildcard' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=XLOOKUP("West",$A$1:$A$4,$B$1:$B$4,99)=99',
            apply: { fill: '#xlookup-fallback' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=XLOOKUP("North",$A$1:$A$4,$B$1:$B$4,0,0,-1)=14',
            apply: { fill: '#xlookup-reverse' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=XLOOKUP("East",$A$6:$C$6,$A$7:$C$7)=4',
            apply: { fill: '#xlookup-horizontal' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=ISNA(XLOOKUP("South",$A$1:$A$4,$B$1:$B$4,0,0,2))',
            apply: { fill: '#xlookup-unsupported' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=XLOOKUP("North",$A$1:$A$4,$B$1:$B$4,,0,-1)=14',
            apply: { fill: '#xlookup-omitted-if-not-found' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#xlookup-exact');
    expect(overlay.get('0:0:4')?.fill).toBe('#xlookup-wildcard');
    expect(overlay.get('0:0:5')?.fill).toBe('#xlookup-fallback');
    expect(overlay.get('0:0:6')?.fill).toBe('#xlookup-reverse');
    expect(overlay.get('0:0:7')?.fill).toBe('#xlookup-horizontal');
    expect(overlay.get('0:0:8')?.fill).toBe('#xlookup-unsupported');
    expect(overlay.get('0:0:9')?.fill).toBe('#xlookup-omitted-if-not-found');
  });

  it('formula rules evaluate XLOOKUP next smaller and next larger over monotonic ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 5);
    s = seedNumber(s, 1, 0, 10);
    s = seedNumber(s, 2, 0, 20);
    s = seedNumber(s, 3, 0, 40);
    s = seedNumber(s, 0, 1, 50);
    s = seedNumber(s, 1, 1, 100);
    s = seedNumber(s, 2, 1, 200);
    s = seedNumber(s, 3, 1, 400);
    s = seedCell(s, 5, 0, { kind: 'text', value: 'East' });
    s = seedCell(s, 5, 1, { kind: 'text', value: 'North' });
    s = seedCell(s, 5, 2, { kind: 'text', value: 'South' });
    s = seedNumber(s, 6, 0, 4);
    s = seedNumber(s, 6, 1, 12);
    s = seedNumber(s, 6, 2, 8);
    s = seedNumber(s, 0, 5, 10);
    s = seedNumber(s, 1, 5, 5);
    s = seedNumber(s, 2, 5, 20);
    s = seedNumber(s, 0, 6, 100);
    s = seedNumber(s, 1, 6, 50);
    s = seedNumber(s, 2, 6, 200);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=XLOOKUP(17,$A$1:$A$4,$B$1:$B$4,0,-1)=100',
            apply: { fill: '#xlookup-next-smaller' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=XLOOKUP(17,$A$1:$A$4,$B$1:$B$4,0,1)=200',
            apply: { fill: '#xlookup-next-larger' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
            formula: '=XLOOKUP("Nor",$A$6:$C$6,$A$7:$C$7,0,-1)=4',
            apply: { fill: '#xlookup-text-next-smaller' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 3, c0: 2, r1: 3, c1: 2 },
            formula: '=XLOOKUP("Nor",$A$6:$C$6,$A$7:$C$7,0,1)=12',
            apply: { fill: '#xlookup-text-next-larger' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 4, c0: 2, r1: 4, c1: 2 },
            formula: '=XLOOKUP(17,$F$1:$F$3,$G$1:$G$3,"missing",-1)="missing"',
            apply: { fill: '#xlookup-next-smaller-unsorted' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#xlookup-next-smaller');
    expect(overlay.get('0:1:2')?.fill).toBe('#xlookup-next-larger');
    expect(overlay.get('0:2:2')?.fill).toBe('#xlookup-text-next-smaller');
    expect(overlay.get('0:3:2')?.fill).toBe('#xlookup-text-next-larger');
    expect(overlay.get('0:4:2')?.fill).toBe('#xlookup-next-smaller-unsorted');
  });

  it('formula rules evaluate vector LOOKUP over monotonic ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 5);
    s = seedNumber(s, 1, 0, 10);
    s = seedNumber(s, 2, 0, 20);
    s = seedNumber(s, 3, 0, 40);
    s = seedNumber(s, 0, 1, 50);
    s = seedNumber(s, 1, 1, 100);
    s = seedNumber(s, 2, 1, 200);
    s = seedNumber(s, 3, 1, 400);
    s = seedCell(s, 5, 0, { kind: 'text', value: 'East' });
    s = seedCell(s, 5, 1, { kind: 'text', value: 'North' });
    s = seedCell(s, 5, 2, { kind: 'text', value: 'South' });
    s = seedNumber(s, 6, 0, 4);
    s = seedNumber(s, 6, 1, 12);
    s = seedNumber(s, 6, 2, 8);
    s = seedNumber(s, 0, 5, 10);
    s = seedNumber(s, 1, 5, 5);
    s = seedNumber(s, 2, 5, 20);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=LOOKUP(17,$A$1:$A$4,$B$1:$B$4)=100',
            apply: { fill: '#lookup-vector' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=LOOKUP(17,$A$1:$A$4)=10',
            apply: { fill: '#lookup-omitted-result' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
            formula: '=LOOKUP("Nor",$A$6:$C$6,$A$7:$C$7)=4',
            apply: { fill: '#lookup-horizontal-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 3, c0: 2, r1: 3, c1: 2 },
            formula: '=LOOKUP(1,$A$1:$A$4,$B$1:$B$4)=50',
            apply: { fill: '#lookup-too-small' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 4, c0: 2, r1: 4, c1: 2 },
            formula: '=LOOKUP(17,$F$1:$F$3,$B$1:$B$3)=100',
            apply: { fill: '#lookup-unsorted' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 5, c0: 2, r1: 5, c1: 2 },
            formula: '=LOOKUP(17,$A$1:$A$4,$B$1:$B$3)=100',
            apply: { fill: '#lookup-mismatch' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#lookup-vector');
    expect(overlay.get('0:1:2')?.fill).toBe('#lookup-omitted-result');
    expect(overlay.get('0:2:2')?.fill).toBe('#lookup-horizontal-text');
    expect(overlay.get('0:3:2')?.fill).toBeUndefined();
    expect(overlay.get('0:4:2')?.fill).toBeUndefined();
    expect(overlay.get('0:5:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate XLOOKUP and LOOKUP over dynamic ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'East' });
    s = seedCell(s, 3, 0, { kind: 'text', value: 'North' });
    s = seedNumber(s, 0, 1, 12);
    s = seedNumber(s, 1, 1, 8);
    s = seedNumber(s, 2, 1, 4);
    s = seedNumber(s, 3, 1, 14);
    s = seedNumber(s, 0, 4, 5);
    s = seedNumber(s, 1, 4, 10);
    s = seedNumber(s, 2, 4, 20);
    s = seedNumber(s, 3, 4, 40);
    s = seedNumber(s, 0, 5, 50);
    s = seedNumber(s, 1, 5, 100);
    s = seedNumber(s, 2, 5, 200);
    s = seedNumber(s, 3, 5, 400);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=XLOOKUP("South",OFFSET(A1,0,0,4,1),OFFSET(B1,0,0,4,1))=8',
            apply: { fill: '#xlookup-offset' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=XLOOKUP("North",INDIRECT("A1:A4"),INDIRECT("B1:B4"),0,0,-1)=14',
            apply: { fill: '#xlookup-indirect-reverse' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=XLOOKUP(17,OFFSET(E1,0,0,4,1),INDIRECT("F1:F4"),0,-1)=100',
            apply: { fill: '#xlookup-dynamic-approx' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=LOOKUP(17,OFFSET(E1,0,0,4,1),OFFSET(F1,0,0,4,1))=100',
            apply: { fill: '#lookup-offset' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=LOOKUP(17,INDIRECT("E1:E4"))=10',
            apply: { fill: '#lookup-indirect-self' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=XLOOKUP("South",OFFSET(A1,0,0,4,1),OFFSET(B1,0,0,2,1),0)=8',
            apply: { fill: '#xlookup-dynamic-size-mismatch' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#xlookup-offset');
    expect(overlay.get('0:0:3')?.fill).toBe('#xlookup-indirect-reverse');
    expect(overlay.get('0:0:6')?.fill).toBe('#xlookup-dynamic-approx');
    expect(overlay.get('0:0:7')?.fill).toBe('#lookup-offset');
    expect(overlay.get('0:0:8')?.fill).toBe('#lookup-indirect-self');
    expect(overlay.get('0:0:9')?.fill).toBeUndefined();
  });

  it('formula rules evaluate ROW/COLUMN position operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 },
            formula: '=MOD(ROW(),2)=0',
            apply: { fill: '#even-row' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 3 },
            formula: '=COLUMN()=3',
            apply: { fill: '#column-c' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 1, c1: 4 },
            formula: '=ROW(A1)=2',
            apply: { fill: '#relative-row-ref' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=AND(ROWS(A1:C5)=5,COLUMNS(A1:C5)=3,AREAS(A1:C5)=1,AREAS(A1)=1)',
            apply: { fill: '#range-dimensions' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 5, r1: 1, c1: 5 },
            formula:
              '=AND(ROWS(OFFSET(A1,0,0,5,3))=5,COLUMNS(INDIRECT("A1:C5"))=3,AREAS(OFFSET(A1,0,0,5,3))=1)',
            apply: { fill: '#range-dimensions-dynamic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 5, r1: 2, c1: 5 },
            formula: '=AND(ROW(OFFSET(A1,4,0))=5,COLUMN(INDIRECT("C1"))=3)',
            apply: { fill: '#position-dynamic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 3, c0: 5, r1: 3, c1: 5 },
            formula: '=ROW(OFFSET(A1,0,0,2,1))=1',
            apply: { fill: '#position-dynamic-multi' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBeUndefined();
    expect(overlay.get('0:1:0')?.fill).toBe('#even-row');
    expect(overlay.get('0:2:0')?.fill).toBeUndefined();
    expect(overlay.get('0:3:0')?.fill).toBe('#even-row');
    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBe('#column-c');
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
    expect(overlay.get('0:0:4')?.fill).toBeUndefined();
    expect(overlay.get('0:1:4')?.fill).toBe('#relative-row-ref');
    expect(overlay.get('0:0:5')?.fill).toBe('#range-dimensions');
    expect(overlay.get('0:1:5')?.fill).toBe('#range-dimensions-dynamic');
    expect(overlay.get('0:2:5')?.fill).toBe('#position-dynamic');
    expect(overlay.get('0:3:5')?.fill).toBeUndefined();
  });
});
