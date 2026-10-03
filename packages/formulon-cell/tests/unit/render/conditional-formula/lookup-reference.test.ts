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
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 2, c1: 1 },
        '=ISNUMBER(MATCH(A1,$A$1:$A$3,0))',
        '#match-col',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
        '=MATCH(20,$D$1:$F$1,0)=2',
        '#match-row',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
        '=ISNA(MATCH("West",$A$1:$A$3,0))',
        '#match-na',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
        '=ISNA(MATCH("North",$A$1:$A$3,1))',
        '#match-unsupported',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
        '=MATCH("No*",$A$1:$A$3,0)=1',
        '#match-wildcard',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 10, r1: 0, c1: 10 },
        '=MATCH("A~*B",$A$4:$A$4,0)=1',
        '#match-escaped-wildcard',
      ),
    ]);

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
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=MATCH(17,$A$1:$A$4,1)=2',
        '#match-ascending',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
        '=MATCH(17,$B$1:$B$4,-1)=2',
        '#match-descending',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
        '=MATCH("Nor",$D$1:$F$1,1)=1',
        '#match-text-ascending',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 6, r1: 1, c1: 6 },
        '=MATCH(17,$A$1:$A$4)=2',
        '#match-omitted-approx',
      ),
      formulaRule(
        { sheet: 0, r0: 3, c0: 2, r1: 3, c1: 2 },
        '=ISNA(MATCH(17,$A$1:$B$2,1))',
        '#match-not-one-dimensional',
      ),
    ]);

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
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=XMATCH("South",$A$1:$A$4)=2',
        '#xmatch-exact',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=XMATCH("No*",$A$1:$A$4,2)=1',
        '#xmatch-wildcard',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=XMATCH("North",$A$1:$A$4,0,-1)=4',
        '#xmatch-reverse',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
        '=INDEX($B$1:$B$4,XMATCH("North",$A$1:$A$4,0,-1))=14',
        '#xmatch-index',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
        '=XMATCH("East",$A$6:$C$6)=3',
        '#xmatch-horizontal',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
        '=ISNA(XMATCH("South",$A$1:$A$4,-1))',
        '#xmatch-unsupported',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
        '=XMATCH("North",$A$1:$A$4,,-1)=4',
        '#xmatch-omitted-match-mode',
      ),
    ]);

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
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=XMATCH(17,$A$1:$A$4,-1)=2',
        '#xmatch-next-smaller',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
        '=XMATCH(17,$A$1:$A$4,1)=3',
        '#xmatch-next-larger',
      ),
      formulaRule(
        { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
        '=XMATCH("Nor",$B$1:$B$3,-1)=1',
        '#xmatch-text-next-smaller',
      ),
      formulaRule(
        { sheet: 0, r0: 3, c0: 2, r1: 3, c1: 2 },
        '=XMATCH("Nor",$B$1:$B$3,1)=2',
        '#xmatch-text-next-larger',
      ),
    ]);

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
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=MATCH("South",OFFSET(A1,0,0,4,1),0)=2',
        '#match-offset',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=XMATCH("West",INDIRECT("A1:A4"))=4',
        '#xmatch-indirect',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=INDEX(OFFSET(B1,0,0,4,1),XMATCH("West",A1:A4))=14',
        '#index-offset',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=VLOOKUP("South",OFFSET(A1,0,0,4,2),2,FALSE)=8',
        '#vlookup-offset',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
        '=LARGE(INDIRECT("B1:B4"),2)=12',
        '#large-indirect',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
        '=RANK(8,OFFSET(B1,0,0,4,1),0)=3',
        '#rank-offset',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
        '=MATCH("South",OFFSET(A1,0,0,2,2),0)=2',
        '#match-dynamic-not-one-dimensional',
      ),
    ]);

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
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=CHOOSE(2,"North","South","East")="South"',
        '#choose-text',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=CHOOSE(XMATCH("South",$A$1:$A$3),$B$1,$B$2,$B$3)=8',
        '#choose-xmatch',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=CHOOSE(3,1+1,2+2,3+3)=6',
        '#choose-expression',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
        '=ISERROR(CHOOSE(4,"North","South","East"))',
        '#choose-error',
      ),
    ]);

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
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=SWITCH(A1,"North","Region N","South","Region S","Other")="Region N"',
        '#switch-text',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=SWITCH(XMATCH("South",$A$1:$A$3),1,$B$1,2,$B$2,3,$B$3)=8',
        '#switch-xmatch',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=SWITCH("West","North",1,"South",2,99)=99',
        '#switch-default',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
        '=ISNA(SWITCH("West","North",1,"South",2))',
        '#switch-error',
      ),
    ]);

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
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=IFS(A1="North","Region N",A1="South","Region S",TRUE,"Other")="Region N"',
        '#ifs-text',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=IFS(AND(A1="North",B1>10),B1,TRUE,0)=12',
        '#ifs-and',
      ),
      formulaRule(
        { sheet: 0, r0: 2, c0: 3, r1: 2, c1: 3 },
        '=IFS(B3>10,"High",TRUE,"Low")="Low"',
        '#ifs-fallback',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=ISNA(IFS(A1="West","Missing",A1="Central","Missing"))',
        '#ifs-error',
      ),
    ]);

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
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 1, c1: 2 },
        '=IF(A1="North","Region N","Other")="Region N"',
        '#if-text-result',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 1, c1: 3 },
        '=IF(AND(A1="North",B1>10),B1,0)=12',
        '#if-number-result',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=IF(TRUE(),42,1/0)=42',
        '#if-short-circuit-true',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=IF(FALSE(),1/0,"fallback")="fallback"',
        '#if-short-circuit-false',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
        '=IF(A1="North",,99)=0',
        '#if-omitted-true',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 6, r1: 1, c1: 6 },
        '=IF(A2="North",99)=FALSE',
        '#if-omitted-false',
      ),
    ]);

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
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=INDEX($A$1:$A$3,MATCH("South",$A$1:$A$3,0))="South"',
        '#index-column',
      ),
      formulaRule({ sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 }, '=INDEX($D$1:$F$1,2)=20', '#index-row'),
      formulaRule(
        { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
        '=INDEX($D$4:$E$5,2,2)="Gold"',
        '#index-rect',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
        '=ISERROR(INDEX($D$4:$E$5,3,1))',
        '#index-error',
      ),
    ]);

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
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=OFFSET(A1,1,0)="South"',
        '#offset-text',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 1, c1: 3 },
        '=OFFSET(A1,0,1)=12',
        '#offset-relative',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=OFFSET(A1,0,0,1,1)="North"',
        '#offset-sized',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=OFFSET(A1,0,0,2,1)="North"',
        '#offset-range',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
        '=OFFSET(OFFSET(A1,1,0),0,0)="South"',
        '#offset-nested-offset',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
        '=OFFSET(INDIRECT("A1"),1,1)=18',
        '#offset-nested-indirect',
      ),
    ]);

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
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=SUM(OFFSET(A1,0,0,2,1))=30',
        '#offset-range-sum',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=COUNT(OFFSET(A1,0,0,3,1))=2',
        '#offset-range-count',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=SUM(INDIRECT("A1:A2"))=30',
        '#indirect-range-a1',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=SUM(INDIRECT("R1C1:R2C1",FALSE))=30',
        '#indirect-range-r1c1',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=SUM(INDIRECT("RC[-5]:R[1]C[-5]",FALSE))=30',
        '#indirect-range-relative-r1c1',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
        '=SUM(OFFSET(A1,0,0,10001,1))=30',
        '#offset-range-too-large',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
        '=SUM(INDIRECT("Sheet2!A1:A2"))=30',
        '#indirect-range-other-sheet',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
        '=SUM(OFFSET(OFFSET(A1,0,0,2,1),0,0,2,1))=30',
        '#offset-range-nested-offset',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
        '=SUM(OFFSET(INDIRECT("A1:A2"),0,0,2,1))=30',
        '#offset-range-nested-indirect',
      ),
    ]);

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
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=INDIRECT("A1")="North"',
        '#indirect-text',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=INDIRECT(B1,TRUE)="North"',
        '#indirect-ref-text',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 1, c1: 4 },
        '=INDIRECT("A1")="North"',
        '#indirect-fixed',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=INDIRECT("R1C1",FALSE)="North"',
        '#indirect-r1c1',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
        '=INDIRECT("Sheet2!A1")="North"',
        '#indirect-other-sheet',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
        '=INDIRECT("$A$1")="North"',
        '#indirect-absolute',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
        '=INDIRECT("Sheet1!A1")="North"',
        '#indirect-current-sheet',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
        '=INDIRECT("\'Sheet1\'!A1")="North"',
        '#indirect-quoted-current-sheet',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 10, r1: 0, c1: 10 },
        '=INDIRECT("Sheet1!R1C1",FALSE)="North"',
        '#indirect-r1c1-current-sheet',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 11, r1: 0, c1: 11 },
        '=INDIRECT("RC[-11]",FALSE)="North"',
        '#indirect-r1c1-relative',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        '=INDIRECT("RC",FALSE)="North"',
        '#indirect-r1c1-current-cell',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 12, r1: 0, c1: 12 },
        '=INDIRECT("R[-1]C",FALSE)="North"',
        '#indirect-r1c1-out-of-bounds',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 13, r1: 0, c1: 13 },
        '=INDIRECT("Sheet2!R1C1",FALSE)="North"',
        '#indirect-r1c1-other-sheet',
      ),
    ]);

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
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=VLOOKUP("South",$A$1:$B$3,2,FALSE)=8',
        '#vlookup-exact',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=VLOOKUP("Nor*",$A$1:$B$3,2,0)=12',
        '#vlookup-wildcard',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=HLOOKUP("East",$A$5:$C$6,2,FALSE)=4',
        '#hlookup-exact',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
        '=ISNA(VLOOKUP("South",$A$1:$B$3,2,TRUE))',
        '#vlookup-unsupported',
      ),
    ]);

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
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=VLOOKUP(17,$A$1:$B$4,2,TRUE)=100',
        '#vlookup-approx',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 3, r1: 1, c1: 3 },
        '=HLOOKUP("Nor",$A$6:$C$7,2,TRUE)=4',
        '#hlookup-approx-text',
      ),
      formulaRule(
        { sheet: 0, r0: 2, c0: 3, r1: 2, c1: 3 },
        '=ISNA(VLOOKUP(17,$F$1:$G$3,2,TRUE))',
        '#vlookup-approx-unsorted',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=VLOOKUP(17,$A$1:$B$4,2)=100',
        '#vlookup-omitted-approx',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 4, r1: 1, c1: 4 },
        '=HLOOKUP("Nor",$A$6:$C$7,2)=4',
        '#hlookup-omitted-approx',
      ),
    ]);

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
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=XLOOKUP("South",$A$1:$A$4,$B$1:$B$4)=8',
        '#xlookup-exact',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=XLOOKUP("No*",$A$1:$A$4,$B$1:$B$4,0,2)=12',
        '#xlookup-wildcard',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=XLOOKUP("West",$A$1:$A$4,$B$1:$B$4,99)=99',
        '#xlookup-fallback',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
        '=XLOOKUP("North",$A$1:$A$4,$B$1:$B$4,0,0,-1)=14',
        '#xlookup-reverse',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
        '=XLOOKUP("East",$A$6:$C$6,$A$7:$C$7)=4',
        '#xlookup-horizontal',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
        '=ISNA(XLOOKUP("South",$A$1:$A$4,$B$1:$B$4,0,0,2))',
        '#xlookup-unsupported',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
        '=XLOOKUP("North",$A$1:$A$4,$B$1:$B$4,,0,-1)=14',
        '#xlookup-omitted-if-not-found',
      ),
    ]);

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
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=XLOOKUP(17,$A$1:$A$4,$B$1:$B$4,0,-1)=100',
        '#xlookup-next-smaller',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
        '=XLOOKUP(17,$A$1:$A$4,$B$1:$B$4,0,1)=200',
        '#xlookup-next-larger',
      ),
      formulaRule(
        { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
        '=XLOOKUP("Nor",$A$6:$C$6,$A$7:$C$7,0,-1)=4',
        '#xlookup-text-next-smaller',
      ),
      formulaRule(
        { sheet: 0, r0: 3, c0: 2, r1: 3, c1: 2 },
        '=XLOOKUP("Nor",$A$6:$C$6,$A$7:$C$7,0,1)=12',
        '#xlookup-text-next-larger',
      ),
      formulaRule(
        { sheet: 0, r0: 4, c0: 2, r1: 4, c1: 2 },
        '=XLOOKUP(17,$F$1:$F$3,$G$1:$G$3,"missing",-1)="missing"',
        '#xlookup-next-smaller-unsorted',
      ),
    ]);

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
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=LOOKUP(17,$A$1:$A$4,$B$1:$B$4)=100',
        '#lookup-vector',
      ),
      formulaRule(
        { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
        '=LOOKUP(17,$A$1:$A$4)=10',
        '#lookup-omitted-result',
      ),
      formulaRule(
        { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
        '=LOOKUP("Nor",$A$6:$C$6,$A$7:$C$7)=4',
        '#lookup-horizontal-text',
      ),
      formulaRule(
        { sheet: 0, r0: 3, c0: 2, r1: 3, c1: 2 },
        '=LOOKUP(1,$A$1:$A$4,$B$1:$B$4)=50',
        '#lookup-too-small',
      ),
      formulaRule(
        { sheet: 0, r0: 4, c0: 2, r1: 4, c1: 2 },
        '=LOOKUP(17,$F$1:$F$3,$B$1:$B$3)=100',
        '#lookup-unsorted',
      ),
      formulaRule(
        { sheet: 0, r0: 5, c0: 2, r1: 5, c1: 2 },
        '=LOOKUP(17,$A$1:$A$4,$B$1:$B$3)=100',
        '#lookup-mismatch',
      ),
    ]);

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
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=XLOOKUP("South",OFFSET(A1,0,0,4,1),OFFSET(B1,0,0,4,1))=8',
        '#xlookup-offset',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=XLOOKUP("North",INDIRECT("A1:A4"),INDIRECT("B1:B4"),0,0,-1)=14',
        '#xlookup-indirect-reverse',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
        '=XLOOKUP(17,OFFSET(E1,0,0,4,1),INDIRECT("F1:F4"),0,-1)=100',
        '#xlookup-dynamic-approx',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
        '=LOOKUP(17,OFFSET(E1,0,0,4,1),OFFSET(F1,0,0,4,1))=100',
        '#lookup-offset',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
        '=LOOKUP(17,INDIRECT("E1:E4"))=10',
        '#lookup-indirect-self',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
        '=XLOOKUP("South",OFFSET(A1,0,0,4,1),OFFSET(B1,0,0,2,1),0)=8',
        '#xlookup-dynamic-size-mismatch',
      ),
    ]);

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
    s = withConditionalRules(s, [
      formulaRule({ sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 }, '=MOD(ROW(),2)=0', '#even-row'),
      formulaRule({ sheet: 0, r0: 0, c0: 1, r1: 0, c1: 3 }, '=COLUMN()=3', '#column-c'),
      formulaRule({ sheet: 0, r0: 0, c0: 4, r1: 1, c1: 4 }, '=ROW(A1)=2', '#relative-row-ref'),
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=AND(ROWS(A1:C5)=5,COLUMNS(A1:C5)=3,AREAS(A1:C5)=1,AREAS(A1)=1)',
        '#range-dimensions',
      ),
      {
        kind: 'formula',
        range: { sheet: 0, r0: 1, c0: 5, r1: 1, c1: 5 },
        formula:
          '=AND(ROWS(OFFSET(A1,0,0,5,3))=5,COLUMNS(INDIRECT("A1:C5"))=3,AREAS(OFFSET(A1,0,0,5,3))=1)',
        apply: { fill: '#range-dimensions-dynamic' },
      },
      formulaRule(
        { sheet: 0, r0: 2, c0: 5, r1: 2, c1: 5 },
        '=AND(ROW(OFFSET(A1,4,0))=5,COLUMN(INDIRECT("C1"))=3)',
        '#position-dynamic',
      ),
      formulaRule(
        { sheet: 0, r0: 3, c0: 5, r1: 3, c1: 5 },
        '=ROW(OFFSET(A1,0,0,2,1))=1',
        '#position-dynamic-multi',
      ),
    ]);

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
