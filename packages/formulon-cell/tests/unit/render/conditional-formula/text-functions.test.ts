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

  it('formula rules evaluate SEARCH/FIND/LEN text operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North Region' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'south region' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'NE' });
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 2, c1: 1 },
        '=AND(ISNUMBER(SEARCH("region",A1)),LEN(A1)>5)',
        '#search',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 2, c1: 2 },
        '=ISNUMBER(FIND("Region",A1))',
        '#find',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#search');
    expect(overlay.get('0:1:1')?.fill).toBe('#search');
    expect(overlay.get('0:2:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBe('#find');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
  });

  it('formula SEARCH start position is one-based and fails closed when missing', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'north north' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'north south' });
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 1, c1: 1 },
        '=SEARCH("north",A1,7)=7',
        '#search-start',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=SEARCH("north",A1,)=1',
        '#search-omitted-start',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#search-start');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBe('#search-omitted-start');
  });

  it('formula SEARCH supports Excel wildcards while FIND treats them literally', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North Region' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'Q1*Plan' });
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=AND(SEARCH("N*r",A1)=1,SEARCH("r?g",A1)=7)',
        '#search-wildcard',
      ),
      formulaRule({ sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 }, '=SEARCH("~*",A2)=3', '#search-escape'),
      formulaRule({ sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 }, '=FIND("*",A2)=3', '#find-literal'),
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=FIND("Region",A1,)=7',
        '#find-omitted-start',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#search-wildcard');
    expect(overlay.get('0:1:1')?.fill).toBe('#search-escape');
    expect(overlay.get('0:1:2')?.fill).toBe('#find-literal');
    expect(overlay.get('0:0:3')?.fill).toBe('#find-omitted-start');
  });

  it('formula rules evaluate scalar HYPERLINK display values', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = withConditionalRules(s, [
      {
        kind: 'formula',
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        formula:
          '=AND(HYPERLINK("https://example.test","Example")="Example",HYPERLINK("#A1")="#A1",HYPERLINK("#A1",42)=42)',
        apply: { fill: '#hyperlink' },
      },
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=HYPERLINK(NA())=""',
        '#hyperlink-invalid',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#hyperlink');
    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
  });

  it('formula rules evaluate limited CELL information operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedNumber(s, 1, 0, 42);
    s = withConditionalRules(s, [
      {
        kind: 'formula',
        range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        formula:
          '=AND(CELL("address",A1)="$A$1",CELL("row",A1)=1,CELL("col",A1)=1,CELL("contents",A1)="North",CELL("type",A1)="l")',
        apply: { fill: '#cell-info-text' },
      },
      formulaRule(
        { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
        '=AND(CELL("contents",A2)=42,CELL("type",A2)="v")',
        '#cell-info-number',
      ),
      formulaRule(
        { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
        '=AND(CELL("address")="$B$3",CELL("type")="b")',
        '#cell-info-current',
      ),
      formulaRule(
        { sheet: 0, r0: 3, c0: 1, r1: 3, c1: 1 },
        '=CELL("filename",A1)=""',
        '#cell-info-unsupported',
      ),
      {
        kind: 'formula',
        range: { sheet: 0, r0: 4, c0: 1, r1: 4, c1: 1 },
        formula: '=AND(CELL("address",OFFSET(A1,1,0))="$A$2",CELL("contents",INDIRECT("A2"))=42)',
        apply: { fill: '#cell-info-dynamic' },
      },
      formulaRule(
        { sheet: 0, r0: 5, c0: 1, r1: 5, c1: 1 },
        '=CELL("address",OFFSET(A1,0,0,2,1))="$A$1"',
        '#cell-info-dynamic-multi',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#cell-info-text');
    expect(overlay.get('0:1:1')?.fill).toBe('#cell-info-number');
    expect(overlay.get('0:2:1')?.fill).toBe('#cell-info-current');
    expect(overlay.get('0:3:1')?.fill).toBeUndefined();
    expect(overlay.get('0:4:1')?.fill).toBe('#cell-info-dynamic');
    expect(overlay.get('0:5:1')?.fill).toBeUndefined();
  });

  it('formula rules evaluate limited SHEET and SHEETS information operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        '=AND(SHEET()=1,SHEET(A1)=1,SHEET(A1:B2)=1)',
        '#sheet-info',
      ),
      formulaRule({ sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }, '=SHEETS(A1:B2)=1', '#sheets-info'),
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=SHEETS()=1',
        '#sheets-workbook-count',
      ),
      formulaRule({ sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 }, '=SHEET(Sheet2!A1)=2', '#sheet-other'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#sheet-info');
    expect(overlay.get('0:0:1')?.fill).toBe('#sheets-info');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
  });

  it('formula rules evaluate limited SHEET and SHEETS over dynamic ranges', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        '=AND(SHEET(OFFSET(A1,0,0,2,2))=1,SHEETS(INDIRECT("A1:B2"))=1)',
        '#sheet-info-dynamic',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=SHEET(INDIRECT("Sheet2!A1"))=2',
        '#sheet-info-dynamic-other',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#sheet-info-dynamic');
    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
  });

  it('formula rules evaluate ISLOGICAL/ISNONTEXT/ISFORMULA predicates', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'bool', value: true });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'text' });
    const cells = new Map(s.data.cells);
    cells.set('0:2:0', { value: { kind: 'number', value: 3 }, formula: '=SUM(1,2)' });
    s = {
      ...s,
      data: { ...s.data, cells },
      conditional: {
        ...s.conditional,
        rules: [
          formulaRule({ sheet: 0, r0: 0, c0: 1, r1: 2, c1: 1 }, '=ISLOGICAL(A1)', '#logical'),
          formulaRule({ sheet: 0, r0: 0, c0: 2, r1: 2, c1: 2 }, '=ISNONTEXT(A1)', '#nontext'),
          formulaRule({ sheet: 0, r0: 0, c0: 3, r1: 2, c1: 3 }, '=ISFORMULA(A1)', '#formula'),
          formulaRule(
            { sheet: 0, r0: 0, c0: 4, r1: 2, c1: 4 },
            '=EXACT(FORMULATEXT(A1),"=SUM(1,2)")',
            '#formulatext',
          ),
          formulaRule(
            { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            '=AND(TYPE(A1)=4,TYPE(A2)=2,TYPE(A3)=1,TYPE(NA())=16)',
            '#type-codes',
          ),
          formulaRule(
            { sheet: 0, r0: 1, c0: 5, r1: 1, c1: 5 },
            '=EXACT(FORMULATEXT(OFFSET(A1,2,0)),"=SUM(1,2)")',
            '#formulatext-dynamic',
          ),
          formulaRule(
            { sheet: 0, r0: 2, c0: 5, r1: 2, c1: 5 },
            '=FORMULATEXT(OFFSET(A1,0,0,2,1))=""',
            '#formulatext-dynamic-multi',
          ),
          formulaRule(
            { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            '=ISFORMULA(1)',
            '#formula-literal',
          ),
          formulaRule(
            { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            '=ISFORMULA(OFFSET($A$1,2,0))',
            '#formula-dynamic-offset',
          ),
          formulaRule(
            { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            '=ISFORMULA(INDIRECT("$A$3"))',
            '#formula-dynamic-indirect',
          ),
          formulaRule(
            { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            '=ISFORMULA(OFFSET($A$1,0,0,2,1))',
            '#formula-dynamic-multi',
          ),
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#logical');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
    expect(overlay.get('0:2:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBe('#nontext');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
    expect(overlay.get('0:2:2')?.fill).toBe('#nontext');
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
    expect(overlay.get('0:1:3')?.fill).toBeUndefined();
    expect(overlay.get('0:2:3')?.fill).toBe('#formula');
    expect(overlay.get('0:0:4')?.fill).toBeUndefined();
    expect(overlay.get('0:1:4')?.fill).toBeUndefined();
    expect(overlay.get('0:2:4')?.fill).toBe('#formulatext');
    expect(overlay.get('0:0:5')?.fill).toBe('#type-codes');
    expect(overlay.get('0:1:5')?.fill).toBe('#formulatext-dynamic');
    expect(overlay.get('0:2:5')?.fill).toBeUndefined();
    expect(overlay.get('0:0:6')?.fill).toBeUndefined();
    expect(overlay.get('0:0:7')?.fill).toBe('#formula-dynamic-offset');
    expect(overlay.get('0:0:8')?.fill).toBe('#formula-dynamic-indirect');
    expect(overlay.get('0:0:9')?.fill).toBeUndefined();
  });

  it('formula rules evaluate ISREF and NA error operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'error', code: 6, text: '#N/A' });
    s = seedCell(s, 1, 0, { kind: 'error', code: 1, text: '#DIV/0!' });
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=AND(ISREF(A1),ISREF(A1:B2),NOT(ISREF(1)))',
        '#isref',
      ),
      formulaRule({ sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 }, '=ISERROR(NA())', '#na-error'),
      formulaRule({ sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 }, '=NA()=A1', '#na-equals'),
      formulaRule({ sheet: 0, r0: 0, c0: 4, r1: 1, c1: 4 }, '=ISNA(A1)', '#isna'),
      formulaRule({ sheet: 0, r0: 0, c0: 5, r1: 1, c1: 5 }, '=ISERR(A1)', '#iserr'),
      {
        kind: 'formula',
        range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
        formula: '=AND(ERROR.TYPE(A1)=7,ERROR.TYPE(A2)=2,ERROR.TYPE(1/0)=2,ERROR.TYPE(SQRT(-1))=6)',
        apply: { fill: '#error-type' },
      },
      formulaRule(
        { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
        '=ERROR.TYPE(1)=0',
        '#error-type-non-error',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
        '=AND(ISREF(OFFSET(A1,0,0,1,1)),ISREF(INDIRECT("A1:B2")),NOT(ISREF(1)))',
        '#isref-dynamic',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#isref');
    expect(overlay.get('0:0:2')?.fill).toBe('#na-error');
    expect(overlay.get('0:0:3')?.fill).toBe('#na-equals');
    expect(overlay.get('0:0:4')?.fill).toBe('#isna');
    expect(overlay.get('0:1:4')?.fill).toBeUndefined();
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
    expect(overlay.get('0:1:5')?.fill).toBe('#iserr');
    expect(overlay.get('0:0:6')?.fill).toBe('#error-type');
    expect(overlay.get('0:0:7')?.fill).toBeUndefined();
    expect(overlay.get('0:0:8')?.fill).toBe('#isref-dynamic');
  });

  it('formula rules evaluate IFERROR and IFNA fallback operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North Region' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South Area' });
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 1, c1: 1 },
        '=IFERROR(SEARCH("region",A1),0)>0',
        '#iferror-search',
      ),
      formulaRule({ sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 }, '=IFNA(NA(),42)=42', '#ifna-na'),
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=ISERR(IFNA(1/0,42))',
        '#ifna-preserves-other-errors',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#iferror-search');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBe('#ifna-na');
    expect(overlay.get('0:0:3')?.fill).toBe('#ifna-preserves-other-errors');
  });

  it('formula rules evaluate LEFT/RIGHT/MID text operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North Region' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South Region' });
    s = seedCell(s, 2, 0, { kind: 'text', value: 'North Area' });
    s = withConditionalRules(s, [
      formulaRule({ sheet: 0, r0: 0, c0: 1, r1: 2, c1: 1 }, '=LEFT(A1,5)="North"', '#left'),
      formulaRule({ sheet: 0, r0: 0, c0: 2, r1: 2, c1: 2 }, '=RIGHT(A1,6)="Region"', '#right'),
      formulaRule({ sheet: 0, r0: 0, c0: 3, r1: 2, c1: 3 }, '=MID(A1,7,6)="Region"', '#mid'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#left');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
    expect(overlay.get('0:2:1')?.fill).toBe('#left');
    expect(overlay.get('0:0:2')?.fill).toBe('#right');
    expect(overlay.get('0:1:2')?.fill).toBe('#right');
    expect(overlay.get('0:2:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBe('#mid');
    expect(overlay.get('0:1:3')?.fill).toBe('#mid');
    expect(overlay.get('0:2:3')?.fill).toBeUndefined();
  });

  it('formula text slices default to one character and fail closed on invalid arguments', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = withConditionalRules(s, [
      formulaRule({ sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }, '=LEFT(A1)="N"', '#left-default'),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=AND(LEFT(A1,)="N",RIGHT(A1,)="h")',
        '#slice-omitted-count',
      ),
      formulaRule({ sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 }, '=RIGHT(A1,-1)=""', '#right-invalid'),
      formulaRule({ sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 }, '=MID(A1,0,1)="N"', '#mid-invalid'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#left-default');
    expect(overlay.get('0:0:4')?.fill).toBe('#slice-omitted-count');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
  });

  it('formula rules evaluate LOWER/UPPER/TRIM text transforms', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: '  North   Region  ' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'South Region' });
    s = seedCell(s, 2, 0, { kind: 'text', value: ' 1,234.5 ' });
    s = seedCell(s, 3, 0, { kind: 'text', value: '12.5%' });
    s = seedCell(s, 4, 0, { kind: 'text', value: 'north' });
    s = seedCell(s, 5, 0, { kind: 'text', value: '1.234,5' });
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 1, c1: 1 },
        '=LOWER(TRIM(A1))="north region"',
        '#lower-trim',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 1, c1: 2 },
        '=UPPER(LEFT(TRIM(A1),5))="NORTH"',
        '#upper-left',
      ),
      formulaRule(
        { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
        '=VALUE(A3)=1234.5',
        '#value-thousands',
      ),
      formulaRule({ sheet: 0, r0: 3, c0: 1, r1: 3, c1: 1 }, '=VALUE(A4)=0.125', '#value-percent'),
      formulaRule({ sheet: 0, r0: 4, c0: 1, r1: 4, c1: 1 }, '=VALUE(A5)=0', '#value-invalid'),
      formulaRule(
        { sheet: 0, r0: 5, c0: 1, r1: 5, c1: 1 },
        '=NUMBERVALUE(A6,",",".")=1234.5',
        '#numbervalue-separators',
      ),
      formulaRule(
        { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
        '=NUMBERVALUE(A3,,)=1234.5',
        '#numbervalue-omitted-separators',
      ),
      formulaRule(
        { sheet: 0, r0: 5, c0: 2, r1: 5, c1: 2 },
        '=NUMBERVALUE(A6,",",",")=1234.5',
        '#numbervalue-invalid-separators',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#lower-trim');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBe('#upper-left');
    expect(overlay.get('0:1:2')?.fill).toBeUndefined();
    expect(overlay.get('0:2:1')?.fill).toBe('#value-thousands');
    expect(overlay.get('0:3:1')?.fill).toBe('#value-percent');
    expect(overlay.get('0:4:1')?.fill).toBeUndefined();
    expect(overlay.get('0:5:1')?.fill).toBe('#numbervalue-separators');
    expect(overlay.get('0:2:2')?.fill).toBe('#numbervalue-omitted-separators');
    expect(overlay.get('0:5:2')?.fill).toBeUndefined();
  });

  it('formula rules evaluate text concatenation with ampersand', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North Region' });
    s = seedNumber(s, 0, 1, 12);
    s = seedCell(s, 0, 2, { kind: 'bool', value: true });
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=LEFT(A1,5)&"-"&RIGHT(A1,6)="North-Region"',
        '#concat-text',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '="N"&"o"&"r"&"t"&"h"="North"',
        '#concat-left-assoc',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=A1&" "&B1&" "&C1="North Region 12 TRUE"',
        '#concat-coerce',
      ),
      formulaRule({ sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 }, '=D10&"x"="x"', '#concat-blank'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#concat-text');
    expect(overlay.get('0:0:4')?.fill).toBe('#concat-left-assoc');
    expect(overlay.get('0:0:5')?.fill).toBe('#concat-coerce');
    expect(overlay.get('0:0:6')?.fill).toBe('#concat-blank');
  });

  it('formula rules evaluate scalar CONCATENATE and CONCAT text functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedNumber(s, 0, 1, 12);
    s = seedCell(s, 0, 2, { kind: 'bool', value: true });
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=CONCATENATE(A1,"-",B1,"-",C1)="North-12-TRUE"',
        '#concatenate',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=CONCAT(A1,"-",D10)="North-"',
        '#concat-function',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=CONCATENATE(A1,,"-",B1)="North-12"',
        '#concatenate-omitted',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
        '=CONCAT(A1,,D10)="North"',
        '#concat-omitted',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#concatenate');
    expect(overlay.get('0:0:4')?.fill).toBe('#concat-function');
    expect(overlay.get('0:0:5')?.fill).toBe('#concatenate-omitted');
    expect(overlay.get('0:0:6')?.fill).toBe('#concat-omitted');
  });

  it('formula rules evaluate scalar SUBSTITUTE/REPLACE/REPT/TEXTJOIN text functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North Region North' });
    s = seedCell(s, 0, 1, { kind: 'text', value: 'East' });
    s = seedNumber(s, 0, 2, 12);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=SUBSTITUTE(A1,"North","South",2)="North Region South"',
        '#substitute-instance',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=REPLACE(A1,7,6,"Area")="North Area North"',
        '#replace',
      ),
      formulaRule({ sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 }, '=REPT(LEFT(B1,1),3)="EEE"', '#rept'),
      formulaRule(
        { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
        '=TEXTJOIN("-",TRUE,B1,D10,C1)="East-12"',
        '#textjoin-ignore-empty',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
        '=TEXTJOIN("-",FALSE,B1,D10,C1)="East--12"',
        '#textjoin-keep-empty',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
        '=SUBSTITUTE(A1,"North","South",)="South Region South"',
        '#substitute-omitted-instance',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
        '=TEXTJOIN("-",TRUE,B1,,C1)="East-12"',
        '#textjoin-omitted-ignore-empty',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 10, r1: 0, c1: 10 },
        '=TEXTJOIN("-",FALSE,B1,,C1)="East--12"',
        '#textjoin-omitted-keep-empty',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#substitute-instance');
    expect(overlay.get('0:0:4')?.fill).toBe('#replace');
    expect(overlay.get('0:0:5')?.fill).toBe('#rept');
    expect(overlay.get('0:0:6')?.fill).toBe('#textjoin-ignore-empty');
    expect(overlay.get('0:0:7')?.fill).toBe('#textjoin-keep-empty');
    expect(overlay.get('0:0:8')?.fill).toBe('#substitute-omitted-instance');
    expect(overlay.get('0:0:9')?.fill).toBe('#textjoin-omitted-ignore-empty');
    expect(overlay.get('0:0:10')?.fill).toBe('#textjoin-omitted-keep-empty');
  });

  it('formula rules evaluate scalar TEXTBEFORE and TEXTAFTER text functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North-East-West' });
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=AND(TEXTBEFORE(A1,"-")="North",TEXTAFTER(A1,"-")="East-West")',
        '#text-before-after',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=AND(TEXTBEFORE(A1,"-",-1)="North-East",TEXTAFTER(A1,"-",-1)="West")',
        '#text-before-after-negative',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=AND(TEXTBEFORE(A1,"east",1,1)="North-",TEXTAFTER(A1,"east",1,1)="-West")',
        '#text-before-after-ignore-case',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=TEXTAFTER(A1,"/",1,0,0,"missing")="missing"',
        '#textafter-fallback',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#text-before-after');
    expect(overlay.get('0:0:2')?.fill).toBe('#text-before-after-negative');
    expect(overlay.get('0:0:3')?.fill).toBe('#text-before-after-ignore-case');
    expect(overlay.get('0:0:4')?.fill).toBe('#textafter-fallback');
  });

  it('formula rules evaluate scalar TEXT number formatting function', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 1234.567);
    s = seedNumber(s, 0, 1, 0.25);
    s = seedNumber(s, 0, 2, 45651);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=AND(TEXT(A1,"#,##0.00")="1,234.57",TEXT(B1,"0%")="25%")',
        '#text-number-format',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=TEXT(C1,"yyyy-mm-dd")="2024-12-25"',
        '#text-date-format',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=TEXT("North","0")="North"',
        '#text-format-nonnumeric',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#text-number-format');
    expect(overlay.get('0:0:4')?.fill).toBe('#text-date-format');
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
  });

  it('formula rules evaluate scalar DOLLAR and FIXED text formatting functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 1234.567);
    s = seedNumber(s, 0, 1, -1234.567);
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=AND(DOLLAR(A1)="$1,234.57",DOLLAR(A1,0)="$1,235")',
        '#dollar',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=AND(FIXED(A1,1)="1,234.6",FIXED(A1,1,TRUE())="1234.6")',
        '#fixed',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=AND(DOLLAR(A1,-2)="$1,200",FIXED(B1,-2)="-1,200")',
        '#fixed-negative-decimals',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=DOLLAR("North")="$0.00"',
        '#dollar-nonnumeric',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:2')?.fill).toBe('#dollar');
    expect(overlay.get('0:0:3')?.fill).toBe('#fixed');
    expect(overlay.get('0:0:4')?.fill).toBe('#fixed-negative-decimals');
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
  });

  it('formula rules evaluate scalar VALUETOTEXT text conversion', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 12);
    s = seedCell(s, 0, 1, { kind: 'bool', value: true });
    s = seedCell(s, 0, 2, { kind: 'text', value: 'North' });
    s = withConditionalRules(s, [
      {
        kind: 'formula',
        range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        formula:
          '=AND(VALUETOTEXT(A1)="12",VALUETOTEXT(B1)="TRUE",LEN(VALUETOTEXT(C1,1))=7,FIND("North",VALUETOTEXT(C1,1))=2,VALUETOTEXT(NA())="#N/A")',
        apply: { fill: '#value-to-text' },
      },
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=VALUETOTEXT(A1,2)="12"',
        '#value-to-text-invalid',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#value-to-text');
    expect(overlay.get('0:0:4')?.fill).toBeUndefined();
  });

  it('formula scalar text functions fail closed on invalid arguments', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=SUBSTITUTE(A1,"o","0",0)="N0rth"',
        '#substitute-invalid-instance',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
        '=REPLACE(A1,0,1,"S")="Sorth"',
        '#replace-invalid-start',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=REPT(A1,-1)=""',
        '#rept-invalid-count',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=TEXTJOIN("-",A1,A1)="North"',
        '#textjoin-invalid-ignore-empty',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=TEXTBEFORE(A1,"",1)="North"',
        '#textbefore-empty-delimiter',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
        '=TEXTAFTER(A1,"x")="North"',
        '#textafter-missing',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
        '=TEXT(12,"")="12"',
        '#text-empty-pattern',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
    expect(overlay.get('0:0:4')?.fill).toBeUndefined();
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
    expect(overlay.get('0:0:6')?.fill).toBeUndefined();
    expect(overlay.get('0:0:7')?.fill).toBeUndefined();
  });

  it('formula rules evaluate EXACT as a case-sensitive boolean operand', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'north' });
    s = withConditionalRules(s, [
      formulaRule({ sheet: 0, r0: 0, c0: 1, r1: 1, c1: 1 }, '=EXACT(A1,"North")', '#exact'),
      formulaRule(
        { sheet: 0, r0: 0, c0: 2, r1: 1, c1: 2 },
        '=NOT(EXACT(A1,"North"))',
        '#not-exact',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#exact');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:1:2')?.fill).toBe('#not-exact');
  });

  it('formula rules evaluate N/T scalar coercion operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedNumber(s, 0, 0, 12);
    s = seedCell(s, 0, 1, { kind: 'bool', value: true });
    s = seedCell(s, 0, 2, { kind: 'text', value: 'North' });
    s = seedCell(s, 0, 4, { kind: 'error', code: 6, text: '#N/A' });
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=AND(N(A1)=12,N(B1)=1,N(C1)=0,N(D1)=0)',
        '#n-coerce',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
        '=AND(T(C1)="North",T(A1)="",T(D1)="")',
        '#t-coerce',
      ),
      formulaRule({ sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 }, '=ISNA(N(E1))', '#n-error'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:5')?.fill).toBe('#n-coerce');
    expect(overlay.get('0:0:6')?.fill).toBe('#t-coerce');
    expect(overlay.get('0:0:7')?.fill).toBe('#n-error');
  });

  it('formula rules evaluate ADDRESS text references', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = withConditionalRules(s, [
      {
        kind: 'formula',
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        formula:
          '=AND(ADDRESS(3,2)="$B$3",ADDRESS(3,2,4)="B3",ADDRESS(3,2,2,FALSE())="R3C[2]",ADDRESS(3,2,,,"Sheet 1")="\'Sheet 1\'!$B$3")',
        apply: { fill: '#address' },
      },
      formulaRule({ sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }, '=ADDRESS(0,1)=""', '#invalid-address'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#address');
    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
  });

  it('formula text comparisons are case-insensitive unless EXACT is used', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'south' });
    s = withConditionalRules(s, [
      formulaRule({ sheet: 0, r0: 0, c0: 1, r1: 1, c1: 1 }, '=A1="north"', '#case-insensitive-eq'),
      formulaRule({ sheet: 0, r0: 0, c0: 2, r1: 1, c1: 2 }, '=A1<>"NORTH"', '#case-insensitive-ne'),
      formulaRule({ sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 }, '=EXACT(A1,"north")', '#exact-case'),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#case-insensitive-eq');
    expect(overlay.get('0:1:1')?.fill).toBeUndefined();
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:1:2')?.fill).toBe('#case-insensitive-ne');
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
  });

  it('formula rules evaluate character and text cleanup functions', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = withConditionalRules(s, [
      formulaRule(
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        '=AND(CHAR(65)="A",CODE("Apple")=65,UNICODE(UNICHAR(9731))=9731)',
        '#char-code',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
        '=AND(CLEAN("A"&CHAR(10)&"B")="AB",PROPER("north region")="North Region")',
        '#clean-proper',
      ),
      formulaRule({ sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 }, '=CHAR(0)=""', '#char-zero'),
      formulaRule({ sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 }, '=UNICODE("")=0', '#unicode-empty'),
      formulaRule(
        { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
        '=ENCODEURL("North Region/東京")="North%20Region%2F%E6%9D%B1%E4%BA%AC"',
        '#encodeurl',
      ),
      formulaRule(
        { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
        '=ENCODEURL(NA())=""',
        '#encodeurl-error',
      ),
    ]);

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#char-code');
    expect(overlay.get('0:0:1')?.fill).toBe('#clean-proper');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
    expect(overlay.get('0:0:4')?.fill).toBe('#encodeurl');
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
  });
});
