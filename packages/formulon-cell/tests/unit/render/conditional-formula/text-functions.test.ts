import { afterEach, describe, expect, it, vi } from 'vitest';
import { _resetConditionalCache, evaluateConditional } from '../../../../src/render/conditional.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';
import { seedCell, seedNumber } from '../conditional-fixtures.js';

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 2, c1: 1 },
            formula: '=AND(ISNUMBER(SEARCH("region",A1)),LEN(A1)>5)',
            apply: { fill: '#search' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 2, c1: 2 },
            formula: '=ISNUMBER(FIND("Region",A1))',
            apply: { fill: '#find' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 1, c1: 1 },
            formula: '=SEARCH("north",A1,7)=7',
            apply: { fill: '#search-start' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=SEARCH("north",A1,)=1',
            apply: { fill: '#search-omitted-start' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(SEARCH("N*r",A1)=1,SEARCH("r?g",A1)=7)',
            apply: { fill: '#search-wildcard' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=SEARCH("~*",A2)=3',
            apply: { fill: '#search-escape' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
            formula: '=FIND("*",A2)=3',
            apply: { fill: '#find-literal' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=FIND("Region",A1,)=7',
            apply: { fill: '#find-omitted-start' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:1')?.fill).toBe('#search-wildcard');
    expect(overlay.get('0:1:1')?.fill).toBe('#search-escape');
    expect(overlay.get('0:1:2')?.fill).toBe('#find-literal');
    expect(overlay.get('0:0:3')?.fill).toBe('#find-omitted-start');
  });

  it('formula rules evaluate scalar HYPERLINK display values', () => {
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
              '=AND(HYPERLINK("https://example.test","Example")="Example",HYPERLINK("#A1")="#A1",HYPERLINK("#A1",42)=42)',
            apply: { fill: '#hyperlink' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=HYPERLINK(NA())=""',
            apply: { fill: '#hyperlink-invalid' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#hyperlink');
    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
  });

  it('formula rules evaluate limited CELL information operands', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedNumber(s, 1, 0, 42);
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula:
              '=AND(CELL("address",A1)="$A$1",CELL("row",A1)=1,CELL("col",A1)=1,CELL("contents",A1)="North",CELL("type",A1)="l")',
            apply: { fill: '#cell-info-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
            formula: '=AND(CELL("contents",A2)=42,CELL("type",A2)="v")',
            apply: { fill: '#cell-info-number' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
            formula: '=AND(CELL("address")="$B$3",CELL("type")="b")',
            apply: { fill: '#cell-info-current' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 3, c0: 1, r1: 3, c1: 1 },
            formula: '=CELL("filename",A1)=""',
            apply: { fill: '#cell-info-unsupported' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 4, c0: 1, r1: 4, c1: 1 },
            formula:
              '=AND(CELL("address",OFFSET(A1,1,0))="$A$2",CELL("contents",INDIRECT("A2"))=42)',
            apply: { fill: '#cell-info-dynamic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 5, c0: 1, r1: 5, c1: 1 },
            formula: '=CELL("address",OFFSET(A1,0,0,2,1))="$A$1"',
            apply: { fill: '#cell-info-dynamic-multi' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            formula: '=AND(SHEET()=1,SHEET(A1)=1,SHEET(A1:B2)=1)',
            apply: { fill: '#sheet-info' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=SHEETS(A1:B2)=1',
            apply: { fill: '#sheets-info' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=SHEETS()=1',
            apply: { fill: '#sheets-workbook-count' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=SHEET(Sheet2!A1)=2',
            apply: { fill: '#sheet-other' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#sheet-info');
    expect(overlay.get('0:0:1')?.fill).toBe('#sheets-info');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
  });

  it('formula rules evaluate limited SHEET and SHEETS over dynamic ranges', () => {
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
            formula: '=AND(SHEET(OFFSET(A1,0,0,2,2))=1,SHEETS(INDIRECT("A1:B2"))=1)',
            apply: { fill: '#sheet-info-dynamic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=SHEET(INDIRECT("Sheet2!A1"))=2',
            apply: { fill: '#sheet-info-dynamic-other' },
          },
        ],
      },
    };

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
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 2, c1: 1 },
            formula: '=ISLOGICAL(A1)',
            apply: { fill: '#logical' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 2, c1: 2 },
            formula: '=ISNONTEXT(A1)',
            apply: { fill: '#nontext' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 2, c1: 3 },
            formula: '=ISFORMULA(A1)',
            apply: { fill: '#formula' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 2, c1: 4 },
            formula: '=EXACT(FORMULATEXT(A1),"=SUM(1,2)")',
            apply: { fill: '#formulatext' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=AND(TYPE(A1)=4,TYPE(A2)=2,TYPE(A3)=1,TYPE(NA())=16)',
            apply: { fill: '#type-codes' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 1, c0: 5, r1: 1, c1: 5 },
            formula: '=EXACT(FORMULATEXT(OFFSET(A1,2,0)),"=SUM(1,2)")',
            apply: { fill: '#formulatext-dynamic' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 5, r1: 2, c1: 5 },
            formula: '=FORMULATEXT(OFFSET(A1,0,0,2,1))=""',
            apply: { fill: '#formulatext-dynamic-multi' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=ISFORMULA(1)',
            apply: { fill: '#formula-literal' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=ISFORMULA(OFFSET($A$1,2,0))',
            apply: { fill: '#formula-dynamic-offset' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=ISFORMULA(INDIRECT("$A$3"))',
            apply: { fill: '#formula-dynamic-indirect' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=ISFORMULA(OFFSET($A$1,0,0,2,1))',
            apply: { fill: '#formula-dynamic-multi' },
          },
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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(ISREF(A1),ISREF(A1:B2),NOT(ISREF(1)))',
            apply: { fill: '#isref' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=ISERROR(NA())',
            apply: { fill: '#na-error' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=NA()=A1',
            apply: { fill: '#na-equals' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 1, c1: 4 },
            formula: '=ISNA(A1)',
            apply: { fill: '#isna' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 1, c1: 5 },
            formula: '=ISERR(A1)',
            apply: { fill: '#iserr' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula:
              '=AND(ERROR.TYPE(A1)=7,ERROR.TYPE(A2)=2,ERROR.TYPE(1/0)=2,ERROR.TYPE(SQRT(-1))=6)',
            apply: { fill: '#error-type' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=ERROR.TYPE(1)=0',
            apply: { fill: '#error-type-non-error' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=AND(ISREF(OFFSET(A1,0,0,1,1)),ISREF(INDIRECT("A1:B2")),NOT(ISREF(1)))',
            apply: { fill: '#isref-dynamic' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 1, c1: 1 },
            formula: '=IFERROR(SEARCH("region",A1),0)>0',
            apply: { fill: '#iferror-search' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=IFNA(NA(),42)=42',
            apply: { fill: '#ifna-na' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=ISERR(IFNA(1/0,42))',
            apply: { fill: '#ifna-preserves-other-errors' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 2, c1: 1 },
            formula: '=LEFT(A1,5)="North"',
            apply: { fill: '#left' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 2, c1: 2 },
            formula: '=RIGHT(A1,6)="Region"',
            apply: { fill: '#right' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 2, c1: 3 },
            formula: '=MID(A1,7,6)="Region"',
            apply: { fill: '#mid' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=LEFT(A1)="N"',
            apply: { fill: '#left-default' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=AND(LEFT(A1,)="N",RIGHT(A1,)="h")',
            apply: { fill: '#slice-omitted-count' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=RIGHT(A1,-1)=""',
            apply: { fill: '#right-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=MID(A1,0,1)="N"',
            apply: { fill: '#mid-invalid' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 1, c1: 1 },
            formula: '=LOWER(TRIM(A1))="north region"',
            apply: { fill: '#lower-trim' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 1, c1: 2 },
            formula: '=UPPER(LEFT(TRIM(A1),5))="NORTH"',
            apply: { fill: '#upper-left' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
            formula: '=VALUE(A3)=1234.5',
            apply: { fill: '#value-thousands' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 3, c0: 1, r1: 3, c1: 1 },
            formula: '=VALUE(A4)=0.125',
            apply: { fill: '#value-percent' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 4, c0: 1, r1: 4, c1: 1 },
            formula: '=VALUE(A5)=0',
            apply: { fill: '#value-invalid' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 5, c0: 1, r1: 5, c1: 1 },
            formula: '=NUMBERVALUE(A6,",",".")=1234.5',
            apply: { fill: '#numbervalue-separators' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
            formula: '=NUMBERVALUE(A3,,)=1234.5',
            apply: { fill: '#numbervalue-omitted-separators' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 5, c0: 2, r1: 5, c1: 2 },
            formula: '=NUMBERVALUE(A6,",",",")=1234.5',
            apply: { fill: '#numbervalue-invalid-separators' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=LEFT(A1,5)&"-"&RIGHT(A1,6)="North-Region"',
            apply: { fill: '#concat-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '="N"&"o"&"r"&"t"&"h"="North"',
            apply: { fill: '#concat-left-assoc' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=A1&" "&B1&" "&C1="North Region 12 TRUE"',
            apply: { fill: '#concat-coerce' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=D10&"x"="x"',
            apply: { fill: '#concat-blank' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=CONCATENATE(A1,"-",B1,"-",C1)="North-12-TRUE"',
            apply: { fill: '#concatenate' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=CONCAT(A1,"-",D10)="North-"',
            apply: { fill: '#concat-function' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=CONCATENATE(A1,,"-",B1)="North-12"',
            apply: { fill: '#concatenate-omitted' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=CONCAT(A1,,D10)="North"',
            apply: { fill: '#concat-omitted' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=SUBSTITUTE(A1,"North","South",2)="North Region South"',
            apply: { fill: '#substitute-instance' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=REPLACE(A1,7,6,"Area")="North Area North"',
            apply: { fill: '#replace' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=REPT(LEFT(B1,1),3)="EEE"',
            apply: { fill: '#rept' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=TEXTJOIN("-",TRUE,B1,D10,C1)="East-12"',
            apply: { fill: '#textjoin-ignore-empty' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=TEXTJOIN("-",FALSE,B1,D10,C1)="East--12"',
            apply: { fill: '#textjoin-keep-empty' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 },
            formula: '=SUBSTITUTE(A1,"North","South",)="South Region South"',
            apply: { fill: '#substitute-omitted-instance' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 9, r1: 0, c1: 9 },
            formula: '=TEXTJOIN("-",TRUE,B1,,C1)="East-12"',
            apply: { fill: '#textjoin-omitted-ignore-empty' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 10, r1: 0, c1: 10 },
            formula: '=TEXTJOIN("-",FALSE,B1,,C1)="East--12"',
            apply: { fill: '#textjoin-omitted-keep-empty' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(TEXTBEFORE(A1,"-")="North",TEXTAFTER(A1,"-")="East-West")',
            apply: { fill: '#text-before-after' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=AND(TEXTBEFORE(A1,"-",-1)="North-East",TEXTAFTER(A1,"-",-1)="West")',
            apply: { fill: '#text-before-after-negative' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=AND(TEXTBEFORE(A1,"east",1,1)="North-",TEXTAFTER(A1,"east",1,1)="-West")',
            apply: { fill: '#text-before-after-ignore-case' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=TEXTAFTER(A1,"/",1,0,0,"missing")="missing"',
            apply: { fill: '#textafter-fallback' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=AND(TEXT(A1,"#,##0.00")="1,234.57",TEXT(B1,"0%")="25%")',
            apply: { fill: '#text-number-format' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=TEXT(C1,"yyyy-mm-dd")="2024-12-25"',
            apply: { fill: '#text-date-format' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=TEXT("North","0")="North"',
            apply: { fill: '#text-format-nonnumeric' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=AND(DOLLAR(A1)="$1,234.57",DOLLAR(A1,0)="$1,235")',
            apply: { fill: '#dollar' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=AND(FIXED(A1,1)="1,234.6",FIXED(A1,1,TRUE())="1234.6")',
            apply: { fill: '#fixed' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=AND(DOLLAR(A1,-2)="$1,200",FIXED(B1,-2)="-1,200")',
            apply: { fill: '#fixed-negative-decimals' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=DOLLAR("North")="$0.00"',
            apply: { fill: '#dollar-nonnumeric' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula:
              '=AND(VALUETOTEXT(A1)="12",VALUETOTEXT(B1)="TRUE",LEN(VALUETOTEXT(C1,1))=7,FIND("North",VALUETOTEXT(C1,1))=2,VALUETOTEXT(NA())="#N/A")',
            apply: { fill: '#value-to-text' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=VALUETOTEXT(A1,2)="12"',
            apply: { fill: '#value-to-text-invalid' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:3')?.fill).toBe('#value-to-text');
    expect(overlay.get('0:0:4')?.fill).toBeUndefined();
  });

  it('formula scalar text functions fail closed on invalid arguments', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=SUBSTITUTE(A1,"o","0",0)="N0rth"',
            apply: { fill: '#substitute-invalid-instance' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=REPLACE(A1,0,1,"S")="Sorth"',
            apply: { fill: '#replace-invalid-start' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=REPT(A1,-1)=""',
            apply: { fill: '#rept-invalid-count' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=TEXTJOIN("-",A1,A1)="North"',
            apply: { fill: '#textjoin-invalid-ignore-empty' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=TEXTBEFORE(A1,"",1)="North"',
            apply: { fill: '#textbefore-empty-delimiter' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=TEXTAFTER(A1,"x")="North"',
            apply: { fill: '#textafter-missing' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=TEXT(12,"")="12"',
            apply: { fill: '#text-empty-pattern' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 1, c1: 1 },
            formula: '=EXACT(A1,"North")',
            apply: { fill: '#exact' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 1, c1: 2 },
            formula: '=NOT(EXACT(A1,"North"))',
            apply: { fill: '#not-exact' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=AND(N(A1)=12,N(B1)=1,N(C1)=0,N(D1)=0)',
            apply: { fill: '#n-coerce' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 },
            formula: '=AND(T(C1)="North",T(A1)="",T(D1)="")',
            apply: { fill: '#t-coerce' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 7, r1: 0, c1: 7 },
            formula: '=ISNA(N(E1))',
            apply: { fill: '#n-error' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:5')?.fill).toBe('#n-coerce');
    expect(overlay.get('0:0:6')?.fill).toBe('#t-coerce');
    expect(overlay.get('0:0:7')?.fill).toBe('#n-error');
  });

  it('formula rules evaluate ADDRESS text references', () => {
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
              '=AND(ADDRESS(3,2)="$B$3",ADDRESS(3,2,4)="B3",ADDRESS(3,2,2,FALSE())="R3C[2]",ADDRESS(3,2,,,"Sheet 1")="\'Sheet 1\'!$B$3")',
            apply: { fill: '#address' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=ADDRESS(0,1)=""',
            apply: { fill: '#invalid-address' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#address');
    expect(overlay.get('0:0:1')?.fill).toBeUndefined();
  });

  it('formula text comparisons are case-insensitive unless EXACT is used', () => {
    const store = createSpreadsheetStore();
    let s = store.getState();
    s = seedCell(s, 0, 0, { kind: 'text', value: 'North' });
    s = seedCell(s, 1, 0, { kind: 'text', value: 'south' });
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 1, c1: 1 },
            formula: '=A1="north"',
            apply: { fill: '#case-insensitive-eq' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 1, c1: 2 },
            formula: '=A1<>"NORTH"',
            apply: { fill: '#case-insensitive-ne' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=EXACT(A1,"north")',
            apply: { fill: '#exact-case' },
          },
        ],
      },
    };

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
    s = {
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            formula: '=AND(CHAR(65)="A",CODE("Apple")=65,UNICODE(UNICHAR(9731))=9731)',
            apply: { fill: '#char-code' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
            formula: '=AND(CLEAN("A"&CHAR(10)&"B")="AB",PROPER("north region")="North Region")',
            apply: { fill: '#clean-proper' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
            formula: '=CHAR(0)=""',
            apply: { fill: '#char-zero' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
            formula: '=UNICODE("")=0',
            apply: { fill: '#unicode-empty' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 },
            formula: '=ENCODEURL("North Region/東京")="North%20Region%2F%E6%9D%B1%E4%BA%AC"',
            apply: { fill: '#encodeurl' },
          },
          {
            kind: 'formula',
            range: { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 },
            formula: '=ENCODEURL(NA())=""',
            apply: { fill: '#encodeurl-error' },
          },
        ],
      },
    };

    const overlay = evaluateConditional(s);

    expect(overlay.get('0:0:0')?.fill).toBe('#char-code');
    expect(overlay.get('0:0:1')?.fill).toBe('#clean-proper');
    expect(overlay.get('0:0:2')?.fill).toBeUndefined();
    expect(overlay.get('0:0:3')?.fill).toBeUndefined();
    expect(overlay.get('0:0:4')?.fill).toBe('#encodeurl');
    expect(overlay.get('0:0:5')?.fill).toBeUndefined();
  });
});
