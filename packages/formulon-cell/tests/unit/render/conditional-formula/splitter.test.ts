import { describe, expect, it } from 'vitest';
import {
  splitFormulaArgs,
  splitFormulaArgsAllowEmpty,
} from '../../../../src/render/conditional-formula/parser.js';

describe('conditional formula argument containers', () => {
  it('keeps nested calls and doubled quoted delimiters inside their containers', () => {
    expect(splitFormulaArgs('IF(A1,SUM(B1,B2),"x,y"),3')).toEqual(['IF(A1,SUM(B1,B2),"x,y")', '3']);
    expect(splitFormulaArgs(`"a""b,c",'O''Brien, West'!A1,3`)).toEqual([
      '"a""b,c"',
      "'O''Brien, West'!A1",
      '3',
    ]);
  });

  it('keeps array rows and structured-reference headers opaque', () => {
    expect(splitFormulaArgs('{1,2;3,4},5')).toEqual(['{1,2;3,4}', '5']);
    expect(splitFormulaArgs('Table1[[Last, First],[Amount]],3')).toEqual([
      'Table1[[Last, First],[Amount]]',
      '3',
    ]);
    expect(splitFormulaArgs("DeptSales['#Items],3")).toEqual(["DeptSales['#Items]", '3']);
    expect(splitFormulaArgs("Table1[[Col']Name]],3")).toEqual(["Table1[[Col']Name]]", '3']);
  });

  it('preserves explicit interior and trailing blanks only in the permissive splitter', () => {
    expect(splitFormulaArgsAllowEmpty('A1,,{1,2},')).toEqual(['A1', '', '{1,2}', '']);
    expect(splitFormulaArgs('A1,,{1,2},')).toBeNull();
    expect(splitFormulaArgsAllowEmpty('')).toEqual(['']);
  });

  it('rejects crossed or unfinished containers and quotes', () => {
    expect(splitFormulaArgs('A1,(B1],3')).toBeNull();
    expect(splitFormulaArgs('A1,{1,2],3')).toBeNull();
    expect(splitFormulaArgs('A1,[B1,3')).toBeNull();
    expect(splitFormulaArgs('"a,b,3')).toBeNull();
    expect(splitFormulaArgs("'O''Brien,3")).toBeNull();
  });
});
