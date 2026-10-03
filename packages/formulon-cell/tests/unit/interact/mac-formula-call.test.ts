import { describe, expect, it } from 'vitest';
import {
  assembledFormula,
  formulaWithArgumentCount,
  parseOuterCall,
} from '../../../src/interact/mac-formula-call.js';

describe('interact/mac-formula-call', () => {
  it('parses the outer call case-insensitively and keeps empty arguments', () => {
    expect(parseOuterCall('= sum(A1,,B2)', 'SUM')).toEqual({ name: 'SUM', args: ['A1', '', 'B2'] });
    expect(parseOuterCall('=AVERAGE(A1)', 'SUM')).toBeNull();
  });

  it('drops trailing blank arguments when assembling', () => {
    expect(assembledFormula('SUM', ['A1', '', ' '])).toBe('=SUM(A1)');
    expect(assembledFormula('TODAY', ['', ''])).toBe('=TODAY()');
  });

  it('trims trailing blanks when no explicit count is kept', () => {
    expect(formulaWithArgumentCount('SUM', ['A1', '', ''], null)).toBe('=SUM(A1)');
  });

  it('keeps exactly the explicit argument count, blanks included', () => {
    expect(formulaWithArgumentCount('SUM', ['A1', '', ''], 2)).toBe('=SUM(A1,)');
    expect(formulaWithArgumentCount('SUM', ['A1', 'B1', 'C1'], 1)).toBe('=SUM(A1)');
    expect(formulaWithArgumentCount('SUM', ['A1'], 0)).toBe('=SUM()');
  });
});
