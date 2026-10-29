import { describe, expect, it } from 'vitest';
import { formulaCaretContext } from '../../../src/commands/formula-edit-context.js';

describe('formulaCaretContext', () => {
  it.each([
    ['="A1"', 3, 'string'],
    ['="a ""quoted"" value"+A1', 8, 'string'],
    ["='Sheet 1'!A1", 4, 'sheet'],
    ["='Sheet 1'!A1", 11, 'expression'],
    ['=SUMSheet!A1', 3, 'sheet'],
    ['=SUM(A1)', 2, 'function'],
    ['=SUM(A1)', 4, 'function'],
    ['=Table1[[#Headers],[Amount]]', 22, 'bracket'],
    ['=SUM(R[-2]C[-2])', 12, 'expression'],
    ['=SUM(R[-2]C[-2]:R[-1]C[-1])', 20, 'expression'],
    ['=SUM(A1)', 6, 'expression'],
  ] as const)('identifies context in %s at %i', (text, caret, expected) => {
    expect(formulaCaretContext(text, caret)).toBe(expected);
  });
});
