import { describe, expect, it } from 'vitest';
import { scanR1C1References } from '../../../src/commands/r1c1-reference.js';

describe('scanR1C1References', () => {
  it('reports original offsets, qualifiers, endpoints, and axis values', () => {
    const text =
      `='RC'!RC+Sheet1!R[1]C2+R4C5:R6C7:R8C9+LOG10(R3C3)+"R5C5"+Table1[R6C7]` +
      '+FOOR8C9+βR8C9+foo.R8C9+foo\\R8C9+[Book.xlsx]Sheet!R1C1';
    const refs = scanR1C1References(text);

    expect(
      refs.map(({ start, end, qualifier, qualifierStart, qualifierEnd, endpoints }) => ({
        start,
        end,
        qualifier,
        qualifierStart,
        qualifierEnd,
        endpoints,
      })),
    ).toEqual([
      {
        start: text.indexOf("'RC'!"),
        end: text.indexOf("'RC'!") + "'RC'!RC".length,
        qualifier: "'RC'!",
        qualifierStart: text.indexOf("'RC'!"),
        qualifierEnd: text.indexOf("'RC'!") + "'RC'!".length,
        endpoints: [
          {
            start: text.indexOf("'RC'!") + "'RC'!".length,
            end: text.indexOf("'RC'!") + "'RC'!RC".length,
            row: { absolute: false, value: 0 },
            column: { absolute: false, value: 0 },
          },
        ],
      },
      {
        start: text.indexOf('Sheet1!'),
        end: text.indexOf('Sheet1!') + 'Sheet1!R[1]C2'.length,
        qualifier: 'Sheet1!',
        qualifierStart: text.indexOf('Sheet1!'),
        qualifierEnd: text.indexOf('Sheet1!') + 'Sheet1!'.length,
        endpoints: [
          {
            start: text.indexOf('Sheet1!') + 'Sheet1!'.length,
            end: text.indexOf('Sheet1!') + 'Sheet1!R[1]C2'.length,
            row: { absolute: false, value: 1 },
            column: { absolute: true, value: 2 },
          },
        ],
      },
      {
        start: text.indexOf('R4C5'),
        end: text.indexOf('R8C9') + 'R8C9'.length,
        qualifier: '',
        qualifierStart: text.indexOf('R4C5'),
        qualifierEnd: text.indexOf('R4C5'),
        endpoints: [
          {
            start: text.indexOf('R4C5'),
            end: text.indexOf('R4C5') + 'R4C5'.length,
            row: { absolute: true, value: 4 },
            column: { absolute: true, value: 5 },
          },
          {
            start: text.indexOf('R6C7'),
            end: text.indexOf('R6C7') + 'R6C7'.length,
            row: { absolute: true, value: 6 },
            column: { absolute: true, value: 7 },
          },
          {
            start: text.indexOf('R8C9'),
            end: text.indexOf('R8C9') + 'R8C9'.length,
            row: { absolute: true, value: 8 },
            column: { absolute: true, value: 9 },
          },
        ],
      },
      {
        start: text.indexOf('R3C3'),
        end: text.indexOf('R3C3') + 'R3C3'.length,
        qualifier: '',
        qualifierStart: text.indexOf('R3C3'),
        qualifierEnd: text.indexOf('R3C3'),
        endpoints: [
          {
            start: text.indexOf('R3C3'),
            end: text.indexOf('R3C3') + 'R3C3'.length,
            row: { absolute: true, value: 3 },
            column: { absolute: true, value: 3 },
          },
        ],
      },
    ]);
  });

  it('skips escaped strings and quoted external workbook references', () => {
    const text = `="R1C1 "" embedded"+'[Book.xlsx]Sheet 1'!R2C3+R4C5`;
    const refs = scanR1C1References(text);

    expect(
      refs.map(({ qualifier, endpoints }) => ({
        qualifier,
        endpoint: text.slice(endpoints[0]?.start, endpoints[0]?.end),
      })),
    ).toEqual([{ qualifier: '', endpoint: 'R4C5' }]);
  });

  it('skips R1C1-shaped function and identifier names but recognizes an escaped sheet name', () => {
    const text = "=R1C1(A1)+FOOR1C1+βR1C1+Name.R1C1+Name\\R1C1+'O''Brien'!R2C3+R4C5";
    const refs = scanR1C1References(text);

    expect(
      refs.map(({ qualifier, start, end }) => ({
        qualifier,
        token: text.slice(start, end),
      })),
    ).toEqual([
      { qualifier: "'O''Brien'!", token: "'O''Brien'!R2C3" },
      { qualifier: '', token: 'R4C5' },
    ]);
  });

  it('keeps R1C1 references before whitespace parentheses available for intersection formulas', () => {
    const text = '=RC (R[1]C)';
    const refs = scanR1C1References(text);

    expect(refs.map(({ start, end }) => text.slice(start, end))).toEqual(['RC', 'R[1]C']);
  });
});
