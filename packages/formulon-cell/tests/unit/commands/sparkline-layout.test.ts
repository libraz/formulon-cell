import { describe, expect, it } from 'vitest';
import {
  splitDataForLocation,
  switchSparklineSources,
  unionSparklineSources,
} from '../../../src/commands/sparkline-layout.js';

const range = (r0: number, c0: number, r1: number, c1: number) => ({ sheet: 0, r0, c0, r1, c1 });

describe('splitDataForLocation', () => {
  it('splits the rows of the data across a column of cells', () => {
    expect(splitDataForLocation(range(0, 0, 2, 3), range(0, 5, 2, 5))).toEqual([
      range(0, 0, 0, 3),
      range(1, 0, 1, 3),
      range(2, 0, 2, 3),
    ]);
  });

  it('splits the columns of the data across a row of cells', () => {
    expect(splitDataForLocation(range(0, 0, 3, 1), range(5, 0, 5, 1))).toEqual([
      range(0, 0, 3, 0),
      range(0, 1, 3, 1),
    ]);
  });

  it('takes a square block by columns for a row of cells and by rows otherwise', () => {
    expect(splitDataForLocation(range(0, 0, 1, 1), range(5, 0, 5, 1))?.[0]).toEqual(
      range(0, 0, 1, 0),
    );
    expect(splitDataForLocation(range(0, 0, 1, 1), range(5, 0, 6, 0))?.[0]).toEqual(
      range(0, 0, 0, 1),
    );
  });

  it('gives one location cell all of the data and refuses a count that fits neither axis', () => {
    expect(splitDataForLocation(range(0, 0, 2, 3), range(5, 5, 5, 5))).toEqual([range(0, 0, 2, 3)]);
    expect(splitDataForLocation(range(0, 0, 2, 3), range(5, 0, 9, 0))).toBeNull();
  });
});

describe('switchSparklineSources', () => {
  it('turns the rows of a square block into its columns', () => {
    expect(switchSparklineSources(['A1:C1', 'A2:C2', 'A3:C3'])).toEqual([
      'A1:A3',
      'B1:B3',
      'C1:C3',
    ]);
  });

  it('turns the columns of a square block into its rows', () => {
    expect(switchSparklineSources(['A1:A2', 'B1:B2'])).toEqual(['A1:B1', 'A2:B2']);
  });

  it('keeps the sheet qualifier', () => {
    expect(switchSparklineSources(["'My Sheet'!A1:B1", "'My Sheet'!A2:B2"])).toEqual([
      "'My Sheet'!A1:A2",
      "'My Sheet'!B1:B2",
    ]);
  });

  it('does nothing when the member count differs from the other dimension', () => {
    expect(switchSparklineSources(['A1:D1', 'A2:D2'])).toBeNull();
  });

  it('does nothing for sources that do not tile one block', () => {
    expect(switchSparklineSources(['A1:B1', 'A3:B3'])).toBeNull();
    expect(switchSparklineSources(['A1:B1', 'Sheet2!A2:B2'])).toBeNull();
    expect(switchSparklineSources(['A1:B1'])).toBeNull();
  });
});

describe('unionSparklineSources', () => {
  it('reads the block the sources cover', () => {
    expect(unionSparklineSources(['A1:C1', 'A2:C2', 'A3:C3'])).toBe('A1:C3');
    expect(unionSparklineSources(['Sheet2!A1:A4'])).toBe('Sheet2!A1:A4');
  });

  it('falls back to the first source when they differ in sheet', () => {
    expect(unionSparklineSources(['A1:A2', 'S!A1:A2'])).toBe('A1:A2');
  });
});
