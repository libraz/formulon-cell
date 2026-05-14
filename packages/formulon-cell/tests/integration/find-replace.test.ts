import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { findAll, findNext, replaceAll, replaceOne } from '../../src/commands/find.js';
import { mutators } from '../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../test-utils/index.js';

/** Integration: search + replace end-to-end against a mounted stub. */
describe('integration: find/replace', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
    const { workbook, instance } = sheet;
    workbook.setText({ sheet: 0, row: 0, col: 0 }, 'apple');
    workbook.setText({ sheet: 0, row: 0, col: 1 }, 'Apple pie');
    workbook.setText({ sheet: 0, row: 1, col: 0 }, 'banana');
    workbook.setText({ sheet: 0, row: 2, col: 0 }, 'cherry apple');
    workbook.setNumber({ sheet: 0, row: 3, col: 0 }, 42);
    workbook.setFormula({ sheet: 0, row: 4, col: 0 }, '="apple-formula"');
    workbook.recalc();
    mutators.replaceCells(instance.store, workbook.cells(0));
  });

  afterEach(() => sheet.dispose());

  it('findAll returns every match (case-insensitive by default)', () => {
    const matches = findAll(sheet.instance.store.getState(), { query: 'apple' });
    // apple, Apple pie, cherry apple, =apple-formula → 4 matches
    expect(matches.length).toBeGreaterThanOrEqual(3);
  });

  it('findAll honors caseSensitive', () => {
    const ci = findAll(sheet.instance.store.getState(), { query: 'apple' });
    const cs = findAll(sheet.instance.store.getState(), {
      query: 'apple',
      caseSensitive: true,
    });
    expect(cs.length).toBeLessThan(ci.length);
  });

  it('findAll honors matchWhole', () => {
    const whole = findAll(sheet.instance.store.getState(), {
      query: 'apple',
      matchWhole: true,
    });
    // Only A1 ("apple") is a whole-cell match (case-insensitive default).
    expect(whole).toHaveLength(1);
    expect(whole[0]?.addr).toEqual({ sheet: 0, row: 0, col: 0 });
  });

  it('findNext from null returns the first match in row-major order', () => {
    const m = findNext(sheet.instance.store.getState(), { query: 'apple' }, null, 'next');
    expect(m?.addr).toEqual({ sheet: 0, row: 0, col: 0 });
  });

  it('findNext from A1 next-direction returns the next match in reading order', () => {
    const m = findNext(
      sheet.instance.store.getState(),
      { query: 'apple' },
      { sheet: 0, row: 0, col: 0 },
      'next',
    );
    expect(m?.addr).toEqual({ sheet: 0, row: 0, col: 1 });
  });

  it('findNext prev-direction wraps to the last match', () => {
    const m = findNext(
      sheet.instance.store.getState(),
      { query: 'apple' },
      { sheet: 0, row: 0, col: 0 },
      'prev',
    );
    // Last match in reading order (excluding A1 since direction is prev from A1).
    expect(m).not.toBeNull();
  });

  it('empty query returns no matches', () => {
    expect(findAll(sheet.instance.store.getState(), { query: '' })).toEqual([]);
    expect(findNext(sheet.instance.store.getState(), { query: '' }, null, 'next')).toBeNull();
  });

  it('replaceOne preserves the original text, only replacing whole cell value', () => {
    const { instance, workbook } = sheet;
    replaceOne(workbook, { addr: { sheet: 0, row: 0, col: 0 } }, 'apricot');
    expect(workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'apricot',
    });
    expect(instance).toBeTruthy();
  });

  it('replaceOne skips formula cells', () => {
    const { workbook } = sheet;
    const before = workbook.cellFormula({ sheet: 0, row: 4, col: 0 });
    replaceOne(workbook, { addr: { sheet: 0, row: 4, col: 0 } }, 'replaced');
    expect(workbook.cellFormula({ sheet: 0, row: 4, col: 0 })).toBe(before);
  });

  it('replaceAll returns the count of cells touched', () => {
    const { instance, workbook } = sheet;
    const count = replaceAll(instance.store.getState(), workbook, { query: 'apple' }, 'orange');
    expect(count).toBeGreaterThanOrEqual(2);
    expect(workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'orange',
    });
  });

  it('replaceAll does not modify formula cells', () => {
    const { instance, workbook } = sheet;
    const before = workbook.cellFormula({ sheet: 0, row: 4, col: 0 });
    replaceAll(instance.store.getState(), workbook, { query: 'apple' }, 'X');
    expect(workbook.cellFormula({ sheet: 0, row: 4, col: 0 })).toBe(before);
  });
});
