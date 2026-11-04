import { describe, expect, it } from 'vitest';
import { createRangePickFormatter } from '../../../src/interact/range-pick-format.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const NAMES = ['Sheet1', 'Sheet2', 'My Sheet', 'A1'];

const setup = () => {
  const store = createSpreadsheetStore();
  const fmt = createRangePickFormatter({ store, sheetName: (i) => NAMES[i] ?? '' });
  const pick = (sheet: number, r0: number, c0: number, r1 = r0, c1 = c0) => {
    mutators.setSheetIndex(store, sheet);
    mutators.setActive(store, { sheet, row: r0, col: c0 });
    mutators.setRange(store, { sheet, r0, c0, r1, c1 });
  };
  return { store, fmt, pick };
};

describe('createRangePickFormatter', () => {
  it('writes absolute text and never writes a single cell as a range', () => {
    const { fmt, pick } = setup();
    pick(0, 2, 2);
    expect(fmt.cell()).toBe('$C$3');
    expect(fmt.range()).toBe('$C$3');
    pick(0, 1, 1, 4, 3);
    expect(fmt.range()).toBe('$B$2:$D$5');
  });

  it('leaves the pick unqualified on the sheet the dialog opened on', () => {
    const { fmt, pick } = setup();
    pick(0, 2, 2);
    expect(fmt.range()).toBe('$C$3');
  });

  it('qualifies a pick on another sheet and keeps qualifying after returning', () => {
    const { fmt, pick } = setup();
    pick(1, 0, 0);
    expect(fmt.range()).toBe('Sheet2!$A$1');
    pick(0, 2, 2);
    expect(fmt.range()).toBe('Sheet1!$C$3');
  });

  it('latches on a sheet change observed through subscribe even without a pick', () => {
    const { fmt, pick, store } = setup();
    const seen: number[] = [];
    const off = fmt.subscribe(() => seen.push(store.getState().data.sheetIndex));
    pick(1, 5, 5);
    pick(0, 2, 2);
    expect(seen.length).toBeGreaterThan(0);
    expect(fmt.range()).toBe('Sheet1!$C$3');
    off();
  });

  it('begin() recaptures the origin sheet and clears the latch', () => {
    const { fmt, pick } = setup();
    pick(1, 0, 0);
    expect(fmt.range()).toBe('Sheet2!$A$1');
    fmt.begin();
    expect(fmt.range()).toBe('$A$1');
    pick(0, 2, 2);
    expect(fmt.range()).toBe('Sheet1!$C$3');
  });

  it('quotes sheet names that cannot stand bare', () => {
    const { fmt, pick } = setup();
    pick(2, 0, 0);
    expect(fmt.range()).toBe("'My Sheet'!$A$1");
    pick(3, 0, 0);
    expect(fmt.range()).toBe("'A1'!$A$1");
  });

  it('forces or suppresses the sheet prefix', () => {
    const { fmt, pick } = setup();
    pick(0, 2, 2);
    expect(fmt.range({ prefix: 'always' })).toBe('Sheet1!$C$3');
    pick(1, 0, 0);
    expect(fmt.range({ prefix: 'never' })).toBe('$A$1');
    expect(fmt.cell({ prefix: 'always' })).toBe('Sheet2!$A$1');
  });

  it('writes whole-row and whole-column forms', () => {
    const { fmt, pick } = setup();
    pick(0, 2, 2);
    expect(fmt.rows({ prefix: 'never' })).toBe('$3:$3');
    expect(fmt.cols({ prefix: 'never' })).toBe('$C:$C');
    pick(0, 1, 1, 4, 3);
    expect(fmt.rows({ prefix: 'never' })).toBe('$2:$5');
    expect(fmt.cols({ prefix: 'never' })).toBe('$B:$D');
  });

  it('normalizes a reversed selection and accepts an explicit range', () => {
    const { fmt, pick } = setup();
    pick(0, 4, 3, 1, 1);
    expect(fmt.range()).toBe('$B$2:$D$5');
    expect(fmt.range({ range: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 } })).toBe('$A$1:$B$2');
  });
});
