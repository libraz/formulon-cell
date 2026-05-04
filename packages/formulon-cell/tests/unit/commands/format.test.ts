import { beforeEach, describe, expect, it } from 'vitest';
import {
  bumpDecimals,
  clearFormat,
  cycleBorders,
  cycleCurrency,
  cyclePercent,
  formatNumber,
  setAlign,
  setBorders,
  setFillColor,
  setFont,
  setFontColor,
  setNumFmt,
  toggleBold,
  toggleItalic,
  toggleStrike,
  toggleUnderline,
} from '../../../src/commands/format.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import {
  type CellFormat,
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const setRange = (
  store: SpreadsheetStore,
  r0: number,
  c0: number,
  r1: number,
  c1: number,
): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      ...s.selection,
      range: { sheet: 0, r0, c0, r1, c1 },
    },
  }));
};

const fmtAt = (store: SpreadsheetStore, row: number, col: number): CellFormat | undefined =>
  store.getState().format.formats.get(addrKey({ sheet: 0, row, col }));

describe('toggle flags', () => {
  let store: SpreadsheetStore;

  beforeEach(() => {
    store = createSpreadsheetStore();
    setRange(store, 0, 0, 1, 1);
  });

  it('turns the flag on across the whole range when no cell has it', () => {
    toggleBold(store.getState(), store);
    for (let r = 0; r <= 1; r += 1) {
      for (let c = 0; c <= 1; c += 1) {
        expect(fmtAt(store, r, c)?.bold).toBe(true);
      }
    }
  });

  it('turns the flag off only when every cell already has it', () => {
    // First call: enable on every cell.
    toggleBold(store.getState(), store);
    // Second call: disable.
    toggleBold(store.getState(), store);
    expect(fmtAt(store, 0, 0)?.bold).toBe(false);
    expect(fmtAt(store, 1, 1)?.bold).toBe(false);
  });

  it('extends to the rest of the range when at least one cell is missing the flag', () => {
    // Enable on (0,0) only.
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
    // Range covers (0,0)..(1,1). One cell has it, three don't → toggle should
    // enable everywhere, not flip off.
    toggleBold(store.getState(), store);
    expect(fmtAt(store, 0, 0)?.bold).toBe(true);
    expect(fmtAt(store, 1, 1)?.bold).toBe(true);
  });

  it('covers italic / underline / strike via the same path', () => {
    toggleItalic(store.getState(), store);
    toggleUnderline(store.getState(), store);
    toggleStrike(store.getState(), store);
    expect(fmtAt(store, 0, 0)).toMatchObject({ italic: true, underline: true, strike: true });
  });
});

describe('setAlign', () => {
  it('writes the alignment to every cell in the range', () => {
    const store = createSpreadsheetStore();
    setRange(store, 2, 2, 3, 3);
    setAlign(store.getState(), store, 'right');
    expect(fmtAt(store, 2, 2)?.align).toBe('right');
    expect(fmtAt(store, 3, 3)?.align).toBe('right');
  });
});

describe('setNumFmt / cycleCurrency / cyclePercent', () => {
  let store: SpreadsheetStore;

  beforeEach(() => {
    store = createSpreadsheetStore();
    setRange(store, 0, 0, 0, 0);
  });

  it('setNumFmt installs the supplied format', () => {
    setNumFmt(store.getState(), store, { kind: 'fixed', decimals: 3 });
    expect(fmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'fixed', decimals: 3 });
  });

  it('cycleCurrency turns currency on when none is set', () => {
    cycleCurrency(store.getState(), store);
    expect(fmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'currency', decimals: 2, symbol: '$' });
  });

  it('cycleCurrency clears back to general when at least one cell is currency', () => {
    cycleCurrency(store.getState(), store); // on
    cycleCurrency(store.getState(), store); // off
    expect(fmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'general' });
  });

  it('cyclePercent toggles percent on / off', () => {
    cyclePercent(store.getState(), store);
    expect(fmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'percent', decimals: 0 });
    cyclePercent(store.getState(), store);
    expect(fmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'general' });
  });
});

describe('bumpDecimals', () => {
  let store: SpreadsheetStore;

  beforeEach(() => {
    store = createSpreadsheetStore();
    setRange(store, 0, 0, 0, 0);
  });

  it('promotes a general cell to fixed:2 on +1', () => {
    bumpDecimals(store.getState(), store, 1);
    expect(fmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'fixed', decimals: 2 });
  });

  it('does nothing on -1 when the cell is general', () => {
    bumpDecimals(store.getState(), store, -1);
    expect(fmtAt(store, 0, 0)?.numFmt).toBeUndefined();
  });

  it('walks fixed decimals up and down with clamping', () => {
    setNumFmt(store.getState(), store, { kind: 'fixed', decimals: 0 });
    bumpDecimals(store.getState(), store, -1);
    expect(fmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'fixed', decimals: 0 });
    for (let i = 0; i < 12; i += 1) bumpDecimals(store.getState(), store, 1);
    expect(fmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'fixed', decimals: 10 });
  });

  it('preserves currency symbol while bumping decimals', () => {
    setNumFmt(store.getState(), store, { kind: 'currency', decimals: 2, symbol: '€' });
    bumpDecimals(store.getState(), store, 1);
    expect(fmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'currency', decimals: 3, symbol: '€' });
  });

  it('walks percent decimals', () => {
    setNumFmt(store.getState(), store, { kind: 'percent', decimals: 0 });
    bumpDecimals(store.getState(), store, 1);
    expect(fmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'percent', decimals: 1 });
  });
});

describe('borders', () => {
  let store: SpreadsheetStore;

  beforeEach(() => {
    store = createSpreadsheetStore();
    setRange(store, 0, 0, 1, 1);
  });

  it('setBorders writes the supplied sides', () => {
    setBorders(store.getState(), store, { top: true, bottom: true });
    expect(fmtAt(store, 0, 0)?.borders).toMatchObject({ top: true, bottom: true });
  });

  it('cycleBorders paints an outline on first call (perimeter only)', () => {
    cycleBorders(store.getState(), store);
    // Top-left corner: top + left only.
    expect(fmtAt(store, 0, 0)?.borders).toMatchObject({ top: true, left: true });
    expect(fmtAt(store, 0, 0)?.borders?.right).toBeUndefined();
    expect(fmtAt(store, 0, 0)?.borders?.bottom).toBeUndefined();
    // Bottom-right: bottom + right only.
    expect(fmtAt(store, 1, 1)?.borders).toMatchObject({ bottom: true, right: true });
  });

  it('cycleBorders fills all four sides when only an outline exists', () => {
    cycleBorders(store.getState(), store); // outline
    cycleBorders(store.getState(), store); // all
    expect(fmtAt(store, 0, 1)?.borders).toMatchObject({
      top: true,
      right: true,
      bottom: true,
      left: true,
    });
  });

  it('cycleBorders clears all sides when every cell is fully bordered', () => {
    cycleBorders(store.getState(), store); // outline
    cycleBorders(store.getState(), store); // all
    cycleBorders(store.getState(), store); // clear
    const f = fmtAt(store, 0, 0)?.borders;
    expect(f?.top).toBe(false);
    expect(f?.right).toBe(false);
    expect(f?.bottom).toBe(false);
    expect(f?.left).toBe(false);
  });
});

describe('clearFormat / colors / font', () => {
  let store: SpreadsheetStore;

  beforeEach(() => {
    store = createSpreadsheetStore();
    setRange(store, 0, 0, 0, 0);
  });

  it('clearFormat drops the format entry entirely', () => {
    setNumFmt(store.getState(), store, { kind: 'fixed', decimals: 2 });
    expect(fmtAt(store, 0, 0)).toBeDefined();
    clearFormat(store.getState(), store);
    expect(fmtAt(store, 0, 0)).toBeUndefined();
  });

  it('setFontColor / setFillColor write and clear', () => {
    setFontColor(store.getState(), store, '#ff0000');
    expect(fmtAt(store, 0, 0)?.color).toBe('#ff0000');
    setFontColor(store.getState(), store, null);
    expect(fmtAt(store, 0, 0)?.color).toBeUndefined();

    setFillColor(store.getState(), store, '#0f0');
    expect(fmtAt(store, 0, 0)?.fill).toBe('#0f0');
    setFillColor(store.getState(), store, null);
    expect(fmtAt(store, 0, 0)?.fill).toBeUndefined();
  });

  it('setFont updates family / size and clears with null', () => {
    setFont(store.getState(), store, { fontFamily: 'Inter', fontSize: 14 });
    expect(fmtAt(store, 0, 0)?.fontFamily).toBe('Inter');
    expect(fmtAt(store, 0, 0)?.fontSize).toBe(14);

    setFont(store.getState(), store, { fontFamily: null });
    expect(fmtAt(store, 0, 0)?.fontFamily).toBeUndefined();
    // size untouched.
    expect(fmtAt(store, 0, 0)?.fontSize).toBe(14);
  });
});

describe('formatNumber', () => {
  it('returns the input as a string for non-finite numbers', () => {
    expect(formatNumber(Number.NaN, undefined)).toBe('NaN');
    expect(formatNumber(Number.POSITIVE_INFINITY, undefined)).toBe('Infinity');
  });

  it('uses general formatting when no format is supplied', () => {
    expect(formatNumber(1234.5, undefined)).toBe('1,234.5');
  });

  it('applies fixed decimals', () => {
    expect(formatNumber(1.5, { kind: 'fixed', decimals: 3 })).toBe('1.500');
    expect(formatNumber(1.2345, { kind: 'fixed', decimals: 2 })).toBe('1.23');
  });

  it('renders currency with prefix symbol and respects negatives', () => {
    expect(formatNumber(99, { kind: 'currency', decimals: 2, symbol: '$' })).toBe('$99.00');
    expect(formatNumber(-99, { kind: 'currency', decimals: 0, symbol: '€' })).toBe('-€99');
  });

  it('falls back to "$" when currency symbol is missing', () => {
    expect(formatNumber(7.5, { kind: 'currency', decimals: 1 })).toBe('$7.5');
  });

  it('renders percent values', () => {
    expect(formatNumber(0.25, { kind: 'percent', decimals: 0 })).toBe('25%');
    expect(formatNumber(0.1234, { kind: 'percent', decimals: 2 })).toBe('12.34%');
  });
});
