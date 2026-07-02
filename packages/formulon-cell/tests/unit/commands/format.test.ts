import { beforeEach, describe, expect, it } from 'vitest';
import {
  bumpDecimals,
  bumpIndent,
  clearFormat,
  clearVisualFormat,
  cycleBorders,
  cycleCurrency,
  cyclePercent,
  formatNumber,
  setAlign,
  setBorderPreset,
  setBorders,
  setFillColor,
  setFont,
  setFontColor,
  setNumFmt,
  setRotation,
  setVAlign,
  toggleBold,
  toggleItalic,
  toggleStrike,
  toggleUnderline,
  toggleWrap,
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

const effectiveFmtAt = (
  store: SpreadsheetStore,
  row: number,
  col: number,
): CellFormat | undefined => {
  const stored = fmtAt(store, row, col);
  const pending = store.getState().ui.pendingFormat;
  if (pending?.addr.sheet !== 0 || pending.addr.row !== row || pending.addr.col !== col) {
    return stored;
  }
  return { ...(stored ?? {}), ...pending.format };
};

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

  it('toggleBold only changes the bold flag and preserves explicit font fields', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        fontFamily: 'Times New Roman',
        fontSize: 16,
        color: '#445566',
      },
    );
    setRange(store, 0, 0, 0, 0);

    toggleBold(store.getState(), store);

    expect(fmtAt(store, 0, 0)).toMatchObject({
      bold: true,
      fontFamily: 'Times New Roman',
      fontSize: 16,
      color: '#445566',
    });
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

  it('does not scan or materialize huge whole-column toggle ranges', () => {
    mutators.setCellFormat(store, { sheet: 0, row: 4, col: 2 }, { bold: true });
    setRange(store, 0, 2, 1048575, 2);

    toggleBold(store.getState(), store);

    expect(store.getState().format.formats.size).toBe(1);
    expect(fmtAt(store, 4, 2)?.bold).toBe(true);
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

describe('large format ranges', () => {
  it('does not scan or materialize huge whole-column number format toggles', () => {
    const store = createSpreadsheetStore();
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 4, col: 2 },
      { numFmt: { kind: 'percent', decimals: 0 } },
    );
    setRange(store, 0, 2, 1048575, 2);

    cyclePercent(store.getState(), store);

    expect(store.getState().format.formats.size).toBe(1);
    expect(fmtAt(store, 4, 2)?.numFmt).toEqual({ kind: 'percent', decimals: 0 });
  });

  it('does not scan or materialize huge whole-column indent bumps', () => {
    const store = createSpreadsheetStore();
    mutators.setCellFormat(store, { sheet: 0, row: 4, col: 2 }, { indent: 3 });
    setRange(store, 0, 2, 1048575, 2);

    bumpIndent(store.getState(), store, 1);

    expect(store.getState().format.formats.size).toBe(1);
    expect(fmtAt(store, 4, 2)?.indent).toBe(3);
  });

  it('does not scan or materialize huge whole-column decimal bumps', () => {
    const store = createSpreadsheetStore();
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 4, col: 2 },
      { numFmt: { kind: 'fixed', decimals: 2 } },
    );
    setRange(store, 0, 2, 1048575, 2);

    bumpDecimals(store.getState(), store, 1);

    expect(store.getState().format.formats.size).toBe(1);
    expect(fmtAt(store, 4, 2)?.numFmt).toEqual({ kind: 'fixed', decimals: 2 });
  });

  it('does not scan or materialize huge whole-column border presets', () => {
    const store = createSpreadsheetStore();
    mutators.setCellFormat(store, { sheet: 0, row: 4, col: 2 }, { bold: true });
    setRange(store, 0, 2, 1048575, 2);

    setBorderPreset(store.getState(), store, 'outline');
    cycleBorders(store.getState(), store);

    expect(store.getState().format.formats.size).toBe(1);
    expect(fmtAt(store, 4, 2)).toEqual({ bold: true });
  });

  it('clears huge whole-column formats by visiting only existing format entries', () => {
    const store = createSpreadsheetStore();
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 4, col: 2 },
      { bold: true, validation: { kind: 'list', source: ['A', 'B'] } },
    );
    mutators.setCellFormat(store, { sheet: 0, row: 4, col: 3 }, { bold: true });
    setRange(store, 0, 2, 1048575, 2);

    clearFormat(store.getState(), store);

    expect(fmtAt(store, 4, 2)).toBeUndefined();
    expect(fmtAt(store, 4, 3)).toEqual({ bold: true });
  });

  it('clears huge whole-column visual formats while preserving metadata on existing entries', () => {
    const store = createSpreadsheetStore();
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 4, col: 2 },
      {
        bold: true,
        fill: '#ffff00',
        validation: { kind: 'list', source: ['A', 'B'] },
        comment: 'keep',
      },
    );
    mutators.setCellFormat(store, { sheet: 0, row: 4, col: 3 }, { bold: true });
    setRange(store, 0, 2, 1048575, 2);

    clearVisualFormat(store.getState(), store);

    expect(fmtAt(store, 4, 2)).toEqual({
      validation: { kind: 'list', source: ['A', 'B'] },
      comment: 'keep',
    });
    expect(fmtAt(store, 4, 3)).toEqual({ bold: true });
  });
});

describe('alignment ribbon formatting', () => {
  let store: SpreadsheetStore;

  beforeEach(() => {
    store = createSpreadsheetStore();
    setRange(store, 1, 1, 2, 2);
  });

  it('writes vertical alignment and wrap across the selected range', () => {
    setVAlign(store.getState(), store, 'middle');
    toggleWrap(store.getState(), store);

    expect(fmtAt(store, 1, 1)).toMatchObject({ vAlign: 'middle', wrap: true });
    expect(fmtAt(store, 2, 2)).toMatchObject({ vAlign: 'middle', wrap: true });
  });

  it('bumps indent for every selected cell and clamps at Excel-style bounds', () => {
    for (let i = 0; i < 20; i += 1) bumpIndent(store.getState(), store, 1);

    expect(fmtAt(store, 1, 1)?.indent).toBe(15);
    expect(fmtAt(store, 2, 2)?.indent).toBe(15);

    for (let i = 0; i < 20; i += 1) bumpIndent(store.getState(), store, -1);

    expect(fmtAt(store, 1, 1)?.indent).toBe(0);
    expect(fmtAt(store, 2, 2)?.indent).toBe(0);
  });

  it('stages indent as pending input format for a single empty active cell', () => {
    const single = createSpreadsheetStore();
    setRange(single, 0, 0, 0, 0);
    mutators.setActive(single, { sheet: 0, row: 0, col: 0 });

    bumpIndent(single.getState(), single, 1);
    bumpIndent(single.getState(), single, 1);

    expect(single.getState().ui.pendingFormat).toEqual({
      addr: { sheet: 0, row: 0, col: 0 },
      format: { indent: 2 },
    });
    expect(fmtAt(single, 0, 0)).toBeUndefined();
    expect(effectiveFmtAt(single, 0, 0)?.indent).toBe(2);
  });

  it('sets text rotation across the range and clamps to the supported angle range', () => {
    setRotation(store.getState(), store, 45);
    expect(fmtAt(store, 1, 1)?.rotation).toBe(45);
    expect(fmtAt(store, 2, 2)?.rotation).toBe(45);

    setRotation(store.getState(), store, 120);
    expect(fmtAt(store, 1, 1)?.rotation).toBe(90);

    setRotation(store.getState(), store, -120);
    expect(fmtAt(store, 1, 1)?.rotation).toBe(-90);
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
    expect(effectiveFmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'fixed', decimals: 3 });
  });

  it('cycleCurrency turns currency on when none is set', () => {
    cycleCurrency(store.getState(), store);
    expect(effectiveFmtAt(store, 0, 0)?.numFmt).toEqual({
      kind: 'currency',
      decimals: 2,
      symbol: '$',
    });
  });

  it('cycleCurrency uses the active locale currency symbol', () => {
    cycleCurrency(store.getState(), store, 'ja');
    expect(effectiveFmtAt(store, 0, 0)?.numFmt).toEqual({
      kind: 'currency',
      decimals: 2,
      symbol: '¥',
    });
  });

  it('cycleCurrency clears back to general when at least one cell is currency', () => {
    cycleCurrency(store.getState(), store); // on
    cycleCurrency(store.getState(), store); // off
    expect(effectiveFmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'general' });
  });

  it('cyclePercent toggles percent on / off', () => {
    cyclePercent(store.getState(), store);
    expect(effectiveFmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'percent', decimals: 0 });
    cyclePercent(store.getState(), store);
    expect(effectiveFmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'general' });
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
    expect(effectiveFmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'fixed', decimals: 2 });
  });

  it('does nothing on -1 when the cell is general', () => {
    bumpDecimals(store.getState(), store, -1);
    expect(effectiveFmtAt(store, 0, 0)?.numFmt).toBeUndefined();
  });

  it('walks fixed decimals up and down with clamping', () => {
    setNumFmt(store.getState(), store, { kind: 'fixed', decimals: 0 });
    bumpDecimals(store.getState(), store, -1);
    expect(effectiveFmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'fixed', decimals: 0 });
    for (let i = 0; i < 12; i += 1) bumpDecimals(store.getState(), store, 1);
    expect(effectiveFmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'fixed', decimals: 10 });
  });

  it('preserves currency symbol while bumping decimals', () => {
    setNumFmt(store.getState(), store, { kind: 'currency', decimals: 2, symbol: '€' });
    bumpDecimals(store.getState(), store, 1);
    expect(effectiveFmtAt(store, 0, 0)?.numFmt).toEqual({
      kind: 'currency',
      decimals: 3,
      symbol: '€',
    });
  });

  it('walks percent decimals', () => {
    setNumFmt(store.getState(), store, { kind: 'percent', decimals: 0 });
    bumpDecimals(store.getState(), store, 1);
    expect(effectiveFmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'percent', decimals: 1 });
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

  it('setBorderPreset carries the selected line color into styled border sides', () => {
    setRange(store, 0, 0, 0, 0);
    setBorderPreset(store.getState(), store, 'outline', 'thick', '#c00000');

    expect(effectiveFmtAt(store, 0, 0)?.borders).toEqual({
      top: { style: 'thick', color: '#c00000' },
      bottom: { style: 'thick', color: '#c00000' },
      left: { style: 'thick', color: '#c00000' },
      right: { style: 'thick', color: '#c00000' },
    });
  });

  it('setBorderPreset applies inside borders only between cells in a range', () => {
    setRange(store, 0, 0, 1, 1);
    setBorderPreset(store.getState(), store, 'inside', 'thin', '#00a000');

    expect(fmtAt(store, 0, 0)?.borders).toBeUndefined();
    expect(fmtAt(store, 0, 1)?.borders).toEqual({
      left: { style: 'thin', color: '#00a000' },
    });
    expect(fmtAt(store, 1, 0)?.borders).toEqual({
      top: { style: 'thin', color: '#00a000' },
    });
    expect(fmtAt(store, 1, 1)?.borders).toEqual({
      top: { style: 'thin', color: '#00a000' },
      left: { style: 'thin', color: '#00a000' },
    });
  });

  it('setBorderPreset applies diagonal borders with the selected style and color', () => {
    setRange(store, 0, 0, 0, 0);
    setBorderPreset(store.getState(), store, 'diagonalDown', 'dashed', '#4472c4');
    setBorderPreset(store.getState(), store, 'diagonalUp', 'double', '#c00000');

    expect(effectiveFmtAt(store, 0, 0)?.borders).toEqual({
      diagonalDown: { style: 'dashed', color: '#4472c4' },
      diagonalUp: { style: 'double', color: '#c00000' },
    });
  });

  it('setBorderPreset merges with pending input format for a single empty active cell', () => {
    setRange(store, 0, 0, 0, 0);
    toggleBold(store.getState(), store);
    setBorderPreset(store.getState(), store, 'bottom', 'thin', '#4472c4');

    expect(fmtAt(store, 0, 0)).toBeUndefined();
    expect(store.getState().ui.pendingFormat).toEqual({
      addr: { sheet: 0, row: 0, col: 0 },
      format: {
        bold: true,
        borders: { bottom: { style: 'thin', color: '#4472c4' } },
      },
    });
    expect(effectiveFmtAt(store, 0, 0)?.borders?.bottom).toEqual({
      style: 'thin',
      color: '#4472c4',
    });
  });

  it('cycleBorders uses pending input format for a single empty active cell', () => {
    setRange(store, 0, 0, 0, 0);
    toggleItalic(store.getState(), store);
    cycleBorders(store.getState(), store);

    expect(fmtAt(store, 0, 0)).toBeUndefined();
    expect(store.getState().ui.pendingFormat).toEqual({
      addr: { sheet: 0, row: 0, col: 0 },
      format: {
        italic: true,
        borders: { top: true, right: true, bottom: true, left: true },
      },
    });
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
    expect(effectiveFmtAt(store, 0, 0)).toBeDefined();
    clearFormat(store.getState(), store);
    expect(fmtAt(store, 0, 0)).toBeUndefined();
    expect(store.getState().ui.pendingFormat).toBeNull();
  });

  it('clearVisualFormat preserves metadata while removing visual fields', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        bold: true,
        cellStyle: 'good',
        fill: '#ff0000',
        numFmt: { kind: 'fixed', decimals: 2 },
        comment: 'keep',
        hyperlink: 'https://example.com',
        validation: { kind: 'list', source: ['A', 'B'] },
        locked: false,
      },
    );

    clearVisualFormat(store.getState(), store);

    expect(fmtAt(store, 0, 0)).toEqual({
      comment: 'keep',
      hyperlink: 'https://example.com',
      validation: { kind: 'list', source: ['A', 'B'] },
      locked: false,
    });
  });

  it('setFontColor / setFillColor write and clear', () => {
    setFontColor(store.getState(), store, '#ff0000');
    expect(effectiveFmtAt(store, 0, 0)?.color).toBe('#ff0000');
    setFontColor(store.getState(), store, null);
    expect(effectiveFmtAt(store, 0, 0)?.color).toBeUndefined();

    setFillColor(store.getState(), store, '#0f0');
    expect(effectiveFmtAt(store, 0, 0)?.fill).toBe('#0f0');
    setFillColor(store.getState(), store, null);
    expect(effectiveFmtAt(store, 0, 0)?.fill).toBeUndefined();
  });

  it('setFont updates family / size and clears with null', () => {
    setFont(store.getState(), store, { fontFamily: 'Inter', fontSize: 14 });
    expect(effectiveFmtAt(store, 0, 0)?.fontFamily).toBe('Inter');
    expect(effectiveFmtAt(store, 0, 0)?.fontSize).toBe(14);

    setFont(store.getState(), store, { fontFamily: null });
    expect(effectiveFmtAt(store, 0, 0)?.fontFamily).toBeUndefined();
    // size untouched.
    expect(effectiveFmtAt(store, 0, 0)?.fontSize).toBe(14);
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
    expect(formatNumber(1234.5, { kind: 'fixed', decimals: 2, thousands: true })).toBe('1,234.50');
  });

  it('solves fraction formats instead of printing ?/? verbatim (H-3)', () => {
    expect(formatNumber(1.5, { kind: 'custom', pattern: '# ?/?' })).toBe('1 1/2');
    expect(formatNumber(0.5, { kind: 'custom', pattern: '# ?/?' })).toBe('1/2');
    expect(formatNumber(2.25, { kind: 'custom', pattern: '# ??/??' })).toBe('2 1/4');
    // Fixed denominator: 0.625 in eighths → 5/8.
    expect(formatNumber(0.625, { kind: 'custom', pattern: '# ?/8' })).toBe('5/8');
    // Whole number → no fraction shown.
    expect(formatNumber(3, { kind: 'custom', pattern: '# ?/?' })).toBe('3');
    // Negative keeps its sign.
    expect(formatNumber(-1.5, { kind: 'custom', pattern: '# ?/?' })).toBe('-1 1/2');
  });

  it('zero-pads the scientific exponent to two digits (H-4)', () => {
    expect(formatNumber(12345, { kind: 'scientific', decimals: 2 })).toBe('1.23E+04');
    expect(formatNumber(0.00012, { kind: 'scientific', decimals: 2 })).toBe('1.20E-04');
    // Already ≥ 2 exponent digits stays unchanged.
    expect(formatNumber(1.23e15, { kind: 'scientific', decimals: 2 })).toBe('1.23E+15');
  });

  it('renders elapsed-time formats without wrapping at 24h/60 (H-5)', () => {
    // 1.5 days = 36 hours. [h] must show 36, not 12.
    expect(formatNumber(1.5, { kind: 'custom', pattern: '[h]:mm:ss' })).toBe('36:00:00');
    // 90 minutes = 0.0625 days → [m] = 90, [mm]:ss stays within its unit.
    expect(formatNumber(90 / 1440, { kind: 'custom', pattern: '[m]:ss' })).toBe('90:00');
    // 1h30m45s → total seconds for [s].
    const serial = (1 * 3600 + 30 * 60 + 45) / 86400;
    expect(formatNumber(serial, { kind: 'custom', pattern: '[s]' })).toBe('5445');
    // [hh] zero-pads to at least two digits.
    expect(formatNumber(5 / 24, { kind: 'custom', pattern: '[hh]:mm' })).toBe('05:00');
  });

  it('renders five-m month tokens as a single month initial', () => {
    expect(formatNumber(45651, { kind: 'custom', pattern: 'mmmm' }, 'en-US')).toBe('December');
    expect(formatNumber(45651, { kind: 'custom', pattern: 'mmmmm' }, 'en-US')).toBe('D');
  });

  it('matches Excel 1900 date serials before the leap-year bug boundary', () => {
    expect(formatNumber(1, { kind: 'custom', pattern: 'yyyy-mm-dd' })).toBe('1900-01-01');
    expect(formatNumber(59, { kind: 'custom', pattern: 'yyyy-mm-dd' })).toBe('1900-02-28');
    expect(formatNumber(61, { kind: 'custom', pattern: 'yyyy-mm-dd' })).toBe('1900-03-01');
  });

  it('renders scientific and engineering notation instead of E+0 verbatim (H-6)', () => {
    // Engineering: 3 integer placeholders → exponent stepped in multiples of 3.
    expect(formatNumber(12345, { kind: 'custom', pattern: '##0.0E+0' })).toBe('12.3E+3');
    expect(formatNumber(0.00012345, { kind: 'custom', pattern: '##0.0E+0' })).toBe('123.5E-6');
    // Plain scientific custom pattern.
    expect(formatNumber(12345, { kind: 'custom', pattern: '0.00E+00' })).toBe('1.23E+04');
    // Negative keeps its sign; exponent sign spec `E-` omits the plus.
    expect(formatNumber(-12345, { kind: 'custom', pattern: '0.0E-0' })).toBe('-1.2E4');
  });

  it('distributes digits across multi-run integer patterns (M-1)', () => {
    // Phone number: digits spread right-to-left across the placeholder runs.
    expect(formatNumber(1234567890, { kind: 'custom', pattern: '000-000-0000' })).toBe(
      '123-456-7890',
    );
    // SSN-style grouping.
    expect(formatNumber(123456789, { kind: 'custom', pattern: '000-00-0000' })).toBe('123-45-6789');
    // A plain grouped pattern is NOT treated as multi-run.
    expect(formatNumber(1234567, { kind: 'custom', pattern: '#,##0' })).toBe('1,234,567');
  });

  it('renders currency with prefix symbol and respects negatives', () => {
    expect(formatNumber(99, { kind: 'currency', decimals: 2, symbol: '$' })).toBe('$99.00');
    expect(formatNumber(-99, { kind: 'currency', decimals: 0, symbol: '€' })).toBe('-€99');
  });

  it('renders Excel-style negative number variants', () => {
    expect(formatNumber(-99, { kind: 'fixed', decimals: 0, negativeStyle: 'parens' })).toBe('(99)');
    expect(formatNumber(-99, { kind: 'fixed', decimals: 0, negativeStyle: 'red' })).toBe('-99');
    expect(
      formatNumber(-99, {
        kind: 'currency',
        decimals: 0,
        symbol: '$',
        negativeStyle: 'red-parens',
      }),
    ).toBe('($99)');
  });

  it('falls back to "$" when currency symbol is missing', () => {
    expect(formatNumber(7.5, { kind: 'currency', decimals: 1 })).toBe('$7.5');
  });

  it('renders percent values', () => {
    expect(formatNumber(0.25, { kind: 'percent', decimals: 0 })).toBe('25%');
    expect(formatNumber(0.1234, { kind: 'percent', decimals: 2 })).toBe('12.34%');
  });

  it('does not scale for percent signs inside quoted custom literals', () => {
    expect(formatNumber(0.25, { kind: 'custom', pattern: '0%' })).toBe('25%');
    expect(formatNumber(0.25, { kind: 'custom', pattern: '0"%"' })).toBe('0%');
  });

  it('renders spreadsheet locale-tagged currency custom formats', () => {
    const fmt = { kind: 'custom' as const, pattern: '[$¥-411]#,##0;[Red]-[$¥-411]#,##0' };
    expect(formatNumber(1234, fmt, 'ja-JP')).toBe('¥1,234');
    expect(formatNumber(-1234, fmt, 'ja-JP')).toBe('-¥1,234');
  });

  it('renders Special category masks without losing the category semantics', () => {
    expect(formatNumber(12345, { kind: 'special', pattern: '00000' })).toBe('12345');
    expect(formatNumber(123456789, { kind: 'special', pattern: '00000-0000' })).toBe('12345-6789');
    expect(formatNumber(123456789, { kind: 'special', pattern: '000-00-0000' })).toBe(
      '123-45-6789',
    );
  });

  it('hides spreadsheet accounting spacing and fill directives in custom output', () => {
    const fmt = {
      kind: 'custom' as const,
      pattern: '_-"$"* #,##0.00_-;-"$"* #,##0.00_-;_-"$"* "-"??_-;_-@_-',
    };
    expect(formatNumber(1234, fmt)).toBe('$1,234.00');
    expect(formatNumber(-1234, fmt)).toBe('-$1,234.00');
  });
});
