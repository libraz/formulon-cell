import { beforeEach, describe, expect, it } from 'vitest';
import {
  applySelectionFormatAction,
  applySelectionFormatPatch,
  bumpDecimals,
  bumpIndent,
  clearFormat,
  clearVisualFormat,
  cycleBorders,
  cycleCurrency,
  cyclePercent,
  planSelectionFormat,
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
  withSelectionFormatOrigin,
} from '../../../src/commands/format.js';
import { History, recordRepeatableFormatChange } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { fixedFormPolicy } from '../../../src/commands/interaction-policy.js';
import { setCellLocked, setProtectedSheet } from '../../../src/commands/protection.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
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

const setSelection = (
  store: SpreadsheetStore,
  range: { sheet: number; r0: number; c0: number; r1: number; c1: number },
  extraRanges: (typeof range)[] = [],
): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      ...s.selection,
      active: { sheet: range.sheet, row: range.r0, col: range.c0 },
      anchor: { sheet: range.sheet, row: range.r0, col: range.c0 },
      range,
      extraRanges,
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

  it('applies all borders across every selected fragment while leaving a hole untouched', () => {
    const rectangles = [
      { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 },
      { sheet: 0, r0: 0, c0: 2, r1: 2, c1: 2 },
      { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
      { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
      { sheet: 0, r0: 4, c0: 4, r1: 4, c1: 4 },
    ];
    const [primary, ...extras] = rectangles;
    if (!primary) throw new Error('Missing primary border range');
    setSelection(store, primary, extras);

    setBorderPreset(store.getState(), store, 'all');

    for (const [row, col] of [
      [0, 0],
      [1, 0],
      [2, 0],
      [0, 2],
      [1, 2],
      [2, 2],
      [0, 1],
      [2, 1],
      [4, 4],
    ] as [number, number][]) {
      expect(fmtAt(store, row, col)?.borders).toEqual({
        top: { style: 'thin' },
        right: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
      });
    }
    expect(fmtAt(store, 1, 1)).toBeUndefined();
  });

  it('keeps border edges independent for separate areas and does not bridge the gap', () => {
    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 }, [
      { sheet: 0, r0: 0, c0: 3, r1: 1, c1: 4 },
    ]);
    setBorderPreset(store.getState(), store, 'outline');

    expect(fmtAt(store, 0, 1)?.borders?.right).toEqual({ style: 'thin' });
    expect(fmtAt(store, 0, 3)?.borders?.left).toEqual({ style: 'thin' });
    expect(fmtAt(store, 0, 2)).toBeUndefined();

    const inside = createSpreadsheetStore();
    setSelection(inside, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 }, [
      { sheet: 0, r0: 0, c0: 3, r1: 1, c1: 4 },
    ]);
    setBorderPreset(inside.getState(), inside, 'inside');

    expect(fmtAt(inside, 0, 0)).toBeUndefined();
    expect(fmtAt(inside, 1, 1)?.borders).toEqual({
      top: { style: 'thin' },
      left: { style: 'thin' },
    });
    expect(fmtAt(inside, 0, 3)).toBeUndefined();
    expect(fmtAt(inside, 1, 4)?.borders).toEqual({
      top: { style: 'thin' },
      left: { style: 'thin' },
    });
    expect(fmtAt(inside, 0, 2)).toBeUndefined();
  });

  it('cycles the complete disjoint union through outline, all, and clear', () => {
    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 }, [
      { sheet: 0, r0: 0, c0: 3, r1: 1, c1: 4 },
    ]);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { borders: { diagonalDown: true } },
    );

    cycleBorders(store.getState(), store);
    expect(fmtAt(store, 0, 1)?.borders?.right).toBe(true);
    expect(fmtAt(store, 0, 3)?.borders?.left).toBe(true);
    cycleBorders(store.getState(), store);
    for (const [row, col] of [
      [0, 0],
      [0, 1],
      [1, 0],
      [1, 1],
      [0, 3],
      [0, 4],
      [1, 3],
      [1, 4],
    ] as [number, number][]) {
      expect(fmtAt(store, row, col)?.borders).toMatchObject({
        top: true,
        right: true,
        bottom: true,
        left: true,
      });
    }
    cycleBorders(store.getState(), store);
    for (const [row, col] of [
      [0, 0],
      [0, 1],
      [1, 0],
      [1, 1],
      [0, 3],
      [0, 4],
      [1, 3],
      [1, 4],
    ] as [number, number][]) {
      expect(fmtAt(store, row, col)?.borders).toMatchObject({
        top: false,
        right: false,
        bottom: false,
        left: false,
      });
    }
    expect(fmtAt(store, 0, 2)).toBeUndefined();
    expect(fmtAt(store, 1, 2)).toBeUndefined();
    expect(fmtAt(store, 0, 0)?.borders?.diagonalDown).toBe(true);
  });

  it('rejects a capped union atomically and publishes one update for a valid multi-area preset', () => {
    const capped = createSpreadsheetStore();
    setSelection(capped, { sheet: 0, r0: 0, c0: 0, r1: 999, c1: 99 }, [
      { sheet: 0, r0: 1000, c0: 0, r1: 1000, c1: 0 },
    ]);
    mutators.setCellFormat(capped, { sheet: 0, row: 0, col: 0 }, { bold: true });
    mutators.setPendingFormat(capped, {
      addr: { sheet: 0, row: 0, col: 0 },
      format: { italic: true },
    });
    const beforeFormats = new Map(capped.getState().format.formats);
    const beforePending = capped.getState().ui.pendingFormat;

    setBorderPreset(capped.getState(), capped, 'outline');

    expect(capped.getState().format.formats).toEqual(beforeFormats);
    expect(capped.getState().ui.pendingFormat).toEqual(beforePending);

    const published = createSpreadsheetStore();
    setSelection(published, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 }, [
      { sheet: 0, r0: 0, c0: 3, r1: 1, c1: 4 },
    ]);
    let updates = 0;
    const unsubscribe = published.subscribe(() => {
      updates += 1;
    });
    setBorderPreset(published.getState(), published, 'outline');
    unsubscribe();
    expect(updates).toBe(1);
  });

  it('rejects restricted mixed selections and skips locked fragments while formatting unlocked cells', async () => {
    const restricted = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const controller = new InteractionController({
      store: restricted,
      getWb: () => workbook,
      history: new History(),
    });
    const unregister = registerInteractionController(restricted, controller);
    try {
      setSelection(restricted, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, [
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
      ]);
      controller.setPolicy({
        ...fixedFormPolicy([{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }]),
        operations: { format: true },
      });
      setBorderPreset(restricted.getState(), restricted, 'outline');
      expect(restricted.getState().format.formats.size).toBe(0);
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }

    const protectedStore = createSpreadsheetStore();
    setSelection(protectedStore, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 }, [
      { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 4 },
    ]);
    setCellLocked(protectedStore, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, false);
    setCellLocked(protectedStore, { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 }, false);
    setProtectedSheet(protectedStore, 0, true);
    setBorderPreset(protectedStore.getState(), protectedStore, 'outline');
    expect(fmtAt(protectedStore, 0, 0)?.borders).toBeDefined();
    expect(fmtAt(protectedStore, 0, 3)?.borders).toBeDefined();
    expect(fmtAt(protectedStore, 0, 1)).toBeUndefined();
    expect(fmtAt(protectedStore, 0, 4)).toBeUndefined();
  });

  it('rejects partial merge coverage and formats a fully covered merge plus remote area', () => {
    const partial = createSpreadsheetStore();
    const merge = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    mutators.mergeRange(partial, merge);
    setSelection(partial, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 }, [
      { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
    ]);
    setBorderPreset(partial.getState(), partial, 'outline');
    expect(partial.getState().format.formats.size).toBe(0);

    const complete = createSpreadsheetStore();
    mutators.mergeRange(complete, merge);
    setSelection(complete, merge, [{ sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 }]);
    setBorderPreset(complete.getState(), complete, 'outline');
    expect(fmtAt(complete, 0, 0)?.borders?.top).toEqual({ style: 'thin' });
    expect(fmtAt(complete, 1, 1)?.borders?.bottom).toEqual({ style: 'thin' });
    expect(fmtAt(complete, 0, 3)?.borders?.top).toEqual({ style: 'thin' });
  });

  it('keeps one-cell inside borders pending and materializes a selection with extras', () => {
    const single = createSpreadsheetStore();
    setSelection(single, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    toggleBold(single.getState(), single);
    const pendingBefore = single.getState().ui.pendingFormat;
    setBorderPreset(single.getState(), single, 'inside');
    expect(single.getState().ui.pendingFormat).toEqual(pendingBefore);
    expect(single.getState().format.formats.size).toBe(0);

    const multi = createSpreadsheetStore();
    setSelection(multi, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, [
      { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
    ]);
    setBorderPreset(multi.getState(), multi, 'outline');
    expect(multi.getState().ui.pendingFormat).toBeNull();
    expect(fmtAt(multi, 0, 0)?.borders).toBeDefined();
    expect(fmtAt(multi, 0, 1)?.borders).toBeDefined();
  });

  it('records a multi-area border preset as one undo entry', () => {
    const history = new History();
    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 }, [
      { sheet: 0, r0: 0, c0: 3, r1: 1, c1: 4 },
    ]);
    recordRepeatableFormatChange(history, store, () => {
      setBorderPreset(store.getState(), store, 'all');
    });

    expect(history.canUndo()).toBe(true);
    expect(history.undo()).toBe(true);
    expect(store.getState().format.formats.size).toBe(0);
    expect(history.undo()).toBe(false);
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

describe('non-contiguous selection formatting', () => {
  it('toggles the complete union while leaving a hole untouched', () => {
    const store = createSpreadsheetStore();
    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 }, [
      { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 4 },
    ]);
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });

    toggleBold(store.getState(), store);
    for (const col of [0, 1, 3, 4]) expect(fmtAt(store, 0, col)?.bold).toBe(true);
    expect(fmtAt(store, 0, 2)).toBeUndefined();

    toggleBold(store.getState(), store);
    for (const col of [0, 1, 3, 4]) expect(fmtAt(store, 0, col)?.bold).toBe(false);
    expect(fmtAt(store, 0, 2)).toBeUndefined();
  });

  it('applies relative changes once to overlapping union areas', () => {
    const store = createSpreadsheetStore();
    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 }, [
      { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 },
    ]);
    bumpIndent(store.getState(), store, 1);
    expect(fmtAt(store, 0, 0)?.indent).toBe(1);
    expect(fmtAt(store, 1, 1)?.indent).toBe(1);
    expect(fmtAt(store, 2, 2)?.indent).toBe(1);
  });

  it('rejects a union over the materialization cap without a partial write', () => {
    const store = createSpreadsheetStore();
    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 999, c1: 99 }, [
      { sheet: 0, r0: 1000, c0: 0, r1: 1000, c1: 0 },
    ]);

    toggleBold(store.getState(), store);

    expect(store.getState().format.formats.size).toBe(0);
  });

  it('stages one blank cell but materializes a blank extra selection', () => {
    const single = createSpreadsheetStore();
    setSelection(single, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    mutators.setActive(single, { sheet: 0, row: 0, col: 0 });
    toggleBold(single.getState(), single);
    expect(single.getState().ui.pendingFormat?.format.bold).toBe(true);
    expect(single.getState().format.formats.size).toBe(0);

    const multi = createSpreadsheetStore();
    setSelection(multi, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, [
      { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
    ]);
    toggleBold(multi.getState(), multi);
    expect(multi.getState().ui.pendingFormat).toBeNull();
    expect(fmtAt(multi, 0, 0)?.bold).toBe(true);
    expect(fmtAt(multi, 0, 1)?.bold).toBe(true);
  });

  it('requires complete merge coverage before changing a union', () => {
    const store = createSpreadsheetStore();
    const merge = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    mutators.mergeRange(store, merge);
    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    toggleBold(store.getState(), store);
    expect(store.getState().format.formats.size).toBe(0);

    setSelection(store, merge, [{ sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 }]);
    toggleBold(store.getState(), store);
    for (const [row, col] of [
      [0, 0],
      [0, 1],
      [1, 0],
      [1, 1],
      [0, 3],
    ] as [number, number][]) {
      expect(fmtAt(store, row, col)?.bold).toBe(true);
    }
  });

  it('keeps locked cells unchanged while formatting unlocked cells', () => {
    const store = createSpreadsheetStore();
    setCellLocked(store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }, false);
    setProtectedSheet(store, 0, true);
    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, [
      { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
    ]);

    toggleBold(store.getState(), store);

    expect(fmtAt(store, 0, 0)).toBeUndefined();
    expect(fmtAt(store, 0, 1)?.bold).toBe(true);
  });

  it('keeps the legacy locked-cell skip with a registered controller without policy', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const controller = new InteractionController({
      store,
      getWb: () => workbook,
      history: new History(),
    });
    const unregister = registerInteractionController(store, controller);
    try {
      setCellLocked(store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }, false);
      setCellLocked(store, { sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 }, false);
      setProtectedSheet(store, 0, true);
      setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 }, [
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 4 },
      ]);

      toggleBold(store.getState(), store);

      expect(fmtAt(store, 0, 0)).toBeUndefined();
      expect(fmtAt(store, 0, 1)?.bold).toBe(true);
      expect(fmtAt(store, 0, 3)).toBeUndefined();
      expect(fmtAt(store, 0, 4)?.bold).toBe(true);
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });

  it('rejects a mixed selection atomically under an explicit format policy', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const controller = new InteractionController({
      store,
      getWb: () => workbook,
      history: new History(),
    });
    const unregister = registerInteractionController(store, controller);
    try {
      setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, [
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
      ]);
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { italic: true });
      const policy = fixedFormPolicy({
        ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }],
      });
      controller.setPolicy({
        ...policy,
        operations: { ...policy.operations, format: true },
      });
      const before = new Map(store.getState().format.formats);
      const pendingBefore = store.getState().ui.pendingFormat;

      toggleBold(store.getState(), store);

      expect(store.getState().format.formats).toEqual(before);
      expect(store.getState().ui.pendingFormat).toBe(pendingBefore);
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });

  it('clears sparse visual formats across a huge primary and extra range', () => {
    const store = createSpreadsheetStore();
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 4, col: 2 },
      {
        bold: true,
        comment: 'keep',
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 8, col: 4 },
      {
        fill: '#ffeeaa',
        validation: { kind: 'list', source: ['A'] },
      },
    );
    mutators.setCellFormat(store, { sheet: 0, row: 4, col: 3 }, { italic: true });
    mutators.setCellFormat(store, { sheet: 1, row: 4, col: 2 }, { fill: '#000000' });
    setSelection(store, { sheet: 0, r0: 0, c0: 2, r1: 1_048_575, c1: 2 }, [
      { sheet: 0, r0: 8, c0: 4, r1: 8, c1: 4 },
    ]);

    clearVisualFormat(store.getState(), store);

    expect(fmtAt(store, 4, 2)).toEqual({ comment: 'keep' });
    expect(fmtAt(store, 8, 4)).toEqual({ validation: { kind: 'list', source: ['A'] } });
    expect(fmtAt(store, 4, 3)?.italic).toBe(true);
    expect(store.getState().format.formats.get('1:4:2')?.fill).toBe('#000000');
  });

  it('does not clear a protected pending format', () => {
    for (const clear of [clearFormat, clearVisualFormat]) {
      const store = createSpreadsheetStore();
      setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 1 }, { italic: true });
      toggleBold(store.getState(), store);
      const pendingBefore = store.getState().ui.pendingFormat;
      const formatsBefore = new Map(store.getState().format.formats);

      setProtectedSheet(store, 0, true);
      clear(store.getState(), store);

      expect(store.getState().ui.pendingFormat).toEqual(pendingBefore);
      expect(store.getState().format.formats).toEqual(formatsBefore);
    }
  });

  it('propagates nested command context and restores it after a throw', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const controller = new InteractionController({
      store,
      getWb: () => workbook,
      history: new History(),
    });
    const unregister = registerInteractionController(store, controller);
    const seen: string[] = [];
    try {
      setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
      controller.setPolicy({
        operations: { format: true },
        editable: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }],
        restrict: ({ intent }) => {
          seen.push(`${intent.origin}:${intent.commandId ?? ''}`);
          return true;
        },
      });

      withSelectionFormatOrigin(
        store,
        'ribbon',
        () => {
          applySelectionFormatPatch(store.getState(), store, { bold: true });
          withSelectionFormatOrigin(store, 'keyboard', () => {
            applySelectionFormatPatch(
              store.getState(),
              store,
              { italic: true },
              { origin: 'instanceApi', commandId: 'italic' },
            );
          });
        },
        'bold',
      );

      expect(() =>
        withSelectionFormatOrigin(
          store,
          'contextMenu',
          () => {
            throw new Error('restore context');
          },
          'underline',
        ),
      ).toThrow('restore context');
      applySelectionFormatPatch(store.getState(), store, { strike: true });

      expect(seen).toEqual(['ribbon:bold', 'instanceApi:italic', 'instanceApi:']);
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });
});

describe('selection format actions', () => {
  it('does not stage a locked blank cell and only stages an unlocked blank cell', () => {
    const locked = createSpreadsheetStore();
    setSelection(locked, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setProtectedSheet(locked, 0, true);
    expect(applySelectionFormatAction(locked.getState(), locked, { patch: { bold: true } })).toBe(
      false,
    );
    expect(locked.getState().ui.pendingFormat).toBeNull();
    expect(locked.getState().format.formats.size).toBe(0);
    expect(
      applySelectionFormatAction(
        locked.getState(),
        locked,
        { patch: { bold: true } },
        { allowPending: false },
      ),
    ).toBe(false);
    expect(locked.getState().ui.pendingFormat).toBeNull();
    expect(locked.getState().format.formats.size).toBe(0);

    const unlocked = createSpreadsheetStore();
    setSelection(unlocked, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    expect(
      applySelectionFormatAction(unlocked.getState(), unlocked, { patch: { bold: true } }),
    ).toBe(true);
    expect(unlocked.getState().ui.pendingFormat?.format).toEqual({ bold: true });
  });

  it('plans the disjoint union and preserves untouched mixed metadata', () => {
    const store = createSpreadsheetStore();
    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 }, [
      { sheet: 0, r0: 4, c0: 4, r1: 4, c1: 4 },
    ]);
    const active = { sheet: 0, row: 0, col: 0 };
    const extra = { sheet: 0, row: 4, col: 4 };
    mutators.setCellFormat(store, active, {
      bold: true,
      hyperlink: 'https://example.test',
      comment: 'keep active metadata',
    });
    mutators.setCellFormat(store, extra, {
      bold: false,
      hyperlink: 'https://extra.test',
      comment: 'keep extra metadata',
    });

    const plan = planSelectionFormat(store.getState());
    expect(plan?.cells).toHaveLength(10);
    expect(
      applySelectionFormatAction(store.getState(), store, {
        patch: { align: 'center' },
      }),
    ).toBe(true);
    expect(fmtAt(store, 0, 0)).toMatchObject({
      align: 'center',
      bold: true,
      hyperlink: 'https://example.test',
    });
    expect(fmtAt(store, 4, 4)).toMatchObject({
      align: 'center',
      bold: false,
      hyperlink: 'https://extra.test',
    });
    expect(fmtAt(store, 1, 1)?.align).toBe('center');
  });

  it('derives metadata permissions and cannot bypass them with optional operations', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const controller = new InteractionController({
      store,
      getWb: () => workbook,
      history: new History(),
    });
    const unregister = registerInteractionController(store, controller);
    try {
      setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
      controller.setPolicy({
        ...fixedFormPolicy([{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }]),
        operations: { format: true },
      });
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
      const before = new Map(store.getState().format.formats);

      expect(
        applySelectionFormatAction(store.getState(), store, {
          patch: { bold: false, comment: 'blocked' },
          operations: ['format'],
        }),
      ).toBe(false);
      expect(store.getState().format.formats).toEqual(before);
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });

  it('preserves each target hyperlink metadata only when its URL is unchanged', () => {
    const store = createSpreadsheetStore();
    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, [
      { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
      { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
    ]);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        hyperlink: 'https://same.example',
        hyperlinkDisplay: 'A',
        hyperlinkTooltip: 'tip A',
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 1 },
      {
        hyperlink: 'https://same.example',
        hyperlinkDisplay: 'B',
        hyperlinkTooltip: 'tip B',
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 2 },
      {
        hyperlink: 'https://old.example',
        hyperlinkDisplay: 'C',
        hyperlinkTooltip: 'tip C',
      },
    );

    expect(
      applySelectionFormatAction(store.getState(), store, {
        patch: { hyperlink: 'https://same.example' },
      }),
    ).toBe(true);
    expect(fmtAt(store, 0, 0)).toMatchObject({
      hyperlink: 'https://same.example',
      hyperlinkDisplay: 'A',
      hyperlinkTooltip: 'tip A',
    });
    expect(fmtAt(store, 0, 1)).toMatchObject({
      hyperlink: 'https://same.example',
      hyperlinkDisplay: 'B',
      hyperlinkTooltip: 'tip B',
    });
    expect(fmtAt(store, 0, 2)).toMatchObject({ hyperlink: 'https://same.example' });
    expect(fmtAt(store, 0, 2)?.hyperlinkDisplay).toBeUndefined();
    expect(fmtAt(store, 0, 2)?.hyperlinkTooltip).toBeUndefined();
  });
});
