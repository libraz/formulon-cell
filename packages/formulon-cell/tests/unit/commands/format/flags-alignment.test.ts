import { beforeEach, describe, expect, it } from 'vitest';
import {
  bumpIndent,
  setAlign,
  setRotation,
  setVAlign,
  toggleBold,
  toggleItalic,
  toggleStrike,
  toggleUnderline,
  toggleWrap,
} from '../../../../src/commands/format.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import { effectiveFmtAt, fmtAt, setRange } from './fixtures.js';

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
