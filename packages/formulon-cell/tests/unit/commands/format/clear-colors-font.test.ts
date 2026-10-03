import { beforeEach, describe, expect, it } from 'vitest';
import {
  clearFormat,
  clearVisualFormat,
  setFillColor,
  setFont,
  setFontColor,
  setNumFmt,
} from '../../../../src/commands/format.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import { effectiveFmtAt, fmtAt, setRange } from './fixtures.js';

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
