import { describe, expect, it } from 'vitest';
import {
  bumpDecimals,
  bumpIndent,
  clearFormat,
  clearVisualFormat,
  cycleBorders,
  cyclePercent,
  setBorderPreset,
} from '../../../../src/commands/format.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';
import { fmtAt, setRange } from './fixtures.js';

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
