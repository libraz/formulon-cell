import { describe, expect, it } from 'vitest';
import { planSelectionFormat } from '../../../src/commands/format.js';
import { addrKey, MAX_COL, MAX_ROW } from '../../../src/engine/address.js';
import type { Range } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { nativeFontSamplesForDialog } from '../../../src/interact/format-dialog.js';
import { summarizeDialogFormats } from '../../../src/interact/format-dialog-state.js';
import type { SpreadsheetStore } from '../../../src/store/store.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const selectRange = (store: SpreadsheetStore, range: Range): void => {
  const active = { sheet: range.sheet, row: range.r0, col: range.c0 };
  store.setState((state) => ({
    ...state,
    selection: { ...state.selection, active, anchor: active, range, extraRanges: [] },
  }));
};

const columnA: Range = { sheet: 0, r0: 0, c0: 0, r1: MAX_ROW, c1: 0 };

const makeWorkbook = (
  styleBases: readonly number[] | null = [0, 1, 1],
  missingXfIndex?: number,
  missingFontIndex?: number,
) => {
  const calls: Range[] = [];
  const workbook = {
    capabilities: { rangeStylePatch: true },
    workbookDefaultFont: { name: 'Arial', size: 11, scheme: 0 },
    workbookDisplayFont: { name: 'Arial', size: 11 },
    getRangeStyleBases(range: Range) {
      calls.push(range);
      return styleBases;
    },
    getCellXf(index: number) {
      if (index === missingXfIndex) return null;
      return { fontIndex: index, fillIndex: 0, borderIndex: 0, numFmtId: 0 };
    },
    getResolvedFontRecord(index: number) {
      if (index === missingFontIndex) return null;
      return {
        name: 'Arial',
        size: 11,
        bold: index === 1,
        italic: false,
        strike: false,
        underline: 0,
        colorArgb: 0xff000000,
        outline: false,
        shadow: false,
        vertAlign: 0,
      };
    },
    getResolvedFillRecord() {
      return { pattern: 0, fgArgb: 0, bgArgb: 0 };
    },
    getResolvedBorderRecord() {
      const side = { style: 0, colorArgb: 0xff000000 };
      return {
        left: side,
        right: side,
        top: side,
        bottom: side,
        diagonal: side,
        diagonalUp: false,
        diagonalDown: false,
      };
    },
  } as unknown as WorkbookHandle;
  return { workbook, calls };
};

describe('interact/format-dialog native font mixed summary', () => {
  it('projects distinct native XFs once per range and resolves the default font', () => {
    const store = createSpreadsheetStore();
    selectRange(store, columnA);
    const plan = planSelectionFormat(store.getState());
    if (!plan) throw new Error('whole-column plan missing');
    const { workbook, calls } = makeWorkbook();

    const samples = nativeFontSamplesForDialog(store.getState(), plan, workbook);

    expect(calls).toEqual([columnA]);
    expect(samples).toHaveLength(2);
    expect(samples[0]).toMatchObject({ fontFamily: 'Arial', fontSize: 14.666666666666666 });
    expect(samples[0]?.bold).toBeUndefined();
    expect(samples[1]?.bold).toBe(true);
  });

  it('does not mark a homogeneous native default font as mixed', () => {
    const store = createSpreadsheetStore();
    selectRange(store, columnA);
    const plan = planSelectionFormat(store.getState());
    if (!plan) throw new Error('whole-column plan missing');
    const { workbook, calls } = makeWorkbook([0, 0]);
    const samples = nativeFontSamplesForDialog(store.getState(), plan, workbook);
    const active = samples[0];
    if (!active) throw new Error('native font sample missing');

    const summary = summarizeDialogFormats(store.getState(), plan, () => active, samples);

    expect(calls).toHaveLength(1);
    expect(summary.mixed.has('bold')).toBe(false);
    expect(summary.mixed.has('fontFamily')).toBe(false);
    expect(summary.mixed.has('fontSize')).toBe(false);
  });

  it('ignores formats outside the selected axis and gates samples for selected formats or pending input', () => {
    const store = createSpreadsheetStore();
    selectRange(store, columnA);
    const outside = { sheet: 0, row: 0, col: 1 };
    mutators.setCellFormat(store, outside, { italic: true });
    const plan = planSelectionFormat(store.getState());
    if (!plan) throw new Error('whole-column plan missing');
    const first = makeWorkbook();
    expect(nativeFontSamplesForDialog(store.getState(), plan, first.workbook)).toHaveLength(2);
    expect(first.calls).toHaveLength(1);

    const selected = { sheet: 0, row: 5, col: 0 };
    mutators.setCellFormat(store, selected, { bold: true });
    const authored = makeWorkbook();
    expect(nativeFontSamplesForDialog(store.getState(), plan, authored.workbook)).toHaveLength(2);
    expect(authored.calls).toEqual([
      { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 },
      { sheet: 0, r0: 6, c0: 0, r1: MAX_ROW, c1: 0 },
    ]);

    store.setState((state) => ({
      ...state,
      format: { ...state.format, formats: new Map([[addrKey(outside), { italic: true }]]) },
      ui: { ...state.ui, pendingFormat: { addr: selected, format: { bold: false } } },
    }));
    const pending = makeWorkbook();
    expect(nativeFontSamplesForDialog(store.getState(), plan, pending.workbook)).toHaveLength(2);
    expect(pending.calls).toEqual([
      { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 },
      { sheet: 0, r0: 6, c0: 0, r1: MAX_ROW, c1: 0 },
    ]);
  });

  it('still samples native bases on a whole sheet that holds a store-formatted cell', () => {
    const store = createSpreadsheetStore();
    selectRange(store, { sheet: 0, r0: 0, c0: 0, r1: MAX_ROW, c1: MAX_COL });
    mutators.setCellFormat(store, { sheet: 0, row: 5, col: 5 }, { italic: true });
    const plan = planSelectionFormat(store.getState());
    if (!plan) throw new Error('whole-sheet plan missing');
    const { workbook, calls } = makeWorkbook();

    const samples = nativeFontSamplesForDialog(store.getState(), plan, workbook);

    expect(samples).toHaveLength(2);
    expect(calls).toEqual([
      { sheet: 0, r0: 0, c0: 0, r1: 4, c1: MAX_COL },
      { sheet: 0, r0: 5, c0: 0, r1: 5, c1: 4 },
      { sheet: 0, r0: 5, c0: 6, r1: 5, c1: MAX_COL },
      { sheet: 0, r0: 6, c0: 0, r1: MAX_ROW, c1: MAX_COL },
    ]);
    const active = samples[0];
    if (!active) throw new Error('native font sample missing');
    const summary = summarizeDialogFormats(store.getState(), plan, () => active, samples);
    expect(summary.mixed.has('bold')).toBe(true);
  });

  it('falls back when the selected sheet has overrides, an intersecting merge, or incomplete XF data', () => {
    const store = createSpreadsheetStore();
    selectRange(store, columnA);
    const plan = planSelectionFormat(store.getState());
    if (!plan) throw new Error('whole-column plan missing');

    for (const field of ['sheetFontOverrides', 'sheetAxisOverrides'] as const) {
      store.setState((state) => ({
        ...state,
        format: { ...state.format, [field]: new Map([[0, {}]]) },
      }));
      const overridden = makeWorkbook();
      expect(nativeFontSamplesForDialog(store.getState(), plan, overridden.workbook)).toEqual([]);
      expect(overridden.calls).toHaveLength(0);
      store.setState((state) => ({
        ...state,
        format: { ...state.format, [field]: undefined },
      }));
    }

    const anchor = { sheet: 0, row: 10, col: 0 };
    const merge: Range = { sheet: 0, r0: 10, c0: 0, r1: 11, c1: 1 };
    store.setState((state) => ({
      ...state,
      merges: {
        ...state.merges,
        byAnchor: new Map([[addrKey(anchor), merge]]),
      },
    }));
    const merged = makeWorkbook();
    expect(nativeFontSamplesForDialog(store.getState(), plan, merged.workbook)).toEqual([]);
    expect(merged.calls).toHaveLength(0);

    const incomplete = makeWorkbook(null);
    store.setState((state) => ({ ...state, merges: { ...state.merges, byAnchor: new Map() } }));
    expect(nativeFontSamplesForDialog(store.getState(), plan, incomplete.workbook)).toEqual([]);
    expect(incomplete.calls).toHaveLength(1);

    const unprojectable = makeWorkbook([0, 1], 1);
    expect(nativeFontSamplesForDialog(store.getState(), plan, unprojectable.workbook)).toEqual([]);
    expect(unprojectable.calls).toHaveLength(1);

    const missingFont = makeWorkbook([0, 1], undefined, 1);
    expect(nativeFontSamplesForDialog(store.getState(), plan, missingFont.workbook)).toEqual([]);
    expect(missingFont.calls).toHaveLength(1);
  });
});
