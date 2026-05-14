import { beforeEach, describe, expect, it } from 'vitest';

import {
  type PasteOperation,
  type PasteSpecialOptions,
  type PasteWhat,
  pasteSpecial,
} from '../../../../src/commands/clipboard/paste-special.js';
import { captureSnapshot } from '../../../../src/commands/clipboard/snapshot.js';
import { addrKey, WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import {
  type CellFormat,
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

type SeedRow = {
  row: number;
  col: number;
  value: number | string;
  formula?: string;
  format?: CellFormat;
};

const seed = (store: SpreadsheetStore, wb: WorkbookHandle, cells: SeedRow[]): void => {
  store.setState((s) => {
    const map = new Map(s.data.cells);
    const fmt = new Map(s.format.formats);
    for (const c of cells) {
      const addr = { sheet: 0, row: c.row, col: c.col };
      const k = addrKey(addr);
      if (c.formula) {
        wb.setFormula(addr, c.formula);
        map.set(k, {
          value:
            typeof c.value === 'number'
              ? { kind: 'number', value: c.value }
              : { kind: 'text', value: c.value },
          formula: c.formula,
        });
      } else if (typeof c.value === 'number') {
        wb.setNumber(addr, c.value);
        map.set(k, { value: { kind: 'number', value: c.value }, formula: null });
      } else {
        wb.setText(addr, c.value);
        map.set(k, { value: { kind: 'text', value: c.value }, formula: null });
      }
      if (c.format) fmt.set(k, c.format);
    }
    return {
      ...s,
      data: { ...s.data, cells: map },
      format: { ...s.format, formats: fmt },
    };
  });
  wb.recalc();
};

const setActive = (store: SpreadsheetStore, row: number, col: number): void => {
  mutators.setActive(store, { sheet: 0, row, col });
};

const opt = (over: Partial<PasteSpecialOptions> = {}): PasteSpecialOptions => ({
  what: 'all',
  operation: 'none',
  skipBlanks: false,
  transpose: false,
  ...over,
});

const cellAt = (wb: WorkbookHandle, row: number, col: number) =>
  wb.getValue({ sheet: 0, row, col });

const fmtAt = (store: SpreadsheetStore, row: number, col: number) =>
  store.getState().format.formats.get(addrKey({ sheet: 0, row, col }));

describe('paste-special — what × operation × transpose × skipBlanks matrix', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  describe('what — coverage of each enum value', () => {
    // Source is a plain numeric cell (no formula) so "values" and "formulas"
    // produce comparable observable state: both write 8. We separately cover
    // formula-shifting in the transpose section.
    const seedPlain = (): void => {
      seed(store, wb, [
        {
          row: 1,
          col: 0,
          value: 8,
          format: {
            bold: true,
            fill: '#ffff00',
            numFmt: { kind: 'currency', decimals: 2, symbol: '$' },
          },
        },
      ]);
    };

    const cases: Array<{
      what: PasteWhat;
      expect: { hasValue: boolean; hasBold: boolean; hasNumFmt: boolean };
    }> = [
      { what: 'all', expect: { hasValue: true, hasBold: true, hasNumFmt: true } },
      { what: 'values', expect: { hasValue: true, hasBold: false, hasNumFmt: false } },
      // No source formula → "formulas" leaves value alone. Format pieces are
      // also not pasted because `formulas` doesn't request them.
      { what: 'formulas', expect: { hasValue: false, hasBold: false, hasNumFmt: false } },
      { what: 'formats', expect: { hasValue: false, hasBold: true, hasNumFmt: true } },
      { what: 'values-and-numfmt', expect: { hasValue: true, hasBold: false, hasNumFmt: true } },
      // No source formula → no value write; numFmt still cherry-picked.
      {
        what: 'formulas-and-numfmt',
        expect: { hasValue: false, hasBold: false, hasNumFmt: true },
      },
    ];

    for (const c of cases) {
      it(`what = "${c.what}" applies the expected subset`, () => {
        seedPlain();
        const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 });
        if (!snap) throw new Error('no snapshot');
        setActive(store, 5, 5);
        const res = pasteSpecial(store.getState(), store, wb, snap, opt({ what: c.what }));
        expect(res).not.toBeNull();
        wb.recalc();

        const v = cellAt(wb, 5, 5);
        if (c.expect.hasValue) {
          expect(v).toMatchObject({ kind: 'number', value: 8 });
        } else {
          expect(v.kind).toBe('blank');
        }
        const fmt = fmtAt(store, 5, 5);
        expect(Boolean(fmt?.bold), `bold for what=${c.what}`).toBe(c.expect.hasBold);
        expect(Boolean(fmt?.numFmt), `numFmt for what=${c.what}`).toBe(c.expect.hasNumFmt);
      });
    }

    it('what="all" pastes the source formula (with refs shifted), not the cached value', () => {
      seed(store, wb, [
        { row: 0, col: 0, value: 7 },
        { row: 1, col: 0, value: 8, formula: '=A1+1' },
      ]);
      const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 });
      if (!snap) throw new Error('no snapshot');
      // Paste at (5, 5). Formula `=A1+1` shifts to `=E5+1`; E5 is blank so the
      // result is 0+1=1, not the original 8.
      setActive(store, 5, 5);
      pasteSpecial(store.getState(), store, wb, snap, opt({ what: 'all' }));
      wb.recalc();
      expect(cellAt(wb, 5, 5)).toMatchObject({ kind: 'number', value: 1 });
    });
  });

  describe('operation — arithmetic against an existing destination', () => {
    const arithCases: Array<{ op: PasteOperation; dest: number; src: number; expected: number }> = [
      { op: 'none', dest: 10, src: 3, expected: 3 },
      { op: 'add', dest: 10, src: 3, expected: 13 },
      { op: 'subtract', dest: 10, src: 3, expected: 7 },
      { op: 'multiply', dest: 10, src: 3, expected: 30 },
      { op: 'divide', dest: 10, src: 2, expected: 5 },
    ];

    for (const c of arithCases) {
      it(`operation = "${c.op}" → ${c.dest} ⊕ ${c.src} = ${c.expected}`, () => {
        seed(store, wb, [
          { row: 0, col: 0, value: c.src },
          { row: 5, col: 5, value: c.dest },
        ]);
        const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
        if (!snap) throw new Error('no snapshot');
        setActive(store, 5, 5);
        pasteSpecial(store.getState(), store, wb, snap, opt({ operation: c.op }));
        wb.recalc();
        const v = cellAt(wb, 5, 5);
        expect(v).toMatchObject({ kind: 'number', value: c.expected });
      });
    }

    it('divide by zero leaves destination untouched (NaN is skipped)', () => {
      seed(store, wb, [
        { row: 0, col: 0, value: 0 },
        { row: 5, col: 5, value: 9 },
      ]);
      const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
      if (!snap) throw new Error('no snapshot');
      setActive(store, 5, 5);
      pasteSpecial(store.getState(), store, wb, snap, opt({ operation: 'divide' }));
      wb.recalc();
      expect(cellAt(wb, 5, 5)).toMatchObject({ kind: 'number', value: 9 });
    });
  });

  describe('transpose — geometry × value placement', () => {
    it('2×3 source becomes 3×2 destination on transpose', () => {
      seed(store, wb, [
        { row: 0, col: 0, value: 1 },
        { row: 0, col: 1, value: 2 },
        { row: 0, col: 2, value: 3 },
        { row: 1, col: 0, value: 4 },
        { row: 1, col: 1, value: 5 },
        { row: 1, col: 2, value: 6 },
      ]);
      const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 2 });
      if (!snap) throw new Error('no snapshot');
      setActive(store, 5, 5);
      pasteSpecial(store.getState(), store, wb, snap, opt({ transpose: true }));
      wb.recalc();
      // Original [[1,2,3],[4,5,6]] → transposed [[1,4],[2,5],[3,6]]
      expect(cellAt(wb, 5, 5)).toMatchObject({ kind: 'number', value: 1 });
      expect(cellAt(wb, 5, 6)).toMatchObject({ kind: 'number', value: 4 });
      expect(cellAt(wb, 6, 5)).toMatchObject({ kind: 'number', value: 2 });
      expect(cellAt(wb, 6, 6)).toMatchObject({ kind: 'number', value: 5 });
      expect(cellAt(wb, 7, 5)).toMatchObject({ kind: 'number', value: 3 });
      expect(cellAt(wb, 7, 6)).toMatchObject({ kind: 'number', value: 6 });
    });

    it('transpose + values mode preserves the transposed positions', () => {
      seed(store, wb, [
        { row: 0, col: 0, value: 'a' },
        { row: 0, col: 1, value: 'b' },
        { row: 1, col: 0, value: 'c' },
        { row: 1, col: 1, value: 'd' },
      ]);
      const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
      if (!snap) throw new Error('no snapshot');
      setActive(store, 5, 5);
      pasteSpecial(store.getState(), store, wb, snap, opt({ what: 'values', transpose: true }));
      // [[a,b],[c,d]] → transposed [[a,c],[b,d]]
      expect(cellAt(wb, 5, 5)).toMatchObject({ kind: 'text', value: 'a' });
      expect(cellAt(wb, 5, 6)).toMatchObject({ kind: 'text', value: 'c' });
      expect(cellAt(wb, 6, 5)).toMatchObject({ kind: 'text', value: 'b' });
      expect(cellAt(wb, 6, 6)).toMatchObject({ kind: 'text', value: 'd' });
    });
  });

  describe('skipBlanks — leaves destination untouched', () => {
    it('preserves destination when source is blank, even in "all" mode', () => {
      // Source has A1=5, B1 blank. Destination C5=99, D5=100.
      seed(store, wb, [
        { row: 0, col: 0, value: 5 },
        { row: 4, col: 2, value: 99 },
        { row: 4, col: 3, value: 100 },
      ]);
      const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
      if (!snap) throw new Error('no snapshot');
      setActive(store, 4, 2);
      pasteSpecial(store.getState(), store, wb, snap, opt({ skipBlanks: true }));
      wb.recalc();
      // C5 overwritten by source 5; D5 should keep its 100 (skip blank).
      expect(cellAt(wb, 4, 2)).toMatchObject({ kind: 'number', value: 5 });
      expect(cellAt(wb, 4, 3)).toMatchObject({ kind: 'number', value: 100 });
    });

    it('snapshot omits unseeded blank cells, so paste leaves destination intact', () => {
      // captureSnapshot only emits entries for cells that exist in the cells
      // map. A truly unseeded B1 is absent from the snapshot grid, so the
      // paste loop skips that column entirely — D5 keeps its 100 regardless
      // of skipBlanks.
      seed(store, wb, [
        { row: 0, col: 0, value: 5 },
        { row: 4, col: 2, value: 99 },
        { row: 4, col: 3, value: 100 },
      ]);
      const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
      if (!snap) throw new Error('no snapshot');
      setActive(store, 4, 2);
      pasteSpecial(store.getState(), store, wb, snap, opt({ skipBlanks: false }));
      wb.recalc();
      expect(cellAt(wb, 4, 3)).toMatchObject({ kind: 'number', value: 100 });
    });
  });

  describe('combinations — what × transpose, what × skipBlanks', () => {
    it('formulas + transpose: every source has a formula → both destinations written', () => {
      // Both cells carry formulas so `what:'formulas'` writes both destinations.
      seed(store, wb, [
        { row: 0, col: 0, value: 10, formula: '=10' },
        { row: 1, col: 0, value: 20, formula: '=A1*2' },
      ]);
      const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 });
      if (!snap) throw new Error('no snapshot');
      setActive(store, 5, 5);
      pasteSpecial(store.getState(), store, wb, snap, opt({ what: 'formulas', transpose: true }));
      // Transposed 2×1 → 1×2. First cell (5,5) gets `=10`; (5,6) gets the
      // shifted version of `=A1*2`.
      wb.recalc();
      expect(cellAt(wb, 5, 5)).toMatchObject({ kind: 'number', value: 10 });
      const v = cellAt(wb, 5, 6);
      expect(v.kind).toBe('number');
    });

    it('formats + skipBlanks: no format write when source is blank', () => {
      seed(store, wb, [
        { row: 0, col: 0, value: 1, format: { bold: true } },
        // (0,1) is blank with no format.
      ]);
      const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
      if (!snap) throw new Error('no snapshot');
      setActive(store, 5, 5);
      pasteSpecial(store.getState(), store, wb, snap, opt({ what: 'formats', skipBlanks: true }));
      expect(fmtAt(store, 5, 5)?.bold).toBe(true);
      expect(fmtAt(store, 5, 6)).toBeUndefined();
    });
  });

  describe('active range update after paste', () => {
    it('writtenRange equals the destination block for non-transpose', () => {
      seed(store, wb, [
        { row: 0, col: 0, value: 1 },
        { row: 0, col: 1, value: 2 },
        { row: 1, col: 0, value: 3 },
        { row: 1, col: 1, value: 4 },
      ]);
      const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
      if (!snap) throw new Error('no snapshot');
      setActive(store, 10, 10);
      const res = pasteSpecial(store.getState(), store, wb, snap, opt());
      expect(res?.writtenRange).toEqual({ sheet: 0, r0: 10, c0: 10, r1: 11, c1: 11 });
    });

    it('writtenRange reflects the transposed dimensions', () => {
      seed(store, wb, [
        { row: 0, col: 0, value: 1 },
        { row: 0, col: 1, value: 2 },
        { row: 0, col: 2, value: 3 },
      ]);
      const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
      if (!snap) throw new Error('no snapshot');
      setActive(store, 10, 10);
      const res = pasteSpecial(store.getState(), store, wb, snap, opt({ transpose: true }));
      // 1×3 → 3×1
      expect(res?.writtenRange).toEqual({ sheet: 0, r0: 10, c0: 10, r1: 12, c1: 10 });
    });
  });
});
