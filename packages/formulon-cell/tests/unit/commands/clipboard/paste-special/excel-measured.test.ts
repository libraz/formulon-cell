import { beforeEach, describe, expect, it } from 'vitest';
import { insertCopiedCellsFromTSV } from '../../../../../src/commands/clipboard/insert-copied-cells.js';
import { pasteLink, pasteSpecial } from '../../../../../src/commands/clipboard/paste-special.js';
import { captureSnapshot } from '../../../../../src/commands/clipboard/snapshot.js';
import { History } from '../../../../../src/commands/history.js';
import { addrKey, type WorkbookHandle } from '../../../../../src/engine/workbook-handle.js';
import {
  type CellFormat,
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../../src/store/store.js';
import { assertSnap, newWb, num, seedAndMirror, setActive } from './fixtures.js';

// Expectations are the Excel results measured through the object model on
// harness/make_clip.py (results.json / results2.json).

const fmt = (store: SpreadsheetStore, row: number, col: number): CellFormat | undefined =>
  store.getState().format.formats.get(addrKey({ sheet: 0, row, col }));

const setFmt = (store: SpreadsheetStore, row: number, col: number, f: Partial<CellFormat>) =>
  mutators.setCellFormat(store, { sheet: 0, row, col }, f);

const opts = (over: Partial<Parameters<typeof pasteSpecial>[4]> = {}) => ({
  what: 'all' as const,
  operation: 'none' as const,
  skipBlanks: false,
  transpose: false,
  ...over,
});

describe('paste special as Excel measures it', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  it('skip blanks skips a blank source cell that only carries a format', () => {
    // A1:D1 = 10, 20, 30, (blank, gray fill) pasted over 1,2,3,4: the last
    // destination keeps 4 and takes no fill.
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 10 },
      { row: 0, col: 1, value: 20 },
      { row: 0, col: 2, value: 30 },
      { row: 10, col: 8, value: 1 },
      { row: 10, col: 9, value: 2 },
      { row: 10, col: 10, value: 3 },
      { row: 10, col: 11, value: 4 },
    ]);
    setFmt(store, 0, 3, { fill: '#808080' });
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 3 });
    assertSnap(snap);
    setActive(store, 10, 8);
    pasteSpecial(store.getState(), store, wb, snap, opts({ skipBlanks: true }));
    wb.recalc();
    expect(num(wb, 10, 8)).toBe(10);
    expect(num(wb, 10, 11)).toBe(4);
    expect(fmt(store, 10, 11)?.fill).toBeUndefined();
  });

  it('transpose re-points references inside the copied block, absolute ones too', () => {
    // A1:D4 transposed to N100.
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 10 },
      { row: 0, col: 1, value: 20, formula: '=A1*2' },
      { row: 1, col: 0, value: 5 },
      { row: 1, col: 1, value: 25, formula: '=A2+B1' },
      { row: 1, col: 2, value: 15, formula: '=SUM(A1:A2)' },
      { row: 1, col: 3, value: 15, formula: '=$A$1+A2' },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 3 });
    assertSnap(snap);
    setActive(store, 99, 13);
    pasteSpecial(store.getState(), store, wb, snap, opts({ transpose: true }));
    expect(wb.cellFormula({ sheet: 0, row: 100, col: 13 })).toBe('=N100*2');
    expect(wb.cellFormula({ sheet: 0, row: 100, col: 14 })).toBe('=O100+N101');
    expect(wb.cellFormula({ sheet: 0, row: 101, col: 14 })).toBe('=SUM(N100:O100)');
    expect(wb.cellFormula({ sheet: 0, row: 102, col: 14 })).toBe('=$N$100+O100');
  });

  it('transpose still shifts references outside the copied block', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 1, value: 0, formula: '=Z1' }]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    assertSnap(snap);
    setActive(store, 9, 0);
    pasteSpecial(store.getState(), store, wb, snap, opts({ what: 'formulas', transpose: true }));
    // B1 lands on A11: offset (+10, -1) from its source.
    expect(wb.cellFormula({ sheet: 0, row: 10, col: 0 })).toBe('=Y11');
  });

  it('pastes a single cell into a merged anchor and keeps the merge', () => {
    // 'left' pasted onto F1 of the merged F1:G2.
    seedAndMirror(store, wb, [{ row: 2, col: 0, value: 'left' }]);
    mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 5, r1: 1, c1: 6 });
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 2, c0: 0, r1: 2, c1: 0 });
    assertSnap(snap);
    setActive(store, 0, 5);
    const result = pasteSpecial(store.getState(), store, wb, snap, opts());
    expect(result).not.toBeNull();
    expect(wb.getValue({ sheet: 0, row: 0, col: 5 })).toEqual({ kind: 'text', value: 'left' });
    expect([...store.getState().merges.byAnchor.values()]).toContainEqual({
      sheet: 0,
      r0: 0,
      c0: 5,
      r1: 1,
      c1: 6,
    });
  });

  it('column widths pastes only the source widths', () => {
    // Source columns 13.17 / 8.17 characters wide (stored px here).
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 1 }]);
    mutators.setColWidth(store, 0, 97);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    assertSnap(snap);
    setActive(store, 61, 13);
    pasteSpecial(store.getState(), store, wb, snap, opts({ what: 'column-widths' }));
    const layout = store.getState().layout;
    expect(layout.colWidths.get(13)).toBe(97);
    expect(layout.colWidths.get(14) ?? layout.defaultColWidth).toBe(layout.defaultColWidth);
    expect(wb.getValue({ sheet: 0, row: 61, col: 13 })).toEqual({ kind: 'blank' });
  });

  it('all except borders carries values and formats but not borders', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 7 },
      { row: 55, col: 13, value: 1 },
    ]);
    setFmt(store, 0, 0, {
      bold: true,
      fill: '#ffff00',
      borders: { bottom: { style: 'thin' } },
    } as Partial<CellFormat>);
    setFmt(store, 55, 13, { borders: { top: { style: 'thick' } } } as Partial<CellFormat>);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    assertSnap(snap);
    setActive(store, 55, 13);
    pasteSpecial(store.getState(), store, wb, snap, opts({ what: 'all-except-borders' }));
    expect(num(wb, 55, 13)).toBe(7);
    expect(fmt(store, 55, 13)?.bold).toBe(true);
    expect(fmt(store, 55, 13)?.fill).toBe('#ffff00');
    expect(fmt(store, 55, 13)?.borders).toEqual({ top: { style: 'thick' } });
  });

  it('comments pastes only the note', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 10 }]);
    setFmt(store, 0, 0, { bold: true, comment: 'note A1', commentAuthor: 'me' });
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    assertSnap(snap);
    setActive(store, 43, 13);
    pasteSpecial(store.getState(), store, wb, snap, opts({ what: 'comments' }));
    expect(wb.getValue({ sheet: 0, row: 43, col: 13 })).toEqual({ kind: 'blank' });
    expect(fmt(store, 43, 13)?.comment).toBe('note A1');
    expect(fmt(store, 43, 13)?.bold).toBeUndefined();
  });

  it('validation pastes only the validation rule', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 10 }]);
    const validation = { kind: 'list', source: ['a', 'b'] } as unknown as CellFormat['validation'];
    setFmt(store, 0, 0, { bold: true, validation });
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    assertSnap(snap);
    setActive(store, 49, 13);
    pasteSpecial(store.getState(), store, wb, snap, opts({ what: 'validation' }));
    expect(fmt(store, 49, 13)?.validation).toEqual(validation);
    expect(fmt(store, 49, 13)?.bold).toBeUndefined();
    expect(wb.getValue({ sheet: 0, row: 49, col: 13 })).toEqual({ kind: 'blank' });
  });

  it('all using the source theme pastes like all within a workbook', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 3 }]);
    setFmt(store, 0, 0, { italic: true });
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    assertSnap(snap);
    setActive(store, 5, 5);
    pasteSpecial(store.getState(), store, wb, snap, opts({ what: 'all-using-source-theme' }));
    expect(num(wb, 5, 5)).toBe(3);
    expect(fmt(store, 5, 5)?.italic).toBe(true);
  });

  it('all merging conditional formats adds the source rules beside the destination ones', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 9 },
      { row: 1, col: 0, value: 1 },
    ]);
    mutators.addConditionalRule(store, {
      kind: 'cell-value',
      range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 },
      op: '>',
      a: 5,
      apply: { fill: '#ffc7ce' },
    });
    mutators.addConditionalRule(store, {
      kind: 'formula',
      range: { sheet: 0, r0: 1, c0: 0, r1: 3, c1: 0 },
      formula: '=A2>0',
      apply: { bold: true },
    });
    mutators.addConditionalRule(store, {
      kind: 'blanks',
      range: { sheet: 0, r0: 0, c0: 5, r1: 1, c1: 5 },
      apply: { fill: '#eeeeee' },
    });
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 });
    assertSnap(snap);
    setActive(store, 0, 5);
    pasteSpecial(
      store.getState(),
      store,
      wb,
      snap,
      opts({ what: 'all-merging-conditional-formats' }),
    );
    expect(num(wb, 0, 5)).toBe(9);
    const rules = store.getState().conditional.rules;
    expect(rules).toHaveLength(5);
    // The destination's own rule stays; the copied part of each source rule
    // lands on the pasted cells, a formula re-anchored with it.
    expect(rules).toContainEqual(
      expect.objectContaining({ kind: 'blanks', range: { sheet: 0, r0: 0, c0: 5, r1: 1, c1: 5 } }),
    );
    expect(rules).toContainEqual(
      expect.objectContaining({
        kind: 'cell-value',
        range: { sheet: 0, r0: 0, c0: 5, r1: 1, c1: 5 },
        a: 5,
      }),
    );
    expect(rules).toContainEqual(
      expect.objectContaining({
        kind: 'formula',
        range: { sheet: 0, r0: 1, c0: 5, r1: 1, c1: 5 },
        formula: '=F2>0',
      }),
    );
  });

  it('paste link writes relative references to the source cells', () => {
    // A1:B2 linked at BP150: =A1, =B1, =A2, =B2.
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 10 },
      { row: 0, col: 1, value: 20 },
      { row: 1, col: 0, value: 5 },
      { row: 1, col: 1, value: 25 },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    assertSnap(snap);
    setActive(store, 149, 67);
    const result = pasteLink(store.getState(), store, wb, snap);
    expect(result?.writtenRange).toEqual({ sheet: 0, r0: 149, c0: 67, r1: 150, c1: 68 });
    expect(wb.cellFormula({ sheet: 0, row: 149, col: 67 })).toBe('=A1');
    expect(wb.cellFormula({ sheet: 0, row: 149, col: 68 })).toBe('=B1');
    expect(wb.cellFormula({ sheet: 0, row: 150, col: 67 })).toBe('=A2');
    expect(wb.cellFormula({ sheet: 0, row: 150, col: 68 })).toBe('=B2');
  });
});

describe('insert cut cells as Excel measures it', () => {
  it('moves the cut cells: AD1:AD2 cut and inserted at AD4 leaves 3,1,2,4', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    seedAndMirror(store, wb, [
      { row: 0, col: 29, value: 1 },
      { row: 1, col: 29, value: 2 },
      { row: 2, col: 29, value: 3 },
      { row: 3, col: 29, value: 4 },
      { row: 0, col: 30, value: 10, formula: '=SUM(AD1:AD4)' },
    ]);
    const snap = captureSnapshot(
      store.getState(),
      { sheet: 0, r0: 0, c0: 29, r1: 1, c1: 29 },
      'cut',
    );
    assertSnap(snap);
    mutators.setCopyRange(store, { sheet: 0, r0: 0, c0: 29, r1: 1, c1: 29 });
    setActive(store, 3, 29);
    const result = insertCopiedCellsFromTSV(store, wb, null, '1\r\n2', 'down', snap);
    wb.recalc();
    expect([0, 1, 2, 3].map((row) => num(wb, row, 29))).toEqual([3, 1, 2, 4]);
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 30 })).toBe('=SUM(AD1:AD4)');
    expect(result?.writtenRange).toEqual({ sheet: 0, r0: 1, c0: 29, r1: 2, c1: 29 });
    expect(store.getState().ui.copyRange ?? null).toBeNull();
  });

  it.each([false, true])(
    'leaves the workbook intact when the insert fails after the close-up (history: %s)',
    async (withHistory) => {
      const store = createSpreadsheetStore();
      const wb = await newWb();
      seedAndMirror(store, wb, [
        { row: 0, col: 0, value: 1 },
        { row: 1, col: 0, value: 2 },
        { row: 2, col: 0, value: 3 },
        { row: 0, col: 1, value: 3, formula: '=A3' },
      ]);
      // C5:D5 straddles the band a downward insert at C3 would shift, so the
      // insert is refused only after the source column has closed up.
      mutators.mergeRange(store, { sheet: 0, r0: 4, c0: 2, r1: 4, c1: 3 });
      const cut = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 };
      const snap = captureSnapshot(store.getState(), cut, 'cut');
      assertSnap(snap);
      mutators.setCopyRange(store, cut, 'cut');
      setActive(store, 2, 2);
      const history = withHistory ? new History() : null;

      const result = insertCopiedCellsFromTSV(store, wb, history, '1\r\n2', 'down', snap);

      expect(result).toBeNull();
      wb.recalc();
      expect([0, 1, 2].map((row) => num(wb, row, 0))).toEqual([1, 2, 3]);
      expect(wb.cellFormula({ sheet: 0, row: 0, col: 1 })).toBe('=A3');
      expect(store.getState().ui.copyRange).toEqual(cut);
      expect(history?.canUndo() ?? false).toBe(false);
    },
  );
});
