import { beforeEach, describe, expect, it } from 'vitest';
import { pasteSpecial } from '../../../../../src/commands/clipboard/paste-special.js';
import { captureSnapshot } from '../../../../../src/commands/clipboard/snapshot.js';
import { addrKey, type WorkbookHandle } from '../../../../../src/engine/workbook-handle.js';
import {
  type CellFormat,
  type ConditionalRule,
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../../src/store/store.js';
import { assertSnap, newWb, seedAndMirror, setActive } from './fixtures.js';

// Excel results on harness/make_cf.py (cf.xlsx): source A1:A3 with a CF rule,
// list validation 'a,b' and a note on A1; destination C1:C5 with its own rule,
// whole-number validation on C1:C3 and notes on C1/C2.

const range = (r0: number, c0: number, r1: number, c1: number) => ({ sheet: 0, r0, c0, r1, c1 });
const srcRule: ConditionalRule = {
  kind: 'cell-value',
  range: range(0, 0, 2, 0),
  op: '>',
  a: 2,
  apply: { fill: '#ffc7ce' },
};
const dstRule: ConditionalRule = {
  kind: 'cell-value',
  range: range(0, 2, 4, 2),
  op: '<',
  a: 5,
  apply: { fill: '#c6efce' },
};
const srcList = { kind: 'list', source: ['a', 'b'] } as unknown as CellFormat['validation'];
const dstWhole = {
  kind: 'whole',
  op: 'between',
  a: 1,
  b: 9,
} as unknown as CellFormat['validation'];

const opts = (what: Parameters<typeof pasteSpecial>[4]['what'], skipBlanks = false) => ({
  what,
  operation: 'none' as const,
  skipBlanks,
  transpose: false,
});

describe('paste special edges as Excel measures them', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  const fmt = (row: number, col: number): CellFormat | undefined =>
    store.getState().format.formats.get(addrKey({ sheet: 0, row, col }));
  const rulesAt = (row: number, col: number) =>
    store
      .getState()
      .conditional.rules.filter(
        (r) => r.range.r0 <= row && row <= r.range.r1 && r.range.c0 <= col && col <= r.range.c1,
      )
      .map((r) => (r.kind === 'cell-value' ? r.a : r.kind));

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 1, col: 0, value: 2 },
      { row: 2, col: 0, value: 3 },
      { row: 0, col: 2, value: 1 },
      { row: 1, col: 2, value: 2 },
      { row: 2, col: 2, value: 3 },
      { row: 3, col: 2, value: 4 },
      // E1:E3: a plain source with no rule, validation or note.
      { row: 0, col: 4, value: 1 },
      { row: 1, col: 4, value: 2 },
      { row: 2, col: 4, value: 3 },
    ]);
    for (const row of [0, 1, 2])
      mutators.setCellFormat(store, { sheet: 0, row, col: 0 }, { validation: srcList });
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { comment: 'src note', commentAuthor: 'me' },
    );
    for (const row of [0, 1, 2])
      mutators.setCellFormat(store, { sheet: 0, row, col: 2 }, { validation: dstWhole });
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 2 },
      { comment: 'dst note', commentAuthor: 'me' },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 1, col: 2 },
      { comment: 'dst note2', commentAuthor: 'me' },
    );
    mutators.addConditionalRule(store, srcRule);
    mutators.addConditionalRule(store, dstRule);
  });

  const pasteFrom = (c0: number, what: Parameters<typeof opts>[0], skipBlanks = false) => {
    const snap = captureSnapshot(store.getState(), range(0, c0, 2, c0));
    assertSnap(snap);
    setActive(store, 0, 2);
    pasteSpecial(store.getState(), store, wb, snap, opts(what, skipBlanks));
  };

  it.each(['all', 'all-except-borders', 'formats'] as const)(
    '%s replaces the destination rules on the pasted cells only',
    (what) => {
      pasteFrom(0, what);
      for (const row of [0, 1, 2]) expect(rulesAt(row, 2)).toEqual([2]);
      expect(rulesAt(3, 2)).toEqual([5]);
    },
  );

  it.each(['all', 'formats'] as const)(
    '%s from a rule-less source removes the destination rules',
    (what) => {
      pasteFrom(4, what);
      for (const row of [0, 1, 2]) expect(rulesAt(row, 2)).toEqual([]);
      expect(rulesAt(3, 2)).toEqual([5]);
    },
  );

  it.each(['values', 'formulas', 'comments', 'validation'] as const)(
    '%s leaves the rules alone',
    (what) => {
      pasteFrom(0, what);
      for (const row of [0, 1, 2, 3]) expect(rulesAt(row, 2)).toEqual([5]);
    },
  );

  it.each([
    ['all', srcList],
    ['all-except-borders', srcList],
    ['validation', srcList],
    ['formats', dstWhole],
    ['values', dstWhole],
    ['formulas', dstWhole],
    ['comments', dstWhole],
  ] as const)('%s leaves the expected validation', (what, expected) => {
    pasteFrom(0, what);
    expect(fmt(0, 2)?.validation).toEqual(expected);
  });

  it('all from a plain source removes validation; formats from it keeps it', () => {
    pasteFrom(4, 'all');
    expect(fmt(0, 2)?.validation).toBeUndefined();
  });

  it('formats from a plain source keeps the validation and the notes', () => {
    pasteFrom(4, 'formats');
    expect(fmt(0, 2)?.validation).toEqual(dstWhole);
    expect(fmt(0, 2)?.comment).toBe('dst note');
    expect(fmt(1, 2)?.comment).toBe('dst note2');
  });

  it.each(['all', 'all-except-borders', 'comments'] as const)(
    '%s puts the source note and deletes destination notes the source lacks',
    (what) => {
      pasteFrom(0, what);
      expect(fmt(0, 2)?.comment).toBe('src note');
      expect(fmt(1, 2)?.comment).toBeUndefined();
    },
  );

  it('comments from a noteless source deletes the destination notes', () => {
    pasteFrom(4, 'comments');
    expect(fmt(0, 2)?.comment).toBeUndefined();
    expect(fmt(1, 2)?.comment).toBeUndefined();
  });

  it('merging conditional formats lists the source rule first, then the destination rule', () => {
    // C1 reads [>2, <5]; C5, outside the paste, keeps only <5.
    pasteFrom(0, 'all-merging-conditional-formats');
    for (const row of [0, 1, 2]) expect(rulesAt(row, 2)).toEqual([2, 5]);
    expect(rulesAt(3, 2)).toEqual([5]);
  });

  it('merging from a rule-less source keeps the destination rules', () => {
    pasteFrom(4, 'all-merging-conditional-formats');
    for (const row of [0, 1, 2]) expect(rulesAt(row, 2)).toEqual([5]);
  });

  it('all with skip blanks from a noteless source keeps the destination note', () => {
    pasteFrom(4, 'all', true);
    expect(fmt(0, 2)?.comment).toBe('dst note');
  });

  // pr-cvh.xlsx: a formats paste carries CF but not validation, notes or links.
  it('formats leaves the destination hyperlink and takes none from the source', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        hyperlink: 'https://example.com/x',
        hyperlinkTooltip: 'tip',
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 1, col: 2 },
      { hyperlink: 'https://dest.example' },
    );
    pasteFrom(0, 'formats');
    expect(fmt(0, 2)?.hyperlink).toBeUndefined();
    expect(fmt(0, 2)?.hyperlinkTooltip).toBeUndefined();
    expect(fmt(1, 2)?.hyperlink).toBe('https://dest.example');
  });

  it('a cut moves the conditional formats with the cells', () => {
    const snap = captureSnapshot(store.getState(), range(0, 0, 2, 0), 'cut');
    assertSnap(snap);
    setActive(store, 0, 6);
    pasteSpecial(store.getState(), store, wb, snap, opts('all'));
    for (const row of [0, 1, 2]) {
      expect(rulesAt(row, 6)).toEqual([2]);
      expect(rulesAt(row, 0)).toEqual([]);
    }
  });
});
