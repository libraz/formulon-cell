import { beforeEach, describe, expect, it } from 'vitest';
import {
  boundingRange,
  findMatchingRanges,
  type GoToSpecialKind,
  type GoToSpecialOptions,
  selectionFromMatches,
  selectionFromRanges,
} from '../../../src/commands/goto-special.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

/** Replicate the engine cells into the store data slice — validation entries
 *  live on the format slice and conditional rules on the conditional slice. */
const sync = (store: SpreadsheetStore, wb: WorkbookHandle): void => {
  store.setState((s) => {
    const cells = new Map<
      string,
      { value: ReturnType<WorkbookHandle['getValue']>; formula: string | null }
    >();
    for (const e of wb.cells(0)) {
      cells.set(addrKey(e.addr), { value: e.value, formula: e.formula });
    }
    return { ...s, data: { ...s.data, cells } };
  });
};

/** Select `r0:c0..r1:c1`; the active cell defaults to its top-left corner. */
const select = (
  store: SpreadsheetStore,
  r0: number,
  c0: number,
  r1 = r0,
  c1 = c0,
  active = { row: r0, col: c0 },
): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      active: { sheet: 0, ...active },
      anchor: { sheet: 0, ...active },
      range: { sheet: 0, r0, c0, r1, c1 },
      extraRanges: [],
    },
  }));
};

const text = (rs: readonly { r0: number; c0: number; r1: number; c1: number }[]): string[] =>
  rs.map((r) => `${r.r0}:${r.c0}-${r.r1}:${r.c1}`);

describe('findMatchingRanges', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  const find = (kind: GoToSpecialKind, options?: GoToSpecialOptions): string[] | string => {
    const result = findMatchingRanges(wb, store, kind, options);
    return result.status === 'found' ? text(result.ranges) : result.status;
  };

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  it('formulas — returns only cells with a formula', () => {
    wb.setFormula({ sheet: 0, row: 0, col: 0 }, '=1+2');
    wb.setText({ sheet: 0, row: 0, col: 1 }, 'hello');
    wb.setNumber({ sheet: 0, row: 1, col: 0 }, 42);
    wb.recalc();
    sync(store, wb);
    expect(find('formulas')).toEqual(['0:0-0:0']);
  });

  it('constants — populated cells without a formula, merged into rectangles', () => {
    wb.setFormula({ sheet: 0, row: 0, col: 0 }, '=1+2');
    wb.setText({ sheet: 0, row: 0, col: 1 }, 'hello');
    wb.setNumber({ sheet: 0, row: 1, col: 0 }, 42);
    wb.setNumber({ sheet: 0, row: 1, col: 1 }, 43);
    wb.recalc();
    sync(store, wb);
    expect(find('constants')).toEqual(['0:1-0:1', '1:0-1:1']);
  });

  it('constants — can be filtered by value kind', () => {
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 42);
    wb.setText({ sheet: 0, row: 0, col: 1 }, 'hello');
    wb.setBool({ sheet: 0, row: 0, col: 2 }, true);
    wb.setText({ sheet: 0, row: 0, col: 3 }, '#REF!');
    wb.setFormula({ sheet: 0, row: 1, col: 0 }, '=10');
    wb.recalc();
    sync(store, wb);
    expect(
      find('constants', {
        filters: { numbers: false, text: true, logical: true, errors: false },
      }),
    ).toEqual(['0:1-0:2']);
  });

  it('formulas — can be filtered by evaluated value kind', () => {
    wb.setFormula({ sheet: 0, row: 0, col: 0 }, '=1+2');
    wb.setFormula({ sheet: 0, row: 0, col: 1 }, '="apple"');
    wb.setFormula({ sheet: 0, row: 0, col: 2 }, '=1/0');
    wb.setNumber({ sheet: 0, row: 1, col: 0 }, 42);
    wb.recalc();
    sync(store, wb);
    expect(
      find('formulas', { filters: { numbers: false, text: true, logical: false, errors: true } }),
    ).toEqual(['0:1-0:2']);
  });

  it('numbers — only number-kind cells', () => {
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    wb.setNumber({ sheet: 0, row: 0, col: 1 }, 2);
    wb.setText({ sheet: 0, row: 0, col: 2 }, 'three');
    wb.setFormula({ sheet: 0, row: 1, col: 0 }, '=10');
    wb.recalc();
    sync(store, wb);
    expect(find('numbers')).toEqual(['0:0-0:1', '1:0-1:0']);
  });

  it('text — only text-kind cells, excluding error sentinels', () => {
    wb.setText({ sheet: 0, row: 0, col: 0 }, 'apple');
    wb.setText({ sheet: 0, row: 0, col: 1 }, '#REF!');
    wb.setNumber({ sheet: 0, row: 1, col: 0 }, 1);
    wb.recalc();
    sync(store, wb);
    expect(find('text')).toEqual(['0:0-0:0']);
  });

  it('errors — formula errors and explicit error sentinels in text', () => {
    wb.setFormula({ sheet: 0, row: 0, col: 0 }, '=1/0');
    wb.setText({ sheet: 0, row: 0, col: 1 }, '#NAME?');
    wb.setText({ sheet: 0, row: 1, col: 0 }, 'fine');
    wb.recalc();
    sync(store, wb);
    expect(find('errors')).toEqual(['0:0-0:1']);
  });

  it('reports none when nothing matches', () => {
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    sync(store, wb);
    expect(find('formulas')).toBe('none');
  });

  it('a single selected cell searches the used range', () => {
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    wb.setNumber({ sheet: 0, row: 5, col: 5 }, 2);
    wb.recalc();
    sync(store, wb);
    select(store, 9, 9);
    expect(find('numbers')).toEqual(['0:0-0:0', '5:5-5:5']);
  });

  it('a multi-cell selection searches only inside itself', () => {
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    wb.setNumber({ sheet: 0, row: 5, col: 5 }, 2);
    wb.recalc();
    sync(store, wb);
    select(store, 0, 0, 2, 2);
    expect(find('numbers')).toEqual(['0:0-0:0']);
  });

  it('merges row runs with identical spans into one rectangle', () => {
    for (let r = 0; r < 3; r += 1) {
      for (let c = 0; c < 2; c += 1) wb.setNumber({ sheet: 0, row: r, col: c }, r * 2 + c);
    }
    wb.setNumber({ sheet: 0, row: 3, col: 0 }, 9);
    wb.recalc();
    sync(store, wb);
    expect(find('numbers')).toEqual(['0:0-2:1', '3:0-3:0']);
  });

  it('blanks — empty cells inside a selection, as rectangles', () => {
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    wb.setNumber({ sheet: 0, row: 2, col: 2 }, 9);
    wb.recalc();
    sync(store, wb);
    select(store, 0, 0, 2, 2);
    expect(find('blanks')).toEqual(['0:1-0:2', '1:0-1:2', '2:0-2:1']);
  });

  it('blanks — a single cell searches the used range', () => {
    wb.setNumber({ sheet: 0, row: 1, col: 1 }, 1);
    wb.setNumber({ sheet: 0, row: 3, col: 3 }, 2);
    wb.recalc();
    sync(store, wb);
    select(store, 8, 8);
    const result = findMatchingRanges(wb, store, 'blanks');
    expect(result.status).toBe('found');
    if (result.status !== 'found') return;
    const area = result.ranges.reduce((n, r) => n + (r.r1 - r.r0 + 1) * (r.c1 - r.c0 + 1), 0);
    expect(area).toBe(9 - 2);
    expect(result.ranges.every((r) => r.r0 >= 1 && r.c0 >= 1 && r.r1 <= 3 && r.c1 <= 3)).toBe(true);
  });

  it('blanks — reports too-large instead of dropping a huge selection', () => {
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    wb.recalc();
    sync(store, wb);
    select(store, 0, 0, 1048575, 0);
    expect(find('blanks')).toBe('too-large');
  });

  it('data-validation — includes blank cells carrying a rule', () => {
    wb.setText({ sheet: 0, row: 0, col: 0 }, 'a');
    sync(store, wb);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 5, col: 1 },
      { validation: { kind: 'list', source: ['a', 'b', 'c'] } },
    );
    expect(find('data-validation')).toEqual(['5:1-5:1']);
  });

  it('data-validation — same keeps cells sharing the active cell rule', () => {
    const list = { kind: 'list' as const, source: ['a', 'b'] };
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { validation: list });
    mutators.setCellFormat(store, { sheet: 0, row: 1, col: 0 }, { validation: { ...list } });
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 2, col: 0 },
      { validation: { kind: 'list', source: ['x'] } },
    );
    select(store, 0, 0);
    expect(find('data-validation')).toEqual(['0:0-2:0']);
    expect(find('data-validation', { same: true })).toEqual(['0:0-1:0']);
  });

  it('conditional-format — covers blank cells inside a rule range', () => {
    wb.setNumber({ sheet: 0, row: 5, col: 5 }, 3);
    wb.recalc();
    sync(store, wb);
    mutators.addConditionalRule(store, {
      kind: 'data-bar',
      range: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
      color: '#638ec6',
    });
    expect(find('conditional-format')).toEqual(['0:0-1:1']);
  });

  it('conditional-format — same keeps cells covered by the same rules', () => {
    mutators.addConditionalRule(store, {
      kind: 'data-bar',
      range: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 },
      color: '#638ec6',
    });
    mutators.addConditionalRule(store, {
      kind: 'data-bar',
      range: { sheet: 0, r0: 2, c0: 0, r1: 3, c1: 0 },
      color: '#ff0000',
    });
    select(store, 0, 0);
    expect(find('conditional-format', { same: true })).toEqual(['0:0-1:0']);
    expect(find('conditional-format')).toEqual(['0:0-3:0']);
  });

  it('comments — cells with a comment', () => {
    mutators.setCellFormat(store, { sheet: 0, row: 2, col: 2 }, { comment: 'note' });
    select(store, 0, 0);
    expect(find('comments')).toEqual(['2:2-2:2']);
  });

  it('current-region — the contiguous block around the active cell', () => {
    wb.setNumber({ sheet: 0, row: 1, col: 1 }, 1);
    wb.setNumber({ sheet: 0, row: 1, col: 2 }, 1);
    wb.setNumber({ sheet: 0, row: 2, col: 1 }, 1);
    wb.setNumber({ sheet: 0, row: 8, col: 8 }, 1);
    wb.recalc();
    sync(store, wb);
    select(store, 1, 1);
    expect(find('current-region')).toEqual(['1:1-2:2']);
    select(store, 5, 5);
    expect(find('current-region')).toBe('none');
  });

  it('last-cell — max used row by max used column', () => {
    wb.setNumber({ sheet: 0, row: 4, col: 1 }, 1);
    wb.setNumber({ sheet: 0, row: 1, col: 6 }, 1);
    wb.recalc();
    sync(store, wb);
    select(store, 0, 0);
    expect(find('last-cell')).toEqual(['4:6-4:6']);
  });

  it('visible-cells — the selection minus hidden rows and columns', () => {
    select(store, 0, 0, 3, 2);
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenRows: new Set([1]), hiddenCols: new Set([1]) },
    }));
    expect(find('visible-cells')).toEqual(['0:0-0:0', '0:2-0:2', '2:0-3:0', '2:2-3:2']);
  });

  it('row-differences — cells that differ from the active row', () => {
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    wb.setNumber({ sheet: 0, row: 0, col: 1 }, 2);
    wb.setNumber({ sheet: 0, row: 1, col: 0 }, 1);
    wb.setNumber({ sheet: 0, row: 1, col: 1 }, 5);
    wb.setNumber({ sheet: 0, row: 2, col: 0 }, 7);
    wb.recalc();
    sync(store, wb);
    select(store, 0, 0, 2, 1);
    expect(find('row-differences')).toEqual(['1:1-2:1', '2:0-2:0']);
  });

  it('column-differences — cells that differ from the active column', () => {
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    wb.setNumber({ sheet: 0, row: 1, col: 0 }, 2);
    wb.setNumber({ sheet: 0, row: 0, col: 1 }, 1);
    wb.setNumber({ sheet: 0, row: 1, col: 1 }, 9);
    wb.recalc();
    sync(store, wb);
    select(store, 0, 0, 1, 2);
    expect(find('column-differences')).toEqual(['0:2-0:2', '1:1-1:2']);
  });

  it('precedents — direct only versus all levels', () => {
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    wb.setFormula({ sheet: 0, row: 0, col: 1 }, '=A1');
    wb.setFormula({ sheet: 0, row: 0, col: 2 }, '=B1');
    wb.recalc();
    sync(store, wb);
    select(store, 0, 2);
    expect(find('precedents')).toEqual(['0:1-0:1']);
    expect(find('precedents', { allLevels: true })).toEqual(['0:0-0:1']);
  });

  it('dependents — direct only versus all levels', () => {
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    wb.setFormula({ sheet: 0, row: 0, col: 1 }, '=A1');
    wb.setFormula({ sheet: 0, row: 0, col: 2 }, '=B1');
    wb.recalc();
    sync(store, wb);
    select(store, 0, 0);
    expect(find('dependents')).toEqual(['0:1-0:1']);
    expect(find('dependents', { allLevels: true })).toEqual(['0:1-0:2']);
  });
});

describe('boundingRange', () => {
  it('returns the inclusive bounding rectangle of a match list', () => {
    const got = boundingRange([
      { sheet: 0, row: 1, col: 2 },
      { sheet: 0, row: 4, col: 0 },
      { sheet: 0, row: 2, col: 5 },
    ]);
    expect(got).toEqual({ sheet: 0, r0: 1, c0: 0, r1: 4, c1: 5 });
  });

  it('throws on empty input', () => {
    expect(() => boundingRange([])).toThrow();
  });
});

describe('selectionFromMatches', () => {
  it('keeps non-contiguous matches as primary + extra ranges, merging neighbours', () => {
    const selection = selectionFromMatches([
      { sheet: 0, row: 1, col: 2 },
      { sheet: 0, row: 4, col: 0 },
      { sheet: 0, row: 1, col: 3 },
      { sheet: 0, row: 2, col: 5 },
    ]);
    expect(selection).toEqual({
      active: { sheet: 0, row: 1, col: 2 },
      anchor: { sheet: 0, row: 1, col: 2 },
      range: { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 3 },
      extraRanges: [
        { sheet: 0, r0: 2, c0: 5, r1: 2, c1: 5 },
        { sheet: 0, r0: 4, c0: 0, r1: 4, c1: 0 },
      ],
    });
  });

  it('throws on empty input', () => {
    expect(() => selectionFromMatches([])).toThrow();
  });
});

describe('selectionFromRanges', () => {
  it('activates the first rectangle top-left cell', () => {
    const selection = selectionFromRanges([
      { sheet: 0, r0: 2, c0: 1, r1: 4, c1: 3 },
      { sheet: 0, r0: 7, c0: 0, r1: 7, c1: 0 },
    ]);
    expect(selection.active).toEqual({ sheet: 0, row: 2, col: 1 });
    expect(selection.anchor).toEqual({ sheet: 0, row: 2, col: 1 });
    expect(selection.range).toEqual({ sheet: 0, r0: 2, c0: 1, r1: 4, c1: 3 });
    expect(selection.extraRanges).toEqual([{ sheet: 0, r0: 7, c0: 0, r1: 7, c1: 0 }]);
  });

  it('throws on empty input', () => {
    expect(() => selectionFromRanges([])).toThrow();
  });
});
