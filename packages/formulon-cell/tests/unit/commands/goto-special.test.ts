import { beforeEach, describe, expect, it } from 'vitest';
import { boundingRange, findMatchingCells } from '../../../src/commands/goto-special.js';
import { WorkbookHandle, addrKey } from '../../../src/engine/workbook-handle.js';
import {
  type SpreadsheetStore,
  createSpreadsheetStore,
  mutators,
} from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

/** Replicate the engine cells into the store data slice — the dialog reads
 *  validation entries off the format slice and conditional rules off the
 *  conditional slice, so we keep both in lockstep with the engine here. */
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

describe('findMatchingCells', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

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
    const got = findMatchingCells(wb, store, 'sheet', 'formulas');
    expect(got).toHaveLength(1);
    expect(got[0]).toEqual({ sheet: 0, row: 0, col: 0 });
  });

  it('constants — populated cells without a formula', () => {
    wb.setFormula({ sheet: 0, row: 0, col: 0 }, '=1+2');
    wb.setText({ sheet: 0, row: 0, col: 1 }, 'hello');
    wb.setNumber({ sheet: 0, row: 1, col: 0 }, 42);
    wb.recalc();
    sync(store, wb);
    const got = findMatchingCells(wb, store, 'sheet', 'constants');
    const keys = got.map((a) => `${a.row}:${a.col}`).sort();
    expect(keys).toEqual(['0:1', '1:0']);
  });

  it('numbers — only number-kind cells', () => {
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    wb.setNumber({ sheet: 0, row: 0, col: 1 }, 2);
    wb.setText({ sheet: 0, row: 0, col: 2 }, 'three');
    wb.setFormula({ sheet: 0, row: 1, col: 0 }, '=10');
    wb.recalc();
    sync(store, wb);
    const got = findMatchingCells(wb, store, 'sheet', 'numbers');
    const keys = got.map((a) => `${a.row}:${a.col}`).sort();
    // Formula cell evaluates to a number — Excel includes it under "Numbers".
    expect(keys).toEqual(['0:0', '0:1', '1:0']);
  });

  it('text — only text-kind cells, excluding error sentinels', () => {
    wb.setText({ sheet: 0, row: 0, col: 0 }, 'apple');
    wb.setText({ sheet: 0, row: 0, col: 1 }, '#REF!');
    wb.setNumber({ sheet: 0, row: 1, col: 0 }, 1);
    wb.recalc();
    sync(store, wb);
    const got = findMatchingCells(wb, store, 'sheet', 'text');
    expect(got).toHaveLength(1);
    expect(got[0]).toEqual({ sheet: 0, row: 0, col: 0 });
  });

  it('errors — formula errors and explicit error sentinels in text', () => {
    wb.setFormula({ sheet: 0, row: 0, col: 0 }, '=1/0');
    wb.setText({ sheet: 0, row: 0, col: 1 }, '#NAME?');
    wb.setText({ sheet: 0, row: 1, col: 0 }, 'fine');
    wb.recalc();
    sync(store, wb);
    const got = findMatchingCells(wb, store, 'sheet', 'errors');
    const keys = got.map((a) => `${a.row}:${a.col}`).sort();
    expect(keys).toEqual(['0:0', '0:1']);
  });

  it('blanks — empty cells inside a selection rectangle', () => {
    // Populate a sparse 3×3 region; the rest stay blank.
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    wb.setNumber({ sheet: 0, row: 2, col: 2 }, 9);
    wb.recalc();
    sync(store, wb);
    // Select the full 3×3 rectangle so blanks are bounded.
    store.setState((s) => ({
      ...s,
      selection: {
        active: { sheet: 0, row: 0, col: 0 },
        anchor: { sheet: 0, row: 0, col: 0 },
        range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
      },
    }));
    const got = findMatchingCells(wb, store, 'selection', 'blanks');
    // 9 cells in the rect minus 2 populated = 7 blank.
    expect(got).toHaveLength(7);
    expect(got).not.toContainEqual({ sheet: 0, row: 0, col: 0 });
    expect(got).not.toContainEqual({ sheet: 0, row: 2, col: 2 });
  });

  it('non-blanks — every populated cell', () => {
    wb.setText({ sheet: 0, row: 0, col: 0 }, 'a');
    wb.setNumber({ sheet: 0, row: 1, col: 1 }, 2);
    wb.recalc();
    sync(store, wb);
    const got = findMatchingCells(wb, store, 'sheet', 'non-blanks');
    expect(got).toHaveLength(2);
  });

  it('data-validation — cells with a validation entry on the format slice', () => {
    wb.setText({ sheet: 0, row: 0, col: 0 }, 'a');
    wb.setText({ sheet: 0, row: 5, col: 1 }, 'b');
    wb.recalc();
    sync(store, wb);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 5, col: 1 },
      { validation: { kind: 'list', source: ['a', 'b', 'c'] } },
    );
    const got = findMatchingCells(wb, store, 'sheet', 'data-validation');
    expect(got).toEqual([{ sheet: 0, row: 5, col: 1 }]);
  });

  it('conditional-format — cells inside any rule range', () => {
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    wb.setNumber({ sheet: 0, row: 1, col: 0 }, 2);
    wb.setNumber({ sheet: 0, row: 5, col: 5 }, 3);
    wb.recalc();
    sync(store, wb);
    mutators.addConditionalRule(store, {
      kind: 'data-bar',
      range: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 },
      color: '#638ec6',
    });
    const got = findMatchingCells(wb, store, 'sheet', 'conditional-format');
    const keys = got.map((a) => `${a.row}:${a.col}`).sort();
    expect(keys).toEqual(['0:0', '1:0']);
  });

  it('selection scope confines the sweep to the selection rect', () => {
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    wb.setNumber({ sheet: 0, row: 5, col: 5 }, 2);
    wb.recalc();
    sync(store, wb);
    store.setState((s) => ({
      ...s,
      selection: {
        active: { sheet: 0, row: 0, col: 0 },
        anchor: { sheet: 0, row: 0, col: 0 },
        range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
      },
    }));
    const got = findMatchingCells(wb, store, 'selection', 'numbers');
    expect(got).toEqual([{ sheet: 0, row: 0, col: 0 }]);
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
