import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { clearSelectedContents } from '../../../src/commands/clear-contents.js';
import { copy, planCopyRanges } from '../../../src/commands/clipboard/copy.js';
import { pasteSpecial } from '../../../src/commands/clipboard/paste-special.js';
import { captureSnapshotFromCopyPlan } from '../../../src/commands/clipboard/snapshot.js';
import { executeRibbonFillAction } from '../../../src/commands/fill.js';
import { History } from '../../../src/commands/history.js';
import { deleteVisibleRows } from '../../../src/commands/structure.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

// List A1:C11 (header row 0, data rows 1..10); rows 1,2,4,5,6 hidden by the filter.
const LIST: Range = { sheet: 0, r0: 0, c0: 0, r1: 10, c1: 2 };
const HIDDEN = [1, 2, 4, 5, 6];

describe('edits on a filtered list act on visible rows only', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let history: History;

  const num = (row: number, col: number): number | null => {
    const v = wb.getValue({ sheet: 0, row, col });
    return v.kind === 'number' ? v.value : null;
  };
  const select = (range: Range): void => {
    store.setState((s) => ({
      ...s,
      selection: {
        ...s.selection,
        active: { sheet: 0, row: range.r0, col: range.c0 },
        anchor: { sheet: 0, row: range.r0, col: range.c0 },
        range,
        extraRanges: [],
      },
    }));
  };
  const refresh = (): void => mutators.replaceCells(store, wb.cells(0));

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await WorkbookHandle.createDefault({ preferStub: true });
    history = new History();
    for (let r = 1; r <= 10; r += 1) {
      for (let c = 0; c < 3; c += 1) wb.setNumber({ sheet: 0, row: r, col: c }, r * 10 + c);
    }
    wb.setNumber({ sheet: 0, row: 0, col: 8 }, 777);
    refresh();
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenRows: new Set(HIDDEN) },
      ui: {
        ...s.ui,
        filterRanges: [LIST],
        filterCriteria: [{ range: LIST, byCol: 0, hiddenValues: ['x'] }],
      },
    }));
  });

  afterEach(() => wb.dispose());

  it('Clear Contents leaves the hidden rows alone', () => {
    select({ sheet: 0, r0: 1, c0: 2, r1: 10, c1: 2 });
    clearSelectedContents({ store, workbook: wb, history });
    expect(num(3, 2)).toBeNull();
    expect(num(7, 2)).toBeNull();
    expect(num(1, 2)).toBe(12);
    expect(num(4, 2)).toBe(42);
    expect(num(6, 2)).toBe(62);
  });

  it('Fill Down copies the first visible row of the selection into visible rows only', () => {
    select({ sheet: 0, r0: 1, c0: 0, r1: 10, c1: 2 });
    executeRibbonFillAction({ store, workbook: wb, history, action: 'down' });
    expect(num(3, 0)).toBe(30);
    expect(num(7, 0)).toBe(30);
    expect(num(9, 1)).toBe(31);
    expect(num(1, 0)).toBe(10);
    expect(num(4, 0)).toBe(40);
  });

  it('Fill Down on one cell copies the row above even when the filter hides it', () => {
    select({ sheet: 0, r0: 3, c0: 2, r1: 3, c1: 2 });
    expect(executeRibbonFillAction({ store, workbook: wb, history, action: 'down' })).toBe(true);
    expect(num(3, 2)).toBe(22);
  });

  it('Delete Sheet Rows removes visible rows only and the filter range shrinks', () => {
    select({ sheet: 0, r0: 1, c0: 0, r1: 10, c1: 2 });
    deleteVisibleRows(store, wb, history, 1, 10);
    refresh();
    // Rows 1,2,4,5,6 survive, in order, at the top of the list.
    expect([1, 2, 3, 4, 5].map((r) => num(r, 0))).toEqual([10, 20, 40, 50, 60]);
    expect(num(6, 0)).toBeNull();
    const [range] = store.getState().ui.filterRanges;
    expect(range?.r1).toBe(5);
    expect(store.getState().ui.filterCriteria).toHaveLength(1);
  });

  it('pasting one copied cell onto a filtered selection writes visible cells only', () => {
    select({ sheet: 0, r0: 0, c0: 8, r1: 0, c1: 8 });
    const plan = planCopyRanges(store.getState());
    const snap = captureSnapshotFromCopyPlan(store.getState(), plan as NonNullable<typeof plan>);
    select({ sheet: 0, r0: 1, c0: 2, r1: 10, c1: 2 });
    pasteSpecial(
      store.getState(),
      store,
      wb,
      snap as NonNullable<typeof snap>,
      { what: 'all', operation: 'none', skipBlanks: false, transpose: false },
      history,
    );
    expect(num(3, 2)).toBe(777);
    expect(num(10, 2)).toBe(777);
    expect(num(1, 2)).toBe(12);
    expect(num(5, 2)).toBe(52);
  });

  it('copying a filtered range copies visible rows, so the paste is contiguous', () => {
    select({ sheet: 0, r0: 1, c0: 0, r1: 10, c1: 0 });
    const result = copy(store.getState());
    expect(result?.tsv).toBe('30\r\n70\r\n80\r\n90\r\n100');
    const plan = planCopyRanges(store.getState());
    const snap = captureSnapshotFromCopyPlan(store.getState(), plan as NonNullable<typeof plan>);
    select({ sheet: 0, r0: 14, c0: 0, r1: 14, c1: 0 });
    pasteSpecial(
      store.getState(),
      store,
      wb,
      snap as NonNullable<typeof snap>,
      { what: 'all', operation: 'none', skipBlanks: false, transpose: false },
      history,
    );
    expect([14, 15, 16, 17, 18].map((r) => num(r, 0))).toEqual([30, 70, 80, 90, 100]);
    expect(num(19, 0)).toBeNull();
  });
});
