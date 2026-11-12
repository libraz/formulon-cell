import { beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { removeDuplicates } from '../../../src/commands/sort.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const seedText = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  row: number,
  col: number,
  value: string,
): void => {
  wb.setText({ sheet: 0, row, col }, value);
  store.setState((s) => {
    const map = new Map(s.data.cells);
    map.set(addrKey({ sheet: 0, row, col }), { value: { kind: 'text', value }, formula: null });
    return { ...s, data: { ...s.data, cells: map } };
  });
};

describe('removeDuplicates', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await WorkbookHandle.createDefault({ preferStub: true });
  });

  it('keeps rows whose column texts only collide when concatenated', () => {
    seedText(store, wb, 0, 0, 'a');
    seedText(store, wb, 0, 1, 'text:b');
    seedText(store, wb, 1, 0, 'atext:');
    seedText(store, wb, 1, 1, 'b');
    const removed = removeDuplicates(store.getState(), store, wb, {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 1,
      c1: 1,
    });
    expect(removed).toEqual({ ok: true, removed: 0, remaining: 2 });
  });

  it('restores formats on undo', () => {
    seedText(store, wb, 0, 0, 'x');
    seedText(store, wb, 1, 0, 'x');
    seedText(store, wb, 2, 0, 'y');
    mutators.setCellFormat(store, { sheet: 0, row: 2, col: 0 }, { bold: true });
    const history = new History();
    wb.attachHistory(history);
    history.begin();
    const removed = removeDuplicates(
      store.getState(),
      store,
      wb,
      { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 },
      {},
      history,
    );
    history.end();
    expect(removed).toEqual({ ok: true, removed: 1, remaining: 2 });
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 0 })),
    ).toMatchObject({ bold: true });

    expect(history.undo()).toBe(true);
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 0 })),
    ).toMatchObject({ bold: true });
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 0 })),
    ).toBeUndefined();
  });
});
