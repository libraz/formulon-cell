import { describe, expect, it } from 'vitest';
import { planSelectionFormat } from '../../../src/commands/format.js';
import { History } from '../../../src/commands/history.js';
import { addrKey } from '../../../src/engine/address.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { runFormatDialogTransaction } from '../../../src/interact/format-dialog-transaction.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const selectRow = (store: SpreadsheetStore): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      active: { sheet: 0, row: 0, col: 0 },
      anchor: { sheet: 0, row: 0, col: 0 },
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
    },
  }));
};

const workbook = (addMerge: () => void): WorkbookHandle =>
  ({
    capabilities: { merges: true },
    engineClearMerges: () => true,
    engineAddMerge: addMerge,
    setBlank: () => undefined,
    setText: () => undefined,
    getValue: () => ({ kind: 'blank' }),
  }) as unknown as WorkbookHandle;

const run = (store: SpreadsheetStore, history: History | null, wb: WorkbookHandle) => {
  const state = store.getState();
  const plan = planSelectionFormat(state);
  if (!plan) throw new Error('no plan');
  return runFormatDialogTransaction({
    store,
    history,
    getWb: () => wb,
    state,
    liveWb: wb,
    plan,
    action: { patch: { bold: true } },
    merge: { action: 'merge', range: state.selection.range },
  });
};

describe('runFormatDialogTransaction', () => {
  it('applies format and merge as one undo step and keeps the format repeat', () => {
    const store = createSpreadsheetStore();
    selectRow(store);
    const history = new History();
    expect(
      run(
        store,
        history,
        workbook(() => true),
      ),
    ).toBe(true);
    const a1 = addrKey({ sheet: 0, row: 0, col: 0 });
    expect(store.getState().format.formats.get(a1)?.bold).toBe(true);
    expect(store.getState().merges.byAnchor.size).toBe(1);

    expect(history.undo()).toBe(true);
    expect(store.getState().format.formats.get(a1)?.bold).toBeUndefined();
    expect(store.getState().merges.byAnchor.size).toBe(0);
    expect(history.canUndo()).toBe(false);
  });

  it('rolls the format back and rethrows when the merge step throws', () => {
    const store = createSpreadsheetStore();
    selectRow(store);
    const history = new History();
    const failure = new Error('engine merge failed');
    expect(() =>
      run(
        store,
        history,
        workbook(() => {
          throw failure;
        }),
      ),
    ).toThrow(failure);
    const a1 = addrKey({ sheet: 0, row: 0, col: 0 });
    expect(store.getState().format.formats.get(a1)?.bold).toBeUndefined();
    expect(history.canUndo()).toBe(false);
  });
});
