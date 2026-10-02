import { afterEach, describe, expect, it } from 'vitest';
import {
  clearSelectedContents,
  collectSelectedContentAddresses,
} from '../../../src/commands/clear-contents.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { fixedFormPolicy } from '../../../src/commands/interaction-policy.js';
import type { Addr, Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const addr = (row: number, col: number, sheet = 0): Addr => ({ sheet, row, col });

describe('clearSelectedContents', () => {
  let workbook: WorkbookHandle | undefined;

  afterEach(() => {
    workbook?.dispose();
    workbook = undefined;
  });

  const setSelection = (store: SpreadsheetStore, range: Range, extraRanges: Range[] = []): void => {
    store.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        active: addr(range.r0, range.c0, range.sheet),
        anchor: addr(range.r0, range.c0, range.sheet),
        range,
        extraRanges,
      },
    }));
  };

  it('collects the sparse union once, skips the hole, and restores formulas in one undo', async () => {
    workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const store = createSpreadsheetStore();
    workbook.setNumber(addr(0, 0), 1);
    workbook.setNumber(addr(0, 1), 2);
    workbook.setNumber(addr(0, 2), 3);
    workbook.setNumber(addr(1, 0), 4);
    workbook.setNumber(addr(1, 1), 5); // the hollow selection hole
    workbook.setNumber(addr(1, 2), 6);
    workbook.setFormula(addr(2, 2), '=""'); // formula result is blank-looking, but is content
    mutators.replaceCells(store, workbook.physicalCells(0));
    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 }, [
      { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 },
      { sheet: 0, r0: 1, c0: 2, r1: 2, c1: 2 },
    ]);

    const selected = collectSelectedContentAddresses(store, workbook);
    expect(selected).toEqual([
      addr(0, 0),
      addr(0, 1),
      addr(0, 2),
      addr(1, 0),
      addr(1, 2),
      addr(2, 2),
    ]);

    const history = new History();
    const result = clearSelectedContents({ store, workbook, history, origin: 'ribbon' });
    expect(result.status).toBe('applied');
    expect(result.applied).toEqual(selected);
    expect(workbook.getValue(addr(1, 1))).toEqual({ kind: 'number', value: 5 });
    expect(workbook.cellFormula(addr(2, 2))).toBeNull();
    expect(history.canUndo()).toBe(true);
    expect(history.undo()).toBe(true);
    expect(workbook.getValue(addr(0, 1))).toEqual({ kind: 'number', value: 2 });
    expect(workbook.cellFormula(addr(2, 2))).toBe('=""');
    expect(history.redo()).toBe(true);
    expect(workbook.getValue(addr(0, 1))).toEqual({ kind: 'blank' });
  });

  it('deduplicates overlapping ranges and clears a merge only when fully covered', async () => {
    workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const store = createSpreadsheetStore();
    workbook.setNumber(addr(0, 1), 10);
    workbook.setNumber(addr(0, 2), 20);
    workbook.setNumber(addr(0, 3), 30);
    mutators.replaceCells(store, workbook.physicalCells(0));

    const merge: Range = { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 2 };
    store.setState((state) => ({
      ...state,
      merges: {
        byAnchor: new Map([['0:0:1', merge]]),
        byCell: new Map([['0:0:2', '0:0:1']]),
      },
    }));
    setSelection(store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }, [
      { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
    ]);
    expect(collectSelectedContentAddresses(store, workbook)).toEqual([]);

    setSelection(store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 3 }, [merge]);
    expect(collectSelectedContentAddresses(store, workbook)).toEqual([addr(0, 1), addr(0, 3)]);

    const result = clearSelectedContents({ store, workbook, history: new History() });
    expect(result.status).toBe('applied');
    expect(result.applied).toEqual([addr(0, 1), addr(0, 3)]);
    expect(workbook.getValue(addr(0, 1))).toEqual({ kind: 'blank' });
    expect(workbook.getValue(addr(0, 3))).toEqual({ kind: 'blank' });
    expect(store.getState().merges.byAnchor).toEqual(new Map([['0:0:1', merge]]));
    expect(store.getState().merges.byCell).toEqual(new Map([['0:0:2', '0:0:1']]));
  });

  it('ignores styled blank cache entries while retaining formula cells with blank-looking values', async () => {
    workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const store = createSpreadsheetStore();
    workbook.setFormula(addr(0, 1), '=""');
    mutators.setCellFormat(store, addr(0, 0), { fill: '#ffeeaa' });
    mutators.replaceCells(store, workbook.physicalCells(0));
    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });

    expect(collectSelectedContentAddresses(store, workbook)).toEqual([addr(0, 1)]);
  });

  it('scans a giant sparse column and sheet through physical cells only', async () => {
    workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const store = createSpreadsheetStore();
    const first = addr(900_000, 0);
    const second = addr(1_048_575, 16_383);
    workbook.setNumber(first, 7);
    workbook.setText(second, 'edge');

    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 1_048_575, c1: 0 });
    expect(collectSelectedContentAddresses(store, workbook)).toEqual([first]);
    setSelection(store, {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 1_048_575,
      c1: 16_383,
    });
    expect(collectSelectedContentAddresses(store, workbook)).toEqual([first, second]);

    const result = clearSelectedContents({ store, workbook });
    expect(result.status).toBe('applied');
    expect(result.applied).toEqual([first, second]);
    expect(workbook.getValue(first)).toEqual({ kind: 'blank' });
    expect(workbook.getValue(second)).toEqual({ kind: 'blank' });
  });

  it('reports a registered controller revision for an empty selection', async () => {
    workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const store = createSpreadsheetStore();
    const history = new History();
    const controller = new InteractionController({
      store,
      getWb: () => workbook as WorkbookHandle,
      history,
    });
    const unregister = registerInteractionController(store, controller);
    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });

    try {
      controller.setPolicy({ defaultOperation: 'allow' });
      const result = clearSelectedContents({ store, workbook, history });
      expect(result).toMatchObject({ status: 'noop', applied: [], revision: 1 });
      expect(history.canUndo()).toBe(false);
    } finally {
      unregister();
      controller.dispose();
    }
  });

  it('rejects a mixed explicit policy batch atomically', async () => {
    workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const store = createSpreadsheetStore();
    const history = new History();
    workbook.setNumber(addr(0, 0), 1);
    workbook.setNumber(addr(0, 1), 2);
    mutators.replaceCells(store, workbook.physicalCells(0));
    const range: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 };
    setSelection(store, range);
    const controller = new InteractionController({
      store,
      getWb: () => workbook as WorkbookHandle,
      history,
    });
    controller.setPolicy(fixedFormPolicy([{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }]));

    try {
      const result = clearSelectedContents({ store, workbook, history, commands: controller });
      expect(result.status).toBe('rejected');
      expect(result.applied).toEqual([]);
      expect(workbook.getValue(addr(0, 0))).toEqual({ kind: 'number', value: 1 });
      expect(workbook.getValue(addr(0, 1))).toEqual({ kind: 'number', value: 2 });
      expect(history.canUndo()).toBe(false);
    } finally {
      controller.dispose();
    }
  });

  it('forwards a composite command id and keeps the standalone default', async () => {
    workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const store = createSpreadsheetStore();
    const history = new History();
    const target = addr(0, 0);
    workbook.setNumber(target, 7);
    mutators.replaceCells(store, workbook.physicalCells(0));
    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const controller = new InteractionController({
      store,
      getWb: () => workbook as WorkbookHandle,
      history,
    });
    const commandIds: string[] = [];
    controller.setPolicy({
      ...fixedFormPolicy([{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }]),
      restrict: ({ commandId }) => {
        commandIds.push(commandId ?? '');
        return true;
      },
    });

    try {
      expect(
        clearSelectedContents({
          store,
          workbook,
          history,
          commands: controller,
          origin: 'ribbon',
          commandId: 'clear-all',
        }).status,
      ).toBe('applied');
      workbook.setNumber(target, 7);
      mutators.replaceCells(store, workbook.physicalCells(0));
      expect(
        clearSelectedContents({
          store,
          workbook,
          history,
          commands: controller,
          origin: 'ribbon',
        }).status,
      ).toBe('applied');
      expect(commandIds[0]).toBe('clear-all');
      expect(commandIds[commandIds.length - 1]).toBe('clear-contents');
      expect(new Set(commandIds)).toEqual(new Set(['clear-all', 'clear-contents']));
    } finally {
      controller.dispose();
    }
  });
});
