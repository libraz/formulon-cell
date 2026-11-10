import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import type { InteractionPolicy } from '../../../src/commands/interaction-policy.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { sortRangeWithHistory } from '../../../src/commands/sort.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

/** A1 header, A2:A4 data. */
const RANGE: Range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };

const allowAll: InteractionPolicy = { defaultOperation: 'allow' };
const denySort: InteractionPolicy = { defaultOperation: 'allow', operations: { sort: false } };
const denyOrigin = (denied: string): InteractionPolicy => ({
  defaultOperation: 'allow',
  restrict: ({ origin }) => origin !== denied,
});

describe('sort under the interaction policy', () => {
  const handles: WorkbookHandle[] = [];
  const controllers: InteractionController[] = [];

  afterEach(() => {
    for (const controller of controllers.splice(0)) controller.dispose();
    for (const handle of handles.splice(0)) handle.dispose();
  });

  const setup = async () => {
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    handles.push(workbook);
    const store = createSpreadsheetStore();
    const history = new History();
    const controller = new InteractionController({ store, getWb: () => workbook, history });
    controllers.push(controller);
    registerInteractionController(store, controller);
    workbook.setText({ sheet: 0, row: 0, col: 0 }, 'Qty');
    for (const [row, value] of [
      [1, 3],
      [2, 1],
      [3, 2],
    ] as const) {
      workbook.setNumber({ sheet: 0, row, col: 0 }, value);
    }
    mutators.replaceCells(store, workbook.cells(0));
    workbook.attachHistory(history);
    const column = (): unknown[] =>
      [1, 2, 3].map((row) => {
        const value = workbook.getValue({ sheet: 0, row, col: 0 });
        return value.kind === 'number' ? value.value : value.kind;
      });
    const sort = (): boolean =>
      sortRangeWithHistory({
        store,
        workbook,
        history,
        range: RANGE,
        options: { byCol: 0, direction: 'asc', hasHeader: true },
        origin: 'ribbon',
        commandId: 'sortAsc',
      });
    const unlockData = (): void => {
      for (const row of [1, 2, 3])
        mutators.setCellFormat(store, { sheet: 0, row, col: 0 }, { locked: false });
    };
    return { store, history, controller, column, sort, unlockData };
  };

  it('sorts under a policy and replays through undo and redo', async () => {
    const { history, controller, column, sort } = await setup();
    controller.setPolicy(allowAll);
    expect(sort()).toBe(true);
    expect(column()).toEqual([1, 2, 3]);
    expect(history.undo()).toBe(true);
    expect(column()).toEqual([3, 1, 2]);
    expect(history.redo()).toBe(true);
    expect(column()).toEqual([1, 2, 3]);
  });

  it('refuses when the policy denies sort', async () => {
    const { history, controller, column, sort } = await setup();
    controller.setPolicy(denySort);
    expect(sort()).toBe(false);
    expect(column()).toEqual([3, 1, 2]);
    expect(history.canUndo()).toBe(false);
  });

  for (const [label, policy] of [
    ['without a policy', undefined],
    ['under an allow-all policy', allowAll],
  ] as const) {
    it(`needs the Sort option and unlocked data on a protected sheet ${label}`, async () => {
      const { store, history, controller, column, sort, unlockData } = await setup();
      controller.setPolicy(policy);
      setProtectedSheet(store, 0, true, { permissions: { sort: true } });
      expect(sort()).toBe(false);

      unlockData();
      setProtectedSheet(store, 0, true);
      expect(sort()).toBe(false);
      expect(column()).toEqual([3, 1, 2]);
      expect(history.canUndo()).toBe(false);

      setProtectedSheet(store, 0, true, { permissions: { sort: true } });
      expect(sort()).toBe(true);
      expect(column()).toEqual([1, 2, 3]);
    });
  }

  it('refuses replay once the policy denies the undo or redo', async () => {
    const { history, controller, column, sort } = await setup();
    controller.setPolicy(allowAll);
    expect(sort()).toBe(true);
    controller.setPolicy(denyOrigin('undo'));
    expect(history.undo()).toBe(false);
    expect(column()).toEqual([1, 2, 3]);
    controller.setPolicy(allowAll);
    expect(history.undo()).toBe(true);
    controller.setPolicy(denySort);
    expect(history.redo()).toBe(false);
    expect(column()).toEqual([3, 1, 2]);
  });
});
