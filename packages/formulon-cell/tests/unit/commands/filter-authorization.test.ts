import { afterEach, describe, expect, it } from 'vitest';
import { canExecuteBuiltIn } from '../../../src/commands/built-in-command-policy.js';
import { applyValueFilter, copyAdvancedFilterResult } from '../../../src/commands/filter.js';
import {
  advancedFilterCopyWrite,
  type FilterActionContext,
  filterUseWrite,
} from '../../../src/commands/filter-actions.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import type { InteractionPolicy } from '../../../src/commands/interaction-policy.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { toggleAutoFilterFromSelection } from '../../../src/interact/ribbon-filter-data.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

/** A1 header, A2:A4 values 1, 2, 1. */
const LIST: Range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };

const allowAll: InteractionPolicy = { defaultOperation: 'allow' };
const denyFilter: InteractionPolicy = { defaultOperation: 'allow', operations: { filter: false } };
const denyOrigin = (denied: string): InteractionPolicy => ({
  defaultOperation: 'allow',
  restrict: ({ origin }) => origin !== denied,
});

describe('filter writes under the interaction policy', () => {
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
    workbook.attachHistory(history);
    workbook.setText({ sheet: 0, row: 0, col: 0 }, 'Qty');
    for (const [row, value] of [
      [1, 1],
      [2, 2],
      [3, 1],
    ] as const) {
      workbook.setNumber({ sheet: 0, row, col: 0 }, value);
    }
    mutators.replaceCells(store, workbook.cells(0));
    mutators.setRange(store, LIST);
    history.clear();
    const ctx = (origin: FilterActionContext['origin'] = 'ribbon'): FilterActionContext => ({
      store,
      history,
      workbook,
      origin,
      commandId: 'filterByValue',
    });
    const hidden = (): number[] => [...store.getState().layout.hiddenRows].sort();
    const filterOutTwos = (origin?: FilterActionContext['origin']) =>
      filterUseWrite(ctx(origin), LIST, false, () => {
        applyValueFilter(store.getState(), store, LIST, 0, ['2'], workbook);
        return true;
      });
    const toggle = (origin: FilterActionContext['origin'] = 'keyboard') =>
      toggleAutoFilterFromSelection({ ...ctx(origin), commandId: 'filter' });
    return { store, workbook, history, controller, ctx, hidden, filterOutTwos, toggle };
  };

  it('filters under a policy and replays through undo and redo', async () => {
    const { history, controller, hidden, filterOutTwos } = await setup();
    controller.setPolicy(allowAll);
    expect(filterOutTwos()).toBe(true);
    expect(hidden()).toEqual([2]);
    expect(history.undo()).toBe(true);
    expect(hidden()).toEqual([]);
    expect(history.redo()).toBe(true);
    expect(hidden()).toEqual([2]);
  });

  it('refuses filter writes and commands when the policy denies filter', async () => {
    const { store, history, controller, hidden, filterOutTwos, toggle } = await setup();
    controller.setPolicy(denyFilter);
    expect(filterOutTwos()).toBe(false);
    toggle();
    expect(hidden()).toEqual([]);
    expect(store.getState().ui.filterRanges).toEqual([]);
    expect(history.canUndo()).toBe(false);
    for (const id of [
      'filter',
      'filterClear',
      'filterReapply',
      'filterByValue',
      'mac.data.filter',
      'mac.data.clear',
      'mac.data.advancedFilter',
      'tableFilterButton',
    ] as const) {
      expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(false);
    }
  });

  for (const [label, policy] of [
    ['without a policy', undefined],
    ['under an allow-all policy', allowAll],
  ] as const satisfies readonly (readonly [string, InteractionPolicy | undefined])[]) {
    it(`uses a filter on a protected sheet only with Use AutoFilter, never toggles ${label}`, async () => {
      const { store, controller, hidden, filterOutTwos, toggle } = await setup();
      controller.setPolicy(policy);
      toggle();
      expect(store.getState().ui.filterRanges).toHaveLength(1);
      setProtectedSheet(store, 0, true);
      expect(canExecuteBuiltIn(store, 'filterClear', 'ribbon').allowed).toBe(false);
      expect(filterOutTwos()).toBe(false);
      expect(hidden()).toEqual([]);

      setProtectedSheet(store, 0, true, { permissions: { autoFilter: true } });
      expect(canExecuteBuiltIn(store, 'filterClear', 'ribbon').allowed).toBe(true);
      expect(canExecuteBuiltIn(store, 'filter', 'ribbon').allowed).toBe(false);
      expect(filterOutTwos()).toBe(true);
      expect(hidden()).toEqual([2]);
      toggle();
      expect(store.getState().ui.filterRanges).toHaveLength(1);
    });
  }

  it('refuses replay once the policy denies the undo or redo', async () => {
    const { history, controller, hidden, filterOutTwos } = await setup();
    controller.setPolicy(allowAll);
    filterOutTwos();
    controller.setPolicy(denyOrigin('undo'));
    expect(history.undo()).toBe(false);
    expect(hidden()).toEqual([2]);
    controller.setPolicy(allowAll);
    expect(history.undo()).toBe(true);
    controller.setPolicy(denyFilter);
    expect(history.redo()).toBe(false);
    expect(hidden()).toEqual([]);
  });

  it('decides Ctrl+Shift+L under the keyboard origin', async () => {
    const { store, controller, toggle } = await setup();
    controller.setPolicy(denyOrigin('keyboard'));
    toggle('keyboard');
    expect(store.getState().ui.filterRanges).toEqual([]);
    toggle('ribbon');
    expect(store.getState().ui.filterRanges).toHaveLength(1);
  });

  it('declares the Advanced Filter copy destination as a value write', async () => {
    const { store, workbook, controller, ctx } = await setup();
    workbook.setText({ sheet: 0, row: 0, col: 3 }, 'Qty');
    workbook.setNumber({ sheet: 0, row: 1, col: 3 }, 1);
    mutators.replaceCells(store, workbook.cells(0));
    const criteria: Range = { sheet: 0, r0: 0, c0: 3, r1: 1, c1: 3 };
    const output: Range = { sheet: 0, r0: 0, c0: 6, r1: 3, c1: 6 };
    const copy = () =>
      advancedFilterCopyWrite(ctx(), LIST, output, null, () =>
        copyAdvancedFilterResult(
          store.getState(),
          store,
          LIST,
          criteria,
          { sheet: 0, row: 0, col: 6 },
          {},
          workbook,
        ),
      );
    controller.setPolicy({ defaultOperation: 'allow', operations: { valueEdit: false } });
    expect(copy()).toBeNull();
    expect(workbook.getValue({ sheet: 0, row: 0, col: 6 }).kind).toBe('blank');
    controller.setPolicy(allowAll);
    expect(copy()).toBe(3);
    expect(workbook.getValue({ sheet: 0, row: 1, col: 6 })).toEqual({ kind: 'number', value: 1 });
  });
});
