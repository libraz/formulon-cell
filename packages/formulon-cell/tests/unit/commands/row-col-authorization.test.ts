import { afterEach, describe, expect, it } from 'vitest';
import { canExecuteBuiltIn } from '../../../src/commands/built-in-command-policy.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import type { InteractionPolicy } from '../../../src/commands/interaction-policy.js';
import { groupCols, groupRows } from '../../../src/commands/outline.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import {
  deleteColumnsAction,
  insertCellsAction,
  insertRowsAction,
  layoutWrite,
} from '../../../src/commands/row-col-actions.js';
import { hideRows } from '../../../src/commands/row-col-layout.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const allowAll: InteractionPolicy = { defaultOperation: 'allow' };
const denyOrigin = (denied: string): InteractionPolicy => ({
  defaultOperation: 'allow',
  restrict: ({ origin }) => origin !== denied,
});
const OUTLINE_STRUCTURE = [
  'outlineGroup',
  'outlineUngroup',
  'rowGroup',
  'rowUngroup',
  'colGroup',
  'colUngroup',
  'mac.data.group',
  'mac.data.ungroup',
  'outlineSummaryBelow',
  'outlineSummaryRight',
] as const;
const DETAIL = [
  'outlineShowDetail',
  'outlineHideDetail',
  'mac.data.showDetail',
  'mac.data.hideDetail',
] as const;

describe('row and column writes under the interaction policy', () => {
  const cleanups: (() => void)[] = [];

  afterEach(() => {
    for (const cleanup of cleanups.splice(0)) cleanup();
  });

  const setup = async () => {
    const workbook = await WorkbookHandle.createDefault();
    for (let row = 0; row < 6; row += 1) workbook.setNumber({ sheet: 0, row, col: 0 }, row + 1);
    const store = createSpreadsheetStore();
    mutators.replaceCells(store, workbook.cells(0));
    const history = new History();
    const controller = new InteractionController({ store, getWb: () => workbook, history });
    registerInteractionController(store, controller);
    cleanups.push(() => {
      controller.dispose();
      workbook.dispose();
    });
    const ctx = (commandId: string) =>
      ({ store, workbook, history, origin: 'ribbon', commandId }) as const;
    const value = (row: number) => workbook.getValue({ sheet: 0, row, col: 0 });
    return { store, workbook, history, controller, ctx, value };
  };

  for (const [label, policy] of [
    ['without a policy', undefined],
    ['under an allow-all policy', allowAll],
  ] as const satisfies readonly (readonly [string, InteractionPolicy | undefined])[]) {
    it(`decides row and column changes on a protected sheet by their options ${label}`, async () => {
      const { store, history, controller, ctx, value } = await setup();
      controller.setPolicy(policy);
      setProtectedSheet(store, 0, true);
      for (const id of [
        'insertRows',
        'deleteRows',
        'colInsertLeft',
        'rowHeight',
        'colWidth',
      ] as const)
        expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(false);
      expect(insertRowsAction(ctx('insertRows'), 1)).toBe(false);
      expect(value(1)).toEqual({ kind: 'number', value: 2 });
      expect(layoutWrite(ctx('rowHide'), 'row', false, () => true)).toBe(false);

      setProtectedSheet(store, 0, true, {
        permissions: { insertRows: true, formatRows: true, formatColumns: true },
      });
      for (const id of ['insertRows', 'rowHeight', 'colWidth', 'rowHide'] as const)
        expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(true);
      expect(canExecuteBuiltIn(store, 'colDelete', 'ribbon').allowed).toBe(false);
      expect(insertRowsAction(ctx('insertRows'), 1)).toBe(true);
      expect(value(2)).toEqual({ kind: 'number', value: 2 });
      expect(
        layoutWrite(ctx('rowHide'), 'row', false, () => {
          hideRows(store, history, 0, 0);
          return true;
        }),
      ).toBe(true);
      expect(store.getState().layout.hiddenRows.has(0)).toBe(true);
    });

    it(`refuses Group, Ungroup and the outline settings under any protection ${label}`, async () => {
      const { store, controller } = await setup();
      controller.setPolicy(policy);
      setProtectedSheet(store, 0, true, { permissions: { formatRows: true, formatColumns: true } });
      for (const id of OUTLINE_STRUCTURE)
        expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(false);
    });

    it(`decides Show/Hide Detail by the axis of the outline ${label}`, async () => {
      const { store, history, controller } = await setup();
      controller.setPolicy(policy);
      groupRows(store, history, 1, 3);
      groupCols(store, history, 1, 3);
      mutators.setRange(store, { sheet: 0, r0: 1, c0: 0, r1: 3, c1: 0 });
      setProtectedSheet(store, 0, true, { permissions: { formatColumns: true } });
      for (const id of DETAIL)
        expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(false);
      setProtectedSheet(store, 0, true, { permissions: { formatRows: true } });
      for (const id of DETAIL)
        expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(true);

      mutators.setRange(store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 3 });
      for (const id of DETAIL)
        expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(false);
      setProtectedSheet(store, 0, true, { permissions: { formatColumns: true } });
      for (const id of DETAIL)
        expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(true);
    });
  }

  it('decides a cell shift as the rows or columns it moves', async () => {
    const { store, controller, ctx, value } = await setup();
    controller.setPolicy({ defaultOperation: 'allow', operations: { insertRows: false } });
    const range = { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 };
    expect(insertCellsAction(ctx('insertCells'), range, 'down')).toBe(false);
    expect(value(1)).toEqual({ kind: 'number', value: 2 });
    expect(canExecuteBuiltIn(store, 'insertRows', 'ribbon').allowed).toBe(false);
    expect(deleteColumnsAction(ctx('colDelete'), 3)).toBe(true);
  });

  it('refuses replay once the policy denies the undo or redo', async () => {
    const { history, controller, ctx, value } = await setup();
    controller.setPolicy(allowAll);
    expect(insertRowsAction(ctx('insertRows'), 1)).toBe(true);
    controller.setPolicy(denyOrigin('undo'));
    expect(history.undo()).toBe(false);
    expect(value(2)).toEqual({ kind: 'number', value: 2 });
    controller.setPolicy(allowAll);
    expect(history.undo()).toBe(true);
    expect(value(1)).toEqual({ kind: 'number', value: 2 });
    controller.setPolicy({ defaultOperation: 'allow', operations: { insertRows: false } });
    expect(history.redo()).toBe(false);
    expect(value(1)).toEqual({ kind: 'number', value: 2 });
  });
});
