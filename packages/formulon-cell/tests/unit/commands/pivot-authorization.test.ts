import { afterEach, describe, expect, it } from 'vitest';
import { canExecuteBuiltIn } from '../../../src/commands/built-in-command-policy.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import type { InteractionPolicy } from '../../../src/commands/interaction-policy.js';
import { applyPivotGrandTotals } from '../../../src/commands/pivot-layout.js';
import { createPivotTableAtPlacement } from '../../../src/commands/pivot-table.js';
import {
  setProtectedSheet,
  setWorkbookStructureProtected,
} from '../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { commitPivotChange } from '../../../src/interact/pivot-field-settings-dialog-apply.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const SOURCE = { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 2 };

const allowAll: InteractionPolicy = { defaultOperation: 'allow' };
const denyPivot: InteractionPolicy = { defaultOperation: 'allow', operations: { pivot: false } };
const denyOrigin = (denied: string): InteractionPolicy => ({
  defaultOperation: 'allow',
  restrict: ({ origin }) => origin !== denied,
});

describe('pivot writes under the interaction policy', () => {
  const handles: WorkbookHandle[] = [];
  const controllers: InteractionController[] = [];

  afterEach(() => {
    for (const controller of controllers.splice(0)) controller.dispose();
    for (const handle of handles.splice(0)) handle.dispose();
  });

  const setup = async () => {
    const wb = await WorkbookHandle.createDefault();
    handles.push(wb);
    const rows: (string | number)[][] = [
      ['Region', 'Sales'],
      ['East', 12],
      ['West', 5],
    ];
    for (const [r, row] of rows.entries()) {
      for (const [c, value] of row.entries()) {
        const addr = { sheet: 0, row: r + 1, col: c + 1 };
        if (typeof value === 'number') wb.setNumber(addr, value);
        else wb.setText(addr, value);
      }
    }
    const store = createSpreadsheetStore();
    mutators.replaceCells(store, wb.cells(0));
    const history = new History();
    const controller = new InteractionController({ store, getWb: () => wb, history });
    controllers.push(controller);
    registerInteractionController(store, controller);
    const create = (placement: { row: number; col: number } | 'new-sheet' = { row: 1, col: 5 }) =>
      createPivotTableAtPlacement(
        { store, workbook: wb, history, origin: 'ribbon', commandId: 'pivotTableInsert' },
        { source: SOURCE, rowField: 'Region', valueField: 'Sales' },
        placement,
      );
    const reports = (): number => wb.getPivotTables().length;
    const grandTotals = (rows: boolean, pivotIndex = 0) =>
      commitPivotChange(
        {
          wb,
          history,
          store,
          origin: 'ribbon',
          commandId: 'pivotGrandTotals',
          onChanged: () => undefined,
        },
        0,
        pivotIndex,
        () => applyPivotGrandTotals(wb, { sheetIndex: 0, pivotIndex }, { rows, cols: rows }),
      );
    return { wb, store, history, controller, create, reports, grandTotals };
  };

  it('creates a report under a policy and replays it through undo and redo', async () => {
    const { history, controller, create, reports } = await setup();
    controller.setPolicy(allowAll);
    expect(create()).toMatchObject({ ok: true });
    expect(reports()).toBe(1);
    expect(history.undo()).toBe(true);
    expect(reports()).toBe(0);
    expect(history.redo()).toBe(true);
    expect(reports()).toBe(1);
  });

  it('refuses creating and editing when the policy denies pivot', async () => {
    const { store, controller, create, reports, grandTotals } = await setup();
    controller.setPolicy(allowAll);
    create();
    controller.setPolicy(denyPivot);
    expect(create({ row: 1, col: 9 })).toMatchObject({ ok: false, reason: 'refused' });
    expect(reports()).toBe(1);
    expect(grandTotals(false)).toBe(false);
    for (const id of ['pivotTableInsert', 'pivotDelete', 'pivotRefresh'] as const) {
      expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(false);
    }
  });

  for (const [label, policy] of [
    ['without a policy', undefined],
    ['under an allow-all policy', allowAll],
  ] as const satisfies readonly (readonly [string, InteractionPolicy | undefined])[]) {
    it(`refuses Insert PivotTable on a protected sheet, and edits need the option, ${label}`, async () => {
      const { store, controller, create, reports, grandTotals } = await setup();
      controller.setPolicy(policy);
      create();
      setProtectedSheet(store, 0, true, { permissions: { pivotTables: true } });
      expect(canExecuteBuiltIn(store, 'pivotTableInsert', 'ribbon').allowed).toBe(false);
      expect(create({ row: 1, col: 9 })).toMatchObject({ ok: false });
      expect(create('new-sheet')).toMatchObject({ ok: false });
      expect(reports()).toBe(1);
      expect(grandTotals(false)).toBe(true);

      setProtectedSheet(store, 0, true);
      expect(grandTotals(true)).toBe(false);
    });
  }

  it('refuses a report on a new sheet under workbook-structure protection', async () => {
    const { store, controller, create, reports } = await setup();
    controller.setPolicy(allowAll);
    setWorkbookStructureProtected(store, true);
    expect(create('new-sheet')).toMatchObject({ ok: false, reason: 'refused' });
    expect(reports()).toBe(0);
    expect(create()).toMatchObject({ ok: true });
  });

  it('refuses replay once the policy denies the undo or redo', async () => {
    const { history, controller, create, reports } = await setup();
    controller.setPolicy(allowAll);
    create();
    controller.setPolicy(denyOrigin('undo'));
    expect(history.undo()).toBe(false);
    expect(reports()).toBe(1);
    controller.setPolicy(allowAll);
    expect(history.undo()).toBe(true);
    controller.setPolicy(denyPivot);
    expect(history.redo()).toBe(false);
    expect(reports()).toBe(0);
  });
});
