import { afterEach, describe, expect, it } from 'vitest';
import {
  builtInCommandIntent,
  canExecuteBuiltIn,
} from '../../../src/commands/built-in-command-policy.js';
import { updateTableOverlay } from '../../../src/commands/format-as-table.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import type { InteractionPolicy } from '../../../src/commands/interaction-policy.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { executeQuickAnalysisAction } from '../../../src/commands/quick-analysis.js';
import {
  renameActiveTableAction,
  setTableColumnFunctionAction,
  type TableOption,
  toggleTableOptionAction,
} from '../../../src/commands/table-design.js';
import { hydrateTableOverlaysFromEngine } from '../../../src/engine/table-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { applyTableResize } from '../../../src/interact/table-resize.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const TABLE = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 };
const OUTSIDE = { sheet: 0, r0: 0, c0: 4, r1: 3, c1: 5 };
const CREATE_IDS = [
  'formatTableInsert',
  'mac.insert.table',
  'mac.automate.makeSubtable',
  'mac.data.dataTable',
] as const;
const REFUSED_ON_PROTECTED = [
  ...CREATE_IDS,
  'formatTableHome',
  'tableHeaderRow',
  'tableTotalRow',
  'tableName',
  'cellStyles',
] as const;
const STYLE_TOGGLES = [
  'tableBandedRows',
  'tableBandedColumns',
  'tableFirstColumn',
  'tableLastColumn',
] as const;

const allowAll: InteractionPolicy = { defaultOperation: 'allow' };
const denyTable: InteractionPolicy = { defaultOperation: 'allow', operations: { table: false } };
const denyOrigin = (denied: string): InteractionPolicy => ({
  defaultOperation: 'allow',
  restrict: ({ origin }) => origin !== denied,
});

describe('table writes under the interaction policy', () => {
  const cleanups: (() => void)[] = [];

  afterEach(() => {
    for (const cleanup of cleanups.splice(0)) cleanup();
  });

  const setup = async () => {
    const wb = await WorkbookHandle.createDefault();
    [
      ['Item', 'Qty'],
      ['A', 1],
      ['B', 2],
      ['C', 3],
    ].forEach(([item, qty], row) => {
      wb.setText({ sheet: 0, row, col: 0 }, item as string);
      if (typeof qty === 'number') wb.setNumber({ sheet: 0, row, col: 1 }, qty);
      else wb.setText({ sheet: 0, row, col: 1 }, qty as string);
    });
    expect(wb.createTable(0, { range: 'A1:B4', name: 'Sales', columns: ['Item', 'Qty'] })).toBe(0);
    const store = createSpreadsheetStore();
    mutators.replaceCells(store, wb.cells(0));
    hydrateTableOverlaysFromEngine(wb, store);
    wb.attachStore(store);
    const history = new History();
    const controller = new InteractionController({ store, getWb: () => wb, history });
    registerInteractionController(store, controller);
    cleanups.push(() => {
      controller.dispose();
      wb.detachStore(store);
      wb.dispose();
    });
    mutators.setActive(store, { sheet: 0, row: 1, col: 0 });
    const table = () => store.getState().tables.tables[0];
    const ctx = { store, history, workbook: wb, origin: 'ribbon' } as const;
    const toggle = (commandId: string, option: TableOption) =>
      toggleTableOptionAction({ ...ctx, commandId }, option);
    return { wb, store, history, controller, table, ctx, toggle };
  };

  for (const [label, policy] of [
    ['without a policy', undefined],
    ['under an allow-all policy', allowAll],
  ] as const satisfies readonly (readonly [string, InteractionPolicy | undefined])[]) {
    it(`refuses every table change on a protected sheet, the style toggles included, ${label}`, async () => {
      const { store, history, controller, table, ctx, toggle } = await setup();
      controller.setPolicy(policy);
      mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 });
      for (let row = 0; row <= 3; row += 1)
        for (let col = 0; col <= 1; col += 1)
          mutators.setCellFormat(store, { sheet: 0, row, col }, { locked: false });
      mutators.setActive(store, { sheet: 0, row: 1, col: 0 });
      setProtectedSheet(store, 0, true, { permissions: { formatCells: true } });

      for (const id of REFUSED_ON_PROTECTED)
        expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(false);
      for (const id of STYLE_TOGGLES)
        expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(false);
      // Refresh All runs data connections, not a table change: it stays available.
      expect(canExecuteBuiltIn(store, 'tableRefresh', 'ribbon').allowed).toBe(true);

      expect(toggle('tableTotalRow', 'showTotal')).toBe('none');
      expect(table()?.showTotal).toBe(false);
      expect(toggle('tableHeaderRow', 'showHeader')).toBe('none');
      expect(table()?.showHeader).toBe(true);
      expect(renameActiveTableAction({ ...ctx, commandId: 'tableName' }, 'Other')).toBe(
        'protected',
      );
      const current = table();
      if (!current) throw new Error('missing table');
      expect(
        applyTableResize({ ...ctx, origin: 'editor', commandId: 'tableResize' }, current, {
          ...TABLE,
          c1: 2,
        }),
      ).toBe(false);
      expect(table()?.range).toEqual(TABLE);

      // The style toggles are refused from the UI, unlocked cells or not.
      expect(toggle('tableBandedRows', 'banded')).toBe('none');
      expect(table()?.banded).toBe(true);
      expect(toggle('tableFirstColumn', 'firstCol')).toBe('none');
      expect(table()?.firstCol).not.toBe(true);
      expect(history.canUndo()).toBe(false);
      // The data API still changes the banding.
      expect(updateTableOverlay(store, table()?.id ?? '', { banded: false })?.banded).toBe(false);
    });

    it(`edits the total cell's function on a protected sheet only when that cell is unlocked ${label}`, async () => {
      const { store, controller, table, ctx, toggle } = await setup();
      controller.setPolicy(policy);
      expect(toggle('tableTotalRow', 'showTotal')).toBe('toggled');
      const range = table()?.range;
      if (!range) throw new Error('missing table');
      const pick = (column: number) =>
        setTableColumnFunctionAction(
          { ...ctx, origin: 'editor', commandId: 'tableTotalRow' },
          table()?.id ?? '',
          column,
          'max',
        );
      mutators.setCellFormat(store, { sheet: 0, row: range.r1, col: 1 }, { locked: false });
      setProtectedSheet(store, 0, true);
      expect(pick(1)).toBe(true);
      expect(table()?.totals?.[1]?.function).toBe('max');
      expect(pick(0)).toBe(false);
      expect(table()?.totals?.[0]?.function).not.toBe('max');
      mutators.setCellFormat(store, { sheet: 0, row: range.r1, col: 0 }, { locked: false });
      expect(pick(0)).toBe(true);
    });
  }

  it('refuses Quick Analysis Table on a protected sheet', async () => {
    const { wb, store, history } = await setup();
    setProtectedSheet(store, 0, true);
    const result = executeQuickAnalysisAction({
      store,
      wb,
      history,
      actionId: 'tables-as-table',
      range: OUTSIDE,
      stats: { sum: 0, count: 0, numericCount: 0, average: null, min: null, max: null } as never,
    });
    expect(result).toEqual({ ok: false, reason: 'protected' });
    expect(store.getState().tables.tables).toHaveLength(1);
  });

  it('refuses creating a table over an existing one on an unprotected sheet', async () => {
    const { store } = await setup();
    mutators.setRange(store, { sheet: 0, r0: 1, c0: 0, r1: 2, c1: 1 });
    for (const id of CREATE_IDS)
      expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(false);
    // Inside a table the Home gallery restyles it.
    expect(canExecuteBuiltIn(store, 'formatTableHome', 'ribbon').allowed).toBe(true);
    mutators.setRange(store, OUTSIDE);
    for (const id of CREATE_IDS)
      expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(true);
  });

  it('decides a Table Design command on the table under the active cell', async () => {
    const { store, controller } = await setup();
    controller.setPolicy(allowAll);
    expect(builtInCommandIntent(store, 'tableTotalRow', 'ribbon')?.effects).toEqual([
      { kind: 'range', range: TABLE },
    ]);
    mutators.setRange(store, OUTSIDE);
    expect(builtInCommandIntent(store, 'formatTableInsert', 'ribbon')?.effects).toEqual([
      { kind: 'range', range: OUTSIDE },
    ]);
  });

  it('refuses every table change, banding included, when the policy denies table', async () => {
    const { store, controller, table, toggle } = await setup();
    controller.setPolicy(denyTable);
    for (const id of [
      ...REFUSED_ON_PROTECTED.filter((id) => id !== 'cellStyles'),
      ...STYLE_TOGGLES,
    ]) {
      if (id === 'mac.data.dataTable') continue;
      expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(false);
    }
    expect(toggle('tableBandedRows', 'banded')).toBe('none');
    expect(table()?.banded).toBe(true);
  });

  it('refuses replay once the policy denies the undo or redo', async () => {
    const { history, controller, table, toggle } = await setup();
    controller.setPolicy(allowAll);
    expect(toggle('tableBandedRows', 'banded')).toBe('toggled');
    controller.setPolicy(denyOrigin('undo'));
    expect(history.undo()).toBe(false);
    expect(table()?.banded).toBe(false);
    controller.setPolicy(allowAll);
    expect(history.undo()).toBe(true);
    expect(table()?.banded).toBe(true);
    controller.setPolicy(denyTable);
    expect(history.redo()).toBe(false);
    expect(table()?.banded).toBe(true);
  });
});
