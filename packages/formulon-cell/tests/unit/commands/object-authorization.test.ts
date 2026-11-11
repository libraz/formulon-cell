import { SparklineType } from '@libraz/formulon';
import { afterEach, describe, expect, it } from 'vitest';
import { canExecuteBuiltIn } from '../../../src/commands/built-in-command-policy.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import type {
  InteractionPolicy,
  OperationIntent,
} from '../../../src/commands/interaction-policy.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { createSessionChart } from '../../../src/commands/session-chart.js';
import { createSessionShape } from '../../../src/commands/session-illustration.js';
import { setSlicerSelectionAction } from '../../../src/commands/slicers.js';
import {
  defaultSparklineGroupSettings,
  setSparklinesAction,
  sparklineAt,
} from '../../../src/commands/sparkline.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const DATA = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 };
const OBJECT_IDS = ['chartInsert', 'pictureInsert', 'shapesInsert', 'mac.insert.slicer'] as const;
const allowAll: InteractionPolicy = { defaultOperation: 'allow' };
const denyOrigin = (denied: string): InteractionPolicy => ({
  defaultOperation: 'allow',
  restrict: ({ origin }) => origin !== denied,
});

describe('drawing-object writes under the interaction policy', () => {
  const cleanups: (() => void)[] = [];

  afterEach(() => {
    for (const cleanup of cleanups.splice(0)) cleanup();
  });

  const setup = async () => {
    const wb = await WorkbookHandle.createDefault();
    for (const [row, [item, qty]] of [
      ['Region', 'Qty'],
      ['East', 1],
      ['West', 2],
      ['East', 3],
    ].entries()) {
      wb.setText({ sheet: 0, row, col: 0 }, item as string);
      if (typeof qty === 'number') wb.setNumber({ sheet: 0, row, col: 1 }, qty);
      else wb.setText({ sheet: 0, row, col: 1 }, qty as string);
    }
    const store = createSpreadsheetStore();
    mutators.replaceCells(store, wb.cells(0));
    const history = new History();
    const controller = new InteractionController({ store, getWb: () => wb, history });
    registerInteractionController(store, controller);
    cleanups.push(() => {
      controller.dispose();
      wb.dispose();
    });
    const chart = (origin: OperationIntent['origin'] = 'ribbon') =>
      createSessionChart(
        store,
        DATA,
        { wb, anchor: { row: 0, col: 4 }, authorization: { origin, commandId: 'chartInsert' } },
        history,
      );
    const charts = () => store.getState().charts.charts.length;
    return { wb, store, history, controller, chart, charts };
  };

  for (const [label, policy] of [
    ['without a policy', undefined],
    ['under an allow-all policy', allowAll],
  ] as const satisfies readonly (readonly [string, InteractionPolicy | undefined])[]) {
    it(`needs Edit objects on a protected sheet, whatever the lock of the anchor ${label}`, async () => {
      const { store, history, controller, chart, charts } = await setup();
      controller.setPolicy(policy);
      setProtectedSheet(store, 0, true);
      for (const id of OBJECT_IDS)
        expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(false);
      expect(chart()).toBeNull();
      expect(charts()).toBe(0);
      const shape = { shape: 'rectangle' as const, x: 10, y: 10, w: 40, h: 20 };
      expect(createSessionShape(store, DATA, shape, history)).toBeNull();

      // Every cell stays locked; Edit objects alone lets the objects through.
      setProtectedSheet(store, 0, true, { permissions: { objects: true } });
      for (const id of OBJECT_IDS)
        expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(true);
      expect(chart()).not.toBeNull();
      expect(charts()).toBe(1);
      expect(createSessionShape(store, DATA, shape, history)).not.toBeNull();
    });
  }

  it('refuses an object change the policy denies, and its replay', async () => {
    const { history, controller, chart, charts } = await setup();
    controller.setPolicy({ defaultOperation: 'allow', operations: { object: false } });
    expect(chart()).toBeNull();
    controller.setPolicy(allowAll);
    expect(chart()).not.toBeNull();
    controller.setPolicy(denyOrigin('undo'));
    expect(history.undo()).toBe(false);
    expect(charts()).toBe(1);
    controller.setPolicy(allowAll);
    expect(history.undo()).toBe(true);
    expect(charts()).toBe(0);
    controller.setPolicy({ defaultOperation: 'allow', operations: { object: false } });
    expect(history.redo()).toBe(false);
    expect(charts()).toBe(0);
  });

  it('declares the sheet and the location cells of a sparkline', async () => {
    const { store, history, controller } = await setup();
    const seen: OperationIntent[] = [];
    controller.setPolicy({
      defaultOperation: 'allow',
      restrict: ({ intent }) => {
        seen.push(intent);
        return true;
      },
    });
    const addr = { sheet: 0, row: 1, col: 3 };
    const ctx = { store, history, origin: 'ribbon', commandId: 'mac.insert.sparkline' } as const;
    expect(
      setSparklinesAction(
        ctx,
        [{ addr, source: 'B2:B4' }],
        defaultSparklineGroupSettings(SparklineType.Line),
      ),
    ).toBe(true);
    expect(sparklineAt(store.getState(), addr)).not.toBeNull();
    expect(seen.find((intent) => intent.origin === 'ribbon')?.effects).toEqual([
      { kind: 'sheet', sheet: 0 },
      { kind: 'cells', cells: [addr] },
    ]);
  });

  it('decides slicer filtering as an object change and a filter use of its table', async () => {
    const { wb, store, history, controller } = await setup();
    expect(wb.createTable(0, { range: 'A1:B4', name: 'Sales', columns: ['Region', 'Qty'] })).toBe(
      0,
    );
    mutators.addSlicer(store, {
      id: 's1',
      tableName: 'Sales',
      column: 'Region',
      selected: [],
      sheet: 0,
      x: 0,
      y: 0,
    });
    const ctx = {
      store,
      history,
      workbook: wb,
      origin: 'editor',
      commandId: 'slicerFilter',
    } as const;
    const selected = () => store.getState().slicers.slicers[0]?.selected;
    const hidden = () => [...store.getState().layout.hiddenRows].sort((a, b) => a - b);

    controller.setPolicy({ defaultOperation: 'allow', operations: { filter: false } });
    expect(setSlicerSelectionAction(ctx, 's1', ['West'])).toBe(false);
    expect(selected()).toEqual([]);

    setProtectedSheet(store, 0, true, { permissions: { objects: true } });
    controller.setPolicy(allowAll);
    expect(setSlicerSelectionAction(ctx, 's1', ['West'])).toBe(false);
    setProtectedSheet(store, 0, true, { permissions: { objects: true, autoFilter: true } });
    expect(setSlicerSelectionAction(ctx, 's1', ['West'])).toBe(true);
    expect(selected()).toEqual(['West']);
    expect(hidden()).toEqual([1, 3]);
    // The selection and the rows it hid undo together.
    expect(history.undo()).toBe(true);
    expect(selected()).toEqual([]);
    expect(hidden()).toEqual([]);
  });
});
