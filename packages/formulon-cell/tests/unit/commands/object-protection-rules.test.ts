import { afterEach, describe, expect, it } from 'vitest';
import { canExecuteBuiltIn } from '../../../src/commands/built-in-command-policy.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { setCellLocked, setProtectedSheet } from '../../../src/commands/protection.js';
import { executeRibbonClearAction } from '../../../src/commands/ribbon-clear.js';
import { removeSlicerAction } from '../../../src/commands/slicers.js';
import { setSparklines, sparklineAt } from '../../../src/commands/sparkline.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const G = (row: number) => ({ sheet: 0, row, col: 6 });

describe('object protection rules', () => {
  const cleanups: (() => void)[] = [];

  afterEach(() => {
    for (const cleanup of cleanups.splice(0)) cleanup();
  });

  const setup = async () => {
    const workbook = await WorkbookHandle.createDefault();
    for (let row = 0; row < 4; row += 1) workbook.setNumber({ sheet: 0, row, col: 0 }, row + 1);
    const store = createSpreadsheetStore();
    mutators.replaceCells(store, workbook.cells(0));
    const history = new History();
    const controller = new InteractionController({ store, getWb: () => workbook, history });
    registerInteractionController(store, controller);
    cleanups.push(() => {
      controller.dispose();
      workbook.dispose();
    });
    // One three-cell group on G1:G3.
    expect(
      setSparklines(
        store,
        [0, 1, 2].map((row) => ({ addr: G(row), spec: { kind: 'line', source: 'A1:A4' } })),
        history,
      ),
    ).toBe(true);
    const clear = (action: 'all' | 'contents' | 'formats', r0: number, r1: number) => {
      mutators.setRange(store, { sheet: 0, r0, c0: 6, r1, c1: 6 });
      executeRibbonClearAction({ store, workbook, history, action });
    };
    return { workbook, store, history, controller, clear };
  };

  it('keeps sparklines through Clear Contents and Clear Formats, and Clear All removes only the cleared cells', async () => {
    const { store, clear } = await setup();
    const group = sparklineAt(store.getState(), G(0))?.group;
    expect(group).toBeDefined();
    clear('contents', 1, 1);
    clear('formats', 1, 1);
    expect(sparklineAt(store.getState(), G(1))).not.toBeNull();

    clear('all', 1, 1);
    expect(sparklineAt(store.getState(), G(1))).toBeNull();
    // G1 and G3 stay one group.
    expect(sparklineAt(store.getState(), G(0))?.group).toBe(group);
    expect(sparklineAt(store.getState(), G(2))?.group).toBe(group);

    clear('all', 0, 2);
    expect([0, 1, 2].map((row) => sparklineAt(store.getState(), G(row)))).toEqual([
      null,
      null,
      null,
    ]);
  });

  it('clears a sparkline on a protected sheet by its cell lock, not Edit objects', async () => {
    const { store, clear } = await setup();
    setCellLocked(store, { sheet: 0, r0: 1, c0: 6, r1: 1, c1: 6 }, false);
    setProtectedSheet(store, 0, true, { permissions: { objects: true } });
    clear('all', 0, 0);
    expect(sparklineAt(store.getState(), G(0))).not.toBeNull();
    setProtectedSheet(store, 0, true);
    clear('all', 1, 1);
    expect(sparklineAt(store.getState(), G(1))).toBeNull();
  });

  it('greys Insert Sparkline under any protection', async () => {
    const { store } = await setup();
    expect(canExecuteBuiltIn(store, 'mac.insert.sparkline', 'ribbon').allowed).toBe(true);
    for (const permissions of [{}, { objects: true }]) {
      setProtectedSheet(store, 0, true, { permissions });
      expect(canExecuteBuiltIn(store, 'mac.insert.sparkline', 'ribbon').allowed).toBe(false);
    }
  });

  it('deletes a slicer by Edit objects alone, Use AutoFilter not mattering', async () => {
    const { workbook, store, history } = await setup();
    workbook.setText({ sheet: 0, row: 0, col: 2 }, 'Qty');
    for (let row = 1; row < 4; row += 1) workbook.setNumber({ sheet: 0, row, col: 2 }, row);
    expect(workbook.createTable(0, { range: 'C1:C4', name: 'Sales', columns: ['Qty'] })).toBe(0);
    mutators.addSlicer(store, {
      id: 's1',
      tableName: 'Sales',
      column: 'Qty',
      selected: [],
      sheet: 0,
      x: 0,
      y: 0,
    });
    const ctx = {
      store,
      history,
      workbook,
      origin: 'editor',
      commandId: 'slicerDelete',
    } as const;
    setProtectedSheet(store, 0, true, { permissions: { autoFilter: true } });
    expect(removeSlicerAction(ctx, 's1')).toBe(false);
    expect(store.getState().slicers.slicers).toHaveLength(1);
    setProtectedSheet(store, 0, true, { permissions: { objects: true } });
    expect(removeSlicerAction(ctx, 's1')).toBe(true);
    expect(store.getState().slicers.slicers).toHaveLength(0);
  });
});
