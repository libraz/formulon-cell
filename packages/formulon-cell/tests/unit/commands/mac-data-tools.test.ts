// @vitest-environment node

import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { InteractionController } from '../../../src/commands/interaction-controller.js';
import type { OperationIntent } from '../../../src/commands/interaction-policy.js';
import {
  type ConsolidateRequest,
  cellSnapshot,
  commitMacConsolidate,
  commitMacGoalSeek,
  commitMacSubtotal,
  MAX_SUBTOTAL_GROUPS,
  MAX_SUBTOTAL_OUTPUT_CELLS,
  planMacConsolidate,
  planMacSubtotal,
  type SubtotalRequest,
  solveMacGoalSeek,
  supportsSubtotal,
} from '../../../src/commands/mac-data-tools.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import type { SpreadsheetInstance } from '../../../src/mount/types.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const allow = (): { allowed: true } => ({ allowed: true });

const makeInstance = (wb: WorkbookHandle, store: SpreadsheetStore): SpreadsheetInstance => {
  const commands = {
    canExecute: () => allow(),
    execute(command: {
      changes: readonly {
        addr: { sheet: number; row: number; col: number };
        value: import('../../../src/engine/types.js').CellValue;
        formula?: string | null;
      }[];
    }) {
      const atomic = wb.applyCellPatchAtomic(command.changes);
      return {
        status: atomic.changed.length > 0 ? 'applied' : 'noop',
        applied: atomic.changed,
        rejected: [],
        revision: 1,
      } as const;
    },
  };
  return { workbook: wb, store, commands } as unknown as SpreadsheetInstance;
};

const makeControllerInstance = (
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  history: History,
): SpreadsheetInstance => {
  const commands = new InteractionController({
    store,
    getWb: () => wb,
    history,
  });
  wb.attachHistory(history);
  return { workbook: wb, store, commands, history } as unknown as SpreadsheetInstance;
};

describe('Mac Data command operations', () => {
  const handles: WorkbookHandle[] = [];

  afterEach(() => {
    for (const wb of handles.splice(0)) wb.dispose();
  });

  it('solves on a scratch workbook and leaves live cells untouched until commit', async () => {
    const wb = await WorkbookHandle.createDefault();
    handles.push(wb);
    const store = createSpreadsheetStore();
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 3);
    wb.setFormula({ sheet: 0, row: 0, col: 1 }, '=A1*2');
    wb.recalc();
    const instance = makeInstance(wb, store);

    const result = await solveMacGoalSeek(instance, {
      formulaCell: { sheet: 0, row: 0, col: 1 },
      targetValue: 10,
      changingCell: { sheet: 0, row: 0, col: 0 },
    });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value.changingValue).toBeCloseTo(5, 8);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 3 });
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 6 });
  });

  it('recalculates each scratch trial in manual calculation mode', async () => {
    const wb = await WorkbookHandle.createDefault();
    handles.push(wb);
    if (wb.isStub || !wb.capabilities.calcMode) return;
    const store = createSpreadsheetStore();
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 3);
    wb.setFormula({ sheet: 0, row: 0, col: 1 }, '=A1*2');
    wb.recalc();
    expect(wb.setCalcMode(1)).toBe(true);
    try {
      const result = await solveMacGoalSeek(makeInstance(wb, store), {
        formulaCell: { sheet: 0, row: 0, col: 1 },
        targetValue: 10,
        changingCell: { sheet: 0, row: 0, col: 0 },
      });
      expect(result.ok).toBe(true);
      if (result.ok) expect(result.value.changingValue).toBeCloseTo(5, 8);
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 3 });
    } finally {
      wb.setCalcMode(0);
    }
  });

  it('rejects Goal Seek before saving or trialing a formula changing cell', async () => {
    const wb = await WorkbookHandle.createDefault();
    handles.push(wb);
    const store = createSpreadsheetStore();
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 3);
    wb.setFormula({ sheet: 0, row: 0, col: 1 }, '=A1*2');
    wb.setFormula({ sheet: 0, row: 0, col: 2 }, '=A1+1');
    const result = await solveMacGoalSeek(makeInstance(wb, store), {
      formulaCell: { sheet: 0, row: 0, col: 1 },
      targetValue: 10,
      changingCell: { sheet: 0, row: 0, col: 2 },
    });
    expect(result).toMatchObject({ ok: false, error: { status: 'invalid' } });
    expect(wb.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({ kind: 'number', value: 4 });
  });

  it('records accepted Goal Seek as one undoable controller intent', async () => {
    const wb = await WorkbookHandle.createDefault();
    handles.push(wb);
    expect(wb.isStub).toBe(false);
    const store = createSpreadsheetStore();
    const history = new History();
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 3);
    wb.setFormula({ sheet: 0, row: 0, col: 1 }, '=A1*2');
    wb.recalc();
    const instance = makeControllerInstance(wb, store, history);
    const request = {
      formulaCell: { sheet: 0, row: 0, col: 1 },
      targetValue: 10,
      changingCell: { sheet: 0, row: 0, col: 0 },
    } as const;
    const solved = await solveMacGoalSeek(instance, request);
    expect(solved.ok).toBe(true);
    if (!solved.ok) return;
    const committed = commitMacGoalSeek(
      instance,
      request,
      solved.value,
      cellSnapshot(wb, request.formulaCell),
      cellSnapshot(wb, request.changingCell),
    );
    expect(committed.ok).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 5 });
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 10 });
    expect(history.undo()).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 3 });
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 6 });
    expect(history.redo()).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 5 });
  });

  it('rejects Goal Seek when an unchanged observed cell has a changed precedent', async () => {
    const wb = await WorkbookHandle.createDefault();
    handles.push(wb);
    expect(wb.isStub).toBe(false);
    const store = createSpreadsheetStore();
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 0);
    wb.setFormula({ sheet: 0, row: 0, col: 1 }, '=A1*C1');
    wb.setNumber({ sheet: 0, row: 0, col: 2 }, 2);
    wb.recalc();
    const request = {
      formulaCell: { sheet: 0, row: 0, col: 1 },
      targetValue: 10,
      changingCell: { sheet: 0, row: 0, col: 0 },
    } as const;
    const solved = await solveMacGoalSeek(makeInstance(wb, store), request);
    expect(solved.ok).toBe(true);
    if (!solved.ok) return;
    const formulaBefore = cellSnapshot(wb, request.formulaCell);
    const changingBefore = cellSnapshot(wb, request.changingCell);
    wb.setNumber({ sheet: 0, row: 0, col: 2 }, -2);
    wb.recalc();
    const history = new History();
    const instance = makeControllerInstance(wb, store, history);

    const committed = commitMacGoalSeek(
      instance,
      request,
      solved.value,
      formulaBefore,
      changingBefore,
    );

    expect(committed).toMatchObject({ ok: false, error: { status: 'stale' } });
    expect(wb.getValue(request.changingCell)).toEqual({ kind: 'number', value: 0 });
    expect(history.canUndo()).toBe(false);
  });

  it('rejects a Goal Seek solution reused with a different changing cell', async () => {
    const wb = await WorkbookHandle.createDefault();
    handles.push(wb);
    expect(wb.isStub).toBe(false);
    const store = createSpreadsheetStore();
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 0);
    wb.setFormula({ sheet: 0, row: 0, col: 1 }, '=A1*C1');
    wb.setNumber({ sheet: 0, row: 0, col: 2 }, 2);
    wb.setNumber({ sheet: 0, row: 0, col: 3 }, 0);
    wb.recalc();
    const originalRequest = {
      formulaCell: { sheet: 0, row: 0, col: 1 },
      targetValue: 10,
      changingCell: { sheet: 0, row: 0, col: 0 },
    } as const;
    const solved = await solveMacGoalSeek(makeInstance(wb, store), originalRequest);
    expect(solved.ok).toBe(true);
    if (!solved.ok) return;
    const reusedRequest = {
      ...originalRequest,
      changingCell: { sheet: 0, row: 0, col: 3 },
    } as const;
    const history = new History();
    const instance = makeControllerInstance(wb, store, history);

    const committed = commitMacGoalSeek(
      instance,
      reusedRequest,
      solved.value,
      cellSnapshot(wb, reusedRequest.formulaCell),
      cellSnapshot(wb, reusedRequest.changingCell),
    );

    expect(committed).toMatchObject({ ok: false, error: { status: 'stale' } });
    expect(wb.getValue(originalRequest.changingCell)).toEqual({ kind: 'number', value: 0 });
    expect(wb.getValue(reusedRequest.changingCell)).toEqual({ kind: 'number', value: 0 });
    expect(history.canUndo()).toBe(false);
  });

  it('captures overlapping consolidate sources before a single output commit', async () => {
    const handle = await WorkbookHandle.createDefault();
    handles.push(handle);
    const store = createSpreadsheetStore();
    handle.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    handle.setNumber({ sheet: 0, row: 1, col: 0 }, 2);
    handle.setNumber({ sheet: 0, row: 0, col: 2 }, 3);
    handle.setNumber({ sheet: 0, row: 1, col: 2 }, 4);
    handle.recalc();
    const history = new History();
    const instance = makeControllerInstance(handle, store, history);
    const request: ConsolidateRequest = {
      sources: ['A1:A2', 'C1:C2'],
      destination: 'A1:A2',
      function: 'sum',
    };
    const plan = planMacConsolidate(instance, request);
    expect(plan.ok).toBe(true);
    if (!plan.ok) return;
    expect(plan.value.values).toEqual([
      { kind: 'number', value: 4 },
      { kind: 'number', value: 6 },
    ]);
    expect(commitMacConsolidate(instance, request, plan.value).ok).toBe(true);
    expect(handle.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 4 });
    expect(handle.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'number', value: 6 });
    expect(history.undo()).toBe(true);
    expect(handle.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 1 });
    expect(history.redo()).toBe(true);
    expect(handle.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'number', value: 6 });
  });

  it('refuses consolidate before writing when authorization denies the output', async () => {
    const wb = await WorkbookHandle.createDefault();
    handles.push(wb);
    const store = createSpreadsheetStore();
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    wb.setNumber({ sheet: 0, row: 0, col: 1 }, 2);
    const instance = makeInstance(wb, store);
    instance.commands.canExecute = (() => ({
      allowed: false,
      code: 'protected',
      reason: 'blocked',
    })) as never;
    const plan = planMacConsolidate(instance, {
      sources: ['A1:B1'],
      destination: 'C1',
      function: 'sum',
    });
    expect(plan).toEqual({ ok: false, error: { status: 'rejected', reason: 'blocked' } });
    expect(wb.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({ kind: 'blank' });
  });

  it('writes native SUBTOTAL rows bottom-up and groups each data block in one history step', async () => {
    const wb = await WorkbookHandle.createDefault();
    handles.push(wb);
    expect(supportsSubtotal(wb)).toBe(true);
    const store = createSpreadsheetStore();
    const history = new History();
    const rows: [string, number][] = [
      ['A', 1],
      ['A', 2],
      ['B', 4],
      ['B', 5],
    ];
    for (const [index, [group, value]] of rows.entries()) {
      wb.setText({ sheet: 0, row: index + 1, col: 0 }, group);
      wb.setNumber({ sheet: 0, row: index + 1, col: 1 }, value);
    }
    wb.recalc();
    const instance = makeControllerInstance(wb, store, history);
    const request: SubtotalRequest = {
      range: 'A1:B5',
      groupByColumn: 0,
      subtotalColumns: [1],
      function: 'sum',
    };
    const plan = planMacSubtotal(instance, request);
    expect(plan.ok).toBe(true);
    if (!plan.ok) return;
    expect(commitMacSubtotal(instance, request, plan.value).ok).toBe(true);
    expect(wb.cellFormula({ sheet: 0, row: 3, col: 1 })).toContain('SUBTOTAL(9');
    expect(wb.cellFormula({ sheet: 0, row: 6, col: 1 })).toContain('SUBTOTAL(9');
    expect(wb.getValue({ sheet: 0, row: 3, col: 1 })).toEqual({ kind: 'number', value: 3 });
    expect(wb.getValue({ sheet: 0, row: 6, col: 1 })).toEqual({ kind: 'number', value: 9 });
    expect(history.canUndo()).toBe(true);
    expect(history.undo()).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 4, col: 1 })).toEqual({ kind: 'number', value: 5 });
  });

  it('aborts every Subtotal mutation when a later structural step fails', async () => {
    const wb = await WorkbookHandle.createDefault();
    handles.push(wb);
    if (wb.isStub) return;
    const store = createSpreadsheetStore();
    const history = new History();
    wb.setText({ sheet: 0, row: 0, col: 0 }, 'Group');
    wb.setText({ sheet: 0, row: 1, col: 0 }, 'A');
    wb.setNumber({ sheet: 0, row: 1, col: 1 }, 1);
    wb.setText({ sheet: 0, row: 2, col: 0 }, 'B');
    wb.setNumber({ sheet: 0, row: 2, col: 1 }, 2);
    wb.recalc();
    const instance = makeControllerInstance(wb, store, history);
    const plan = planMacSubtotal(instance, {
      range: 'A1:B3',
      groupByColumn: 0,
      subtotalColumns: [1],
      function: 'sum',
    });
    expect(plan.ok).toBe(true);
    if (!plan.ok) return;
    const firstGroup = plan.value.groups[0];
    expect(firstGroup).toBeDefined();
    if (!firstGroup) return;
    const failed = {
      ...plan.value,
      groups: [firstGroup, { start: 1_048_575, end: 1_048_575, insertAt: 1_048_575, label: 'bad' }],
    };
    const committed = commitMacSubtotal(
      instance,
      { range: 'A1:B3', groupByColumn: 0, subtotalColumns: [1], function: 'sum' },
      failed,
    );
    expect(committed.ok).toBe(false);
    expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'number', value: 1 });
    expect(wb.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({ kind: 'number', value: 2 });
    expect(history.canUndo()).toBe(false);
  });

  it('bounds subtotal output before materializing a wide Cartesian address set', async () => {
    const wb = await WorkbookHandle.createDefault();
    handles.push(wb);
    expect(wb.isStub).toBe(false);
    const store = createSpreadsheetStore();
    for (let row = 1; row <= 7; row += 1) {
      wb.setText({ sheet: 0, row, col: 0 }, `Group ${row}`);
    }
    wb.recalc();
    const calls: OperationIntent[] = [];
    const instance = makeInstance(wb, store);
    instance.commands.canExecute = ((intent: OperationIntent) => {
      calls.push(intent);
      return allow();
    }) as never;
    const columns = Array.from({ length: 16_384 }, (_, col) => col);
    const result = planMacSubtotal(instance, {
      range: 'A1:XFD8',
      groupByColumn: 0,
      subtotalColumns: columns,
      function: 'sum',
    });

    expect(result).toMatchObject({ ok: false, error: { status: 'unsupported' } });
    expect(calls.some((intent) => intent.effects.some((effect) => effect.kind === 'cells'))).toBe(
      false,
    );
    expect(MAX_SUBTOTAL_OUTPUT_CELLS).toBe(100_000);
  });

  it('bounds subtotal group insertion work before planning output cells', async () => {
    const subtotalGroupLimit = 1_000;
    const wb = await WorkbookHandle.createDefault();
    handles.push(wb);
    expect(wb.isStub).toBe(false);
    const store = createSpreadsheetStore();
    for (let row = 1; row <= subtotalGroupLimit + 1; row += 1) {
      wb.setText({ sheet: 0, row, col: 0 }, `Group ${row}`);
    }
    wb.recalc();
    const calls: OperationIntent[] = [];
    const instance = makeInstance(wb, store);
    instance.commands.canExecute = ((intent: OperationIntent) => {
      calls.push(intent);
      return allow();
    }) as never;
    const result = planMacSubtotal(instance, {
      range: `A1:B${subtotalGroupLimit + 2}`,
      groupByColumn: 0,
      subtotalColumns: [1],
      function: 'sum',
    });
    expect(result).toMatchObject({ ok: false, error: { status: 'unsupported' } });
    expect(calls.some((intent) => intent.effects.some((effect) => effect.kind === 'cells'))).toBe(
      false,
    );
    expect(MAX_SUBTOTAL_GROUPS).toBe(1_000);
  });

  it('authorizes subtotal apply, undo, redo, and outline as one native history item', async () => {
    const wb = await WorkbookHandle.createDefault();
    handles.push(wb);
    expect(wb.isStub).toBe(false);
    const store = createSpreadsheetStore();
    const history = new History();
    for (const [row, group, value] of [
      [1, 'A', 1],
      [2, 'A', 2],
      [3, 'B', 4],
      [4, 'B', 5],
    ] as const) {
      wb.setText({ sheet: 0, row, col: 0 }, group);
      wb.setNumber({ sheet: 0, row, col: 1 }, value);
    }
    wb.recalc();
    const instance = makeControllerInstance(wb, store, history);
    const restrictedOrigins: string[] = [];
    instance.commands.setPolicy({
      defaultOperation: 'deny',
      operations: {
        insertRows: true,
        deleteRows: true,
        valueEdit: true,
        formulaEdit: true,
        format: true,
      },
      restrict: ({ intent }) => {
        if (intent.commandId === 'mac.data.subtotal') {
          restrictedOrigins.push(`${intent.operation}:${intent.origin}`);
        }
        return true;
      },
    });
    const calls: OperationIntent[] = [];
    const originalCanExecute = instance.commands.canExecute.bind(instance.commands);
    instance.commands.canExecute = ((intent: OperationIntent) => {
      calls.push(intent);
      return originalCanExecute(intent);
    }) as never;
    const request: SubtotalRequest = {
      range: 'A1:B5',
      groupByColumn: 0,
      subtotalColumns: [1],
      function: 'sum',
    };
    const plan = planMacSubtotal(instance, request);
    expect(plan.ok).toBe(true);
    if (!plan.ok) return;
    expect(commitMacSubtotal(instance, request, plan.value).ok).toBe(true);
    expect(history.canUndo()).toBe(true);
    expect(wb.cellFormula({ sheet: 0, row: 3, col: 1 })).toContain('SUBTOTAL(9');
    expect(wb.cellFormula({ sheet: 0, row: 6, col: 1 })).toContain('SUBTOTAL(9');
    expect(store.getState().layout.outlineRows.size).toBeGreaterThan(0);
    expect(calls.some((intent) => intent.operation === 'format')).toBe(true);
    expect(
      calls.some((intent) => intent.operation === 'format' && intent.origin === 'ribbon'),
    ).toBe(true);
    expect(restrictedOrigins).toEqual(
      expect.arrayContaining([
        'insertRows:ribbon',
        'valueEdit:ribbon',
        'formulaEdit:ribbon',
        'format:ribbon',
      ]),
    );

    expect(history.undo()).toBe(true);
    expect(history.canUndo()).toBe(false);
    expect(wb.cellFormula({ sheet: 0, row: 3, col: 1 })).toBeNull();
    expect(store.getState().layout.outlineRows.size).toBe(0);
    expect(calls.some((intent) => intent.operation === 'deleteRows')).toBe(true);
    expect(calls.some((intent) => intent.operation === 'format' && intent.origin === 'undo')).toBe(
      true,
    );

    expect(history.redo()).toBe(true);
    expect(wb.cellFormula({ sheet: 0, row: 6, col: 1 })).toContain('SUBTOTAL(9');
    expect(store.getState().layout.outlineRows.size).toBeGreaterThan(0);
    expect(
      calls.some((intent) => intent.operation === 'insertRows' && intent.origin === 'redo'),
    ).toBe(true);
    expect(calls.some((intent) => intent.operation === 'format' && intent.origin === 'redo')).toBe(
      true,
    );
  });

  it('fails closed before subtotal mutation when real policy denies outline format', async () => {
    const wb = await WorkbookHandle.createDefault();
    handles.push(wb);
    expect(wb.isStub).toBe(false);
    const store = createSpreadsheetStore();
    const history = new History();
    wb.setText({ sheet: 0, row: 1, col: 0 }, 'A');
    wb.setNumber({ sheet: 0, row: 1, col: 1 }, 1);
    wb.setText({ sheet: 0, row: 2, col: 0 }, 'B');
    wb.setNumber({ sheet: 0, row: 2, col: 1 }, 2);
    wb.recalc();
    const instance = makeControllerInstance(wb, store, history);
    instance.commands.setPolicy({
      defaultOperation: 'deny',
      operations: {
        insertRows: true,
        deleteRows: true,
        valueEdit: true,
        formulaEdit: true,
        format: false,
      },
    });
    const request: SubtotalRequest = {
      range: 'A1:B3',
      groupByColumn: 0,
      subtotalColumns: [1],
      function: 'sum',
    };

    const result = planMacSubtotal(instance, request);

    expect(result).toMatchObject({ ok: false, error: { status: 'rejected' } });
    expect(result.ok ? undefined : result.error.reason).toContain('format');
    expect(history.canUndo()).toBe(false);
    expect(wb.cellFormula({ sheet: 0, row: 2, col: 1 })).toBeNull();
    expect(store.getState().layout.outlineRows.size).toBe(0);
  });

  it('rejects a Subtotal plan when the active sheet changes before commit', async () => {
    const wb = await WorkbookHandle.createDefault();
    handles.push(wb);
    expect(wb.isStub).toBe(false);
    const otherSheet = wb.addSheet('Other Data');
    expect(otherSheet).toBe(1);
    const store = createSpreadsheetStore();
    const history = new History();
    wb.setText({ sheet: 0, row: 1, col: 0 }, 'A');
    wb.setNumber({ sheet: 0, row: 1, col: 1 }, 1);
    wb.setText({ sheet: 0, row: 2, col: 0 }, 'B');
    wb.setNumber({ sheet: 0, row: 2, col: 1 }, 2);
    wb.recalc();
    const instance = makeControllerInstance(wb, store, history);
    const request: SubtotalRequest = {
      range: `${wb.sheetName(0)}!A1:B3`,
      groupByColumn: 0,
      subtotalColumns: [1],
      function: 'sum',
    };
    const plan = planMacSubtotal(instance, request);
    expect(plan.ok).toBe(true);
    if (!plan.ok) return;
    mutators.setSheetIndex(store, otherSheet);

    const result = commitMacSubtotal(instance, request, plan.value);

    expect(result).toMatchObject({ ok: false, error: { status: 'stale' } });
    expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'number', value: 1 });
    expect(wb.getValue({ sheet: otherSheet, row: 1, col: 1 })).toEqual({ kind: 'blank' });
    expect(store.getState().layout.outlineRows.size).toBe(0);
    expect(history.canUndo()).toBe(false);
  });

  it('consolidates quoted cross-sheet sources and replays only the destination sheet', async () => {
    const wb = await WorkbookHandle.createDefault();
    handles.push(wb);
    expect(wb.isStub).toBe(false);
    const otherSheet = wb.addSheet('Other Data');
    expect(otherSheet).toBe(1);
    const store = createSpreadsheetStore();
    const history = new History();
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 100);
    wb.setNumber({ sheet: otherSheet, row: 0, col: 0 }, 7);
    wb.recalc();
    const instance = makeControllerInstance(wb, store, history);
    const request: ConsolidateRequest = {
      sources: ["'Other Data'!A1"],
      destination: "'Other Data'!C1",
      function: 'sum',
    };
    const plan = planMacConsolidate(instance, request);
    expect(plan.ok).toBe(true);
    if (!plan.ok) return;
    expect(plan.value.destination.sheet).toBe(otherSheet);
    expect(plan.value.values).toEqual([{ kind: 'number', value: 7 }]);
    expect(commitMacConsolidate(instance, request, plan.value).ok).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({ kind: 'blank' });
    expect(wb.getValue({ sheet: otherSheet, row: 0, col: 2 })).toEqual({
      kind: 'number',
      value: 7,
    });
    expect(history.undo()).toBe(true);
    expect(wb.getValue({ sheet: otherSheet, row: 0, col: 2 })).toEqual({ kind: 'blank' });
    expect(history.redo()).toBe(true);
    expect(wb.getValue({ sheet: otherSheet, row: 0, col: 2 })).toEqual({
      kind: 'number',
      value: 7,
    });
  });
});
