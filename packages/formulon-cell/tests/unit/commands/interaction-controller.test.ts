import { afterEach, describe, expect, it, vi } from 'vitest';

import { History } from '../../../src/commands/history.js';
import { InteractionController } from '../../../src/commands/interaction-controller.js';
import type { OperationIntent } from '../../../src/commands/interaction-policy.js';
import {
  fixedFormPolicy,
  type InteractionPolicy,
  viewerPolicy,
} from '../../../src/commands/interaction-policy.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { addrKey } from '../../../src/engine/address.js';
import { type Addr, type CellValue, type Range, ValueKind } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const A1 = { sheet: 0, row: 0, col: 0 } as const;
const B1 = { sheet: 0, row: 0, col: 1 } as const;
const SUBTOTAL_RANGE = { sheet: 0, r0: 0, c0: 0, r1: 9, c1: 3 } as const;

const subtotalStructuralIntent = (
  operation: 'insertRows' | 'deleteRows',
  origin: OperationIntent['origin'],
  range: Range = SUBTOTAL_RANGE,
  commandId = 'mac.data.subtotal',
  effects: OperationIntent['effects'] = [{ kind: 'workbook' }, { kind: 'range', range }],
): OperationIntent => ({
  operation,
  origin,
  commandId,
  effects,
});

describe('InteractionController', () => {
  let workbook: WorkbookHandle | undefined;
  let controller: InteractionController | undefined;

  afterEach(() => {
    controller?.dispose();
    workbook?.dispose();
    controller = undefined;
    workbook = undefined;
  });

  async function createController(
    options: { onChanged?: (result: unknown) => void; getBounds?: () => Range } = {},
  ) {
    const store = createSpreadsheetStore();
    const currentWorkbook = await WorkbookHandle.createDefault({ preferStub: true });
    workbook = currentWorkbook;
    const history = new History();
    controller = new InteractionController({
      store,
      getWb: () => currentWorkbook,
      history,
      onChanged: options.onChanged,
      getBounds: options.getBounds,
    });
    return { store, history, controller, workbook };
  }

  const edit = (addr: typeof A1, input: string) => ({
    type: 'cellBatch' as const,
    operation: 'valueEdit' as const,
    origin: 'editor' as const,
    changes: [{ addr, input }],
  });

  it('rejects user mutation in a viewer while allowing trusted host reset', async () => {
    const { controller: service, workbook: wb, history } = await createController();
    service.setPolicy(viewerPolicy());

    const rejected = service.execute(edit(A1, '42'));
    expect(rejected.status).toBe('rejected');
    expect(rejected.rejected[0]?.code).toBe('readOnly');
    expect(wb.getValue(A1)).toEqual({ kind: 'blank' });
    expect(history.canUndo()).toBe(false);

    const host = service.applyChanges([{ addr: A1, input: '42' }]);
    expect(host.status).toBe('applied');
    expect(wb.getValue(A1)).toEqual({ kind: 'number', value: 42 });
    expect(history.canUndo()).toBe(false);
  });

  it('supports reject and explicit skip modes for mixed fixed-form paste', async () => {
    const { controller: service, workbook: wb } = await createController();
    service.setPolicy(fixedFormPolicy({ ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }] }));

    const command = {
      type: 'cellBatch' as const,
      operation: 'paste' as const,
      origin: 'clipboard' as const,
      changes: [
        { addr: A1, input: 'left' },
        { addr: B1, input: 'right' },
      ],
    };
    const rejected = service.execute(command);
    expect(rejected.status).toBe('rejected');
    expect(rejected.applied).toEqual([]);
    expect(wb.getValue(A1)).toEqual({ kind: 'blank' });

    const skipped = service.execute({ ...command, denied: 'skipIneligible' });
    expect(skipped.status).toBe('applied');
    expect(skipped.applied).toEqual([A1]);
    expect(skipped.rejected[0]?.addr).toEqual(B1);
    expect(wb.getValue(A1)).toEqual({ kind: 'text', value: 'left' });
    expect(wb.getValue(B1)).toEqual({ kind: 'blank' });
  });

  it('requires formula permission and reauthorizes undo after policy changes', async () => {
    const { controller: service, workbook: wb, history } = await createController();
    service.setPolicy(fixedFormPolicy({ ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }] }));

    const formula = service.execute({
      type: 'cellBatch',
      operation: 'valueEdit',
      origin: 'editor',
      changes: [{ addr: A1, input: '=1+1' }],
    });
    expect(formula.status).toBe('rejected');
    expect(wb.cellFormula(A1)).toBeNull();

    const applied = service.execute(edit(A1, '7'));
    expect(applied.status).toBe('applied');
    expect(history.canUndo()).toBe(true);
    service.setPolicy(viewerPolicy());
    expect(history.undo()).toBe(false);
    expect(history.canUndo()).toBe(true);
    expect(wb.getValue(A1)).toEqual({ kind: 'number', value: 7 });
  });

  it('denies a composite replay atomically when any authorization intent fails', async () => {
    const { controller: service, history } = await createController();
    const callbackLog: string[] = [];
    const valueIntent: OperationIntent = {
      operation: 'valueEdit',
      origin: 'redo',
      commandId: 'composite-value',
      effects: [{ kind: 'workbook' }],
    };
    const formatIntent: OperationIntent = {
      operation: 'format',
      origin: 'undo',
      commandId: 'composite-format',
      effects: [{ kind: 'workbook' }],
    };
    history.begin({
      replayAuthorization: {
        undo: [valueIntent, formatIntent],
        redo: [valueIntent, formatIntent],
      },
    });
    history.push({
      undo: () => callbackLog.push('first-undo'),
      redo: () => callbackLog.push('first-redo'),
    });
    history.push({
      undo: () => callbackLog.push('second-undo'),
      redo: () => callbackLog.push('second-redo'),
    });
    history.end();

    service.setPolicy({ operations: { valueEdit: true, format: false } });

    expect(history.undo()).toBe(false);
    expect(callbackLog).toEqual([]);
    expect(history.canUndo()).toBe(true);
    expect(history.canRedo()).toBe(false);

    service.setPolicy({ operations: { valueEdit: true, format: true } });
    expect(history.undo()).toBe(true);
    expect(callbackLog).toEqual(['second-undo', 'first-undo']);
  });

  it('routes only the scoped Subtotal structural intents through restricted policy', async () => {
    const { controller: service, store } = await createController();
    const seen: string[] = [];
    service.setPolicy({
      defaultOperation: 'deny',
      operations: { insertRows: true, deleteRows: true },
      restrict: ({ intent, addr }) => {
        expect(addr).toBeUndefined();
        seen.push(`${intent.operation}:${intent.origin}`);
        return true;
      },
    });

    expect(service.canExecute(subtotalStructuralIntent('insertRows', 'ribbon'))).toEqual({
      allowed: true,
    });
    expect(service.canExecute(subtotalStructuralIntent('deleteRows', 'undo'))).toEqual({
      allowed: true,
    });
    expect(service.canExecute(subtotalStructuralIntent('insertRows', 'redo'))).toEqual({
      allowed: true,
    });
    expect(seen).toEqual(['insertRows:ribbon', 'deleteRows:undo', 'insertRows:redo']);

    expect(
      service.canExecute(subtotalStructuralIntent('insertRows', 'ribbon', SUBTOTAL_RANGE, 'other')),
    ).toMatchObject({ allowed: false, code: 'unsupported' });
    expect(service.canExecute(subtotalStructuralIntent('insertRows', 'keyboard'))).toMatchObject({
      allowed: false,
      code: 'unsupported',
    });
    expect(
      service.canExecute(
        subtotalStructuralIntent('insertRows', 'ribbon', SUBTOTAL_RANGE, 'mac.data.subtotal', [
          { kind: 'workbook' },
        ]),
      ),
    ).toMatchObject({ allowed: false, code: 'invalid' });

    service.setPolicy({ defaultOperation: 'allow' });
    expect(service.canExecute(subtotalStructuralIntent('insertRows', 'ribbon'))).toMatchObject({
      allowed: false,
      code: 'operationDenied',
    });

    service.setPolicy({ readOnly: true, operations: { insertRows: true } });
    expect(service.canExecute(subtotalStructuralIntent('insertRows', 'ribbon'))).toMatchObject({
      allowed: false,
      code: 'readOnly',
    });

    service.setPolicy({
      operations: { insertRows: true },
      editable: { ranges: [SUBTOTAL_RANGE] },
      restrict: () => true,
    });
    expect(service.canExecute(subtotalStructuralIntent('insertRows', 'ribbon'))).toMatchObject({
      allowed: false,
      code: 'operationDenied',
    });

    service.setPolicy({ operations: { insertRows: true }, restrict: () => true });
    setProtectedSheet(store, 0, true);
    expect(service.canExecute(subtotalStructuralIntent('insertRows', 'ribbon'))).toMatchObject({
      allowed: false,
      code: 'protected',
    });
    setProtectedSheet(store, 0, false);
  });

  it('pins Subtotal structural scope to the active worksheet and full bounds', async () => {
    const partial = await createController({
      getBounds: () => ({ sheet: 0, r0: 0, c0: 0, r1: 9, c1: 2 }),
    });
    partial.controller.setPolicy({ operations: { insertRows: true } });
    expect(
      partial.controller.canExecute(subtotalStructuralIntent('insertRows', 'ribbon')),
    ).toMatchObject({ allowed: false, code: 'outOfBounds' });
    partial.controller.dispose();
    partial.workbook.dispose();
    controller = undefined;
    workbook = undefined;

    const { controller: service, store } = await createController();
    service.setPolicy({ operations: { insertRows: true } });
    mutators.setSheetIndex(store, 1);
    expect(service.canExecute(subtotalStructuralIntent('insertRows', 'ribbon'))).toMatchObject({
      allowed: false,
      code: 'outOfBounds',
    });
  });

  it('rechecks recognized Subtotal replay after policy clear without affecting ordinary history', async () => {
    const { controller: service, store, history } = await createController();
    const callbacks: string[] = [];
    service.setPolicy({ operations: { insertRows: true, deleteRows: true } });
    const undoIntent = subtotalStructuralIntent('deleteRows', 'undo');
    const redoIntent = subtotalStructuralIntent('insertRows', 'redo');
    history.begin({ replayAuthorization: { undo: [undoIntent], redo: [redoIntent] } });
    history.push({
      undo: () => callbacks.push('subtotal-undo'),
      redo: () => callbacks.push('subtotal-redo'),
    });
    history.end();

    service.setPolicy(undefined);
    mutators.setSheetIndex(store, 1);
    expect(history.undo()).toBe(false);
    expect(callbacks).toEqual([]);
    mutators.setSheetIndex(store, 0);
    setProtectedSheet(store, 0, true);
    expect(history.undo()).toBe(false);
    expect(callbacks).toEqual([]);
    setProtectedSheet(store, 0, false);
    expect(history.undo()).toBe(true);
    expect(callbacks).toEqual(['subtotal-undo']);

    history.push({
      undo: () => callbacks.push('ordinary-undo'),
      redo: () => callbacks.push('ordinary-redo'),
    });
    expect(history.undo()).toBe(true);
    expect(callbacks).toEqual(['subtotal-undo', 'ordinary-undo']);
  });

  it('rejects an empty composite replay authorization bundle', async () => {
    const { history } = await createController();
    let callbacks = 0;
    history.begin({ replayAuthorization: { undo: [], redo: [] } });
    history.push({
      undo: () => {
        callbacks += 1;
      },
      redo: () => {
        callbacks += 1;
      },
    });
    history.end();

    expect(history.undo()).toBe(false);
    expect(callbacks).toBe(0);
    expect(history.canUndo()).toBe(true);
  });

  it('authorizes format intents against editable cells and protection', async () => {
    const { store, controller: service } = await createController();
    const policy = fixedFormPolicy({ ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }] });
    service.setPolicy({ ...policy, operations: { ...policy.operations, format: true } });
    const intent = (addr: Addr) => ({
      operation: 'format' as const,
      origin: 'ribbon' as const,
      commandId: 'bold',
      effects: [{ kind: 'cells' as const, cells: [addr] }],
    });

    expect(service.canExecute(intent(A1))).toEqual({ allowed: true });
    expect(service.canExecute(intent(B1))).toMatchObject({
      allowed: false,
      code: 'cellIneligible',
      addr: B1,
    });

    setProtectedSheet(store, 0, true);
    expect(service.canExecute(intent(A1))).toMatchObject({
      allowed: false,
      code: 'protected',
      addr: A1,
    });
  });

  it('keeps pending editor format out of foreign-origin writes', async () => {
    const { store, controller: service, workbook: wb, history } = await createController();
    const pending = {
      addr: A1,
      format: { bold: true, borders: { bottom: { style: 'thin' as const } } },
    };
    mutators.setPendingFormat(store, pending);

    const result = service.execute({
      type: 'cellBatch',
      operation: 'valueEdit',
      origin: 'keyboard',
      changes: [{ addr: A1, input: '42' }],
    });

    expect(result.status).toBe('applied');
    expect(wb.getValue(A1)).toEqual({ kind: 'number', value: 42 });
    expect(store.getState().format.formats.get(addrKey(A1))).toBeUndefined();
    expect(store.getState().ui.pendingFormat).toEqual(pending);
    expect(history.canUndo()).toBe(true);
  });

  it('denies a pending format before a value write when the format operation is restricted', async () => {
    const { store, controller: service, workbook: wb, history } = await createController();
    mutators.setPendingFormat(store, { addr: A1, format: { italic: true } });
    service.setPolicy({
      operations: { valueEdit: true, format: true },
      restrict: ({ operation, addr }) => operation !== 'format' || addr === undefined,
    });

    const result = service.execute({
      type: 'cellBatch',
      operation: 'valueEdit',
      origin: 'editor',
      changes: [{ addr: A1, input: '42' }],
    });

    expect(result.status).toBe('rejected');
    expect(result.rejected[0]?.code).toBe('operationDenied');
    expect(wb.getValue(A1)).toEqual({ kind: 'blank' });
    expect(store.getState().format.formats.get(addrKey(A1))).toBeUndefined();
    expect(store.getState().ui.pendingFormat).toEqual({ addr: A1, format: { italic: true } });
    expect(history.canUndo()).toBe(false);
  });

  it.each(['editor', 'formulaBar', 'keyboard'] as const)(
    'keeps implicit percent coercion working for %s without pending format',
    async (origin) => {
      const { store, controller: service, workbook: wb, history } = await createController();
      service.setPolicy(fixedFormPolicy({ ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }] }));

      const result = service.execute({
        type: 'cellBatch',
        operation: 'valueEdit',
        origin,
        changes: [{ addr: A1, input: '5%' }],
      });

      expect(result.status).toBe('applied');
      expect(wb.getValue(A1)).toEqual({ kind: 'number', value: 0.05 });
      expect(store.getState().format.formats.get(addrKey(A1))?.numFmt).toEqual({
        kind: 'percent',
        decimals: 0,
      });
      expect(history.undo()).toBe(true);
      expect(wb.getValue(A1)).toEqual({ kind: 'blank' });
      expect(store.getState().format.formats.get(addrKey(A1))).toBeUndefined();
      expect(history.redo()).toBe(true);
      expect(wb.getValue(A1)).toEqual({ kind: 'number', value: 0.05 });
      expect(store.getState().format.formats.get(addrKey(A1))?.numFmt).toEqual({
        kind: 'percent',
        decimals: 0,
      });
    },
  );

  it('keeps implicit percent coercion working through the trusted API path', async () => {
    const { store, controller: service, workbook: wb, history } = await createController();
    service.setPolicy(fixedFormPolicy({ ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }] }));

    const result = service.applyChanges([{ addr: A1, input: '5%' }], {
      history: 'record',
      origin: 'instanceApi',
    });

    expect(result.status).toBe('applied');
    expect(wb.getValue(A1)).toEqual({ kind: 'number', value: 0.05 });
    expect(store.getState().format.formats.get(addrKey(A1))?.numFmt).toEqual({
      kind: 'percent',
      decimals: 0,
    });
    expect(history.undo()).toBe(true);
    expect(wb.getValue(A1)).toEqual({ kind: 'blank' });
    expect(history.redo()).toBe(true);
    expect(wb.getValue(A1)).toEqual({ kind: 'number', value: 0.05 });
  });

  it('deduplicates every coordinate in a merged format authorization effect', async () => {
    const { store, controller: service } = await createController();
    mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    const requested = [A1, B1, { sheet: 0, row: 1, col: 0 }, { sheet: 0, row: 1, col: 1 }];
    let checks = 0;
    service.setPolicy({
      operations: { format: true },
      editable: {
        predicate: ({ operation }) => {
          if (operation === 'format') checks += 1;
          return true;
        },
      },
    });

    expect(
      service.canExecute({
        operation: 'format',
        origin: 'ribbon',
        commandId: 'bold',
        effects: [{ kind: 'cells', cells: requested }],
      }),
    ).toEqual({ allowed: true });
    expect(checks).toBe(4);
  });

  it('applies navigation bounds to format authorization', async () => {
    const store = createSpreadsheetStore();
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    workbook = wb;
    const history = new History();
    const service = new InteractionController({
      store,
      getWb: () => wb,
      history,
      getBounds: () => ({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }),
    });
    controller = service;
    const policy = fixedFormPolicy({
      ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 }],
    });
    service.setPolicy({ ...policy, operations: { ...policy.operations, format: true } });

    expect(
      service.canExecute({
        operation: 'format',
        origin: 'ribbon',
        commandId: 'bold',
        effects: [{ kind: 'cells', cells: [{ sheet: 0, row: 1, col: 0 }] }],
      }),
    ).toMatchObject({ allowed: false, code: 'outOfBounds' });
  });

  it('checks formula permission on each formula cell in a mixed batch', async () => {
    const { controller: service, workbook: wb } = await createController();
    service.setPolicy({
      defaultOperation: 'allow',
      editable: {
        predicate: ({ addr, operation }) =>
          !(
            operation === 'formulaEdit' &&
            addr.sheet === A1.sheet &&
            addr.row === A1.row &&
            addr.col === A1.col
          ),
      },
    });
    const result = service.execute({
      type: 'cellBatch',
      operation: 'paste',
      origin: 'clipboard',
      changes: [
        { addr: A1, input: '=1+1' },
        { addr: B1, input: '2' },
      ],
    });
    expect(result.status).toBe('rejected');
    expect(result.rejected[0]?.code).toBe('cellIneligible');
    expect(wb.cellFormula(A1)).toBeNull();
    expect(wb.getValue(B1)).toEqual({ kind: 'blank' });
  });

  it('preflights the inverse formula and lets trusted updates cross navigation bounds', async () => {
    const store = createSpreadsheetStore();
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    wb.setFormula(A1, '=1+1');
    const history = new History();
    const service = new InteractionController({
      store,
      getWb: () => wb,
      history,
      getBounds: () => ({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }),
    });
    service.setPolicy(fixedFormPolicy({ ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 }] }));
    const deniedRecord = service.applyChanges([{ addr: A1, input: '7' }], { history: 'record' });
    expect(deniedRecord.status).toBe('rejected');
    expect(wb.cellFormula(A1)).toBe('=1+1');
    const hiddenHost = service.applyChanges([{ addr: { sheet: 0, row: 1, col: 0 }, input: '8' }]);
    expect(hiddenHost.status).toBe('applied');
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'number', value: 8 });
    service.dispose();
    wb.dispose();
  });

  it('rolls back a failed atomic engine patch without emitting value events', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const events: unknown[] = [];
    const unsubscribe = wb.subscribe((event) => events.push(event));
    const raw = (
      wb as unknown as {
        wb: {
          setNumber: (
            sheet: number,
            row: number,
            col: number,
            value: number,
          ) => {
            ok: boolean;
            message?: string;
          };
        };
      }
    ).wb;
    const original = raw.setNumber.bind(raw);
    let writes = 0;
    raw.setNumber = (sheet, row, col, value) => {
      writes += 1;
      if (writes === 2) return { ok: false, message: 'forced failure' };
      return original(sheet, row, col, value);
    };

    expect(() =>
      wb.applyCellPatchAtomic([
        { addr: A1, value: { kind: 'number', value: 1 } },
        { addr: B1, value: { kind: 'number', value: 2 } },
      ]),
    ).toThrow('forced failure');
    expect(wb.getValue(A1)).toEqual({ kind: 'blank' });
    expect(wb.getValue(B1)).toEqual({ kind: 'blank' });
    expect(events).toEqual([]);
    unsubscribe();
    wb.dispose();
  });

  it('isolates throwing host and renderer observers after a committed batch', async () => {
    let notified = 0;
    const { controller: service, workbook: wb } = await createController({
      onChanged: () => {
        throw new Error('host observer');
      },
    });
    service.subscribe(() => {
      throw new Error('stale renderer');
    });
    service.subscribe(() => {
      notified += 1;
    });
    service.setPolicy({ operations: { valueEdit: true } });
    notified = 0;
    const result = service.execute(edit(A1, '3'));
    expect(result.status).toBe('applied');
    expect(notified).toBe(1);
    expect(wb.getValue(A1)).toEqual({ kind: 'number', value: 3 });
  });

  it('isolates throwing workbook and history observers after commit', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const workbookEvents: unknown[] = [];
    wb.subscribe(() => {
      throw new Error('stale workbook observer');
    });
    wb.subscribe((event) => workbookEvents.push(event));
    expect(() =>
      wb.applyCellPatchAtomic([{ addr: A1, value: { kind: 'number', value: 5 } }]),
    ).not.toThrow();
    expect(workbookEvents.length).toBeGreaterThan(0);

    const history = new History();
    let historyNotifications = 0;
    history.subscribe(() => {
      throw new Error('stale history observer');
    });
    history.subscribe(() => {
      historyNotifications += 1;
    });
    expect(() =>
      history.push({
        undo: () => undefined,
        redo: () => undefined,
      }),
    ).not.toThrow();
    expect(history.canUndo()).toBe(true);
    expect(historyNotifications).toBe(1);
    wb.dispose();
  });

  it('repairs dependent formula caches after a failed recalculation', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    wb.setNumber(A1, 1);
    wb.setFormula(B1, '=A1+1');
    expect(wb.getValue(B1)).toEqual({ kind: 'number', value: 2 });
    const raw = (
      wb as unknown as {
        wb: {
          recalc: () => { ok: boolean; message?: string };
        };
      }
    ).wb;
    const original = raw.recalc.bind(raw);
    let recalculations = 0;
    raw.recalc = () => {
      recalculations += 1;
      if (recalculations === 1) return { ok: false, message: 'forced recalc failure' };
      return original();
    };
    expect(() =>
      wb.applyCellPatchAtomic([{ addr: A1, value: { kind: 'number', value: 9 } }]),
    ).toThrow('forced recalc failure');
    expect(wb.getValue(A1)).toEqual({ kind: 'number', value: 1 });
    expect(wb.getValue(B1)).toEqual({ kind: 'number', value: 2 });
    expect(recalculations).toBe(2);
    wb.dispose();
  });
  it('fails closed when a JavaScript host supplies an async restriction hook', async () => {
    const { controller: service, workbook: wb } = await createController();
    service.setPolicy({
      operations: { valueEdit: true },
      restrict: (() => Promise.reject(new Error('async hook failure'))) as unknown as NonNullable<
        InteractionPolicy['restrict']
      >,
    });
    const result = service.execute(edit(A1, 'Disallowed'));
    expect(result.status).toBe('rejected');
    expect(result.rejected[0]?.reason).toBe('restriction hook must be synchronous');
    expect(wb.getValue(A1).kind).toBe('blank');
  });

  it('leaves an existing history guard installed for ephemeral controllers', async () => {
    const store = createSpreadsheetStore();
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const history = new History();
    history.setGuard(() => false);
    const ephemeral = new InteractionController({
      store,
      getWb: () => wb,
      history,
      manageHistoryGuard: false,
    });

    ephemeral.setPolicy(viewerPolicy());
    history.push({ undo: () => undefined, redo: () => undefined });
    expect(history.undo()).toBe(false);
    ephemeral.setPolicy(undefined);
    history.push({ undo: () => undefined, redo: () => undefined });
    expect(history.undo()).toBe(false);
    ephemeral.dispose();
    expect(history.undo()).toBe(false);
    wb.dispose();
  });

  it('keeps unrestricted clear batches above the authorization cap atomic', () => {
    const count = 100_001;
    const changes = Array.from({ length: count }, (_, row) => ({
      addr: { sheet: 0, row, col: 0 },
      value: { kind: 'blank' as const },
    }));

    const makeAtomicWorkbook = (): {
      workbook: WorkbookHandle;
      valueAt: (addr: Addr) => CellValue;
      atomicCalls: () => number;
      formulaScanCalls: () => number;
    } => {
      const values = new Map<string, { value: CellValue; formula: string | null }>();
      for (let row = 0; row < count; row += 1) {
        values.set(`0:${row}:0`, { value: { kind: 'number', value: row }, formula: null });
      }
      let atomicCalls = 0;
      let formulaScanCalls = 0;
      const snapshot = (a: Addr): { addr: Addr; value: CellValue; formula: string | null } => {
        const current = values.get(`0:${a.row}:0`);
        return {
          addr: { ...a },
          value: current?.value ?? { kind: 'blank' },
          formula: current?.formula ?? null,
        };
      };
      const workbook = {
        sheetCount: 1,
        cellFormula: (_a: Addr): string | null => {
          throw new Error('scalar formula lookup is not supported by this workbook');
        },
        cellFormulas: (addrs: readonly Addr[]): ReadonlyMap<string, string | null> => {
          formulaScanCalls += 1;
          return new Map(
            addrs.map((addr) => [addrKey(addr), values.get(`0:${addr.row}:0`)?.formula ?? null]),
          );
        },
        applyCellPatchAtomic: (
          patches: readonly {
            addr: Addr;
            value: CellValue;
            formula?: string | null;
          }[],
        ) => {
          atomicCalls += 1;
          const before = patches.map((patch) => snapshot(patch.addr));
          const changed: Addr[] = [];
          for (const patch of patches) {
            const key = `0:${patch.addr.row}:0`;
            const next = { value: patch.value, formula: patch.formula ?? null };
            const current = values.get(key);
            if (current?.value.kind !== next.value.kind || current?.formula !== next.formula) {
              changed.push({ ...patch.addr });
            }
            values.set(key, next);
          }
          const after = patches.map((patch) => snapshot(patch.addr));
          return { before, after, changed };
        },
      } as unknown as WorkbookHandle;
      return {
        workbook,
        valueAt: (a) => snapshot(a).value,
        atomicCalls: () => atomicCalls,
        formulaScanCalls: () => formulaScanCalls,
      };
    };

    const unrestricted = makeAtomicWorkbook();
    const unrestrictedStore = createSpreadsheetStore();
    const unrestrictedHistory = new History();
    const unrestrictedController = new InteractionController({
      store: unrestrictedStore,
      getWb: () => unrestricted.workbook,
      history: unrestrictedHistory,
    });
    const command = {
      type: 'cellBatch' as const,
      operation: 'clear' as const,
      origin: 'keyboard' as const,
      changes,
    };
    const applied = unrestrictedController.execute(command);
    expect(applied.status).toBe('applied');
    expect(applied.applied).toHaveLength(count);
    expect(unrestricted.atomicCalls()).toBe(1);
    expect(unrestricted.formulaScanCalls()).toBe(1);
    expect(unrestricted.valueAt({ sheet: 0, row: count - 1, col: 0 })).toEqual({ kind: 'blank' });
    expect(unrestrictedHistory.undo()).toBe(true);
    expect(unrestricted.atomicCalls()).toBe(2);
    expect(unrestricted.formulaScanCalls()).toBe(1);
    expect(unrestricted.valueAt({ sheet: 0, row: count - 1, col: 0 })).toEqual({
      kind: 'number',
      value: count - 1,
    });
    expect(unrestrictedHistory.redo()).toBe(true);
    expect(unrestricted.atomicCalls()).toBe(3);
    expect(unrestricted.formulaScanCalls()).toBe(1);
    expect(unrestricted.valueAt({ sheet: 0, row: count - 1, col: 0 })).toEqual({ kind: 'blank' });
    unrestrictedController.dispose();

    const restricted = makeAtomicWorkbook();
    const restrictedStore = createSpreadsheetStore();
    const restrictedHistory = new History();
    const restrictedController = new InteractionController({
      store: restrictedStore,
      getWb: () => restricted.workbook,
      history: restrictedHistory,
    });
    restrictedController.setPolicy({ defaultOperation: 'allow' });
    const rejected = restrictedController.execute(command);
    expect(rejected.status).toBe('rejected');
    expect(rejected.rejected[0]?.code).toBe('unsupported');
    expect(restricted.atomicCalls()).toBe(0);
    expect(restricted.formulaScanCalls()).toBe(1);
    expect(restrictedHistory.canUndo()).toBe(false);
    restrictedController.dispose();

    const denied = makeAtomicWorkbook();
    const deniedController = new InteractionController({
      store: createSpreadsheetStore(),
      getWb: () => denied.workbook,
      history: new History(),
    });
    deniedController.setPolicy(viewerPolicy());
    const policyRejected = deniedController.execute(command);
    expect(policyRejected.status).toBe('rejected');
    expect(policyRejected.rejected[0]?.code).toBe('readOnly');
    expect(denied.atomicCalls()).toBe(0);
    expect(denied.formulaScanCalls()).toBe(0);
    deniedController.dispose();
  });

  it('stops exposing a workbook whose rollback recalculation also fails', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    wb.setNumber(A1, 1);
    const engine = (wb as unknown as { wb: { recalc(): { ok: boolean; message?: string } } }).wb;
    const original = engine.recalc.bind(engine);
    engine.recalc = () => ({ ok: false, message: 'unrecoverable recalc failure' });
    try {
      expect(() =>
        wb.applyCellPatchAtomic([{ addr: A1, value: { kind: 'number', value: 9 } }]),
      ).toThrow('rollback failed');
      expect(() => wb.getValue(A1)).toThrow('rollback failed');
      expect(() => wb.setNumber(A1, 2)).toThrow('rollback failed');
    } finally {
      engine.recalc = original;
      wb.dispose();
    }
  });

  describe('data-validation parity with the direct write path', () => {
    const valueChange = (value: CellValue) => ({
      type: 'cellBatch' as const,
      operation: 'valueEdit' as const,
      origin: 'editor' as const,
      changes: [{ addr: A1, value }],
    });

    it('rejects an input change that fails a custom rule', async () => {
      const { store, controller: service, workbook: wb } = await createController();
      mutators.setCellFormat(store, A1, {
        validation: { kind: 'custom', formula: '=A1>0', allowBlank: true, errorStyle: 'stop' },
      });
      vi.spyOn(wb, 'evalFormula').mockReturnValue({
        status: { status: 0 },
        value: { kind: ValueKind.Number, number: 0 },
      } as unknown as ReturnType<typeof wb.evalFormula>);

      const result = service.execute(edit(A1, '-5'));
      expect(result.status).toBe('rejected');
      expect(result.rejected[0]?.code).toBe('invalid');
    });

    it('lets a warning-severity rule through on a value change', async () => {
      const { store, controller: service, workbook: wb } = await createController();
      mutators.setCellFormat(store, A1, {
        validation: { kind: 'whole', op: 'between', a: 1, b: 10, errorStyle: 'warning' },
      });

      const result = service.execute(valueChange({ kind: 'number', value: 99 }));
      expect(result.status).not.toBe('rejected');
      expect(wb.getValue(A1)).toEqual({ kind: 'number', value: 99 });
    });

    it('records a stop rule silently when its error alert is off', async () => {
      const { store, controller: service, workbook: wb } = await createController();
      mutators.setCellFormat(store, A1, {
        validation: {
          kind: 'whole',
          op: 'between',
          a: 1,
          b: 10,
          errorStyle: 'stop',
          showErrorMessage: false,
        },
      });

      const result = service.execute(valueChange({ kind: 'number', value: 99 }));
      expect(result.status).not.toBe('rejected');
      expect(wb.getValue(A1)).toEqual({ kind: 'number', value: 99 });
    });

    it('still rejects a stop-severity value change', async () => {
      const { store, controller: service } = await createController();
      mutators.setCellFormat(store, A1, {
        validation: { kind: 'whole', op: 'between', a: 1, b: 10, errorStyle: 'stop' },
      });

      const result = service.execute(valueChange({ kind: 'number', value: 99 }));
      expect(result.status).toBe('rejected');
      expect(result.rejected[0]?.code).toBe('invalid');
    });
  });
});
