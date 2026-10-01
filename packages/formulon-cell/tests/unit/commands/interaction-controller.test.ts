import { afterEach, describe, expect, it } from 'vitest';

import { History } from '../../../src/commands/history.js';
import { InteractionController } from '../../../src/commands/interaction-controller.js';
import {
  fixedFormPolicy,
  type InteractionPolicy,
  viewerPolicy,
} from '../../../src/commands/interaction-policy.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const A1 = { sheet: 0, row: 0, col: 0 } as const;
const B1 = { sheet: 0, row: 0, col: 1 } as const;

describe('InteractionController', () => {
  let workbook: WorkbookHandle | undefined;
  let controller: InteractionController | undefined;

  afterEach(() => {
    controller?.dispose();
    workbook?.dispose();
    controller = undefined;
    workbook = undefined;
  });

  async function createController(options: { onChanged?: (result: unknown) => void } = {}) {
    const store = createSpreadsheetStore();
    const currentWorkbook = await WorkbookHandle.createDefault({ preferStub: true });
    workbook = currentWorkbook;
    const history = new History();
    controller = new InteractionController({
      store,
      getWb: () => currentWorkbook,
      history,
      onChanged: options.onChanged,
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
});
