import { afterEach, describe, expect, it } from 'vitest';
import {
  type AuthorizedIntentFactory,
  authorizedWrite,
} from '../../../src/commands/authorized-write.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import type { InteractionPolicy } from '../../../src/commands/interaction-policy.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const intents: AuthorizedIntentFactory = (origin) => [
  {
    operation: 'format',
    origin,
    commandId: 'test.authorizedWrite',
    effects: [{ kind: 'cells', cells: [{ sheet: 0, row: 0, col: 0 }] }],
  },
];

const denyOrigin = (denied: string): InteractionPolicy => ({
  defaultOperation: 'allow',
  restrict: ({ origin }) => origin !== denied,
});

describe('authorizedWrite', () => {
  const handles: WorkbookHandle[] = [];
  const controllers: InteractionController[] = [];

  afterEach(() => {
    for (const controller of controllers.splice(0)) controller.dispose();
    for (const handle of handles.splice(0)) handle.dispose();
  });

  const setup = async (options: { readonly register?: boolean } = {}) => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    handles.push(wb);
    const store = createSpreadsheetStore();
    const history = new History();
    const log: string[] = [];
    const attach = (): InteractionController => {
      const controller = new InteractionController({ store, getWb: () => wb, history });
      controllers.push(controller);
      registerInteractionController(store, controller);
      return controller;
    };
    const controller = options.register === false ? undefined : attach();
    const recordStep = (label: string): void => {
      log.push(label);
      history.push({
        undo: () => log.push(`undo ${label}`),
        redo: () => log.push(`redo ${label}`),
      });
    };
    const run = (write: () => number) =>
      authorizedWrite({ store, history, intents, origin: 'ribbon', write });
    return { store, history, log, controller, attach, recordStep, run };
  };

  it('runs an allowed write as one history step and returns its value', async () => {
    const { history, log, controller, recordStep, run } = await setup();
    controller?.setPolicy({ defaultOperation: 'allow' });
    const result = run(() => {
      recordStep('a');
      recordStep('b');
      return 7;
    });
    expect(result).toEqual({ ok: true, value: 7 });
    expect(history.undo()).toBe(true);
    expect(history.canUndo()).toBe(false);
    expect(log).toEqual(['a', 'b', 'undo b', 'undo a']);
    expect(history.redo()).toBe(true);
    expect(log.slice(-2)).toEqual(['redo a', 'redo b']);
  });

  it('refuses before writing when the apply intent is denied', async () => {
    const { history, log, controller, recordStep, run } = await setup();
    controller?.setPolicy(denyOrigin('ribbon'));
    const result = run(() => {
      recordStep('a');
      return 1;
    });
    expect(result).toMatchObject({
      ok: false,
      decision: { allowed: false, code: 'operationDenied' },
    });
    expect(log).toEqual([]);
    expect(history.canUndo()).toBe(false);
  });

  it('refuses up front a write whose undo or redo would be denied', async () => {
    const { history, log, controller, recordStep, run } = await setup();
    for (const origin of ['undo', 'redo']) {
      controller?.setPolicy(denyOrigin(origin));
      const result = run(() => {
        recordStep('a');
        return 1;
      });
      expect(result.ok).toBe(false);
    }
    expect(log).toEqual([]);
    expect(history.canUndo()).toBe(false);
  });

  it('aborts what a throwing write recorded and rethrows', async () => {
    const { history, log, recordStep, run } = await setup();
    expect(() =>
      run(() => {
        recordStep('a');
        throw new Error('write failed');
      }),
    ).toThrow('write failed');
    expect(log).toEqual(['a', 'undo a']);
    expect(history.canUndo()).toBe(false);
  });

  it('joins an open transaction instead of opening a second entry', async () => {
    const { history, log, controller, recordStep, run } = await setup();
    controller?.setPolicy({ defaultOperation: 'allow' });
    const token = history.begin({ replayAuthorization: { undo: intents('undo'), redo: [] } });
    recordStep('outer');
    expect(
      run(() => {
        recordStep('inner');
        return 1;
      }),
    ).toEqual({ ok: true, value: 1 });
    history.end(token);
    expect(history.undo()).toBe(true);
    expect(history.canUndo()).toBe(false);
    expect(log).toEqual(['outer', 'inner', 'undo inner', 'undo outer']);
    // The outer transaction's authorization governs: it declared no redo intent.
    expect(history.redo()).toBe(false);
  });

  it('refuses undo and redo replay when the policy changes after the write', async () => {
    const { history, log, controller, recordStep, run } = await setup();
    controller?.setPolicy({ defaultOperation: 'allow' });
    run(() => {
      recordStep('a');
      return 1;
    });
    controller?.setPolicy(denyOrigin('undo'));
    expect(history.undo()).toBe(false);
    controller?.setPolicy({ defaultOperation: 'allow' });
    expect(history.undo()).toBe(true);
    controller?.setPolicy(denyOrigin('redo'));
    expect(history.redo()).toBe(false);
    expect(log).toEqual(['a', 'undo a']);
  });

  it('records the intents on the entry even without a controller', async () => {
    const { history, log, attach, recordStep, run } = await setup({ register: false });
    expect(
      run(() => {
        recordStep('a');
        return 1;
      }),
    ).toEqual({ ok: true, value: 1 });
    attach().setPolicy(denyOrigin('undo'));
    expect(history.undo()).toBe(false);
    expect(log).toEqual(['a']);
  });

  it('keeps the legacy protection and workbook gate on an unrestricted controller', async () => {
    const { history, log, store, recordStep } = await setup();
    const result = authorizedWrite({
      store,
      history,
      intents: (origin) => [{ operation: 'sheet', origin, effects: [{ kind: 'sheet', sheet: 9 }] }],
      origin: 'ribbon',
      write: () => recordStep('a'),
    });
    expect(result).toMatchObject({ ok: false, decision: { code: 'outOfBounds' } });
    expect(log).toEqual([]);
  });
});
