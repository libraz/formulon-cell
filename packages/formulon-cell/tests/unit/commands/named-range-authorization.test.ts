import { afterEach, describe, expect, it } from 'vitest';
import { canExecuteBuiltIn } from '../../../src/commands/built-in-command-policy.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import type { InteractionPolicy } from '../../../src/commands/interaction-policy.js';
import {
  type NamedRangeWriteContext,
  namedRangeWrite,
} from '../../../src/commands/named-range-actions.js';
import { listDefinedNames, upsertDefinedName } from '../../../src/commands/named-ranges.js';
import {
  setProtectedSheet,
  setWorkbookStructureProtected,
} from '../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { attachNamedRangeDialog } from '../../../src/interact/named-range-dialog.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const NAME_IDS = [
  'defineName',
  'namedRanges',
  'mac.formulas.defineName',
  'mac.formulas.namesManager',
  'mac.formulas.createNames.topRow',
] as const;

const allowAll: InteractionPolicy = { defaultOperation: 'allow' };
const denyNames: InteractionPolicy = {
  defaultOperation: 'allow',
  operations: { namedRange: false },
};
const denyOrigin = (denied: string): InteractionPolicy => ({
  defaultOperation: 'allow',
  restrict: ({ origin }) => origin !== denied,
});

describe('defined-name writes under the interaction policy', () => {
  const handles: WorkbookHandle[] = [];
  const controllers: InteractionController[] = [];

  afterEach(() => {
    for (const controller of controllers.splice(0)) controller.dispose();
    for (const handle of handles.splice(0)) handle.dispose();
    document.body.replaceChildren();
  });

  const setup = async () => {
    const workbook = await WorkbookHandle.createDefault();
    handles.push(workbook);
    const store = createSpreadsheetStore();
    const history = new History();
    const controller = new InteractionController({ store, getWb: () => workbook, history });
    controllers.push(controller);
    registerInteractionController(store, controller);
    const ctx: NamedRangeWriteContext = {
      store,
      workbook,
      history,
      origin: 'formulaBar',
      commandId: 'defineName',
    };
    const names = (): string[] => listDefinedNames(workbook).map((entry) => entry.name);
    const define = (name: string) =>
      namedRangeWrite(ctx, () => upsertDefinedName(workbook, name, '=Sheet1!$A$1'));
    return { store, workbook, history, controller, names, define };
  };

  it('defines a name under a policy and replays it through undo and redo', async () => {
    const { history, controller, names, define } = await setup();
    controller.setPolicy(allowAll);
    expect(define('Total')).toMatchObject({ ok: true, value: { ok: true } });
    expect(names()).toEqual(['Total']);
    expect(history.undo()).toBe(true);
    expect(names()).toEqual([]);
    expect(history.redo()).toBe(true);
    expect(names()).toEqual(['Total']);
  });

  it('refuses the write and every name command when the policy denies namedRange', async () => {
    const { store, history, controller, names, define } = await setup();
    controller.setPolicy(denyNames);
    expect(define('Total')).toMatchObject({ ok: false, decision: { code: 'operationDenied' } });
    expect(names()).toEqual([]);
    expect(history.canUndo()).toBe(false);
    for (const id of NAME_IDS) {
      expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(false);
    }
  });

  for (const [label, policy] of [
    ['without a policy', undefined],
    ['under an allow-all policy', allowAll],
  ] as const satisfies readonly (readonly [string, InteractionPolicy | undefined])[]) {
    it(`ignores sheet and workbook-structure protection ${label}`, async () => {
      const { store, controller, names, define } = await setup();
      controller.setPolicy(policy);
      setProtectedSheet(store, 0, true);
      setWorkbookStructureProtected(store, true);
      for (const id of NAME_IDS) {
        expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(true);
      }
      expect(define('Total')).toMatchObject({ ok: true });
      expect(names()).toEqual(['Total']);
    });
  }

  it('refuses replay once the policy denies the undo or redo', async () => {
    const { history, controller, names, define } = await setup();
    controller.setPolicy(allowAll);
    define('Total');
    controller.setPolicy(denyOrigin('undo'));
    expect(history.undo()).toBe(false);
    expect(names()).toEqual(['Total']);
    controller.setPolicy(allowAll);
    expect(history.undo()).toBe(true);
    controller.setPolicy(denyNames);
    expect(history.redo()).toBe(false);
    expect(names()).toEqual([]);
  });

  it('decides the Name Manager under the entry point that opened it', async () => {
    const { store, workbook, history, controller, names } = await setup();
    const host = document.createElement('div');
    document.body.appendChild(host);
    const handle = attachNamedRangeDialog({ host, store, wb: workbook, history, strings: en });
    const newButton = (): HTMLButtonElement | undefined =>
      document.querySelectorAll<HTMLButtonElement>('.fc-namedlg__actions button')[0];
    try {
      controller.setPolicy(denyOrigin('contextMenu'));
      handle.open({ origin: 'contextMenu', commandId: 'defineName' });
      expect(newButton()?.disabled).toBe(true);
      handle.close();
      handle.open({ origin: 'ribbon', commandId: 'namedRanges' });
      expect(newButton()?.disabled).toBe(false);
      handle.close();
      expect(names()).toEqual([]);
    } finally {
      handle.detach();
    }
  });
});
