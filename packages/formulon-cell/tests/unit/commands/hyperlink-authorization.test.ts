import { afterEach, describe, expect, it } from 'vitest';
import { canExecuteBuiltIn } from '../../../src/commands/built-in-command-policy.js';
import { History } from '../../../src/commands/history.js';
import {
  type HyperlinkActionContext,
  setHyperlinkAction,
} from '../../../src/commands/hyperlink-actions.js';
import { hyperlinkAt } from '../../../src/commands/hyperlinks.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import type { InteractionPolicy } from '../../../src/commands/interaction-policy.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { attachHyperlinkDialog } from '../../../src/interact/hyperlink-dialog.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const A1 = { sheet: 0, row: 0, col: 0 };
const URL = 'https://example.com/';
const HYPERLINK_IDS = ['hyperlinkInsert', 'insertHyperlink', 'mac.insert.link'] as const;

const allowAll: InteractionPolicy = { defaultOperation: 'allow' };
const denyHyperlink: InteractionPolicy = {
  defaultOperation: 'allow',
  operations: { hyperlink: false },
};
const denyOrigin = (denied: string): InteractionPolicy => ({
  defaultOperation: 'allow',
  restrict: ({ origin }) => origin !== denied,
});

describe('hyperlink writes under the interaction policy', () => {
  const handles: WorkbookHandle[] = [];
  const controllers: InteractionController[] = [];

  afterEach(() => {
    for (const controller of controllers.splice(0)) controller.dispose();
    for (const handle of handles.splice(0)) handle.dispose();
    document.body.replaceChildren();
  });

  const setup = async () => {
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    handles.push(workbook);
    const store = createSpreadsheetStore();
    const history = new History();
    const controller = new InteractionController({ store, getWb: () => workbook, history });
    controllers.push(controller);
    registerInteractionController(store, controller);
    mutators.setActive(store, A1);
    const ctx: HyperlinkActionContext = {
      store,
      history,
      workbook,
      origin: 'ribbon',
      commandId: 'hyperlinkInsert',
    };
    const link = () => hyperlinkAt(store.getState(), A1);
    return { store, workbook, history, controller, ctx, link };
  };

  it('sets and removes a link under a policy and replays through undo and redo', async () => {
    const { history, controller, ctx, link } = await setup();
    controller.setPolicy(allowAll);
    expect(setHyperlinkAction(ctx, A1, URL)).toBe(true);
    expect(link()).toBe(URL);
    expect(history.undo()).toBe(true);
    expect(link()).toBeNull();
    expect(history.redo()).toBe(true);
    expect(setHyperlinkAction(ctx, A1, undefined)).toBe(true);
    expect(link()).toBeNull();
  });

  it('refuses the write and the hyperlink commands when the policy denies hyperlink', async () => {
    const { store, history, controller, ctx, link } = await setup();
    controller.setPolicy(denyHyperlink);
    expect(setHyperlinkAction(ctx, A1, URL)).toBe(false);
    expect(link()).toBeNull();
    expect(history.canUndo()).toBe(false);
    for (const id of HYPERLINK_IDS) {
      expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(false);
    }
  });

  for (const [label, policy] of [
    ['without a policy', undefined],
    ['under an allow-all policy', allowAll],
  ] as const satisfies readonly (readonly [string, InteractionPolicy | undefined])[]) {
    it(`needs Insert hyperlinks and an unlocked cell on a protected sheet ${label}`, async () => {
      const { store, controller, ctx, link } = await setup();
      controller.setPolicy(policy);
      mutators.setCellFormat(store, A1, { locked: false });
      setProtectedSheet(store, 0, true);
      expect(canExecuteBuiltIn(store, 'hyperlinkInsert', 'ribbon').allowed).toBe(false);
      expect(setHyperlinkAction(ctx, A1, URL)).toBe(false);

      setProtectedSheet(store, 0, true, { permissions: { insertHyperlinks: true } });
      expect(canExecuteBuiltIn(store, 'hyperlinkInsert', 'ribbon').allowed).toBe(true);
      expect(setHyperlinkAction(ctx, A1, URL)).toBe(true);
      expect(link()).toBe(URL);

      const B1 = { sheet: 0, row: 0, col: 1 };
      expect(setHyperlinkAction(ctx, B1, URL)).toBe(false);
      expect(hyperlinkAt(store.getState(), B1)).toBeNull();
    });
  }

  it('refuses replay once the policy denies the undo or redo', async () => {
    const { history, controller, ctx, link } = await setup();
    controller.setPolicy(allowAll);
    setHyperlinkAction(ctx, A1, URL);
    controller.setPolicy(denyOrigin('undo'));
    expect(history.undo()).toBe(false);
    expect(link()).toBe(URL);
    controller.setPolicy(allowAll);
    expect(history.undo()).toBe(true);
    controller.setPolicy(denyHyperlink);
    expect(history.redo()).toBe(false);
    expect(link()).toBeNull();
  });

  it('decides the Insert Hyperlink dialog under the entry point that opened it', async () => {
    const { store, workbook, history, controller, link } = await setup();
    const host = document.createElement('div');
    document.body.appendChild(host);
    const handle = attachHyperlinkDialog({
      host,
      store,
      history,
      strings: en,
      getWb: () => workbook,
    });
    const submit = (): void => {
      const panel = document.querySelector<HTMLElement>('.fc-hldlg');
      const url = panel?.querySelector<HTMLInputElement>('input.fc-fmtdlg__input');
      if (url) url.value = URL;
      panel?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    };
    try {
      controller.setPolicy(denyOrigin('contextMenu'));
      handle.open({ origin: 'contextMenu', commandId: 'insertHyperlink' });
      submit();
      expect(link()).toBeNull();
      handle.close();
      handle.open({ origin: 'keyboard', commandId: 'hyperlinkInsert' });
      submit();
      expect(link()).toBe(URL);
    } finally {
      handle.detach();
    }
  });
});
