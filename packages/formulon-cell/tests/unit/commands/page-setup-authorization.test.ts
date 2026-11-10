import { afterEach, describe, expect, it } from 'vitest';
import { canExecuteBuiltIn } from '../../../src/commands/built-in-command-policy.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import type { InteractionPolicy } from '../../../src/commands/interaction-policy.js';
import { setPageOrientation } from '../../../src/commands/page-setup.js';
import {
  type PageSetupActionContext,
  pageSetupWrite,
} from '../../../src/commands/page-setup-actions.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachPageSetupDialog } from '../../../src/interact/page-setup-dialog.js';
import { createSpreadsheetStore, getPageSetup } from '../../../src/store/store.js';

const PAGE_SETUP_IDS = [
  'pageSetup',
  'printArea',
  'pageBreaks',
  'sheetBackground',
  'marginsPreset',
  'orientationPreset',
  'paperSizePreset',
  'scalePercent',
  'printTitles',
  'pageLayoutGridlinesPrint',
  'mac.page.orientation.landscape',
  'mac.page.break.insert',
] as const;

const allowAll: InteractionPolicy = { defaultOperation: 'allow' };
const denyPageSetup: InteractionPolicy = {
  defaultOperation: 'allow',
  operations: { pageSetup: false },
};
const denyOrigin = (denied: string): InteractionPolicy => ({
  defaultOperation: 'allow',
  restrict: ({ origin }) => origin !== denied,
});

describe('page-setup writes under the interaction policy', () => {
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
    const ctx: PageSetupActionContext = {
      store,
      history,
      origin: 'ribbon',
      commandId: 'orientationPreset',
    };
    const orientation = () => getPageSetup(store.getState(), 0).orientation;
    const landscape = () =>
      pageSetupWrite(ctx, 0, false, () => {
        setPageOrientation(store, 0, 'landscape', history);
        return true;
      });
    return { store, history, controller, orientation, landscape };
  };

  it('changes page setup under a policy and replays it through undo and redo', async () => {
    const { history, controller, orientation, landscape } = await setup();
    controller.setPolicy(allowAll);
    expect(landscape()).toBe(true);
    expect(orientation()).toBe('landscape');
    expect(history.undo()).toBe(true);
    expect(orientation()).not.toBe('landscape');
    expect(history.redo()).toBe(true);
    expect(orientation()).toBe('landscape');
  });

  it('refuses the write and every page-setup command when the policy denies pageSetup', async () => {
    const { store, history, controller, orientation, landscape } = await setup();
    controller.setPolicy(denyPageSetup);
    expect(landscape()).toBe(false);
    expect(orientation()).not.toBe('landscape');
    expect(history.canUndo()).toBe(false);
    for (const id of PAGE_SETUP_IDS) {
      expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(false);
    }
  });

  for (const [label, policy] of [
    ['without a policy', undefined],
    ['under an allow-all policy', allowAll],
  ] as const satisfies readonly (readonly [string, InteractionPolicy | undefined])[]) {
    it(`allows page setup on a protected sheet ${label}`, async () => {
      const { store, controller, orientation, landscape } = await setup();
      controller.setPolicy(policy);
      setProtectedSheet(store, 0, true);
      for (const id of PAGE_SETUP_IDS) {
        if (id === 'sheetBackground') continue;
        expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(true);
      }
      expect(landscape()).toBe(true);
      expect(orientation()).toBe('landscape');
    });

    it(`refuses the sheet background on a protected sheet ${label}`, async () => {
      const { store, history, controller } = await setup();
      controller.setPolicy(policy);
      setProtectedSheet(store, 0, true);
      for (const id of ['sheetBackground', 'mac.page.background'] as const) {
        expect(canExecuteBuiltIn(store, id, 'ribbon'), id).toMatchObject({
          allowed: false,
          code: 'protected',
        });
      }
      let wrote = false;
      const background = () =>
        pageSetupWrite(
          { store, history, origin: 'ribbon', commandId: 'sheetBackground' },
          0,
          false,
          () => {
            wrote = true;
            return true;
          },
        );
      expect(background()).toBe(false);
      expect(wrote).toBe(false);
      setProtectedSheet(store, 0, false);
      expect(background()).toBe(true);
      expect(wrote).toBe(true);
    });
  }

  it('refuses replay once the policy denies the undo or redo', async () => {
    const { history, controller, orientation, landscape } = await setup();
    controller.setPolicy(allowAll);
    landscape();
    controller.setPolicy(denyOrigin('undo'));
    expect(history.undo()).toBe(false);
    expect(orientation()).toBe('landscape');
    controller.setPolicy(allowAll);
    expect(history.undo()).toBe(true);
    controller.setPolicy(denyPageSetup);
    expect(history.redo()).toBe(false);
    expect(orientation()).not.toBe('landscape');
  });

  it('decides the Page Setup dialog OK under the entry point that opened it', async () => {
    const { store, history, controller } = await setup();
    const host = document.createElement('div');
    document.body.appendChild(host);
    const handle = attachPageSetupDialog({ host, store, history });
    const shell = (): HTMLElement | null => document.querySelector<HTMLElement>('.fc-pgsetup');
    const isOpen = (): boolean => {
      const panel = shell();
      return !!panel && panel.closest('[hidden]') === null;
    };
    const ok = (): void =>
      shell()?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    try {
      controller.setPolicy(denyOrigin('contextMenu'));
      handle.open(undefined, { origin: 'contextMenu', commandId: 'pageSetup' });
      ok();
      expect(isOpen()).toBe(true);
      handle.close();
      handle.open(undefined, { origin: 'ribbon', commandId: 'pageSetup' });
      ok();
      expect(isOpen()).toBe(false);
    } finally {
      handle.detach();
    }
  });
});
