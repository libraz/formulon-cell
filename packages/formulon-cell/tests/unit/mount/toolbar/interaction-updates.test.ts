import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  InteractionController,
  registerInteractionController,
} from '../../../../src/commands/interaction-controller.js';
import { addrKey } from '../../../../src/engine/address.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { stubHelpers } from './fixtures.js';

const A1 = { sheet: 0, row: 0, col: 0 } as const;

const countDomCreations = (run: () => void): number => {
  const createElement = vi.spyOn(document, 'createElement');
  const createElementNS = vi.spyOn(document, 'createElementNS');
  try {
    run();
    return createElement.mock.calls.length + createElementNS.mock.calls.length;
  } finally {
    createElement.mockRestore();
    createElementNS.mockRestore();
  }
};

describe('mounted toolbar interaction updates', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;
  let toolbar: ReturnType<typeof Spreadsheet.mountToolbar> | undefined;

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    host = document.createElement('div');
    document.body.appendChild(host);
    toolbar = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
  });

  afterEach(() => {
    toolbar?.dispose();
    toolbar = undefined;
    sheet.dispose();
    host.remove();
  });

  const executeInput = (input: string) =>
    sheet.instance.commands.execute({
      type: 'cellBatch',
      operation: 'valueEdit',
      origin: 'editor',
      changes: [{ addr: A1, input }],
    });

  it('keeps ribbon nodes while an edit projects the committed value and format state', () => {
    sheet.instance.store.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        active: A1,
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      },
      ui: {
        ...state.ui,
        pendingFormat: { addr: A1, format: { bold: true } },
      },
    }));

    const bold = host.querySelector<HTMLButtonElement>('[data-ribbon-command="bold"]');
    const ribbonButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command]');
    expect(bold?.getAttribute('aria-pressed')).toBe('true');
    expect(bold).toBeTruthy();
    expect(ribbonButton).toBeTruthy();

    let result: ReturnType<typeof executeInput> | undefined;
    const created = countDomCreations(() => {
      result = executeInput('7');
    });

    expect(result?.status).toBe('applied');
    expect(sheet.workbook.getValue(A1)).toEqual({ kind: 'number', value: 7 });
    expect(sheet.instance.store.getState().format.formats.get(addrKey(A1))).toMatchObject({
      bold: true,
    });
    expect(host.querySelector('[data-ribbon-command="bold"]')).toBe(bold);
    expect(host.querySelector('[data-ribbon-command]')).toBe(ribbonButton);
    expect(bold?.getAttribute('aria-pressed')).toBe('true');
    expect(created).toBe(0);
  });

  it('projects ordinary no-op and rejected batches without rebuilding the ribbon', () => {
    expect(executeInput('7').status).toBe('applied');
    let bold = host.querySelector<HTMLButtonElement>('[data-ribbon-command="bold"]');
    let created = countDomCreations(() => {
      expect(executeInput('7').status).toBe('noop');
    });
    expect(host.querySelector('[data-ribbon-command="bold"]')).toBe(bold);
    expect(created).toBe(0);

    sheet.instance.setPolicy({ operations: { valueEdit: false } });
    bold = host.querySelector<HTMLButtonElement>('[data-ribbon-command="bold"]');
    created = countDomCreations(() => {
      expect(executeInput('9').status).toBe('rejected');
    });
    expect(host.querySelector('[data-ribbon-command="bold"]')).toBe(bold);
    expect(created).toBe(0);
  });

  it('does not treat a nested ordinary no-op as a policy update', () => {
    toolbar?.dispose();
    let nestedNoopStatus: string | undefined;
    const unsubscribeBeforeToolbar = sheet.instance.commands.subscribe((result) => {
      if (result.status === 'applied' && nestedNoopStatus === undefined) {
        nestedNoopStatus = executeInput('7').status;
      }
    });
    toolbar = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    const bold = host.querySelector<HTMLButtonElement>('[data-ribbon-command="bold"]');

    const created = countDomCreations(() => {
      expect(executeInput('7').status).toBe('applied');
    });
    unsubscribeBeforeToolbar();

    expect(nestedNoopStatus).toBe('noop');
    expect(host.querySelector('[data-ribbon-command="bold"]')).toBe(bold);
    expect(created).toBe(0);
  });

  it('renders once for a nested same-object policy reapplication', () => {
    toolbar?.dispose();
    const policy = { defaultOperation: 'allow' as const, operations: { format: true } };
    sheet.instance.setPolicy(policy);
    let policyWasReapplied = false;
    const unsubscribeBeforeToolbar = sheet.instance.commands.subscribe((result) => {
      if (result.status === 'applied' && !policyWasReapplied) {
        policyWasReapplied = true;
        policy.operations.format = false;
        sheet.instance.setPolicy(policy);
      }
    });
    toolbar = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    const singleRenderCost = countDomCreations(() => toolbar?.rerender());

    const created = countDomCreations(() => {
      expect(executeInput('7').status).toBe('applied');
    });
    unsubscribeBeforeToolbar();

    expect(policyWasReapplied).toBe(true);
    expect(created).toBe(singleRenderCost);
    expect(host.querySelector<HTMLButtonElement>('[data-ribbon-command="bold"]')?.disabled).toBe(
      true,
    );
  });

  it('ignores a copied old-controller callback after rebinding', () => {
    toolbar?.dispose();
    const oldController = sheet.instance.commands;
    const replacement = new InteractionController({
      store: sheet.instance.store,
      getWb: () => sheet.workbook,
      history: sheet.instance.history,
      manageHistoryGuard: false,
    });
    let unregisterReplacement: (() => void) | undefined;
    const registerReplacement = (): void => {
      unregisterReplacement = registerInteractionController(sheet.instance.store, replacement);
    };
    let rebind = false;
    const unsubscribeBeforeToolbar = oldController.subscribe(() => {
      if (rebind) return;
      rebind = true;
      registerReplacement();
      toolbar?.rerender();
    });
    toolbar = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    const singleRenderCost = countDomCreations(() => toolbar?.rerender());
    const bold = host.querySelector<HTMLButtonElement>('[data-ribbon-command="bold"]');

    try {
      const created = countDomCreations(() => {
        oldController.setPolicy({ operations: { format: false } });
      });
      expect(rebind).toBe(true);
      expect(created).toBe(singleRenderCost);
      expect(host.querySelector('[data-ribbon-command="bold"]')).not.toBe(bold);
    } finally {
      unsubscribeBeforeToolbar();
      unregisterReplacement?.();
      replacement.dispose();
    }
  });

  it('rebuilds on policy replacement and same-object policy reapplication', () => {
    const policy = { operations: { format: false } };
    let bold = host.querySelector<HTMLButtonElement>('[data-ribbon-command="bold"]');

    sheet.instance.setPolicy(policy);
    let nextBold = host.querySelector<HTMLButtonElement>('[data-ribbon-command="bold"]');
    expect(nextBold).not.toBe(bold);
    expect(nextBold?.disabled).toBe(true);

    policy.operations.format = true;
    bold = nextBold;
    sheet.instance.setPolicy(policy);
    nextBold = host.querySelector<HTMLButtonElement>('[data-ribbon-command="bold"]');
    expect(nextBold).not.toBe(bold);
    expect(nextBold?.disabled).toBe(false);

    bold = nextBold;
    sheet.instance.setPolicy({ operations: { format: false } });
    nextBold = host.querySelector<HTMLButtonElement>('[data-ribbon-command="bold"]');
    expect(nextBold).not.toBe(bold);
    expect(nextBold?.disabled).toBe(true);

    sheet.instance.setPolicy(undefined);
    expect(host.querySelector<HTMLButtonElement>('[data-ribbon-command="bold"]')?.disabled).toBe(
      false,
    );
  });
});
