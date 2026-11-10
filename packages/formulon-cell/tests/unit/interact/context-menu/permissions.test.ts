import { describe, expect, it, vi } from 'vitest';
import { History } from '../../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../../src/commands/interaction-controller.js';
import { rowShiftBand } from '../../../../src/commands/interaction-effects.js';
import { type OperationIntent, viewerPolicy } from '../../../../src/commands/interaction-policy.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import type { ContextMenuInteractionController } from '../../../../src/interact/context-menu-options.js';
import { createContextMenuPermissions } from '../../../../src/interact/context-menu-permissions.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';

const build = (interactionController?: ContextMenuInteractionController) =>
  createContextMenuPermissions({
    store: createSpreadsheetStore(),
    interactionController,
  });

const controller = (
  over: Partial<ContextMenuInteractionController> = {},
): ContextMenuInteractionController => ({
  canExecute: vi.fn(() => ({ allowed: true as const })),
  execute: vi.fn(),
  ...over,
});

describe('context-menu permissions', () => {
  it('allows everything and reports no policy without a controller', () => {
    const permissions = build();
    expect(permissions.hasExplicitPolicy()).toBe(false);
    expect(permissions.decisionForItem('bold')).toEqual({ allowed: true });
    expect(permissions.navigationBlocks('rowDelete')).toBe(false);
  });

  it('decides every item through the built-in policy once the controller is restricted', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const restricted = new InteractionController({
      store,
      getWb: () => workbook,
      history: new History(),
    });
    const unregister = registerInteractionController(store, restricted);
    try {
      restricted.setPolicy(viewerPolicy());
      const permissions = createContextMenuPermissions({
        store,
        interactionController: restricted,
      });
      expect(permissions.hasExplicitPolicy()).toBe(true);
      expect(permissions.decisionForItem('rowDelete')).toMatchObject({ allowed: false });
      expect(permissions.decisionForItem('rowDelete')).not.toMatchObject({ code: 'invalid' });
      expect(permissions.decisionForItem('selectAll')).toEqual({ allowed: true });
      expect(permissions.decisionForItem('toggleWatch')).toEqual({ allowed: true });
    } finally {
      unregister();
      restricted.dispose();
      workbook.dispose();
    }
  });

  it('denies copy when the host policy turns it off', () => {
    const permissions = build(controller({ policy: { copy: false } }));
    expect(permissions.decisionForItem('copy')).toMatchObject({
      allowed: false,
      code: 'operationDenied',
    });
  });

  it('routes built-in items through canExecute as context-menu intents', () => {
    const canExecute = vi.fn((_intent: OperationIntent) => ({ allowed: true as const }));
    const permissions = build(controller({ canExecute }));
    permissions.decisionForItem('paste');
    permissions.decisionForItem('rowInsertAbove');
    const [paste, rowInsert] = canExecute.mock.calls.map(([intent]) => intent);
    expect(paste).toMatchObject({
      operation: 'paste',
      origin: 'contextMenu',
      commandId: 'paste',
      effects: [{ kind: 'range' }],
    });
    // A row insert declares the band it shifts: from the selection to the sheet end.
    expect(rowInsert).toMatchObject({
      operation: 'insertRows',
      effects: [{ kind: 'range', range: rowShiftBand(0, 0) }],
    });
  });

  it('lets canSelect override the selection policy flag', () => {
    expect(build(controller({ policy: { selection: false } })).canChangeSelection()).toBe(false);
    expect(build(controller()).canChangeSelection()).toBe(true);
    const canSelect = vi.fn(() => ({ allowed: true as const }));
    expect(
      build(controller({ canSelect, policy: { selection: false } })).canChangeSelection(),
    ).toBe(true);
  });
});
