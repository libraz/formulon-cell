import { describe, expect, it, vi } from 'vitest';
import type { OperationIntent } from '../../../../src/commands/interaction-policy.js';
import type { ContextMenuInteractionController } from '../../../../src/interact/context-menu-options.js';
import { createContextMenuPermissions } from '../../../../src/interact/context-menu-permissions.js';
import type { MenuKind } from '../../../../src/interact/context-menu-spec.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';

const build = (
  interactionController?: ContextMenuInteractionController,
  menuKind: MenuKind = 'cell',
) =>
  createContextMenuPermissions({
    store: createSpreadsheetStore(),
    interactionController,
    menuKind: () => menuKind,
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

  it('denies items outside the safe set once the controller is restricted', () => {
    const permissions = build(controller({ restricted: true }));
    expect(permissions.hasExplicitPolicy()).toBe(true);
    const decision = permissions.decisionForItem('rowDelete');
    expect(decision).toMatchObject({ allowed: false, code: 'unsupported' });
    expect(permissions.decisionReason(decision)).toMatch(/restricted mode/);
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
    expect(rowInsert).toMatchObject({ operation: 'insertRows', effects: [{ kind: 'workbook' }] });
  });

  it('treats insertCopiedCells as a workbook effect only on row and column menus', () => {
    const effectKind = (kind: MenuKind): string | undefined => {
      const canExecute = vi.fn((_intent: OperationIntent) => ({ allowed: true as const }));
      build(controller({ canExecute }), kind).decisionForItem('insertCopiedCells');
      return canExecute.mock.calls[0]?.[0].effects[0]?.kind;
    };
    expect(effectKind('cell')).toBe('range');
    expect(effectKind('row')).toBe('workbook');
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
