import {
  builtInCommandIntent,
  builtInProtectionDecision,
  canExecuteBuiltIn,
} from '../commands/built-in-command-policy.js';
import type {
  OperationIntent,
  PermissionCode,
  PermissionDecision,
} from '../commands/interaction-policy.js';
import type { Range } from '../engine/types.js';
import type { SpreadsheetStore } from '../store/store.js';
import type { ContextMenuInteractionController, ContextMenuItem } from './context-menu-options.js';
import type { ItemId } from './context-menu-spec.js';
import { navigationBoundsFor } from './navigation-policy.js';

export interface ContextMenuPermissionsContext {
  readonly store: SpreadsheetStore;
  readonly interactionController: ContextMenuInteractionController | undefined;
}

/** Host policy and bounded-viewport decisions for context-menu items. */
export interface ContextMenuPermissions {
  decisionForItem(item: ItemId | ContextMenuItem | OperationIntent): PermissionDecision;
  decisionReason(decision: PermissionDecision): string | null;
  hasExplicitPolicy(): boolean;
  canPasteToRange(range: Range | null): boolean;
  canChangeSelection(): boolean;
  /** True when `id` is a structural item a bounded viewport cannot run. */
  navigationBlocks(id: string): boolean;
}

export function createContextMenuPermissions(
  ctx: ContextMenuPermissionsContext,
): ContextMenuPermissions {
  const { store, interactionController } = ctx;
  const boundedViewportStructuralIds = new Set<ItemId>([
    'insertCopiedCells',
    'insertCells',
    'insertTableRows',
    'insertTableColumns',
    'deleteCells',
    'rowInsertAbove',
    'rowInsertBelow',
    'rowDelete',
    'colInsertLeft',
    'colInsertRight',
    'colDelete',
  ]);

  const navigationBounded = (): boolean => navigationBoundsFor(store) !== undefined;

  const canPasteToRange = (range: Range | null): boolean => {
    const bounds = navigationBoundsFor(store);
    if (!bounds) return true;
    if (
      !range ||
      range.sheet !== bounds.sheet ||
      range.r0 < bounds.r0 ||
      range.c0 < bounds.c0 ||
      range.r1 > bounds.r1 ||
      range.c1 > bounds.c1
    ) {
      return false;
    }
    if (!interactionController) return true;
    return interactionController.canExecute({
      operation: 'paste',
      origin: 'contextMenu',
      commandId: 'paste',
      effects: [{ kind: 'range', range }],
    }).allowed;
  };

  const hasExplicitPolicy = (): boolean => {
    const policy = interactionController?.policy;
    return (
      interactionController?.restricted === true ||
      (policy !== undefined && Object.keys(policy).length > 0)
    );
  };

  const deniedDecision = (code: PermissionCode, reason?: string): PermissionDecision => ({
    allowed: false,
    code,
    ...(reason ? { reason } : {}),
  });

  const decisionForIntent = (intent: OperationIntent): PermissionDecision => {
    if (!interactionController) return { allowed: true };
    return interactionController.canExecute(intent);
  };

  const decisionReason = (decision: PermissionDecision): string | null =>
    decision.allowed ? null : (decision.reason ?? null);

  const decisionForItem = (
    item: ItemId | ContextMenuItem | OperationIntent,
  ): PermissionDecision => {
    if (typeof item === 'object' && 'operation' in item && 'origin' in item) {
      return decisionForIntent(item);
    }
    const itemId = typeof item === 'string' ? item : item.builtIn;
    if (!itemId) {
      const command = typeof item === 'string' ? undefined : item.command;
      if (!command) return { allowed: true };
      return decisionForIntent({
        operation: command.operation,
        origin: 'contextMenu',
        commandId: command.commandId ?? command.type,
        effects: [
          {
            kind: 'cells',
            cells: command.changes.map((change) => change.addr),
            includesFormula: command.changes.some(
              (change) => 'formula' in change && change.formula != null,
            ),
          },
        ],
      });
    }
    if (itemId === 'copy' && interactionController?.policy?.copy === false) {
      return deniedDecision('operationDenied', 'Copy is disabled by the host policy.');
    }
    if (navigationBounded() && boundedViewportStructuralIds.has(itemId)) {
      return deniedDecision(
        'outOfBounds',
        'This structural operation is unavailable with a bounded viewport.',
      );
    }
    const protection = builtInProtectionDecision(store, itemId);
    if (!protection.allowed) return protection;
    // Insert… / Delete… open a dialog, available while one of its choices is.
    if (hasExplicitPolicy() || itemId === 'insertCells' || itemId === 'deleteCells')
      return canExecuteBuiltIn(store, itemId, 'contextMenu');
    const intent = builtInCommandIntent(store, itemId, 'contextMenu');
    return intent ? decisionForIntent(intent) : { allowed: true };
  };

  const canChangeSelection = (): boolean => {
    if (interactionController?.canSelect) return interactionController.canSelect().allowed;
    return interactionController?.policy?.selection !== false;
  };

  return {
    decisionForItem,
    decisionReason,
    hasExplicitPolicy,
    canPasteToRange,
    canChangeSelection,
    navigationBlocks: (id) => navigationBounded() && boundedViewportStructuralIds.has(id as ItemId),
  };
}
