import { canExecuteBuiltIn } from '../commands/built-in-command-policy.js';
import type {
  InteractionOperation,
  OperationEffect,
  OperationIntent,
  PermissionCode,
  PermissionDecision,
} from '../commands/interaction-policy.js';
import type { Range } from '../engine/types.js';
import type { SpreadsheetStore } from '../store/store.js';
import type { ContextMenuInteractionController, ContextMenuItem } from './context-menu-options.js';
import type { ItemId, MenuKind } from './context-menu-spec.js';
import { navigationBoundsFor } from './navigation-policy.js';

export interface ContextMenuPermissionsContext {
  readonly store: SpreadsheetStore;
  readonly interactionController: ContextMenuInteractionController | undefined;
  /** Kind of the menu currently being built, read at call time. */
  readonly menuKind: () => MenuKind;
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
  const { store, interactionController, menuKind } = ctx;
  const restrictedMenuSafeIds = new Set<ItemId>(['copy', 'paste', 'clear']);
  const navigationUnsupportedIds = new Set<ItemId>([
    'insertCopiedCells',
    'insertCells',
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

  const operationForItem = (id: string): InteractionOperation | null => {
    if (id === 'copy' || id === 'selectAll' || id === 'rowHeight' || id === 'colWidth') {
      return null;
    }
    if (
      id === 'paste' ||
      id === 'pasteSpecial' ||
      id === 'pasteAll' ||
      id === 'pasteFormulas' ||
      id === 'pasteFormulasNumFmt' ||
      id === 'pasteValues' ||
      id === 'pasteValuesNumFmt' ||
      id === 'pasteFormatsOnly' ||
      id === 'pasteTranspose'
    ) {
      return 'paste';
    }
    if (id === 'clear') return 'clear';
    if (
      id === 'cut' ||
      id === 'insertCopiedCells' ||
      id === 'insertCells' ||
      id === 'deleteCells'
    ) {
      return 'moveCells';
    }
    if (id === 'rowInsertAbove' || id === 'rowInsertBelow') return 'insertRows';
    if (id === 'rowDelete') return 'deleteRows';
    if (id === 'colInsertLeft' || id === 'colInsertRight') return 'insertColumns';
    if (id === 'colDelete') return 'deleteColumns';
    if (
      id === 'bold' ||
      id === 'italic' ||
      id === 'underline' ||
      id === 'alignLeft' ||
      id === 'alignCenter' ||
      id === 'alignRight' ||
      id === 'borders' ||
      id === 'formatCells' ||
      id === 'editPhonetic'
    ) {
      return 'format';
    }
    if (id === 'filterClear' || id === 'filterReapply' || id === 'filterByValue') {
      return 'filter';
    }
    if (id === 'sortAsc' || id === 'sortDesc') return 'sort';
    if (id === 'rowHide' || id === 'rowUnhide') return 'resizeRows';
    if (id === 'colHide' || id === 'colUnhide') return 'resizeColumns';
    if (id === 'rowGroup' || id === 'rowUngroup') return 'insertRows';
    if (id === 'colGroup' || id === 'colUngroup') return 'insertColumns';
    if (id === 'insertComment' || id === 'deleteComment') return 'comment';
    if (id === 'insertHyperlink' || id === 'openHyperlink') return 'hyperlink';
    if (id === 'defineName') return 'namedRange';
    if (id === 'toggleWatch') return 'object';
    return 'object';
  };

  const operationEffect = (id: string): OperationEffect => {
    // Row/column structure commands operate on the worksheet axis. Feeding a
    // full row or column range into the authorization materializer would
    // exceed its bounded cell budget before the command's own protection and
    // overflow checks run. Treat these as workbook-structure effects; the
    // navigation gate above still applies to embedded bounded views.
    const structural =
      (id === 'insertCopiedCells' && menuKind() !== 'cell') ||
      id === 'rowInsertAbove' ||
      id === 'rowInsertBelow' ||
      id === 'rowDelete' ||
      id === 'colInsertLeft' ||
      id === 'colInsertRight' ||
      id === 'colDelete';
    return structural
      ? { kind: 'workbook' }
      : { kind: 'range', range: { ...store.getState().selection.range } };
  };

  const intentForBuiltIn = (id: string): OperationIntent | null => {
    const operation = operationForItem(id);
    if (!operation) return null;
    return {
      operation,
      origin: 'contextMenu',
      commandId: id,
      effects: [operationEffect(id)],
    };
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
    if (navigationBounded() && navigationUnsupportedIds.has(itemId)) {
      return deniedDecision(
        'unsupported',
        'This structural operation is unavailable with a bounded viewport.',
      );
    }
    if (hasExplicitPolicy() && !restrictedMenuSafeIds.has(itemId)) {
      return deniedDecision(
        'unsupported',
        'This context-menu operation is unavailable in restricted mode.',
      );
    }
    if (hasExplicitPolicy()) return canExecuteBuiltIn(store, itemId, 'contextMenu');
    const intent = intentForBuiltIn(itemId);
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
    navigationBlocks: (id) => navigationBounded() && navigationUnsupportedIds.has(id as ItemId),
  };
}
