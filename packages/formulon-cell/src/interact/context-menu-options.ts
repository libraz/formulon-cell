import type {
  CellBatchCommand,
  InteractionOperation,
  InteractionOrigin,
  OperationIntent,
  PermissionDecision,
} from '../commands/interaction-policy.js';
import type { Addr, Range } from '../engine/types.js';
import type { ItemId, MenuKind } from './context-menu-spec.js';

export type { ItemId } from './context-menu-spec.js';

/**
 * The information supplied to a context-menu host or composition callback.
 *
 * `defaultItems` is a snapshot for this invocation. Consumers should keep the
 * item IDs and pass this context back to an item action instead of retaining a
 * DOM event or reading selection state later.
 */
export interface ContextMenuContext {
  readonly kind: MenuKind;
  readonly cell: Addr;
  readonly selection: Range;
  readonly defaultItems: readonly ContextMenuItem[];
  readonly event: MouseEvent | KeyboardEvent;
  /** Query the current authorization for a built-in item or operation. */
  readonly canExecute: (item: ItemId | ContextMenuItem | OperationIntent) => PermissionDecision;
  /** Alias for hosts that prefer an explicit permission-oriented name. */
  readonly permission: (item: ItemId | ContextMenuItem | OperationIntent) => PermissionDecision;
}

/**
 * A stable, composable menu item. Built-in actions are identified by
 * `builtIn`; custom IDs never enter the built-in switch by accident.
 */
export interface ContextMenuItem {
  readonly id: string;
  readonly label?: string;
  readonly hint?: string;
  readonly separator?: boolean;
  readonly children?: readonly ContextMenuItem[];
  readonly disabled?: boolean;
  readonly disabledReason?: string;
  readonly builtIn?: ItemId;
  /** A typed cell-batch request executed by the shared interaction service. */
  readonly command?: CellBatchCommand;
  /** Trusted host-owned action. It may open a host dialog or dispatch an app event. */
  readonly action?: (context: ContextMenuContext) => void;
}

export type ContextMenuTransform = (context: ContextMenuContext) => readonly ContextMenuItem[];

/** Small structural view used by the menu; the command service owns the full
 * controller implementation. Keeping this view here avoids giving custom
 * items a generic callback executor. */
export interface ContextMenuInteractionController {
  readonly policy?: import('../commands/interaction-policy.js').InteractionPolicy;
  readonly restricted?: boolean;
  canSelect?(): PermissionDecision;
  canCopy?(): PermissionDecision;
  canExecute(intent: OperationIntent): PermissionDecision;
  execute(command: CellBatchCommand): unknown;
  subscribe?(listener: () => void): () => void;
}

export type ContextMenuOptions =
  | { readonly mode: 'disabled' }
  | {
      readonly mode: 'builtIn';
      /** Select a subset of registered built-in IDs. Omitted keeps legacy items. */
      readonly items?: readonly ItemId[];
      /** Transform the selected built-in snapshot and/or add trusted host items. */
      readonly transform?: ContextMenuTransform;
    }
  | {
      readonly mode: 'host';
      /** The library suppresses its/browser menu and gives the host the context. */
      readonly onOpen: (context: ContextMenuContext) => void;
    };

/**
 * The command origin used for built-in context-menu requests. Kept exported so
 * host command builders can use the same finite origin vocabulary as the
 * interaction policy without widening it to an arbitrary string.
 */
export type ContextMenuCommandOrigin = Extract<InteractionOrigin, 'contextMenu'>;

/** The operation set a custom cell-batch command may request. */
export type ContextMenuCommandOperation =
  | Extract<InteractionOperation, 'valueEdit'>
  | Extract<InteractionOperation, 'formulaEdit'>
  | Extract<InteractionOperation, 'clear'>
  | Extract<InteractionOperation, 'paste'>
  | Extract<InteractionOperation, 'fill'>
  | Extract<InteractionOperation, 'moveCells'>;
