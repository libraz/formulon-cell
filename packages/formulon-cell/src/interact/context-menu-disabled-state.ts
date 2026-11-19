import type { PermissionDecision } from '../commands/interaction-policy.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import type { ContextMenuItem } from './context-menu-options.js';
import type { ContextMenuPermissions } from './context-menu-permissions.js';

export type ContextMenuDisabledStateDeps = Pick<
  ContextMenuPermissions,
  'decisionForItem' | 'decisionReason'
>;

/** Tracks which menu items are disabled by the host or by the policy, and projects that onto buttons. */
export interface ContextMenuDisabledState {
  /** Forget all remembered sources; call when the menu is rebuilt. */
  reset(): void;
  /** Disable a button on behalf of the host policy, remembering that the policy did it. */
  markPolicyDisabled(btn: HTMLButtonElement, decision: PermissionDecision): void;
  /** Disabled state for a plain item or a submenu parent built from `source`. */
  applySource(btn: HTMLButtonElement, source: ContextMenuItem | undefined): void;
  /** Disable items the policy denies, carrying explicit-disable marks over to the copies. */
  decorate(items: readonly ContextMenuItem[]): ContextMenuItem[];
  /** Remember the items the host disabled explicitly, children included. */
  markExplicit(items: readonly ContextMenuItem[]): void;
  /** Re-project a fresh policy decision onto an already-rendered button. */
  refresh(
    button: HTMLButtonElement,
    source: ContextMenuItem | undefined,
    decision: PermissionDecision,
  ): void;
}

export function setContextMenuItemDisabled(
  button: HTMLButtonElement,
  disabled: boolean,
  reason: string | null,
): void {
  projectDisabledState(button, disabled, reason, {
    datasetKey: 'disabledReason',
  });
}

export function createContextMenuDisabledState(
  deps: ContextMenuDisabledStateDeps,
): ContextMenuDisabledState {
  const { decisionForItem, decisionReason } = deps;
  let explicitDisabledSources = new WeakSet<ContextMenuItem>();
  let policyDisabledSources = new WeakSet<ContextMenuItem>();

  const markPolicyDisabled = (btn: HTMLButtonElement, decision: PermissionDecision): void => {
    btn.dataset.policyDisabled = 'true';
    setContextMenuItemDisabled(btn, true, decisionReason(decision));
  };

  /** An explicit disable keeps its own reason, and a policy denial is remembered either way. */
  const applySourceDisabledState = (
    btn: HTMLButtonElement,
    source: ContextMenuItem | undefined,
  ): void => {
    const gated = !!(source?.builtIn || source?.command);
    if (source && explicitDisabledSources.has(source)) {
      if (gated && !decisionForItem(source).allowed) btn.dataset.policyDisabled = 'true';
      setContextMenuItemDisabled(btn, true, source.disabledReason ?? null);
    } else if (source && gated) {
      const decision = decisionForItem(source);
      if (!decision.allowed) markPolicyDisabled(btn, decision);
    }
  };

  const decorateContextItems = (items: readonly ContextMenuItem[]): ContextMenuItem[] =>
    items.map((item) => {
      const children = item.children ? decorateContextItems(item.children) : undefined;
      if (item.disabled) {
        const decorated = children ? { ...item, children } : item;
        if (explicitDisabledSources.has(item)) explicitDisabledSources.add(decorated);
        return decorated;
      }
      const decision = item.builtIn || item.command ? decisionForItem(item) : null;
      if (decision && !decision.allowed) {
        const decorated = {
          ...item,
          ...(children ? { children } : {}),
          disabled: true,
          disabledReason: decisionReason(decision) ?? undefined,
        };
        policyDisabledSources.add(decorated);
        if (explicitDisabledSources.has(item)) explicitDisabledSources.add(decorated);
        return decorated;
      }
      return children ? { ...item, children } : item;
    });

  const markExplicitDisabledSources = (items: readonly ContextMenuItem[]): void => {
    for (const item of items) {
      if (item.disabled && !policyDisabledSources.has(item)) explicitDisabledSources.add(item);
      if (item.children) markExplicitDisabledSources(item.children);
    }
  };

  return {
    reset: () => {
      explicitDisabledSources = new WeakSet<ContextMenuItem>();
      policyDisabledSources = new WeakSet<ContextMenuItem>();
    },
    markPolicyDisabled,
    applySource: applySourceDisabledState,
    decorate: decorateContextItems,
    markExplicit: markExplicitDisabledSources,
    refresh: (button, source, decision) => {
      if (!decision.allowed) {
        markPolicyDisabled(button, decision);
      } else if (button.dataset.policyDisabled === 'true') {
        delete button.dataset.policyDisabled;
        if (source && explicitDisabledSources.has(source)) {
          setContextMenuItemDisabled(button, true, source.disabledReason ?? null);
        } else {
          setContextMenuItemDisabled(button, false, null);
        }
      }
    },
  };
}
