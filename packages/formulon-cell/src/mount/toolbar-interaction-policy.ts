// Built-in command availability for a mounted ribbon: `project` disables the
// buttons the built-in decision refuses (sheet protection always, the host
// policy when there is one) and re-enables them when it allows them again;
// under a policy it also disables form controls, and capture-phase listeners
// on the host block any other ribbon interaction. `detach` removes them.

import { canExecuteBuiltInRequest } from '../commands/built-in-command-policy.js';
import { interactionControllerFor } from '../commands/interaction-controller.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { MAC_FORMULAS_MORE_MENU_ID } from '../toolbar/ribbon/mac/model.js';
import type { SpreadsheetInstance } from './types.js';

/** Dataset key holding the reason this projection disabled a button with. */
const PROJECTED_REASON = 'ribbonDecisionReason';

export interface RibbonInteractionPolicyDeps {
  getInstance: () => SpreadsheetInstance | null;
  closeDynamicDropdowns: () => void;
}

export interface RibbonInteractionPolicy {
  project: () => void;
  detach: () => void;
}

export function attachRibbonInteractionPolicy(
  host: HTMLElement,
  deps: RibbonInteractionPolicyDeps,
): RibbonInteractionPolicy {
  const { getInstance, closeDynamicDropdowns } = deps;

  const project = (): void => {
    const current = getInstance();
    if (!current) return;
    // Sheet protection refuses some built-ins with or without a host policy, so
    // every button follows the one built-in decision; the rest is policy-only.
    const restricted = interactionControllerFor(current.store)?.policy !== undefined;
    if (restricted) closeDynamicDropdowns();
    for (const button of host.querySelectorAll<HTMLButtonElement>('[data-ribbon-command]')) {
      // Intrinsic engine availability has priority over embedding policy. The
      // menu projector already supplied its localized reason; a policy tick
      // must not replace it with a generic denial message.
      if (button.dataset.functionUnavailable === 'true') continue;
      const decision = canExecuteBuiltInRequest(
        current.store,
        button.dataset.ribbonCommand ?? '',
        'ribbon',
      );
      const projected = button.dataset[PROJECTED_REASON];
      if (decision.allowed) {
        if (projected === undefined) continue;
        delete button.dataset[PROJECTED_REASON];
        // Re-enable only what this projection disabled and nothing has re-disabled since.
        if (button.getAttribute('aria-description') === projected)
          projectDisabledState(button, false, null);
        continue;
      }
      // A button another projector already disabled keeps its own (localized) reason.
      if (projected === undefined && button.disabled) continue;
      const reason = decision.reason ?? decision.code;
      projectDisabledState(button, true, reason);
      button.dataset[PROJECTED_REASON] = reason;
    }
    if (!restricted) return;
    for (const input of host.querySelectorAll<HTMLInputElement | HTMLSelectElement>(
      'input, select',
    )) {
      projectDisabledState(input, true, 'Unavailable in restricted embedding');
    }
  };

  // Capture before directly-bound control/dropdown handlers, including hosts
  // using a separately mounted toolbar. Unknown restricted routes fail closed.
  const guardInteraction = (event: Event): void => {
    const current = getInstance();
    if (!current || interactionControllerFor(current.store)?.policy === undefined) return;
    const target = event.target;
    if (!(target instanceof Element)) return;
    const command = target.closest<HTMLElement>('[data-ribbon-command]');
    const inMenu = target.closest('.fc-tb__menu');
    const isInput = target.closest('input, select, textarea');
    if (
      event.type === 'click' &&
      target.closest(`#${MAC_FORMULAS_MORE_MENU_ID} [data-function-category-submenu]`)
    )
      return;
    if (
      command &&
      canExecuteBuiltInRequest(current.store, command.dataset.ribbonCommand ?? '', 'ribbon')
        .allowed &&
      !isInput
    )
      return;
    if (!command && !inMenu && !isInput) return;
    event.preventDefault();
    event.stopImmediatePropagation();
  };

  host.addEventListener('click', guardInteraction, true);
  host.addEventListener('change', guardInteraction, true);
  host.addEventListener('input', guardInteraction, true);

  return {
    project,
    detach: () => {
      host.removeEventListener('click', guardInteraction, true);
      host.removeEventListener('change', guardInteraction, true);
      host.removeEventListener('input', guardInteraction, true);
    },
  };
}
