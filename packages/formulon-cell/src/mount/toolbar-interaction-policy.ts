// Restricted-embedding policy for a mounted ribbon: `project` disables the
// buttons and form controls the policy denies, and capture-phase listeners on
// the host block any other ribbon interaction. `detach` removes the listeners.

import { canExecuteBuiltIn } from '../commands/built-in-command-policy.js';
import { interactionControllerFor } from '../commands/interaction-controller.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { MAC_FORMULAS_MORE_MENU_ID } from '../toolbar/ribbon/mac/model.js';
import type { SpreadsheetInstance } from './types.js';

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
    if (!current || interactionControllerFor(current.store)?.policy === undefined) return;
    closeDynamicDropdowns();
    for (const button of host.querySelectorAll<HTMLButtonElement>('[data-ribbon-command]')) {
      // Intrinsic engine availability has priority over embedding policy. The
      // menu projector already supplied its localized reason; a policy tick
      // must not replace it with a generic denial message.
      if (button.dataset.functionUnavailable === 'true') continue;
      const decision = canExecuteBuiltIn(
        current.store,
        button.dataset.ribbonCommand ?? '',
        'ribbon',
      );
      if (decision.allowed) continue;
      projectDisabledState(button, true, decision.reason ?? decision.code);
    }
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
      canExecuteBuiltIn(current.store, command.dataset.ribbonCommand ?? '', 'ribbon').allowed &&
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
