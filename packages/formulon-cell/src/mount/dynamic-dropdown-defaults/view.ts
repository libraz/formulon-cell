import { type FreezeAction, handleFreezeAction, type SpreadsheetInstance } from '../../index.js';
import type { DynamicDropdownsCtx } from '../../toolbar/ribbon/dynamic-dropdowns.js';
import { setMenuIcon } from '../../toolbar/ribbon/menus/general.js';
import { setMenuControlDisabled } from './menu-feedback.js';

const freezeActionFromMenu = (action: string): FreezeAction | null => {
  if (action === 'row') return 'topRow';
  if (action === 'col') return 'firstColumn';
  if (action === 'selection') return 'panes';
  if (action === 'off') return 'none';
  return null;
};

const buildFreezeAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyFreezeAction'] =>
  (action) => {
    const next = freezeActionFromMenu(action);
    if (!next) return;
    handleFreezeAction(instance, next);
    instance.host.focus();
  };

const updateFreezeMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateFreezeMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const hasFreeze = state.layout.freezeRows > 0 || state.layout.freezeCols > 0;
    const strings = instance.i18n.strings;
    const toolbarStrings = strings.toolbar as typeof strings.toolbar & { unfreezePanes?: string };
    const primary = menu.querySelector<HTMLButtonElement>(
      '[data-freeze="selection"], [data-freeze="off"]',
    );
    if (primary) {
      primary.dataset.freeze = hasFreeze ? 'off' : 'selection';
      setMenuControlDisabled(primary, false);
      const text = primary.querySelector<HTMLElement>('.fc-tb__menu-item__text');
      if (text) {
        text.textContent = hasFreeze
          ? (toolbarStrings.unfreezePanes ?? strings.toolbar.unfreeze)
          : strings.viewToolbar.freezePanes;
      }
      const icon = primary.querySelector<HTMLElement>('.fc-tb__menu-icon');
      if (icon) setMenuIcon(icon, hasFreeze ? 'freeze-off' : 'freeze-panes');
    }
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-freeze]')) {
      if (button === primary) continue;
      setMenuControlDisabled(button, false);
    }
  };

type ViewDropdownDefaults = Pick<DynamicDropdownsCtx, 'applyFreezeAction' | 'updateFreezeMenu'>;

export function createViewDropdownDefaults(instance: SpreadsheetInstance): ViewDropdownDefaults {
  return {
    applyFreezeAction: buildFreezeAction(instance),
    updateFreezeMenu: updateFreezeMenu(instance),
  };
}
