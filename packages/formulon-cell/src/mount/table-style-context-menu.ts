import { customPivotTableStyleById, customTableStyleById } from '../commands/format-as-table.js';
import { anySheetProtected } from '../commands/protection.js';
import {
  activeStyleTarget,
  deleteCustomStyle,
  setStyleDefault,
  type TableStyleKind,
} from '../commands/table-style-actions.js';
import { focusSheet } from '../interact/dialog-focus.js';
import { onOutsidePress } from '../interact/outside-press.js';
import { focusMenuItem, handleMenuKeydown, projectDisabledState } from '../toolbar/menu-a11y.js';
import { menuSeparator, menuTextChip } from '../toolbar/ribbon/menus/general.js';
import type { SpreadsheetInstance } from './types.js';

/** The in-ribbon strip chips; the expanded galleries have no context menu. */
const ITEM_SELECTOR = '[data-table-quick-style], [data-pivot-style]';

type Action = 'applyMaintain' | 'applyClear' | 'modify' | 'duplicate' | 'delete' | 'setDefault';

/** The custom style `styleId` names among `kind`'s custom styles, or null. */
const customStyleOf = (instance: SpreadsheetInstance, kind: TableStyleKind, styleId: string) =>
  kind === 'table'
    ? customTableStyleById(instance.store.getState(), styleId)
    : customPivotTableStyleById(instance.store.getState(), styleId);

/**
 * The context menu of a table or PivotTable style gallery item, in the galleries' menus and
 * strips: apply keeping or clearing the cells' formatting, Set As Default, and Delete for a
 * custom style. Modify and Duplicate open the style dialog.
 */
export function attachTableStyleContextMenu(
  host: HTMLElement,
  getInstance: () => SpreadsheetInstance | null,
  openStyleDialog: (kind: TableStyleKind, mode: 'modify' | 'duplicate', styleId: string) => void,
): () => void {
  let open: { menu: HTMLDivElement; restore: HTMLElement } | null = null;
  let detachOutside: (() => void) | null = null;

  const close = (restoreFocus: boolean): void => {
    if (!open) return;
    const { menu, restore } = open;
    open = null;
    menu.remove();
    detachOutside?.();
    detachOutside = null;
    document.removeEventListener('keydown', onKey, true);
    if (restoreFocus && restore.isConnected) restore.focus({ preventScroll: true });
  };
  const onKey = (event: KeyboardEvent): void => {
    if (!open) return;
    handleMenuKeydown(event, open.menu, { close });
  };

  const run = (
    instance: SpreadsheetInstance,
    kind: TableStyleKind,
    action: Action,
    styleId: string,
    from: string,
  ): void => {
    const entry = { origin: 'ribbon', commandId: from } as const;
    const ctx = { store: instance.store, history: instance.history, ...entry };
    const custom = customStyleOf(instance, kind, styleId);
    switch (action) {
      case 'applyMaintain':
      case 'applyClear': {
        const clearFormatting = action === 'applyClear';
        activeStyleTarget({ ...ctx, workbook: instance.workbook }, kind)?.apply(styleId, {
          clearFormatting,
        });
        break;
      }
      case 'modify':
        // While any sheet is protected Modify stays enabled and does nothing, as on the desktop.
        if (anySheetProtected(instance.store.getState())) break;
        openStyleDialog(kind, action, styleId);
        return;
      case 'duplicate':
        openStyleDialog(kind, action, styleId);
        return;
      case 'delete':
        if (custom) deleteCustomStyle(ctx, instance.workbook, custom.label);
        break;
      case 'setDefault':
        setStyleDefault(instance.workbook, kind, custom?.label ?? styleId, instance.history);
        break;
    }
    focusSheet(instance.host);
  };

  const onContextMenu = (event: MouseEvent): void => {
    const target = event.target instanceof Element ? event.target : null;
    const item = target?.closest<HTMLElement>(ITEM_SELECTOR);
    const instance = getInstance();
    if (!item || !instance || !host.contains(item)) return;
    // A greyed chip takes no right-click: no menu of ours and none of the browser's.
    if (item instanceof HTMLButtonElement && item.disabled) {
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    const kind: TableStyleKind = item.dataset.pivotStyle !== undefined ? 'pivot' : 'table';
    const styleId = item.dataset.tableQuickStyle ?? item.dataset.pivotStyle ?? '';
    // None is a gallery entry without a style to edit.
    const none = styleId === '';
    if (kind === 'table' && none) return;
    event.preventDefault();
    event.stopPropagation();
    close(false);
    const from = kind === 'pivot' ? 'pivotQuickStyles' : 'tableQuickStyles';
    const t = instance.i18n.strings.ribbonMenu;
    const customStyle = customStyleOf(instance, kind, styleId);
    const custom = customStyle !== null;
    const menu = document.createElement('div');
    menu.className = 'fc-tb__menu fc-tb__tablestyle-ctx';
    menu.setAttribute('role', 'menu');
    menu.style.position = 'fixed';
    menu.style.left = `${event.clientX}px`;
    menu.style.top = `${event.clientY}px`;
    const button = (
      action: Action,
      label: string,
      disabled = false,
      reason: string | null = null,
    ): HTMLButtonElement => {
      const node = menuTextChip({
        label,
        attr: 'tableStyleContext',
        value: action,
        className: 'fc-tb__menu-item',
        labelClassName: 'fc-tb__menu-item__text',
      });
      projectDisabledState(node, disabled, reason);
      node.addEventListener('click', () => {
        close(false);
        run(instance, kind, action, styleId, from);
      });
      return node;
    };
    // Built-in styles cannot be modified or deleted; the defaults need an engine that keeps them.
    menu.append(
      button('applyClear', t.tableStyleApplyClear),
      button('applyMaintain', t.tableStyleApplyMaintain),
      menuSeparator(),
      button('modify', t.tableStyleModify, !custom),
      button('duplicate', t.tableStyleDuplicate, none),
      button('delete', t.tableStyleDelete, !custom),
      menuSeparator(),
      button(
        'setDefault',
        t.tableStyleSetDefault,
        none || instance.workbook.capabilities.customTableStyles !== true,
      ),
    );
    host.appendChild(menu);
    open = { menu, restore: item };
    detachOutside = onOutsidePress(
      document,
      () => [open?.menu],
      () => {
        if (open) close(false);
      },
      { event: 'pointerdown' },
    );
    document.addEventListener('keydown', onKey, true);
    focusMenuItem(menu, 'first');
  };

  host.addEventListener('contextmenu', onContextMenu, true);
  return () => {
    close(false);
    host.removeEventListener('contextmenu', onContextMenu, true);
  };
}
