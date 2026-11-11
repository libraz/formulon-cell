import { customPivotTableStyleById, customTableStyleById } from '../commands/format-as-table.js';
import {
  activeStyleTarget,
  deleteCustomStyle,
  setStyleDefault,
  type TableStyleKind,
} from '../commands/table-style-actions.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { menuSeparator, menuTextChip } from '../toolbar/ribbon/menus/general.js';
import { showCreateTableDialog } from './dynamic-dropdown-defaults/styles.js';
import type { SpreadsheetInstance } from './types.js';

const ITEM_SELECTOR =
  '[data-table-style], [data-table-quick-style], [data-pivot-table-style], [data-pivot-style]';

type Action = 'applyMaintain' | 'applyClear' | 'modify' | 'duplicate' | 'delete' | 'setDefault';

/**
 * The context menu of a table or PivotTable style gallery item, in the galleries' menus and
 * strips: apply keeping or clearing the cells' formatting, Set As Default, and Delete for a
 * custom style. Modify and Duplicate open the style dialog.
 */
/** The custom style `styleId` names among `kind`'s custom styles, or null. */
const customStyleOf = (instance: SpreadsheetInstance, kind: TableStyleKind, styleId: string) =>
  kind === 'table'
    ? customTableStyleById(instance.store.getState(), styleId)
    : customPivotTableStyleById(instance.store.getState(), styleId);

export function attachTableStyleContextMenu(
  host: HTMLElement,
  getInstance: () => SpreadsheetInstance | null,
  openStyleDialog: (kind: TableStyleKind, mode: 'modify' | 'duplicate', styleId: string) => void,
): () => void {
  let open: { menu: HTMLDivElement; restore: HTMLElement } | null = null;

  const close = (restoreFocus: boolean): void => {
    if (!open) return;
    const { menu, restore } = open;
    open = null;
    menu.remove();
    document.removeEventListener('pointerdown', onOutside, true);
    document.removeEventListener('keydown', onKey, true);
    if (restoreFocus && restore.isConnected) restore.focus({ preventScroll: true });
  };
  const onOutside = (event: PointerEvent): void => {
    if (open && event.target instanceof Node && !open.menu.contains(event.target)) close(false);
  };
  const onKey = (event: KeyboardEvent): void => {
    if (!open) return;
    const items = [...open.menu.querySelectorAll<HTMLButtonElement>('button:not(:disabled)')];
    const at = items.indexOf(document.activeElement as HTMLButtonElement);
    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      close(true);
    } else if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
      event.preventDefault();
      const step = event.key === 'ArrowDown' ? 1 : -1;
      items[(at + step + items.length) % items.length]?.focus();
    }
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
        const target = activeStyleTarget({ ...ctx, workbook: instance.workbook }, kind);
        if (target) target.apply(styleId, { clearFormatting });
        else if (kind === 'table')
          void showCreateTableDialog(instance, { style: styleId, clearFormatting }, entry);
        break;
      }
      case 'modify':
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
    instance.host.focus();
  };

  const onContextMenu = (event: MouseEvent): void => {
    const target = event.target instanceof Element ? event.target : null;
    const item = target?.closest<HTMLElement>(ITEM_SELECTOR);
    const instance = getInstance();
    if (!item || !instance || !host.contains(item)) return;
    const kind: TableStyleKind =
      item.dataset.pivotTableStyle !== undefined || item.dataset.pivotStyle !== undefined
        ? 'pivot'
        : 'table';
    const styleId =
      item.dataset.tableStyle ??
      item.dataset.tableQuickStyle ??
      item.dataset.pivotTableStyle ??
      item.dataset.pivotStyle ??
      '';
    // None is a gallery entry without a style to edit.
    const none = styleId === '';
    if (kind === 'table' && none) return;
    event.preventDefault();
    event.stopPropagation();
    close(false);
    const from =
      kind === 'pivot'
        ? 'pivotQuickStyles'
        : item.dataset.tableQuickStyle !== undefined || item.closest('#menu-table-quick-styles')
          ? 'tableQuickStyles'
          : 'formatTableHome';
    const t = instance.i18n.strings.ribbonMenu;
    const custom = customStyleOf(instance, kind, styleId) !== null;
    const menu = document.createElement('div');
    menu.className = 'fc-tb__menu fc-tb__tablestyle-ctx';
    menu.setAttribute('role', 'menu');
    menu.style.position = 'fixed';
    menu.style.left = `${event.clientX}px`;
    menu.style.top = `${event.clientY}px`;
    const button = (action: Action, label: string, disabled = false): HTMLButtonElement => {
      const node = menuTextChip({
        label,
        attr: 'tableStyleContext',
        value: action,
        className: 'fc-tb__menu-item',
        labelClassName: 'fc-tb__menu-item__text',
      });
      projectDisabledState(node, disabled, null);
      node.addEventListener('click', () => {
        close(false);
        run(instance, kind, action, styleId, from);
      });
      return node;
    };
    // Built-in styles cannot be modified or deleted; the defaults need an engine that keeps them.
    menu.append(
      button('applyMaintain', t.tableStyleApplyMaintain),
      button('applyClear', t.tableStyleApplyClear),
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
    document.addEventListener('pointerdown', onOutside, true);
    document.addEventListener('keydown', onKey, true);
    menu.querySelector<HTMLButtonElement>('button:not(:disabled)')?.focus({ preventScroll: true });
  };

  host.addEventListener('contextmenu', onContextMenu, true);
  return () => {
    close(false);
    host.removeEventListener('contextmenu', onContextMenu, true);
  };
}
