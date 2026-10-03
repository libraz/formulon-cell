import { createInteractionButton } from './chip-button.js';
import type { ContextMenuItem } from './context-menu-options.js';
import { type ItemId, isBuiltinItemId, SUBMENU_ICON_ACTION } from './context-menu-spec.js';

/** Rendering entries deliberately carry a string ID. Built-ins are narrowed
 * at execution time; host-defined IDs must never fall through that switch. */
export type RenderMenuEntry =
  | {
      kind: 'item';
      id: string;
      label: string;
      hint?: string;
      source?: ContextMenuItem;
    }
  | {
      kind: 'submenu';
      id: string;
      label: string;
      children: RenderMenuEntry[];
      source?: ContextMenuItem;
    }
  | { kind: 'sep'; id: string; source?: ContextMenuItem };

const isMenuSeparator = (item: ContextMenuItem): boolean => item.separator === true;

export const contextItemToRenderEntry = (item: ContextMenuItem): RenderMenuEntry => {
  if (isMenuSeparator(item)) return { kind: 'sep', id: item.id, source: item };
  if (item.children && item.children.length > 0) {
    return {
      kind: 'submenu',
      id: item.id,
      label: item.label ?? item.id,
      children: item.children.map(contextItemToRenderEntry),
      source: item,
    };
  }
  return {
    kind: 'item',
    id: item.id,
    label: item.label ?? item.id,
    ...(item.hint === undefined ? {} : { hint: item.hint }),
    source: item,
  };
};

export const contextItemIds = (items: readonly ContextMenuItem[]): Set<string> => {
  const ids = new Set<string>();
  const visit = (item: ContextMenuItem): void => {
    ids.add(item.id);
    if (item.builtIn) ids.add(item.builtIn);
    item.children?.forEach(visit);
  };
  items.forEach(visit);
  return ids;
};

const compactContextItems = (items: readonly ContextMenuItem[]): ContextMenuItem[] => {
  const out: ContextMenuItem[] = [];
  for (const item of items) {
    if (item.separator) {
      if (out.length === 0 || out[out.length - 1]?.separator) continue;
      out.push(item);
      continue;
    }
    const children = item.children ? compactContextItems(item.children) : undefined;
    if (item.children && children && children.length === 0) continue;
    out.push(children ? { ...item, children } : item);
  }
  while (out[out.length - 1]?.separator) out.pop();
  return out;
};

/** Drop duplicate IDs and reserved built-in spoofs while preserving order. */
export const sanitizeContextItems = (items: readonly ContextMenuItem[]): ContextMenuItem[] => {
  const seen = new Set<string>();
  const visit = (item: ContextMenuItem): ContextMenuItem | null => {
    if (seen.has(item.id)) return null;
    if (isBuiltinItemId(item.id) && item.builtIn !== item.id) return null;
    if (item.builtIn !== undefined && !isBuiltinItemId(item.builtIn)) return null;
    seen.add(item.id);
    const children = item.children
      ?.map(visit)
      .filter((child): child is ContextMenuItem => child !== null);
    if (item.children && (!children || children.length === 0)) return null;
    return children ? { ...item, children } : item;
  };
  return items.map(visit).filter((item): item is ContextMenuItem => item !== null);
};

export const selectBuiltInItems = (
  items: readonly ContextMenuItem[],
  selected: ReadonlySet<string>,
): ContextMenuItem[] => {
  const visit = (item: ContextMenuItem): ContextMenuItem | null => {
    if (item.separator) return item;
    if (item.builtIn && selected.has(item.builtIn)) return item;
    if (selected.has(item.id)) return item;
    if (item.children) {
      const children = item.children
        .map(visit)
        .filter((child): child is ContextMenuItem => child !== null);
      return children.length > 0 ? { ...item, children } : null;
    }
    return null;
  };
  return compactContextItems(
    items.map(visit).filter((item): item is ContextMenuItem => item !== null),
  );
};

const appendContextMenuLabel = (
  button: HTMLButtonElement,
  labelText: string,
  hintText?: string,
): void => {
  const label = button.ownerDocument.createElement('span');
  label.className = 'fc-ctxmenu__label';
  label.textContent = labelText;
  button.appendChild(label);
  if (hintText !== undefined) {
    const hint = button.ownerDocument.createElement('span');
    hint.className = 'fc-ctxmenu__hint';
    hint.textContent = hintText;
    button.appendChild(hint);
  }
};

export const createContextMenuItemButton = (
  entry: Extract<RenderMenuEntry, { kind: 'item' }>,
): HTMLButtonElement => {
  const button = createInteractionButton({
    className: 'fc-ctxmenu__item',
    dataset: { fcAction: entry.id },
    role: 'menuitem',
    tabIndex: -1,
  });
  appendContextMenuLabel(button, entry.label, entry.hint ?? '');
  return button;
};

export const createContextSubmenuButton = (
  entry: Extract<RenderMenuEntry, { kind: 'submenu' }>,
): HTMLButtonElement => {
  const button = createInteractionButton({
    className: 'fc-ctxmenu__item fc-ctxmenu__item--parent',
    dataset: {
      fcSubmenu: entry.id,
      fcAction: SUBMENU_ICON_ACTION[entry.id] ?? entry.id,
    },
    role: 'menuitem',
    tabIndex: -1,
  });
  button.setAttribute('aria-haspopup', 'menu');
  button.setAttribute('aria-expanded', 'false');
  appendContextMenuLabel(button, entry.label);
  const arrow = button.ownerDocument.createElement('span');
  arrow.className = 'fc-ctxmenu__arrow';
  arrow.setAttribute('aria-hidden', 'true');
  button.appendChild(arrow);
  return button;
};

export const createContextMiniToolbarButton = (item: {
  id: ItemId;
  label: string;
}): HTMLButtonElement => {
  return createInteractionButton({
    className: 'fc-ctxmenu__mini-btn',
    dataset: { fcAction: item.id },
    ariaLabel: item.label,
    tabIndex: -1,
  });
};
