// Generic menu primitives shared across all ribbon dropdown factories.
// Extracted from main.ts so the per-tab menu modules can build off the same
// building blocks without dragging the whole playground entry along with them.

import { prepareMenu } from '../../menu-a11y.js';
import { createRibbonButton } from '../button.js';
import { ribbonDropdownMenuIdForCommand } from '../dynamic-dropdowns.js';

const menuDiv = (
  className: string,
  opts: { id?: string; text?: string; role?: string; ariaLabel?: string; hidden?: boolean } = {},
): HTMLDivElement => {
  const div = document.createElement('div');
  div.className = className;
  if (opts.id) div.id = opts.id;
  if (opts.text !== undefined) div.textContent = opts.text;
  if (opts.role) div.setAttribute('role', opts.role);
  if (opts.ariaLabel) div.setAttribute('aria-label', opts.ariaLabel);
  if (opts.hidden) div.hidden = true;
  return div;
};

/** Maps a ribbon command id to the DOM id its dropdown menu should use.
 *  Falls back to the command id when no entry exists (callers can pass any
 *  string and get a stable id back). */
export const menuIdForCommand = (commandId: string): string =>
  ribbonDropdownMenuIdForCommand(commandId) ?? commandId;

export const createMenu = (id: string): HTMLDivElement => {
  const menu = menuDiv('app__menu', { id, hidden: true });
  prepareMenu(menu);
  return menu;
};

export type MenuButtonOptions = {
  className: string;
  attr: string;
  value: string;
  title?: string;
  ariaLabel?: string;
};

export const createMenuButton = (opts: MenuButtonOptions): HTMLButtonElement => {
  return createRibbonButton({
    className: opts.className,
    role: 'menuitem',
    dataset: { [opts.attr]: opts.value },
    title: opts.title,
    ariaLabel: opts.ariaLabel,
  });
};

const menuSpan = (
  className: string,
  opts: { text?: string; ariaHidden?: boolean } = {},
): HTMLSpanElement => {
  const span = document.createElement('span');
  span.className = className;
  if (opts.text !== undefined) span.textContent = opts.text;
  if (opts.ariaHidden) span.setAttribute('aria-hidden', 'true');
  return span;
};

export const menuIconButton = (
  label: string,
  attr: string,
  value: string,
  icon: string,
): HTMLButtonElement => {
  const button = createMenuButton({
    className: 'app__menu-item app__menu-item--iconic',
    attr,
    value,
  });

  button.append(
    menuSpan(`app__menu-icon app__menu-icon--${icon}`, { ariaHidden: true }),
    menuSpan('app__menu-item__text', { text: label }),
  );
  return button;
};

export const menuPresetButton = (
  label: string,
  attr: string,
  value: string,
  leading: Node,
): HTMLButtonElement => {
  const button = createMenuButton({
    className: 'app__menu-item app__menu-item--preset',
    attr,
    value,
  });

  button.append(leading, menuSpan('app__menu-item__text', { text: label }));
  return button;
};

export type MenuTextChipOptions = {
  label: string;
  attr: string;
  value: string;
  className: string;
  labelClassName?: string;
};

export const menuTextChip = (opts: MenuTextChipOptions): HTMLButtonElement => {
  const button = createMenuButton({
    className: opts.className,
    attr: opts.attr,
    value: opts.value,
    title: opts.label,
    ariaLabel: opts.label,
  });
  button.append(
    menuSpan(opts.labelClassName ?? 'app__menu-text-chip__label', { text: opts.label }),
  );
  return button;
};

export const menuIconSpacer = (): HTMLSpanElement => {
  return menuSpan('app__menu-item__icon-spacer');
};

export const menuSubmenuTrigger = (
  button: HTMLButtonElement,
  dataset?: Record<string, string>,
  opts: { controlsId?: string } = {},
): HTMLButtonElement => {
  button.classList.add('app__menu-item--submenu');
  button.setAttribute('aria-haspopup', 'menu');
  button.setAttribute('aria-expanded', 'false');
  if (opts.controlsId) button.setAttribute('aria-controls', opts.controlsId);
  for (const [key, value] of Object.entries(dataset ?? {})) button.dataset[key] = value;

  button.appendChild(menuSpan('app__menu-item__caret', { text: '▶', ariaHidden: true }));
  return button;
};

export type ColorSwatchButtonOptions = {
  label: string;
  attr: string;
  value: string;
  color: string | null;
};

export const colorSwatchButton = (opts: ColorSwatchButtonOptions): HTMLButtonElement => {
  const button = createMenuButton({
    className: ['app__color-swatch', opts.color === null ? 'app__color-swatch--none' : '']
      .filter(Boolean)
      .join(' '),
    attr: opts.attr,
    value: opts.value,
    title: opts.label,
    ariaLabel: opts.label,
  });
  if (opts.color !== null) button.style.setProperty('--app-menu-swatch-color', opts.color);

  button.append(menuSpan('app__color-swatch__chip', { ariaHidden: true }));
  return button;
};

export const colorSwatchGrid = (className?: string): HTMLDivElement => {
  return menuDiv(['app__color-swatch-grid', className].filter(Boolean).join(' '), {
    role: 'presentation',
  });
};

export const symbolMenuTile = (symbol: string): HTMLButtonElement => {
  const button = createMenuButton({
    className: 'app__symbol-tile',
    attr: 'symbol',
    value: symbol,
    title: symbol,
    ariaLabel: symbol,
  });

  button.append(menuSpan('app__symbol-tile__glyph', { text: symbol, ariaHidden: true }));
  return button;
};

export const symbolMenuGrid = (groupLabel: string, symbols: readonly string[]): HTMLDivElement => {
  const grid = menuDiv('app__symbol-grid', { role: 'presentation', ariaLabel: groupLabel });
  grid.append(...symbols.map((symbol) => symbolMenuTile(symbol)));
  return grid;
};

export type VisualMenuTileOptions = {
  label: string;
  attr: string;
  value: string;
  icon: string;
  className?: string;
};

export const visualMenuTile = (opts: VisualMenuTileOptions): HTMLButtonElement => {
  const button = createMenuButton({
    className: ['app__visual-tile', opts.className].filter(Boolean).join(' '),
    attr: opts.attr,
    value: opts.value,
    title: opts.label,
    ariaLabel: opts.label,
  });

  button.append(
    menuSpan(`app__visual-tile__icon app__visual-tile__icon--${opts.icon}`, {
      ariaHidden: true,
    }),
    menuSpan('app__visual-tile__label', { text: opts.label }),
  );
  return button;
};

export const visualMenuGrid = (className?: string): HTMLDivElement => {
  return menuDiv(['app__visual-grid', className].filter(Boolean).join(' '), {
    role: 'presentation',
  });
};

export const visualMenuTileGrid = (
  className: string,
  tiles: readonly VisualMenuTileOptions[],
): HTMLDivElement => {
  const grid = visualMenuGrid(className);
  grid.append(...tiles.map((tile) => visualMenuTile(tile)));
  return grid;
};

export const menuSeparator = (): HTMLDivElement => {
  return menuDiv('app__menu-sep', { role: 'separator' });
};

export const menuScrollBody = (className: string, ariaLabel?: string): HTMLDivElement => {
  return menuDiv(className, { role: 'group', ariaLabel });
};

export const menuSectionHeader = (label: string): HTMLDivElement => {
  return menuDiv('app__menu-heading', { role: 'presentation', text: label });
};

export type MenuLabeledGridOptions = {
  label: string;
  headingClassName: string;
  gridClassName: string;
  children: readonly Node[];
};

export const menuLabeledGrid = (opts: MenuLabeledGridOptions): [HTMLDivElement, HTMLDivElement] => {
  const heading = menuDiv(opts.headingClassName, { text: opts.label });
  const grid = menuDiv(opts.gridClassName, { role: 'group', ariaLabel: opts.label });
  grid.append(...opts.children);

  return [heading, grid];
};

export type SubmenuOptions = {
  id?: string;
  className: string;
  label: string;
  dataset?: Record<string, string>;
};

export const createSubmenu = (opts: SubmenuOptions): HTMLDivElement => {
  const submenu = menuDiv(opts.className, {
    id: opts.id,
    role: 'menu',
    ariaLabel: opts.label,
    hidden: true,
  });
  for (const [key, value] of Object.entries(opts.dataset ?? {})) submenu.dataset[key] = value;
  return submenu;
};

export const submenuItemText = (text: string): HTMLSpanElement => {
  return menuSpan('app__submenu-item__text', { text });
};
