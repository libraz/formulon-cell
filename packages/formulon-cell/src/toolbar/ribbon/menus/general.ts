// Generic menu primitives shared across all ribbon dropdown factories.
// Extracted from main.ts so the per-tab menu modules can build off the same
// building blocks without dragging the whole playground entry along with them.

import { prepareMenu } from '../../menu-a11y.js';

export const createMenu = (id: string): HTMLDivElement => {
  const menu = document.createElement('div');
  menu.className = 'app__menu';
  menu.id = id;
  menu.hidden = true;
  prepareMenu(menu);
  return menu;
};

export const menuButton = (label: string, attr: string, value: string): HTMLButtonElement => {
  const button = document.createElement('button');
  button.className = 'app__menu-item';
  button.type = 'button';
  button.setAttribute('role', 'menuitem');
  button.dataset[attr] = value;
  button.textContent = label;
  return button;
};

export const menuSeparator = (): HTMLDivElement => {
  const sep = document.createElement('div');
  sep.className = 'app__menu-sep';
  sep.setAttribute('role', 'separator');
  return sep;
};

export const menuSectionHeader = (label: string): HTMLDivElement => {
  const el = document.createElement('div');
  el.className = 'app__menu-heading';
  el.setAttribute('role', 'presentation');
  el.textContent = label;
  return el;
};
