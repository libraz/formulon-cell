// Home tab Borders dropdown. Reuses the SVG preview factories and preset table
// from ribbon/border-icons.ts, and renders three flavors of menu items:
// preset → fire `data-border-preset`, draw action → fire `data-border-draw`,
// submenu trigger → toggle a nested line-color / line-style submenu beside the
// main dropdown. The line-color submenu hosts a shared color palette; picks
// update the playground-owned `selectedBorderColor` via `onPickColor`.

import { createColorPalette, type ToolbarText } from '@libraz/formulon-cell';

import {
  BORDER_PRESETS,
  type BorderPreviewSpec,
  createBorderPreview,
  createLineSamplePreview,
  LINE_STYLES_ALL,
} from '../border-icons.js';
import { createMenu, menuSectionHeader, menuSeparator } from './general.js';

export interface BordersMenuDeps {
  ribbonText: ToolbarText;
  getBorderColor: () => string;
  onPickColor: (color: string) => void;
}

const presetMenuItem = (presetKey: string, label: string): HTMLButtonElement => {
  const btn = document.createElement('button');
  btn.className = 'app__menu-item app__menu-item--preset';
  btn.type = 'button';
  btn.setAttribute('role', 'menuitem');
  btn.dataset.borderPreset = presetKey;
  const spec = BORDER_PRESETS[presetKey];
  if (spec) btn.appendChild(createBorderPreview(spec));
  else {
    const spacer = document.createElement('span');
    spacer.className = 'app__menu-item__icon-spacer';
    btn.appendChild(spacer);
  }
  const text = document.createElement('span');
  text.className = 'app__menu-item__text';
  text.textContent = label;
  btn.appendChild(text);
  return btn;
};

const drawActionItem = (
  action: string,
  label: string,
  icon?: BorderPreviewSpec,
): HTMLButtonElement => {
  const btn = document.createElement('button');
  btn.className = 'app__menu-item app__menu-item--preset';
  btn.type = 'button';
  btn.setAttribute('role', 'menuitemcheckbox');
  btn.setAttribute('aria-checked', 'false');
  btn.dataset.borderDraw = action;
  if (icon) btn.appendChild(createBorderPreview(icon));
  else {
    const spacer = document.createElement('span');
    spacer.className = 'app__menu-item__icon-spacer';
    btn.appendChild(spacer);
  }
  const text = document.createElement('span');
  text.className = 'app__menu-item__text';
  text.textContent = label;
  btn.appendChild(text);
  return btn;
};

const submenuTrigger = (
  submenuKey: 'lineColor' | 'lineStyle',
  label: string,
): HTMLButtonElement => {
  const btn = document.createElement('button');
  btn.className = 'app__menu-item app__menu-item--preset app__menu-item--submenu';
  btn.type = 'button';
  btn.setAttribute('role', 'menuitem');
  btn.setAttribute('aria-haspopup', 'menu');
  btn.setAttribute('aria-expanded', 'false');
  btn.dataset.borderSubmenu = submenuKey;
  const spacer = document.createElement('span');
  spacer.className = 'app__menu-item__icon-spacer';
  btn.appendChild(spacer);
  const text = document.createElement('span');
  text.className = 'app__menu-item__text';
  text.textContent = label;
  btn.appendChild(text);
  const caret = document.createElement('span');
  caret.className = 'app__menu-item__caret';
  caret.setAttribute('aria-hidden', 'true');
  caret.textContent = '▶';
  btn.appendChild(caret);
  return btn;
};

const createLineStyleSubmenu = (label: string, noneLabel: string): HTMLDivElement => {
  const submenu = document.createElement('div');
  submenu.className = 'app__submenu app__submenu--line-style';
  submenu.id = 'menu-borders-line-style';
  submenu.setAttribute('role', 'menu');
  submenu.setAttribute('aria-label', label);
  submenu.hidden = true;
  for (const value of LINE_STYLES_ALL) {
    const btn = document.createElement('button');
    btn.className = 'app__submenu-item';
    btn.type = 'button';
    btn.setAttribute('role', 'menuitemradio');
    btn.setAttribute('aria-checked', value === 'thin' ? 'true' : 'false');
    btn.dataset.borderLineStyle = value;
    if (value === 'none') {
      const span = document.createElement('span');
      span.textContent = noneLabel;
      span.className = 'app__submenu-item__text';
      btn.appendChild(span);
    } else {
      btn.appendChild(createLineSamplePreview(value));
    }
    submenu.appendChild(btn);
  }
  return submenu;
};

const createLineColorSubmenu = (label: string, deps: BordersMenuDeps): HTMLDivElement => {
  const submenu = document.createElement('div');
  submenu.className = 'app__submenu app__submenu--line-color';
  submenu.id = 'menu-borders-line-color';
  submenu.setAttribute('role', 'menu');
  submenu.setAttribute('aria-label', label);
  submenu.hidden = true;
  const palette = createColorPalette({
    themeLabel: deps.ribbonText.themeColors,
    standardLabel: deps.ribbonText.standardColors,
    ariaLabel: label,
    value: deps.getBorderColor(),
    automatic: { label: deps.ribbonText.automatic, color: '#000000' },
    onPick: (color) => deps.onPickColor(color),
  });
  submenu.appendChild(palette.el);
  return submenu;
};

export const createBordersMenu = (deps: BordersMenuDeps): HTMLDivElement => {
  const t = deps.ribbonText;
  const menu = createMenu('menu-borders');
  menu.classList.add('app__menu--borders');
  menu.append(
    // Section 1: single-side edges
    presetMenuItem('bottom', t.bottomBorder),
    presetMenuItem('top', t.topBorder),
    presetMenuItem('left', t.leftBorder),
    presetMenuItem('right', t.rightBorder),
    menuSeparator(),
    // Section 2: frame presets
    presetMenuItem('clear', t.noBorder),
    presetMenuItem('all', t.allBorders),
    presetMenuItem('outline', t.outsideBorders),
    presetMenuItem('thickOutline', t.thickOutsideBorders),
    menuSeparator(),
    // Section 3: combined
    presetMenuItem('doubleBottom', t.doubleBottomBorder),
    presetMenuItem('thickBottom', t.thickBottomBorder),
    presetMenuItem('topAndBottom', t.topAndBottomBorder),
    presetMenuItem('topAndThickBottom', t.topAndThickBottomBorder),
    presetMenuItem('topAndDoubleBottom', t.topAndDoubleBottomBorder),
    // Section 4: heading + draw actions
    menuSectionHeader(t.drawBordersHeading),
    drawActionItem('draw', t.drawBorder, { bottom: 'thin' }),
    drawActionItem('grid', t.drawBorderGrid, {
      top: 'thin',
      right: 'thin',
      bottom: 'thin',
      left: 'thin',
      innerGrid: true,
      showBase: false,
    }),
    drawActionItem('erase', t.eraseBorder),
    submenuTrigger('lineColor', t.lineColor),
    submenuTrigger('lineStyle', t.lineStyle),
    menuSeparator(),
    // Footer
    presetMenuItem('format', t.moreBorders),
  );
  // Submenus sit beside the main dropdown.
  menu.appendChild(createLineColorSubmenu(t.lineColor, deps));
  menu.appendChild(createLineStyleSubmenu(t.lineStyle, t.lineStyleNone));
  return menu;
};
