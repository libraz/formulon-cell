// Home tab Borders dropdown. Reuses the SVG preview factories and preset table
// from ribbon/border-icons.ts, and renders three flavors of menu items:
// preset → fire `data-border-preset`, draw action → fire `data-border-draw`,
// submenu trigger → toggle a nested line-color / line-style submenu beside the
// main dropdown. The line-color submenu hosts a shared color palette; picks
// update the playground-owned `selectedBorderColor` via `onPickColor`.

import { createColorPalette, type ToolbarText } from '../../../index.js';

import { RIBBON_BORDERS_MENU_ID } from '../activation.js';
import {
  BORDER_PRESETS,
  type BorderPreviewSpec,
  createBorderEraserPreview,
  createBorderLineColorPreview,
  createBorderLineStylePreview,
  createBorderPreview,
  createLineSamplePreview,
  LINE_STYLES_ALL,
} from '../border-icons.js';
import {
  createMenu,
  createMenuButton,
  createSubmenu,
  menuIconSpacer,
  menuPresetButton,
  menuSectionHeader,
  menuSeparator,
  menuSubmenuTrigger,
  submenuItemText,
} from './general.js';

export interface BordersMenuDeps {
  ribbonText: ToolbarText;
  getBorderColor: () => string;
  onPickColor: (color: string) => void;
}

const presetMenuItem = (presetKey: string, label: string): HTMLButtonElement => {
  const spec = BORDER_PRESETS[presetKey];
  const leading = spec ? createBorderPreview(spec) : menuIconSpacer();
  return menuPresetButton(label, 'borderPreset', presetKey, leading);
};

const drawActionItem = (
  action: string,
  label: string,
  icon?: BorderPreviewSpec,
  leading?: Node,
): HTMLButtonElement => {
  const btn = menuPresetButton(
    label,
    'borderDraw',
    action,
    leading ?? (icon ? createBorderPreview(icon) : menuIconSpacer()),
  );
  btn.setAttribute('role', 'menuitemcheckbox');
  btn.setAttribute('aria-checked', 'false');
  return btn;
};

const submenuTrigger = (
  submenuKey: 'lineColor' | 'lineStyle',
  label: string,
  leading: Node = menuIconSpacer(),
): HTMLButtonElement => {
  const btn = menuPresetButton(label, 'borderSubmenu', submenuKey, leading);
  return menuSubmenuTrigger(btn, undefined, { controlsId: borderSubmenuId(submenuKey) });
};

const borderSubmenuId = (submenuKey: 'lineColor' | 'lineStyle'): string =>
  submenuKey === 'lineColor' ? 'menu-borders-line-color' : 'menu-borders-line-style';

const createLineStyleSubmenu = (label: string, noneLabel: string): HTMLDivElement => {
  const submenu = createSubmenu({
    id: borderSubmenuId('lineStyle'),
    className: 'app__submenu app__submenu--line-style',
    label,
  });
  for (const value of LINE_STYLES_ALL) {
    const btn = createMenuButton({
      className: 'app__submenu-item',
      attr: 'borderLineStyle',
      value,
    });
    btn.setAttribute('role', 'menuitemradio');
    btn.setAttribute('aria-checked', value === 'thin' ? 'true' : 'false');
    if (value === 'none') {
      btn.classList.add('app__submenu-item--line-style-none');
      btn.appendChild(submenuItemText(noneLabel));
    } else {
      btn.appendChild(createLineSamplePreview(value));
    }
    submenu.appendChild(btn);
  }
  return submenu;
};

const createLineColorSubmenu = (label: string, deps: BordersMenuDeps): HTMLDivElement => {
  const submenu = createSubmenu({
    id: borderSubmenuId('lineColor'),
    className: 'app__submenu app__submenu--line-color',
    label,
  });
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
  const menu = createMenu(RIBBON_BORDERS_MENU_ID);
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
    drawActionItem('erase', t.eraseBorder, undefined, createBorderEraserPreview()),
    submenuTrigger('lineColor', t.lineColor, createBorderLineColorPreview()),
    submenuTrigger('lineStyle', t.lineStyle, createBorderLineStylePreview()),
    menuSeparator(),
    // Footer
    presetMenuItem('format', t.moreBorders),
  );
  // Submenus sit beside the main dropdown.
  menu.appendChild(createLineColorSubmenu(t.lineColor, deps));
  menu.appendChild(createLineStyleSubmenu(t.lineStyle, t.lineStyleNone));
  return menu;
};
