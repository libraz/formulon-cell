// Shape Format tab menus: Shape Fill and Shape Outline. The parent checks the
// selected shape's colour just before the menu opens.

import { createColorPalette } from '../../../components/color-palette.js';
import type { ToolbarMenuText } from '../../menu-text.js';
import type { ToolbarText } from '../../ribbon-model.js';

import { createMenu, menuIconSpacer, menuPresetButton, menuSeparator } from './general.js';

export interface ShapeFormatMenuFactories {
  createShapeFillMenu: () => HTMLDivElement;
  createShapeOutlineMenu: () => HTMLDivElement;
}

export const createShapeFormatMenuFactories = (
  ribbonText: ToolbarText,
  ribbonMenuText: ToolbarMenuText,
): ShapeFormatMenuFactories => {
  /** The colour palette, then the entry that takes the colour away. Every
   *  swatch carries its colour as the menu value. */
  const colorMenu = (
    id: string,
    attr: 'shapeFill' | 'shapeOutline',
    title: string,
    noneLabel: string,
  ): HTMLDivElement => {
    const menu = createMenu(id);
    const palette = createColorPalette({
      themeLabel: ribbonText.themeColors,
      standardLabel: ribbonText.standardColors,
      ariaLabel: title,
      onPick: () => undefined,
    });
    for (const swatch of palette.el.querySelectorAll<HTMLButtonElement>('[data-color]')) {
      swatch.dataset[attr] = swatch.dataset.color;
    }
    const none = menuPresetButton(noneLabel, attr, 'none', menuIconSpacer());
    none.setAttribute('role', 'menuitemradio');
    none.setAttribute('aria-checked', 'false');
    menu.append(palette.el, menuSeparator(), none);
    return menu;
  };

  return {
    createShapeFillMenu: () =>
      colorMenu('menu-shape-fill', 'shapeFill', ribbonText.shapeFill, ribbonMenuText.shapeNoFill),
    createShapeOutlineMenu: () =>
      colorMenu(
        'menu-shape-outline',
        'shapeOutline',
        ribbonText.shapeOutline,
        ribbonMenuText.shapeNoOutline,
      ),
  };
};
