// Shape Format tab menus: Shape Fill and Shape Outline. The parent checks the
// selected shape's colour, weight and dash just before the menu opens.

import {
  SHAPE_LINE_DASHES,
  SHAPE_LINE_WEIGHTS,
  shapeDashStroke,
} from '../../../commands/shape-format.js';
import { createColorPalette } from '../../../components/color-palette.js';
import type { SessionShapeLineDash } from '../../../store/store.js';
import type { ToolbarMenuText } from '../../menu-text.js';
import type { ToolbarText } from '../../ribbon-model.js';

import {
  createMenu,
  createSubmenu,
  menuIconSpacer,
  menuPresetButton,
  menuSeparator,
  menuSubmenuTrigger,
} from './general.js';

export interface ShapeFormatMenuFactories {
  createShapeFillMenu: () => HTMLDivElement;
  createShapeOutlineMenu: () => HTMLDivElement;
}

const SVG_NS = 'http://www.w3.org/2000/svg';
const SAMPLE_W = 64;
const SAMPLE_H = 12;
const PX_PER_PT = 96 / 72;

/** A weight in points the way the menu writes it: 2.25 is 2¼. */
const FRACTIONS: Record<string, string> = { '0.25': '¼', '0.5': '½', '0.75': '¾' };
const pointsLabel = (weight: number): string => {
  const whole = Math.trunc(weight);
  const fraction = FRACTIONS[String(weight - whole)] ?? '';
  return fraction ? `${whole || ''}${fraction}` : String(weight);
};

/** A horizontal line sample `weight` points thick in `dash`. */
const lineSample = (weight: number, dash: SessionShapeLineDash): SVGSVGElement => {
  const svg = document.createElementNS(SVG_NS, 'svg');
  svg.setAttribute('class', 'fc-tb__shape-line-sample');
  svg.setAttribute('viewBox', `0 0 ${SAMPLE_W} ${SAMPLE_H}`);
  svg.setAttribute('width', String(SAMPLE_W));
  svg.setAttribute('height', String(SAMPLE_H));
  svg.setAttribute('aria-hidden', 'true');
  const width = weight * PX_PER_PT;
  const stroke = shapeDashStroke(dash, width);
  const line = document.createElementNS(SVG_NS, 'line');
  for (const [key, value] of Object.entries({
    x1: stroke.linecap === 'round' ? width / 2 : 0,
    y1: SAMPLE_H / 2,
    x2: SAMPLE_W,
    y2: SAMPLE_H / 2,
    stroke: 'currentColor',
    'stroke-width': width,
    'stroke-dasharray': stroke.dasharray,
    'stroke-linecap': stroke.linecap,
  })) {
    line.setAttribute(key, String(value));
  }
  svg.appendChild(line);
  return svg;
};

const radioItem = (
  label: string,
  attr: string,
  value: string,
  leading: Node,
): HTMLButtonElement => {
  const item = menuPresetButton(label, attr, value, leading);
  item.setAttribute('role', 'menuitemradio');
  item.setAttribute('aria-checked', 'false');
  return item;
};

export const createShapeFormatMenuFactories = (
  ribbonText: ToolbarText,
  ribbonMenuText: ToolbarMenuText,
): ShapeFormatMenuFactories => {
  const t = ribbonMenuText;
  const DASH_LABELS: Record<SessionShapeLineDash, string> = {
    solid: t.shapeDashSolid,
    roundDot: t.shapeDashRoundDot,
    squareDot: t.shapeDashSquareDot,
    dash: t.shapeDashDash,
    dashDot: t.shapeDashDashDot,
    longDash: t.shapeDashLongDash,
    longDashDot: t.shapeDashLongDashDot,
    longDashDotDot: t.shapeDashLongDashDotDot,
  };

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
    menu.append(palette.el, menuSeparator(), radioItem(noneLabel, attr, 'none', menuIconSpacer()));
    return menu;
  };

  /** A Weight or Dashes entry and the submenu it opens beside the menu. */
  const outlineSubmenu = (
    key: 'weight' | 'dash',
    label: string,
    items: readonly HTMLButtonElement[],
  ): [HTMLButtonElement, HTMLDivElement] => {
    const id = `menu-shape-outline-${key}`;
    const trigger = menuSubmenuTrigger(
      menuPresetButton(label, 'formatSubmenu', key, menuIconSpacer()),
      undefined,
      { controlsId: id },
    );
    const submenu = createSubmenu({
      id,
      className: 'fc-tb__submenu fc-tb__submenu--format fc-tb__submenu--shape-outline',
      label,
      dataset: { formatPanel: key },
    });
    submenu.append(...items);
    return [trigger, submenu];
  };

  const createShapeOutlineMenu = (): HTMLDivElement => {
    const menu = colorMenu(
      'menu-shape-outline',
      'shapeOutline',
      ribbonText.shapeOutline,
      t.shapeNoOutline,
    );
    const weights = SHAPE_LINE_WEIGHTS.map((weight) =>
      radioItem(
        t.shapeOutlinePoints.replace('{n}', pointsLabel(weight)),
        'shapeOutlineWeight',
        String(weight),
        lineSample(weight, 'solid'),
      ),
    );
    const dashes = SHAPE_LINE_DASHES.map((dash) =>
      radioItem(DASH_LABELS[dash], 'shapeOutlineDash', dash, lineSample(1.5, dash)),
    );
    menu.append(
      menuSeparator(),
      ...outlineSubmenu('weight', t.shapeOutlineWeight, weights),
      ...outlineSubmenu('dash', t.shapeOutlineDashes, dashes),
    );
    return menu;
  };

  return {
    createShapeFillMenu: () =>
      colorMenu('menu-shape-fill', 'shapeFill', ribbonText.shapeFill, t.shapeNoFill),
    createShapeOutlineMenu,
  };
};
