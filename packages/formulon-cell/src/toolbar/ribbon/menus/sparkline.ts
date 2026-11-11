// Sparkline tab menus: Edit Data, the Style gallery, the two colour menus, Axis and Clear.

import { SHAPE_LINE_WEIGHTS } from '../../../commands/shape-format.js';
import { createColorPalette } from '../../../components/color-palette.js';
import { projectDisabledState } from '../../menu-a11y.js';
import type { ToolbarMenuText } from '../../menu-text.js';
import type { ToolbarText } from '../../ribbon-model.js';
import { SPARKLINE_MENU_IDS } from '../sparkline-tabs.js';
import {
  colorSwatchGrid,
  createMenu,
  createMenuButton,
  createSubmenu,
  menuIconSpacer,
  menuPresetButton,
  menuSectionHeader,
  menuSeparator,
  menuSubmenuTrigger,
} from './general.js';
import { lineSample, pointsLabel } from './shape-format.js';

/** The points a sparkline colours apart from its series, in menu order. */
export const SPARKLINE_MARKER_PARTS = [
  'negative',
  'markers',
  'high',
  'low',
  'first',
  'last',
] as const;
export type SparklineMarkerPart = (typeof SPARKLINE_MARKER_PARTS)[number];

/** Axis menu entries, by the value their button carries. */
export const SPARKLINE_AXIS_ITEMS = [
  'general',
  'date',
  'show',
  'rtl',
  'min-individual',
  'min-group',
  'min-custom',
  'max-individual',
  'max-group',
  'max-custom',
] as const;
export type SparklineAxisItem = (typeof SPARKLINE_AXIS_ITEMS)[number];

/** The preset families of the Style gallery, six accents each, then the two numbered sets. */
const STYLE_FAMILIES = ['accentDarker50', 'accentDarker25', 'accent', 'accentLighter40'] as const;
const STYLE_SET_SIZE = 6;

type StyleName = keyof ToolbarText['sparkline']['styleNames'];

/** One slot of the Style gallery: its name template and the number it fills in. */
export interface SparklineStyleSlot {
  id: string;
  name: StyleName;
  n: number;
}

export const SPARKLINE_STYLE_SLOTS: readonly SparklineStyleSlot[] = [
  ...STYLE_FAMILIES.flatMap((name) =>
    Array.from({ length: STYLE_SET_SIZE }, (_, i) => ({ id: `${name}-${i + 1}`, name, n: i + 1 })),
  ),
  ...(['dark', 'colorful'] as const).flatMap((name) =>
    Array.from({ length: STYLE_SET_SIZE }, (_, i) => ({ id: `${name}-${i + 1}`, name, n: i + 1 })),
  ),
];

/** The colours of each Style slot are measured data and have not been taken yet, so every slot
 *  stays disabled and this table stays empty. */
export const SPARKLINE_STYLE_PRESETS: Readonly<Record<string, never>> = {};

export interface SparklineMenuFactories {
  createSparklineEditDataMenu: () => HTMLDivElement;
  createSparklineStyleMenu: () => HTMLDivElement;
  createSparklineColorMenu: () => HTMLDivElement;
  createSparklineMarkerColorMenu: () => HTMLDivElement;
  createSparklineAxisMenu: () => HTMLDivElement;
  createSparklineClearMenu: () => HTMLDivElement;
}

export const createSparklineMenuFactories = (
  ribbonText: ToolbarText,
  ribbonMenuText: ToolbarMenuText,
): SparklineMenuFactories => {
  const t = ribbonText.sparkline;
  const item = (label: string, attr: string, value: string): HTMLButtonElement =>
    menuPresetButton(label, attr, value, menuIconSpacer());
  const checkable = (
    button: HTMLButtonElement,
    role: 'menuitemradio' | 'menuitemcheckbox',
  ): HTMLButtonElement => {
    button.setAttribute('role', role);
    button.setAttribute('aria-checked', 'false');
    return button;
  };

  const palette = (title: string, attr: string, prefix = ''): HTMLElement => {
    const colors = createColorPalette({
      themeLabel: ribbonText.themeColors,
      standardLabel: ribbonText.standardColors,
      ariaLabel: title,
      onPick: () => undefined,
    });
    for (const swatch of colors.el.querySelectorAll<HTMLButtonElement>('[data-color]')) {
      swatch.dataset[attr] = `${prefix}${swatch.dataset.color}`;
    }
    return colors.el;
  };

  /** A trigger and the panel it opens beside the menu. */
  const submenu = (
    menuId: string,
    key: string,
    label: string,
    children: readonly Node[],
    leading: Node = menuIconSpacer(),
  ): [HTMLButtonElement, HTMLDivElement] => {
    const id = `${menuId}-${key}`;
    const trigger = menuSubmenuTrigger(
      menuPresetButton(label, 'formatSubmenu', key, leading),
      undefined,
      { controlsId: id },
    );
    const panel = createSubmenu({
      id,
      className: 'fc-tb__submenu fc-tb__submenu--format',
      label,
      dataset: { formatPanel: key },
    });
    panel.append(...children);
    return [trigger, panel];
  };

  const createSparklineEditDataMenu = (): HTMLDivElement => {
    const menu = createMenu(SPARKLINE_MENU_IDS.editData);
    menu.append(
      item(t.editGroup, 'sparklineEdit', 'group'),
      item(t.editSingle, 'sparklineEdit', 'single'),
      item(t.hiddenEmpty, 'sparklineEdit', 'hidden-empty'),
      item(t.switchRowCol, 'sparklineEdit', 'switch'),
    );
    return menu;
  };

  const createSparklineStyleMenu = (): HTMLDivElement => {
    const menu = createMenu(SPARKLINE_MENU_IDS.style);
    const grid = colorSwatchGrid('fc-tb__sparkline-styles');
    for (const slot of SPARKLINE_STYLE_SLOTS) {
      const name = t.styleNames[slot.name].replace('{n}', String(slot.n));
      const swatch = createMenuButton({
        className: 'fc-tb__sparkline-style',
        attr: 'sparklineStyle',
        value: slot.id,
        title: name,
        ariaLabel: name,
      });
      projectDisabledState(swatch, true, null);
      grid.appendChild(swatch);
    }
    menu.appendChild(grid);
    return menu;
  };

  const createSparklineColorMenu = (): HTMLDivElement => {
    const menu = createMenu(SPARKLINE_MENU_IDS.color);
    const weights = SHAPE_LINE_WEIGHTS.map((weight) =>
      checkable(
        menuPresetButton(
          ribbonMenuText.shapeOutlinePoints.replace('{n}', pointsLabel(weight)),
          'sparklineWeight',
          String(weight),
          lineSample(weight, 'solid'),
        ),
        'menuitemradio',
      ),
    );
    menu.append(
      palette(t.color, 'sparklineColor'),
      menuSeparator(),
      ...submenu(SPARKLINE_MENU_IDS.color, 'weight', t.weight, [
        ...weights,
        menuSeparator(),
        item(t.customValue, 'sparklineWeight', 'custom'),
      ]),
    );
    return menu;
  };

  const createSparklineMarkerColorMenu = (): HTMLDivElement => {
    const menu = createMenu(SPARKLINE_MENU_IDS.markerColor);
    const labels: Record<SparklineMarkerPart, string> = {
      negative: t.negativePoints,
      markers: t.markers,
      high: t.highPoint,
      low: t.lowPoint,
      first: t.firstPoint,
      last: t.lastPoint,
    };
    for (const part of SPARKLINE_MARKER_PARTS) {
      menu.append(
        ...submenu(SPARKLINE_MENU_IDS.markerColor, part, labels[part], [
          palette(labels[part], 'sparklineMarkerColor', `${part}:`),
        ]),
      );
    }
    return menu;
  };

  const createSparklineAxisMenu = (): HTMLDivElement => {
    const menu = createMenu(SPARKLINE_MENU_IDS.axis);
    const radio = (label: string, value: SparklineAxisItem) =>
      checkable(item(label, 'sparklineAxis', value), 'menuitemradio');
    // The engine models no right-to-left plot, so the entry stays greyed.
    const rtl = checkable(item(t.plotRtl, 'sparklineAxis', 'rtl'), 'menuitemcheckbox');
    projectDisabledState(rtl, true, null);
    menu.append(
      menuSectionHeader(t.horizontalAxis),
      radio(t.generalAxis, 'general'),
      // Date Axis Type opens a dialog, so it checks like the other axis types.
      radio(t.dateAxis, 'date'),
      checkable(item(t.showAxis, 'sparklineAxis', 'show'), 'menuitemcheckbox'),
      rtl,
      menuSeparator(),
      menuSectionHeader(t.verticalMin),
      radio(t.autoEach, 'min-individual'),
      radio(t.sameAll, 'min-group'),
      radio(t.customValue, 'min-custom'),
      menuSeparator(),
      menuSectionHeader(t.verticalMax),
      radio(t.autoEach, 'max-individual'),
      radio(t.sameAll, 'max-group'),
      radio(t.customValue, 'max-custom'),
    );
    return menu;
  };

  const createSparklineClearMenu = (): HTMLDivElement => {
    const menu = createMenu(SPARKLINE_MENU_IDS.clear);
    menu.append(
      item(t.clearSelected, 'sparklineClear', 'selected'),
      item(t.clearGroups, 'sparklineClear', 'groups'),
    );
    return menu;
  };

  return {
    createSparklineEditDataMenu,
    createSparklineStyleMenu,
    createSparklineColorMenu,
    createSparklineMarkerColorMenu,
    createSparklineAxisMenu,
    createSparklineClearMenu,
  };
};
