// Sparkline tab menus: Edit Data, the Style gallery, the two colour menus, Axis and Clear.

import { BlanksAs, type Color } from '@libraz/formulon';
import { SHAPE_LINE_WEIGHTS } from '../../../commands/shape-format.js';
import type { SparklineStylePreset } from '../../../commands/sparkline-presets.js';
import { createColorPalette } from '../../../components/color-palette.js';
import type { MacStrings } from '../../../i18n/strings/mac/types.js';
import type { ToolbarMenuText } from '../../menu-text.js';
import type { ToolbarText } from '../../ribbon-model.js';
import { createPenColorIcon } from '../border-icons.js';
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
import type { MenuVariant } from './mac-variants.js';
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

/** The accessible name of the Style preset at gallery `index`: its colour word before its name. */
export const sparklineStyleLabel = (t: ToolbarText['sparkline'], index: number): string => {
  const slot = SPARKLINE_STYLE_SLOTS[index];
  if (!slot) return '';
  const name = t.styleNames[slot.name].replace('{n}', String(slot.n));
  return `${t.styleColorNames[index]}, ${name}`;
};

/** A small line sparkline in a preset's colours, with its high, low, first and last points. */
export const sparklineStyleThumbnail = (
  preset: SparklineStylePreset,
  css: (color: Color) => string,
): SVGSVGElement => {
  const ns = 'http://www.w3.org/2000/svg';
  const svg = document.createElementNS(ns, 'svg');
  svg.setAttribute('class', 'fc-tb__sparkline-style__thumb');
  svg.setAttribute('viewBox', '0 0 28 16');
  svg.setAttribute('aria-hidden', 'true');
  const points: readonly [number, number][] = [
    [3, 10],
    [8, 4],
    [13, 8],
    [18, 13],
    [25, 6],
  ];
  const line = document.createElementNS(ns, 'polyline');
  line.setAttribute('points', points.map(([x, y]) => `${x},${y}`).join(' '));
  line.setAttribute('fill', 'none');
  line.setAttribute('stroke', css(preset.series));
  line.setAttribute('stroke-width', '1.5');
  svg.appendChild(line);
  const marks: readonly [[number, number] | undefined, Color][] = [
    [points[0], preset.first],
    [points[1], preset.high],
    [points[3], preset.low],
    [points[4], preset.last],
  ];
  for (const [point, color] of marks) {
    if (!point) continue;
    const dot = document.createElementNS(ns, 'circle');
    dot.setAttribute('cx', String(point[0]));
    dot.setAttribute('cy', String(point[1]));
    dot.setAttribute('r', '1.8');
    dot.setAttribute('fill', css(color));
    svg.appendChild(dot);
  }
  return svg;
};

/** The menu value of a colour palette's More Colors row. */
export const MORE_COLORS = 'more';

export interface SparklineMenuFactories {
  createSparklineEditDataMenu: (commandId?: string, variant?: MenuVariant) => HTMLDivElement;
  createSparklineStyleMenu: () => HTMLDivElement;
  createSparklineColorMenu: () => HTMLDivElement;
  createSparklineMarkerColorMenu: () => HTMLDivElement;
  createSparklineAxisMenu: () => HTMLDivElement;
  createSparklineClearMenu: () => HTMLDivElement;
}

export const createSparklineMenuFactories = (
  ribbonText: ToolbarText,
  ribbonMenuText: ToolbarMenuText,
  macText: MacStrings['sparkline'],
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
      highContrastOnlyLabel: ribbonText.highContrastOnly,
      moreColorsLabel: ribbonText.moreColors,
      ariaLabel: title,
      onPick: () => undefined,
    });
    for (const swatch of colors.el.querySelectorAll<HTMLButtonElement>('[data-color]')) {
      swatch.dataset[attr] = `${prefix}${swatch.dataset.color}`;
    }
    // More Colors runs through the menu like a swatch; the defaults open the system picker.
    const more = colors.el.querySelector<HTMLButtonElement>('.fc-colorpalette__action--more');
    if (more) more.dataset[attr] = `${prefix}${MORE_COLORS}`;
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

  // Hidden & Empty Cells: on the Mac a submenu of the settings, which a click opens as hover does;
  // elsewhere an entry that opens the Hidden and Empty Cell Settings dialog.
  const createSparklineEditDataMenu = (
    _commandId?: string,
    variant?: MenuVariant,
  ): HTMLDivElement => {
    const menu = createMenu(SPARKLINE_MENU_IDS.editData);
    const empty = (label: string, value: string) =>
      checkable(item(label, 'sparklineEmpty', value), 'menuitemradio');
    const hidden: Node[] = variant?.mac
      ? submenu(SPARKLINE_MENU_IDS.editData, 'hidden-empty', macText.hiddenEmptyMenu, [
          empty(t.emptyAsGaps, String(BlanksAs.Gap)),
          empty(t.emptyAsZero, String(BlanksAs.Zero)),
          empty(t.connectPoints, String(BlanksAs.Span)),
          menuSeparator(),
          checkable(item(t.showHiddenData, 'sparklineEmpty', 'hidden'), 'menuitemcheckbox'),
        ])
      : [item(t.hiddenEmpty, 'sparklineEdit', 'hidden-empty')];
    menu.append(
      item(t.editGroup, 'sparklineEdit', 'group'),
      item(t.editSingle, 'sparklineEdit', 'single'),
      menuSeparator(),
      ...hidden,
      item(t.switchRowCol, 'sparklineEdit', 'switch'),
    );
    return menu;
  };

  const createSparklineStyleMenu = (): HTMLDivElement => {
    const menu = createMenu(SPARKLINE_MENU_IDS.style);
    const grid = colorSwatchGrid('fc-tb__sparkline-styles');
    SPARKLINE_STYLE_SLOTS.forEach((slot, index) => {
      const name = sparklineStyleLabel(t, index);
      const swatch = createMenuButton({
        className: 'fc-tb__sparkline-style',
        attr: 'sparklineStyle',
        value: slot.id,
        title: name,
        ariaLabel: name,
      });
      swatch.setAttribute('role', 'menuitemradio');
      swatch.setAttribute('aria-checked', 'false');
      grid.appendChild(swatch);
    });
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
        item(t.customWeight, 'sparklineWeight', 'custom'),
      ]),
    );
    return menu;
  };

  // Each point is a submenu whose icon shows the point's colour; Markers closes the first block.
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
      const [trigger, panel] = submenu(
        SPARKLINE_MENU_IDS.markerColor,
        part,
        labels[part],
        [palette(labels[part], 'sparklineMarkerColor', `${part}:`)],
        createPenColorIcon('fc-tb__sparkline-marker-icon'),
      );
      menu.append(trigger, panel);
      if (part === 'markers') menu.append(menuSeparator());
    }
    return menu;
  };

  const createSparklineAxisMenu = (): HTMLDivElement => {
    const menu = createMenu(SPARKLINE_MENU_IDS.axis);
    const radio = (label: string, value: SparklineAxisItem) =>
      checkable(item(label, 'sparklineAxis', value), 'menuitemradio');
    menu.append(
      menuSectionHeader(t.horizontalAxis),
      radio(t.generalAxis, 'general'),
      // Date Axis Type opens a dialog, so it checks like the other axis types.
      radio(t.dateAxis, 'date'),
      checkable(item(t.showAxis, 'sparklineAxis', 'show'), 'menuitemcheckbox'),
      checkable(item(t.plotRtl, 'sparklineAxis', 'rtl'), 'menuitemcheckbox'),
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
