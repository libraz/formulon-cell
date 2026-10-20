// Chart Design tab menus: Add Chart Element and Change Colors. The parent
// sets each entry's checked and disabled state for the selected chart just
// before the menu opens.

import { CHART_PALETTE_PRESETS, type ChartPalettePreset } from '../../../charts/series-colors.js';
import type { ChartElementAction } from '../../../commands/chart-design.js';
import type { ToolbarMenuText } from '../../menu-text.js';

import {
  createMenu,
  menuIconSpacer,
  menuPresetButton,
  menuSectionHeader,
  menuSeparator,
} from './general.js';

export interface ChartDesignMenuFactories {
  createChartAddElementMenu: () => HTMLDivElement;
  createChartColorsMenu: () => HTMLDivElement;
}

/** A menu entry that shows whether the chart already looks the way it makes it. */
const checkedItem = (
  label: string,
  attr: string,
  value: string,
  role: 'menuitemcheckbox' | 'menuitemradio',
  leading: Node = menuIconSpacer(),
): HTMLButtonElement => {
  const button = menuPresetButton(label, attr, value, leading);
  button.setAttribute('role', role);
  button.setAttribute('aria-checked', 'false');
  return button;
};

/** The six colours of a palette, side by side. */
const paletteSwatches = (preset: ChartPalettePreset): HTMLSpanElement => {
  const strip = document.createElement('span');
  strip.className = 'fc-tb__chart-palette';
  strip.setAttribute('aria-hidden', 'true');
  for (const color of CHART_PALETTE_PRESETS[preset]) {
    const swatch = document.createElement('span');
    swatch.className = 'fc-tb__chart-palette-swatch';
    swatch.style.backgroundColor = color;
    strip.appendChild(swatch);
  }
  return strip;
};

export const createChartDesignMenuFactories = (
  ribbonMenuText: ToolbarMenuText,
): ChartDesignMenuFactories => {
  const t = ribbonMenuText;
  const element = (
    label: string,
    action: ChartElementAction,
    role: 'menuitemcheckbox' | 'menuitemradio',
  ): HTMLButtonElement => checkedItem(label, 'chartElement', action, role);

  const createChartAddElementMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-chart-add-element');
    menu.append(
      menuSectionHeader(t.chartElementAxisTitles),
      element(t.chartElementPrimaryHorizontal, 'axis-title-horizontal', 'menuitemcheckbox'),
      element(t.chartElementPrimaryVertical, 'axis-title-vertical', 'menuitemcheckbox'),
      menuSectionHeader(t.chartElementChartTitle),
      element(t.chartElementNone, 'title-none', 'menuitemradio'),
      element(t.chartElementAboveChart, 'title-above', 'menuitemradio'),
      menuSectionHeader(t.chartElementDataLabels),
      element(t.chartElementNone, 'labels-none', 'menuitemradio'),
      element(t.chartElementShow, 'labels-show', 'menuitemradio'),
      menuSectionHeader(t.chartElementGridlines),
      element(t.chartElementMajorHorizontal, 'gridlines-horizontal', 'menuitemcheckbox'),
      element(t.chartElementMajorVertical, 'gridlines-vertical', 'menuitemcheckbox'),
      menuSectionHeader(t.chartElementLegend),
      element(t.chartElementNone, 'legend-none', 'menuitemradio'),
      element(t.chartElementRight, 'legend-right', 'menuitemradio'),
      element(t.chartElementTop, 'legend-top', 'menuitemradio'),
      element(t.chartElementLeft, 'legend-left', 'menuitemradio'),
      element(t.chartElementBottom, 'legend-bottom', 'menuitemradio'),
      menuSeparator(),
      menuPresetButton(t.chartElementMoreOptions, 'chartElement', 'more', menuIconSpacer()),
    );
    return menu;
  };

  const createChartColorsMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-chart-colors');
    const colors = (label: string, preset: ChartPalettePreset): HTMLButtonElement =>
      checkedItem(label, 'chartColors', preset, 'menuitemradio', paletteSwatches(preset));
    menu.append(
      colors(t.chartColorsColorful, 'office'),
      colors(t.chartColorsBlue, 'blue'),
      colors(t.chartColorsMonochrome, 'mono'),
    );
    return menu;
  };

  return { createChartAddElementMenu, createChartColorsMenu };
};
