// Conditional Formatting dropdown. Top-level menu lists hover-to-open submenu
// triggers (highlight / topBottom / dataBar / colorScale / iconSet / clear) and
// flat actions (newRule / manage). Each submenu panel renders the matching
// preset list — swatches for data bars and color scales, glyph rows for icon
// sets, plain menu items for highlight rules. All clicks emit a `data-cf-*`
// attribute the playground main dispatcher reads.

import { conditionalMenuText, type ToolbarLang } from '@libraz/formulon-cell';

import { createMenu, menuSectionHeader, menuSeparator } from './general.js';

export type CfSubmenuKey =
  | 'highlight'
  | 'topBottom'
  | 'dataBar'
  | 'colorScale'
  | 'iconSet'
  | 'clear';

type CfIconKind = 'rule' | 'top' | 'bar' | 'scale' | 'icon' | 'new' | 'clear' | 'manage';

const cfIcon = (kind: CfIconKind): HTMLSpanElement => {
  const span = document.createElement('span');
  span.className = `app__cf-icon app__cf-icon--${kind}`;
  span.setAttribute('aria-hidden', 'true');
  return span;
};

const cfMenuItem = (
  label: string,
  action: string,
  icon: CfIconKind = 'rule',
): HTMLButtonElement => {
  const btn = document.createElement('button');
  btn.className = 'app__menu-item app__menu-item--preset';
  btn.type = 'button';
  btn.setAttribute('role', 'menuitem');
  btn.dataset.cfAction = action;
  btn.appendChild(cfIcon(icon));
  const text = document.createElement('span');
  text.className = 'app__menu-item__text';
  text.textContent = label;
  btn.appendChild(text);
  return btn;
};

const cfSubmenuTrigger = (
  key: CfSubmenuKey,
  label: string,
  icon: CfIconKind,
): HTMLButtonElement => {
  const btn = cfMenuItem(label, `submenu-${key}`, icon);
  btn.classList.add('app__menu-item--submenu');
  btn.dataset.cfSubmenu = key;
  const caret = document.createElement('span');
  caret.className = 'app__menu-item__caret';
  caret.textContent = '▶';
  btn.appendChild(caret);
  return btn;
};

const cfSwatchButton = (action: string, colors: string[], title: string): HTMLButtonElement => {
  const btn = document.createElement('button');
  btn.className = 'app__cf-choice';
  btn.type = 'button';
  btn.title = title;
  btn.setAttribute('aria-label', title);
  btn.dataset.cfAction = action;
  const grid = document.createElement('span');
  grid.className = 'app__cf-choice-grid';
  for (const color of colors) {
    const cell = document.createElement('span');
    cell.style.background = color;
    grid.appendChild(cell);
  }
  btn.appendChild(grid);
  return btn;
};

const cfIconChoice = (action: string, glyphs: string[], title: string): HTMLButtonElement => {
  const btn = document.createElement('button');
  btn.className = 'app__cf-icon-choice';
  btn.type = 'button';
  btn.title = title;
  btn.setAttribute('aria-label', title);
  btn.dataset.cfAction = action;
  for (const glyph of glyphs) {
    const span = document.createElement('span');
    span.textContent = glyph;
    btn.appendChild(span);
  }
  return btn;
};

export const buildCfMenuText = (ribbonLang: ToolbarLang) => {
  const t = conditionalMenuText(ribbonLang);
  return {
    highlight: t.highlight,
    topBottom: t.topBottom,
    dataBar: t.dataBars,
    colorScale: t.colorScales,
    iconSet: t.iconSets,
    newRule: t.newRule,
    clear: t.clear,
    manage: t.manage,
    greater: t.greater,
    less: t.less,
    between: t.between,
    equal: t.equal,
    text: t.textContains,
    date: t.dateOccurring,
    duplicate: t.duplicates,
    unique: t.unique,
    top10: t.top10,
    bottom10: t.bottom10,
    top10Percent: t.top10Percent,
    bottom10Percent: t.bottom10Percent,
    aboveAvg: t.aboveAvg,
    belowAvg: t.belowAvg,
    textPrompt: t.textPrompt,
    datePrompt: t.datePrompt,
    otherRules: t.otherRules,
    gradient: t.gradientFill,
    solid: t.solidFill,
    direction: t.direction,
    shapes: t.shapes,
    indicators: t.indicators,
    ratings: t.ratings,
    flags: t.flags,
    bars: t.bars,
    clearSelection: t.clearSelection,
    clearSheet: t.clearSheet,
    dataBarGradientBlue: t.dataBarGradientBlue,
    dataBarGradientGreen: t.dataBarGradientGreen,
    dataBarGradientRed: t.dataBarGradientRed,
    dataBarGradientOrange: t.dataBarGradientOrange,
    dataBarGradientPurple: t.dataBarGradientPurple,
    dataBarGradientTeal: t.dataBarGradientTeal,
    dataBarSolidBlue: t.dataBarSolidBlue,
    dataBarSolidGreen: t.dataBarSolidGreen,
    dataBarSolidRed: t.dataBarSolidRed,
    dataBarSolidOrange: t.dataBarSolidOrange,
    dataBarSolidPurple: t.dataBarSolidPurple,
    dataBarSolidGray: t.dataBarSolidGray,
    colorScaleGreenYellowRed: t.colorScaleGreenYellowRed,
    colorScaleRedYellowGreen: t.colorScaleRedYellowGreen,
    colorScaleGreenWhite: t.colorScaleGreenWhite,
    colorScaleRedWhite: t.colorScaleRedWhite,
    colorScaleBlueWhiteRed: t.colorScaleBlueWhiteRed,
    colorScaleRedWhiteBlue: t.colorScaleRedWhiteBlue,
    colorScaleGreenWhiteGreen: t.colorScaleGreenWhiteGreen,
    colorScaleYellowWhiteGreen: t.colorScaleYellowWhiteGreen,
    colorScaleRedWhiteRed: t.colorScaleRedWhiteRed,
    colorScaleBlueWhiteBlue: t.colorScaleBlueWhiteBlue,
    colorScaleYellowRedGreen: t.colorScaleYellowRedGreen,
    colorScaleGreenYellowGreen: t.colorScaleGreenYellowGreen,
    iconArrows3: t.iconArrows3,
    iconArrows5: t.iconArrows5,
    iconTriangles3: t.iconTriangles3,
    iconTraffic3: t.iconTraffic3,
    iconTrafficRim3: t.iconTrafficRim3,
    iconSymbols3: t.iconSymbols3,
    iconFlags3: t.iconFlags3,
    iconStars3: t.iconStars3,
    iconQuarters5: t.iconQuarters5,
    iconRatings5: t.iconRatings5,
    iconBars5: t.iconBars5,
    iconBoxes5: t.iconBoxes5,
  };
};

type CfMenuText = ReturnType<typeof buildCfMenuText>;

const buildHighlightSubmenu = (submenu: HTMLDivElement, t: CfMenuText): void => {
  submenu.append(
    cfMenuItem(t.greater, 'cell-gt', 'rule'),
    cfMenuItem(t.less, 'cell-lt', 'rule'),
    cfMenuItem(t.between, 'cell-between', 'rule'),
    cfMenuItem(t.equal, 'cell-eq', 'rule'),
    cfMenuItem(t.text, 'text-contains', 'rule'),
    cfMenuItem(t.date, 'date-occurring', 'rule'),
    cfMenuItem(t.duplicate, 'duplicates', 'rule'),
    cfMenuItem(t.unique, 'unique', 'rule'),
    menuSeparator(),
    cfMenuItem(t.otherRules, 'new-rule', 'manage'),
  );
};

const buildTopBottomSubmenu = (submenu: HTMLDivElement, t: CfMenuText): void => {
  submenu.append(
    cfMenuItem(t.top10, 'top10', 'top'),
    cfMenuItem(t.bottom10, 'bottom10', 'top'),
    cfMenuItem(t.top10Percent, 'top10-percent', 'top'),
    cfMenuItem(t.bottom10Percent, 'bottom10-percent', 'top'),
    cfMenuItem(t.aboveAvg, 'above-avg', 'top'),
    cfMenuItem(t.belowAvg, 'below-avg', 'top'),
    menuSeparator(),
    cfMenuItem(t.otherRules, 'new-rule', 'manage'),
  );
};

const buildDataBarSubmenu = (submenu: HTMLDivElement, t: CfMenuText): void => {
  submenu.append(menuSectionHeader(t.gradient));
  const gradient = document.createElement('div');
  gradient.className = 'app__cf-choice-row';
  const gradientBars: readonly (readonly [string, string, string])[] = [
    ['data-blue', '#638ec6', t.dataBarGradientBlue],
    ['data-green', '#63a95c', t.dataBarGradientGreen],
    ['data-red', '#c45a5a', t.dataBarGradientRed],
    ['data-orange', '#d6a440', t.dataBarGradientOrange],
    ['data-purple', '#8a74b9', t.dataBarGradientPurple],
    ['data-teal', '#4ba1a8', t.dataBarGradientTeal],
  ];
  gradientBars.forEach(([action, color, label]) => {
    gradient.appendChild(cfSwatchButton(action, ['#fff', color], label));
  });
  submenu.appendChild(gradient);
  submenu.append(menuSectionHeader(t.solid));
  const solid = document.createElement('div');
  solid.className = 'app__cf-choice-row';
  const solidBars: readonly (readonly [string, string, string])[] = [
    ['data-solid-blue', '#4472c4', t.dataBarSolidBlue],
    ['data-solid-green', '#70ad47', t.dataBarSolidGreen],
    ['data-solid-red', '#c00000', t.dataBarSolidRed],
    ['data-solid-orange', '#ed7d31', t.dataBarSolidOrange],
    ['data-solid-purple', '#8064a2', t.dataBarSolidPurple],
    ['data-solid-gray', '#7f7f7f', t.dataBarSolidGray],
  ];
  solidBars.forEach(([action, color, label]) => {
    solid.appendChild(cfSwatchButton(action, [color, color], label));
  });
  submenu.append(solid, menuSeparator(), cfMenuItem(t.otherRules, 'new-rule', 'manage'));
};

const buildColorScaleSubmenu = (submenu: HTMLDivElement, t: CfMenuText): void => {
  const scales = document.createElement('div');
  scales.className = 'app__cf-choice-grid-panel';
  const colorScales: readonly (readonly [string, readonly string[], string])[] = [
    ['scale-gyr', ['#63be7b', '#ffeb84', '#f8696b'], t.colorScaleGreenYellowRed],
    ['scale-ryg', ['#f8696b', '#ffeb84', '#63be7b'], t.colorScaleRedYellowGreen],
    ['scale-gw', ['#63be7b', '#ffffff'], t.colorScaleGreenWhite],
    ['scale-rw', ['#f8696b', '#ffffff'], t.colorScaleRedWhite],
    ['scale-bwr', ['#5a8dee', '#ffffff', '#f8696b'], t.colorScaleBlueWhiteRed],
    ['scale-rwb', ['#f8696b', '#ffffff', '#5a8dee'], t.colorScaleRedWhiteBlue],
    ['scale-gwg', ['#63be7b', '#ffffff', '#00a651'], t.colorScaleGreenWhiteGreen],
    ['scale-ywg', ['#ffeb84', '#ffffff', '#63be7b'], t.colorScaleYellowWhiteGreen],
    ['scale-rwr', ['#f8696b', '#ffffff', '#c00000'], t.colorScaleRedWhiteRed],
    ['scale-bwb', ['#5a8dee', '#ffffff', '#4472c4'], t.colorScaleBlueWhiteBlue],
    ['scale-yry', ['#ffeb84', '#f8696b', '#63be7b'], t.colorScaleYellowRedGreen],
    ['scale-gyg', ['#63be7b', '#ffeb84', '#00a651'], t.colorScaleGreenYellowGreen],
  ];
  colorScales.forEach(([action, colors, label]) => {
    scales.appendChild(cfSwatchButton(action, [...colors], label));
  });
  submenu.append(scales, menuSeparator(), cfMenuItem(t.otherRules, 'new-rule', 'manage'));
};

const buildIconSetSubmenu = (submenu: HTMLDivElement, t: CfMenuText): void => {
  submenu.append(menuSectionHeader(t.direction));
  const directions = document.createElement('div');
  directions.className = 'app__cf-icon-panel';
  directions.append(
    cfIconChoice('icons-arrows3', ['▲', '▶', '▼'], t.iconArrows3),
    cfIconChoice('icons-arrows5', ['▲', '↗', '▶', '↘', '▼'], t.iconArrows5),
    cfIconChoice('icons-triangles3', ['▲', '▬', '▼'], t.iconTriangles3),
  );
  submenu.appendChild(directions);
  submenu.append(menuSectionHeader(t.shapes));
  const shapes = document.createElement('div');
  shapes.className = 'app__cf-icon-panel';
  shapes.append(
    cfIconChoice('icons-traffic3', ['●', '●', '●'], t.iconTraffic3),
    cfIconChoice('icons-trafficRim3', ['●', '●', '●'], t.iconTrafficRim3),
    cfIconChoice('icons-stars3', ['★', '★', '★'], t.iconStars3),
  );
  submenu.append(shapes, menuSectionHeader(t.indicators));
  const indicators = document.createElement('div');
  indicators.className = 'app__cf-icon-panel';
  indicators.append(
    cfIconChoice('icons-symbols3', ['✓', '!', '×'], t.iconSymbols3),
    cfIconChoice('icons-flags3', ['⚑', '⚑', '⚑'], t.iconFlags3),
  );
  submenu.append(indicators, menuSectionHeader(t.ratings));
  const ratings = document.createElement('div');
  ratings.className = 'app__cf-icon-panel';
  ratings.append(
    cfIconChoice('icons-stars3', ['★', '★', '★'], t.ratings),
    cfIconChoice('icons-quarters5', ['◔', '◑', '◕', '●', '●'], t.iconQuarters5),
    cfIconChoice('icons-ratings5', ['●', '●', '●', '●', '●'], t.iconRatings5),
    cfIconChoice('icons-bars5', ['▮', '▮', '▮', '▮', '▮'], t.iconBars5),
    cfIconChoice('icons-boxes5', ['■', '■', '■', '■', '■'], t.iconBoxes5),
  );
  submenu.append(ratings, menuSeparator(), cfMenuItem(t.otherRules, 'new-rule', 'manage'));
};

const buildClearSubmenu = (submenu: HTMLDivElement, t: CfMenuText): void => {
  submenu.append(
    cfMenuItem(t.clearSelection, 'clear-selection', 'clear'),
    cfMenuItem(t.clearSheet, 'clear-sheet', 'clear'),
  );
};

const createCfPanelSubmenu = (key: CfSubmenuKey, label: string, t: CfMenuText): HTMLDivElement => {
  const submenu = document.createElement('div');
  submenu.className = `app__submenu app__submenu--cf app__submenu--cf-${key}`;
  submenu.dataset.cfPanel = key;
  submenu.setAttribute('role', 'menu');
  submenu.setAttribute('aria-label', label);
  submenu.hidden = true;
  switch (key) {
    case 'highlight':
      buildHighlightSubmenu(submenu, t);
      break;
    case 'topBottom':
      buildTopBottomSubmenu(submenu, t);
      break;
    case 'dataBar':
      buildDataBarSubmenu(submenu, t);
      break;
    case 'colorScale':
      buildColorScaleSubmenu(submenu, t);
      break;
    case 'iconSet':
      buildIconSetSubmenu(submenu, t);
      break;
    case 'clear':
      buildClearSubmenu(submenu, t);
      break;
  }
  return submenu;
};

export const createConditionalMenu = (ribbonLang: ToolbarLang): HTMLDivElement => {
  const t = buildCfMenuText(ribbonLang);
  const menu = createMenu('menu-conditional');
  menu.classList.add('app__menu--conditional');
  menu.append(
    cfSubmenuTrigger('highlight', t.highlight, 'rule'),
    cfSubmenuTrigger('topBottom', t.topBottom, 'top'),
    menuSeparator(),
    cfSubmenuTrigger('dataBar', t.dataBar, 'bar'),
    cfSubmenuTrigger('colorScale', t.colorScale, 'scale'),
    cfSubmenuTrigger('iconSet', t.iconSet, 'icon'),
    menuSeparator(),
    cfMenuItem(t.newRule, 'new-rule', 'new'),
    cfSubmenuTrigger('clear', t.clear, 'clear'),
    cfMenuItem(t.manage, 'manage', 'manage'),
  );
  menu.append(
    createCfPanelSubmenu('highlight', t.highlight, t),
    createCfPanelSubmenu('topBottom', t.topBottom, t),
    createCfPanelSubmenu('dataBar', t.dataBar, t),
    createCfPanelSubmenu('colorScale', t.colorScale, t),
    createCfPanelSubmenu('iconSet', t.iconSet, t),
    createCfPanelSubmenu('clear', t.clear, t),
  );
  return menu;
};
