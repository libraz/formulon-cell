// Conditional Formatting dropdown. Top-level menu lists hover-to-open submenu
// triggers (highlight / topBottom / dataBar / colorScale / iconSet / clear) and
// flat actions (newRule / manage). Each submenu panel renders the matching
// preset list — swatches for data bars and color scales, glyph rows for icon
// sets, plain menu items for highlight rules. All clicks emit a `data-cf-*`
// attribute the playground main dispatcher reads.

import { conditionalMenuText, type ToolbarLang } from '../../../index.js';

import {
  createMenu,
  createMenuButton,
  createSubmenu,
  menuPresetButton,
  menuSectionHeader,
  menuSeparator,
  menuSubmenuTrigger,
} from './general.js';

export type CfSubmenuKey =
  | 'highlight'
  | 'topBottom'
  | 'dataBar'
  | 'colorScale'
  | 'iconSet'
  | 'clear';

type CfIconKind = 'rule' | 'top' | 'bar' | 'scale' | 'icon' | 'new' | 'clear' | 'manage';

const cfSpan = (className?: string, text?: string): HTMLSpanElement => {
  const span = document.createElement('span');
  if (className) span.className = className;
  if (text !== undefined) span.textContent = text;
  return span;
};

const cfIcon = (kind: CfIconKind): HTMLSpanElement => {
  const span = cfSpan(`fc-tb__cf-icon fc-tb__cf-icon--${kind}`);
  span.setAttribute('aria-hidden', 'true');
  return span;
};

const cfMenuItem = (
  label: string,
  action: string,
  icon: CfIconKind = 'rule',
): HTMLButtonElement => {
  return menuPresetButton(label, 'cfAction', action, cfIcon(icon));
};

const cfSubmenuTrigger = (
  key: CfSubmenuKey,
  label: string,
  icon: CfIconKind,
): HTMLButtonElement => {
  const btn = cfMenuItem(label, `submenu-${key}`, icon);
  return menuSubmenuTrigger(btn, { cfSubmenu: key }, { controlsId: cfSubmenuId(key) });
};

const cfChoiceButton = (className: string, action: string, title: string): HTMLButtonElement => {
  return createMenuButton({
    className,
    attr: 'cfAction',
    value: action,
    title,
    ariaLabel: title,
  });
};

const cfSwatchButton = (action: string, colors: string[], title: string): HTMLButtonElement => {
  const btn = cfChoiceButton('fc-tb__cf-choice', action, title);
  const grid = cfSpan('fc-tb__cf-choice-grid');
  for (const color of colors) {
    const cell = cfSpan();
    cell.style.background = color;
    grid.appendChild(cell);
  }
  btn.appendChild(grid);
  return btn;
};

const cfIconChoice = (action: string, count: number, title: string): HTMLButtonElement => {
  const btn = cfChoiceButton(
    `fc-tb__cf-icon-choice fc-tb__cf-icon-choice--${action}`,
    action,
    title,
  );
  for (let index = 0; index < count; index += 1) {
    btn.appendChild(cfSpan());
  }
  return btn;
};

const cfPanel = (className: string, children: readonly Node[] = []): HTMLDivElement => {
  const panel = document.createElement('div');
  panel.className = className;
  panel.append(...children);
  return panel;
};

export const buildCfMenuText = (ribbonLang: ToolbarLang) => {
  const t = conditionalMenuText(ribbonLang) as ReturnType<typeof conditionalMenuText> & {
    dateYesterday: string;
    dateToday: string;
    dateTomorrow: string;
    dateLast7: string;
    dateLastWeek: string;
    dateThisWeek: string;
    dateNextWeek: string;
    dateLastMonth: string;
    dateThisMonth: string;
    dateNextMonth: string;
    dateUnsupported: string;
    greaterPrompt: string;
    lessPrompt: string;
    betweenPrompt: string;
    betweenAndPrompt: string;
    equalPrompt: string;
    topBottomPrompt: string;
    formatPreview: string;
    customFormat: string;
    customFormatTitle: string;
    customFillColor: string;
    customTextColor: string;
    customBold: string;
    customItalic: string;
    customUnderline: string;
    customStrike: string;
    formatLightRed: string;
    formatYellow: string;
    formatGreen: string;
    formatLightRedFill: string;
    formatRedText: string;
    formatRedBorder: string;
    formatRedFill: string;
    formatRedTextFill: string;
    ok: string;
    cancel: string;
  };
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
    greaterPrompt: t.greaterPrompt,
    lessPrompt: t.lessPrompt,
    betweenPrompt: t.betweenPrompt,
    betweenAndPrompt: t.betweenAndPrompt,
    equalPrompt: t.equalPrompt,
    topBottomPrompt: t.topBottomPrompt,
    formatPreview: t.formatPreview,
    customFormat: t.customFormat,
    customFormatTitle: t.customFormatTitle,
    customFillColor: t.customFillColor,
    customTextColor: t.customTextColor,
    customBold: t.customBold,
    customItalic: t.customItalic,
    customUnderline: t.customUnderline,
    customStrike: t.customStrike,
    formatLightRed: t.formatLightRed,
    formatYellow: t.formatYellow,
    formatGreen: t.formatGreen,
    formatLightRedFill: t.formatLightRedFill,
    formatRedText: t.formatRedText,
    formatRedBorder: t.formatRedBorder,
    formatRedFill: t.formatRedFill,
    formatRedTextFill: t.formatRedTextFill,
    datePrompt: t.datePrompt,
    datePeriods: {
      yesterday: t.dateYesterday,
      today: t.dateToday,
      tomorrow: t.dateTomorrow,
      last7: t.dateLast7,
      'last-week': t.dateLastWeek,
      'this-week': t.dateThisWeek,
      'next-week': t.dateNextWeek,
      'last-month': t.dateLastMonth,
      'this-month': t.dateThisMonth,
      'next-month': t.dateNextMonth,
    },
    dateUnsupported: t.dateUnsupported,
    ok: t.ok,
    cancel: t.cancel,
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
  const gradientBars: readonly (readonly [string, string, string])[] = [
    ['data-blue', '#638ec6', t.dataBarGradientBlue],
    ['data-green', '#63a95c', t.dataBarGradientGreen],
    ['data-red', '#c45a5a', t.dataBarGradientRed],
    ['data-orange', '#d6a440', t.dataBarGradientOrange],
    ['data-purple', '#8a74b9', t.dataBarGradientPurple],
    ['data-teal', '#4ba1a8', t.dataBarGradientTeal],
  ];
  const gradient = cfPanel(
    'fc-tb__cf-choice-row',
    gradientBars.map(([action, color, label]) => cfSwatchButton(action, ['#fff', color], label)),
  );
  submenu.appendChild(gradient);
  submenu.append(menuSectionHeader(t.solid));
  const solidBars: readonly (readonly [string, string, string])[] = [
    ['data-solid-blue', '#4472c4', t.dataBarSolidBlue],
    ['data-solid-green', '#70ad47', t.dataBarSolidGreen],
    ['data-solid-red', '#c00000', t.dataBarSolidRed],
    ['data-solid-orange', '#ed7d31', t.dataBarSolidOrange],
    ['data-solid-purple', '#8064a2', t.dataBarSolidPurple],
    ['data-solid-gray', '#7f7f7f', t.dataBarSolidGray],
  ];
  const solid = cfPanel(
    'fc-tb__cf-choice-row',
    solidBars.map(([action, color, label]) => cfSwatchButton(action, [color, color], label)),
  );
  submenu.append(solid, menuSeparator(), cfMenuItem(t.otherRules, 'new-rule', 'manage'));
};

const buildColorScaleSubmenu = (submenu: HTMLDivElement, t: CfMenuText): void => {
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
  const scales = cfPanel(
    'fc-tb__cf-choice-grid-panel',
    colorScales.map(([action, colors, label]) => cfSwatchButton(action, [...colors], label)),
  );
  submenu.append(scales, menuSeparator(), cfMenuItem(t.otherRules, 'new-rule', 'manage'));
};

const buildIconSetSubmenu = (submenu: HTMLDivElement, t: CfMenuText): void => {
  submenu.append(menuSectionHeader(t.direction));
  const directions = cfPanel('fc-tb__cf-icon-panel', [
    cfIconChoice('icons-arrows3', 3, t.iconArrows3),
    cfIconChoice('icons-arrows5', 5, t.iconArrows5),
    cfIconChoice('icons-triangles3', 3, t.iconTriangles3),
  ]);
  submenu.appendChild(directions);
  submenu.append(menuSectionHeader(t.shapes));
  const shapes = cfPanel('fc-tb__cf-icon-panel', [
    cfIconChoice('icons-traffic3', 3, t.iconTraffic3),
    cfIconChoice('icons-trafficRim3', 3, t.iconTrafficRim3),
    cfIconChoice('icons-stars3', 3, t.iconStars3),
  ]);
  submenu.append(shapes, menuSectionHeader(t.indicators));
  const indicators = cfPanel('fc-tb__cf-icon-panel', [
    cfIconChoice('icons-symbols3', 3, t.iconSymbols3),
    cfIconChoice('icons-flags3', 3, t.iconFlags3),
  ]);
  submenu.append(indicators, menuSectionHeader(t.ratings));
  const ratings = cfPanel('fc-tb__cf-icon-panel', [
    cfIconChoice('icons-stars3', 3, t.ratings),
    cfIconChoice('icons-quarters5', 5, t.iconQuarters5),
    cfIconChoice('icons-ratings5', 5, t.iconRatings5),
    cfIconChoice('icons-bars5', 5, t.iconBars5),
    cfIconChoice('icons-boxes5', 5, t.iconBoxes5),
  ]);
  submenu.append(ratings, menuSeparator(), cfMenuItem(t.otherRules, 'new-rule', 'manage'));
};

const buildClearSubmenu = (submenu: HTMLDivElement, t: CfMenuText): void => {
  submenu.append(
    cfMenuItem(t.clearSelection, 'clear-selection', 'clear'),
    cfMenuItem(t.clearSheet, 'clear-sheet', 'clear'),
  );
};

const createCfPanelSubmenu = (key: CfSubmenuKey, label: string, t: CfMenuText): HTMLDivElement => {
  const submenu = createSubmenu({
    id: cfSubmenuId(key),
    className: `fc-tb__submenu fc-tb__submenu--cf fc-tb__submenu--cf-${key}`,
    label,
    dataset: { cfPanel: key },
  });
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

const cfSubmenuId = (key: CfSubmenuKey): string => `menu-conditional-${key}`;

export const createConditionalMenu = (ribbonLang: ToolbarLang): HTMLDivElement => {
  const t = buildCfMenuText(ribbonLang);
  const menu = createMenu('menu-conditional');
  menu.classList.add('fc-tb__menu--conditional');
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
