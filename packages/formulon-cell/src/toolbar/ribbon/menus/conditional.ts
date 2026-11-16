// Conditional Formatting dropdown. Top-level menu lists hover-to-open submenu
// triggers (highlight / topBottom / dataBar / colorScale / iconSet / clear) and
// flat actions (newRule / manage). Each submenu panel renders the matching
// preset list — swatches for data bars and color scales, glyph rows for icon
// sets, plain menu items for highlight rules. All clicks emit a `data-cf-*`
// attribute the conditional-format dispatcher reads.

import type { ConditionalPresetAction } from '../../../commands/conditional-format.js';
import { colorScaleStops, dataBarColor } from '../../../commands/conditional-format.js';
import { projectDisabledState } from '../../menu-a11y.js';
import { conditionalMenuText } from '../../menu-text.js';
import type { ToolbarLang } from '../../ribbon-model.js';
import {
  createMenu,
  createMenuButton,
  createSubmenu,
  menuPresetButton,
  menuSectionHeader,
  menuSeparator,
  menuSubmenuTrigger,
} from './general.js';
import type { MenuVariant } from './mac-variants.js';

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
    ...t,
    dataBar: t.dataBars,
    colorScale: t.colorScales,
    iconSet: t.iconSets,
    text: t.textContains,
    date: t.dateOccurring,
    duplicate: t.duplicates,
    gradient: t.gradientFill,
    solid: t.solidFill,
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
  };
};

type CfMenuText = ReturnType<typeof buildCfMenuText>;

const buildHighlightSubmenu = (submenu: HTMLDivElement, t: CfMenuText, mac: boolean): void => {
  submenu.append(
    cfMenuItem(t.greater, 'cell-gt', 'rule'),
    cfMenuItem(t.less, 'cell-lt', 'rule'),
    cfMenuItem(t.between, 'cell-between', 'rule'),
    cfMenuItem(t.equal, 'cell-eq', 'rule'),
    cfMenuItem(t.text, 'text-contains', 'rule'),
    cfMenuItem(t.date, 'date-occurring', 'rule'),
    cfMenuItem(t.duplicate, 'duplicates', 'rule'),
    ...(mac ? [] : [cfMenuItem(t.unique, 'unique', 'rule')]),
    menuSeparator(),
    cfMenuItem(t.otherRules, 'new-rule', 'manage'),
  );
};

const buildTopBottomSubmenu = (submenu: HTMLDivElement, t: CfMenuText, mac: boolean): void => {
  // The Mac menu orders the rules top 10, top %, bottom 10, bottom %, and each opens a dialog.
  const prompt = (label: string): string => (label.endsWith('...') ? label : `${label}...`);
  const label = (text: string): string => (mac ? prompt(text) : text);
  const rules: ReadonlyArray<readonly [string, string]> = mac
    ? [
        [t.top10, 'top10'],
        [t.top10Percent, 'top10-percent'],
        [t.bottom10, 'bottom10'],
        [t.bottom10Percent, 'bottom10-percent'],
        [t.aboveAvg, 'above-avg'],
        [t.belowAvg, 'below-avg'],
      ]
    : [
        [t.top10, 'top10'],
        [t.bottom10, 'bottom10'],
        [t.top10Percent, 'top10-percent'],
        [t.bottom10Percent, 'bottom10-percent'],
        [t.aboveAvg, 'above-avg'],
        [t.belowAvg, 'below-avg'],
      ];
  submenu.append(
    ...rules.map(([text, action]) => cfMenuItem(label(text), action, 'top')),
    menuSeparator(),
    cfMenuItem(t.otherRules, 'new-rule', 'manage'),
  );
};

type CfSwatch = readonly [action: string, colors: readonly string[], label: string];

const buildDataBarSubmenu = (submenu: HTMLDivElement, t: CfMenuText): void => {
  const swatch = (action: ConditionalPresetAction, label: string): CfSwatch => {
    const color = dataBarColor(action);
    return [action, [action.startsWith('data-solid') ? color : '#fff', color], label];
  };
  const gradientBars: readonly CfSwatch[] = [
    swatch('data-blue', t.dataBarGradientBlue),
    swatch('data-green', t.dataBarGradientGreen),
    swatch('data-red', t.dataBarGradientRed),
    swatch('data-orange', t.dataBarGradientOrange),
    swatch('data-teal', t.dataBarGradientTeal),
    swatch('data-purple', t.dataBarGradientPurple),
  ];
  const solidBars: readonly CfSwatch[] = [
    swatch('data-solid-blue', t.dataBarSolidBlue),
    swatch('data-solid-green', t.dataBarSolidGreen),
    swatch('data-solid-red', t.dataBarSolidRed),
    swatch('data-solid-orange', t.dataBarSolidOrange),
    swatch('data-solid-teal', t.dataBarSolidTeal),
    swatch('data-solid-purple', t.dataBarSolidPurple),
  ];
  const row = (bars: readonly CfSwatch[]): HTMLDivElement =>
    cfPanel(
      'fc-tb__cf-choice-row',
      bars.map(([action, colors, label]) => cfSwatchButton(action, [...colors], label)),
    );
  submenu.append(
    menuSectionHeader(t.gradient),
    row(gradientBars),
    menuSectionHeader(t.solid),
    row(solidBars),
    menuSeparator(),
    cfMenuItem(t.otherRules, 'new-rule', 'manage'),
  );
};

const buildColorScaleSubmenu = (submenu: HTMLDivElement, t: CfMenuText): void => {
  const scale = (action: ConditionalPresetAction, label: string): CfSwatch => [
    action,
    colorScaleStops(action),
    label,
  ];
  const colorScales: readonly CfSwatch[] = [
    scale('scale-gyr', t.colorScaleGreenYellowRed),
    scale('scale-ryg', t.colorScaleRedYellowGreen),
    scale('scale-gwr', t.colorScaleGreenWhiteRed),
    scale('scale-rwg', t.colorScaleRedWhiteGreen),
    scale('scale-bwr', t.colorScaleBlueWhiteRed),
    scale('scale-rwb', t.colorScaleRedWhiteBlue),
    scale('scale-wr', t.colorScaleWhiteRed),
    scale('scale-rw', t.colorScaleRedWhite),
    scale('scale-gw', t.colorScaleGreenWhite),
    scale('scale-wg', t.colorScaleWhiteGreen),
    scale('scale-gy', t.colorScaleGreenYellow),
    scale('scale-yg', t.colorScaleYellowGreen),
  ];
  const scales = cfPanel(
    'fc-tb__cf-choice-grid-panel',
    colorScales.map(([action, colors, label]) => cfSwatchButton(action, [...colors], label)),
  );
  submenu.append(scales, menuSeparator(), cfMenuItem(t.otherRules, 'new-rule', 'manage'));
};

type CfIconRow = readonly [action: string, count: number, label: string];

const buildIconSetSubmenu = (submenu: HTMLDivElement, t: CfMenuText): void => {
  const sections: ReadonlyArray<readonly [header: string, rows: readonly CfIconRow[]]> = [
    [
      t.direction,
      [
        ['icons-arrows3', 3, t.iconArrows3],
        ['icons-arrowsGray3', 3, t.iconArrowsGray3],
        ['icons-triangles3', 3, t.iconTriangles3],
        ['icons-arrowsGray4', 4, t.iconArrowsGray4],
        ['icons-arrows4', 4, t.iconArrows4],
        ['icons-arrowsGray5', 5, t.iconArrowsGray5],
        ['icons-arrows5', 5, t.iconArrows5],
      ],
    ],
    [
      t.shapes,
      [
        ['icons-traffic3', 3, t.iconTraffic3],
        ['icons-trafficRim3', 3, t.iconTrafficRim3],
        ['icons-signs3', 3, t.iconSigns3],
        ['icons-traffic4', 4, t.iconTraffic4],
        ['icons-redToBlack4', 4, t.iconRedToBlack4],
      ],
    ],
    [
      t.indicators,
      [
        ['icons-symbolsCircled3', 3, t.iconSymbolsCircled3],
        ['icons-symbols3', 3, t.iconSymbols3],
        ['icons-flags3', 3, t.iconFlags3],
      ],
    ],
    [
      t.ratings,
      [
        ['icons-stars3', 3, t.iconStars3],
        ['icons-ratings4', 4, t.iconRatings4],
        ['icons-quarters5', 5, t.iconQuarters5],
        ['icons-ratings5', 5, t.iconRatings5],
        ['icons-boxes5', 5, t.iconBoxes5],
      ],
    ],
  ];
  for (const [header, rows] of sections) {
    submenu.append(
      menuSectionHeader(header),
      cfPanel(
        'fc-tb__cf-icon-panel',
        rows.map(([action, count, label]) => cfIconChoice(action, count, label)),
      ),
    );
  }
  submenu.append(menuSeparator(), cfMenuItem(t.otherRules, 'new-rule', 'manage'));
};

const buildClearSubmenu = (submenu: HTMLDivElement, t: CfMenuText, mac: boolean): void => {
  submenu.append(
    cfMenuItem(t.clearSelection, 'clear-selection', 'clear'),
    cfMenuItem(t.clearSheet, 'clear-sheet', 'clear'),
  );
  if (!mac) return;
  // Outside a table or PivotTable these two stay greyed out.
  for (const [label, action] of [
    [t.clearTable, 'clear-table'],
    [t.clearPivot, 'clear-pivot'],
  ] as const) {
    const item = cfMenuItem(label, action, 'clear');
    projectDisabledState(item, true, null);
    submenu.appendChild(item);
  }
};

const createCfPanelSubmenu = (
  key: CfSubmenuKey,
  label: string,
  t: CfMenuText,
  mac: boolean,
): HTMLDivElement => {
  const submenu = createSubmenu({
    id: cfSubmenuId(key),
    className: `fc-tb__submenu fc-tb__submenu--cf fc-tb__submenu--cf-${key}`,
    label,
    dataset: { cfPanel: key },
  });
  switch (key) {
    case 'highlight':
      buildHighlightSubmenu(submenu, t, mac);
      break;
    case 'topBottom':
      buildTopBottomSubmenu(submenu, t, mac);
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
      buildClearSubmenu(submenu, t, mac);
      break;
  }
  return submenu;
};

const cfSubmenuId = (key: CfSubmenuKey): string => `menu-conditional-${key}`;

export const createConditionalMenu = (
  ribbonLang: ToolbarLang,
  variant?: MenuVariant,
): HTMLDivElement => {
  const mac = variant?.mac === true;
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
    createCfPanelSubmenu('highlight', t.highlight, t, mac),
    createCfPanelSubmenu('topBottom', t.topBottom, t, mac),
    createCfPanelSubmenu('dataBar', t.dataBar, t, mac),
    createCfPanelSubmenu('colorScale', t.colorScale, t, mac),
    createCfPanelSubmenu('iconSet', t.iconSet, t, mac),
    createCfPanelSubmenu('clear', t.clear, t, mac),
  );
  return menu;
};
