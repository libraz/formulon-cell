// Shared label resolution for the conditional-formatting dropdown used by
// both the React and Vue toolbar wrappers. The wrappers diverge on how they
// render swatches (React inlines hex tuples, Vue threads through a helper),
// but the action → localized label mapping is identical, so it lives here.

import type { ConditionalPresetAction } from '../commands/conditional-format.js';
import type { Strings } from '../i18n/strings.js';

export type ConditionalIconSetAction = Extract<ConditionalPresetAction, `icons-${string}`>;

type Labels = Strings['conditionalMenu'];

export const conditionalDataBarLabel = (
  action: ConditionalPresetAction | string,
  labels: Labels,
): string => {
  switch (action) {
    case 'data-blue':
      return labels.dataBarGradientBlue;
    case 'data-green':
      return labels.dataBarGradientGreen;
    case 'data-red':
      return labels.dataBarGradientRed;
    case 'data-orange':
      return labels.dataBarGradientOrange;
    case 'data-purple':
      return labels.dataBarGradientPurple;
    case 'data-teal':
      return labels.dataBarGradientTeal;
    case 'data-solid-blue':
      return labels.dataBarSolidBlue;
    case 'data-solid-green':
      return labels.dataBarSolidGreen;
    case 'data-solid-red':
      return labels.dataBarSolidRed;
    case 'data-solid-orange':
      return labels.dataBarSolidOrange;
    case 'data-solid-purple':
      return labels.dataBarSolidPurple;
    case 'data-solid-teal':
      return labels.dataBarSolidTeal;
    default:
      return labels.dataBars;
  }
};

export const conditionalColorScaleLabel = (
  action: ConditionalPresetAction | string,
  labels: Labels,
): string => {
  switch (action) {
    case 'scale-gyr':
      return labels.colorScaleGreenYellowRed;
    case 'scale-ryg':
      return labels.colorScaleRedYellowGreen;
    case 'scale-gw':
      return labels.colorScaleGreenWhite;
    case 'scale-rw':
      return labels.colorScaleRedWhite;
    case 'scale-bwr':
      return labels.colorScaleBlueWhiteRed;
    case 'scale-rwb':
      return labels.colorScaleRedWhiteBlue;
    case 'scale-gwr':
      return labels.colorScaleGreenWhiteRed;
    case 'scale-rwg':
      return labels.colorScaleRedWhiteGreen;
    case 'scale-wr':
      return labels.colorScaleWhiteRed;
    case 'scale-wg':
      return labels.colorScaleWhiteGreen;
    case 'scale-gy':
      return labels.colorScaleGreenYellow;
    case 'scale-yg':
      return labels.colorScaleYellowGreen;
    default:
      return labels.colorScales;
  }
};

export const conditionalIconSetLabel = (
  action: ConditionalIconSetAction | string,
  labels: Labels,
): string => {
  switch (action) {
    case 'icons-arrows3':
      return labels.iconArrows3;
    case 'icons-arrows5':
      return labels.iconArrows5;
    case 'icons-triangles3':
      return labels.iconTriangles3;
    case 'icons-traffic3':
      return labels.iconTraffic3;
    case 'icons-trafficRim3':
      return labels.iconTrafficRim3;
    case 'icons-symbols3':
      return labels.iconSymbols3;
    case 'icons-flags3':
      return labels.iconFlags3;
    case 'icons-stars3':
      return labels.iconStars3;
    case 'icons-quarters5':
      return labels.iconQuarters5;
    case 'icons-ratings5':
      return labels.iconRatings5;
    case 'icons-boxes5':
      return labels.iconBoxes5;
    case 'icons-arrowsGray3':
      return labels.iconArrowsGray3;
    case 'icons-arrows4':
      return labels.iconArrows4;
    case 'icons-arrowsGray4':
      return labels.iconArrowsGray4;
    case 'icons-arrowsGray5':
      return labels.iconArrowsGray5;
    case 'icons-signs3':
      return labels.iconSigns3;
    case 'icons-traffic4':
      return labels.iconTraffic4;
    case 'icons-redToBlack4':
      return labels.iconRedToBlack4;
    case 'icons-symbolsCircled3':
      return labels.iconSymbolsCircled3;
    case 'icons-ratings4':
      return labels.iconRatings4;

    default:
      return labels.iconSets;
  }
};

/** Hex colors for the data-bar preview swatches in the dropdown
 *  (independent of the actual rule's bar color, which is owned by
 *  `applyConditionalPresetAction`). */
export const conditionalDataBarSwatchColor = (action: string): string => {
  if (action.includes('green')) return '#70ad47';
  if (action.includes('red')) return '#c00000';
  if (action.includes('orange')) return '#ed7d31';
  if (action.includes('purple')) return '#8064a2';
  if (action.includes('teal')) return action.includes('solid') ? '#31a9bd' : '#4ba1a8';
  return '#4472c4';
};

/** Hex tuples for the color-scale preview swatches in the dropdown. */
export const conditionalColorScaleSwatchColors = (action: string): readonly string[] => {
  const map: Record<string, readonly string[]> = {
    'scale-gyr': ['#63be7b', '#ffeb84', '#f8696b'],
    'scale-ryg': ['#f8696b', '#ffeb84', '#63be7b'],
    'scale-gw': ['#63be7b', '#ffffff'],
    'scale-rw': ['#f8696b', '#ffffff'],
    'scale-bwr': ['#5a8dee', '#ffffff', '#f8696b'],
    'scale-rwb': ['#f8696b', '#ffffff', '#5a8dee'],
    'scale-gwr': ['#63be7b', '#ffffff', '#f8696b'],
    'scale-rwg': ['#f8696b', '#ffffff', '#63be7b'],
    'scale-wr': ['#ffffff', '#f8696b'],
    'scale-wg': ['#ffffff', '#63be7b'],
    'scale-gy': ['#63be7b', '#ffeb84'],
    'scale-yg': ['#ffeb84', '#63be7b'],
  };
  return map[action] ?? ['#63be7b', '#ffeb84', '#f8696b'];
};

/** Symbol grid for the "Insert → Symbol" dropdown. Grouped in rows of 12 by
 *  the call site (math, Greek, currency, typographic). */
export const TOOLBAR_INSERT_SYMBOLS = [
  '±',
  '×',
  '÷',
  '≤',
  '≥',
  '≠',
  '≈',
  '∞',
  '√',
  '∑',
  '∫',
  'π',
  'Α',
  'Β',
  'Γ',
  'Δ',
  'Θ',
  'Λ',
  'Ξ',
  'Π',
  'Σ',
  'Φ',
  'Ψ',
  'Ω',
  '$',
  '€',
  '¥',
  '£',
  '¢',
  '₩',
  '₹',
  '₽',
  '©',
  '®',
  '™',
  '§',
  '¶',
  '†',
  '‡',
  '•',
] as const;

export type ToolbarInsertSymbol = (typeof TOOLBAR_INSERT_SYMBOLS)[number];
