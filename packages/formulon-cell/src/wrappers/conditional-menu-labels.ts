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
    case 'data-solid-gray':
      return labels.dataBarSolidGray;
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
    case 'scale-gwg':
      return labels.colorScaleGreenWhiteGreen;
    case 'scale-ywg':
      return labels.colorScaleYellowWhiteGreen;
    case 'scale-rwr':
      return labels.colorScaleRedWhiteRed;
    case 'scale-bwb':
      return labels.colorScaleBlueWhiteBlue;
    case 'scale-yry':
      return labels.colorScaleYellowRedGreen;
    case 'scale-gyg':
      return labels.colorScaleGreenYellowGreen;
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
    case 'icons-bars5':
      return labels.iconBars5;
    case 'icons-boxes5':
      return labels.iconBoxes5;
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
  if (action.includes('teal')) return '#4ba1a8';
  if (action.includes('gray')) return '#7f7f7f';
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
    'scale-gwg': ['#63be7b', '#ffffff', '#00a651'],
    'scale-ywg': ['#ffeb84', '#ffffff', '#63be7b'],
    'scale-rwr': ['#f8696b', '#ffffff', '#c00000'],
    'scale-bwb': ['#5a8dee', '#ffffff', '#4472c4'],
    'scale-yry': ['#ffeb84', '#f8696b', '#63be7b'],
    'scale-gyg': ['#63be7b', '#ffeb84', '#00a651'],
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
