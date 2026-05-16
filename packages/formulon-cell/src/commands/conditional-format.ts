import type {
  ConditionalIconSet,
  ConditionalRule,
  ConditionalScalePoint,
  SpreadsheetStore,
  State,
} from '../store/store.js';
import { mutators } from '../store/store.js';

const MAX_SHEET_ROW = 1048575;
const MAX_SHEET_COL = 16383;

export type ConditionalPresetAction =
  | 'clear-selection'
  | 'clear-sheet'
  | 'duplicates'
  | 'unique'
  | 'top10'
  | 'bottom10'
  | 'top10-percent'
  | 'bottom10-percent'
  | 'above-avg'
  | 'below-avg'
  | 'data-blue'
  | 'data-green'
  | 'data-red'
  | 'data-orange'
  | 'data-purple'
  | 'data-teal'
  | 'data-solid-blue'
  | 'data-solid-green'
  | 'data-solid-red'
  | 'data-solid-orange'
  | 'data-solid-purple'
  | 'data-solid-gray'
  | 'scale-gyr'
  | 'scale-ryg'
  | 'scale-gw'
  | 'scale-rw'
  | 'scale-bwr'
  | 'scale-rwb'
  | 'scale-gwg'
  | 'scale-ywg'
  | 'scale-rwr'
  | 'scale-bwb'
  | 'scale-yry'
  | 'scale-gyg'
  | 'icons-arrows3'
  | 'icons-arrows5'
  | 'icons-triangles3'
  | 'icons-traffic3'
  | 'icons-trafficRim3'
  | 'icons-symbols3'
  | 'icons-flags3'
  | 'icons-stars3'
  | 'icons-quarters5'
  | 'icons-ratings5'
  | 'icons-bars5'
  | 'icons-boxes5';

export function listConditionalRules(state: State): readonly ConditionalRule[] {
  return state.conditional.rules;
}

export function addConditionalRule(store: SpreadsheetStore, rule: ConditionalRule): void {
  mutators.addConditionalRule(store, rule);
}

export function removeConditionalRuleAt(store: SpreadsheetStore, index: number): void {
  mutators.removeConditionalRuleAt(store, index);
}

export function clearConditionalRules(store: SpreadsheetStore): void {
  mutators.clearConditionalRules(store);
}

export function clearConditionalRulesInRange(
  store: SpreadsheetStore,
  range: ConditionalRule['range'],
): void {
  mutators.clearConditionalRulesInRange(store, range);
}

export function clearConditionalRulesOnSheet(store: SpreadsheetStore, sheet: number): void {
  mutators.clearConditionalRulesInRange(store, {
    sheet,
    r0: 0,
    c0: 0,
    r1: MAX_SHEET_ROW,
    c1: MAX_SHEET_COL,
  });
}

export function conditionalRulesForRange(
  state: State,
  range: ConditionalRule['range'],
): readonly ConditionalRule[] {
  return state.conditional.rules.filter((rule) => rangesIntersect(rule.range, range));
}

const rangesIntersect = (a: ConditionalRule['range'], b: ConditionalRule['range']): boolean =>
  a.sheet === b.sheet && !(a.r1 < b.r0 || a.r0 > b.r1 || a.c1 < b.c0 || a.c0 > b.c1);

const highlightFill = { fill: '#ffc7ce', color: '#9c0006' } as const;
const topBottomFill = { fill: '#c6efce', color: '#006100' } as const;

const dataBarColor = (action: ConditionalPresetAction): string => {
  const map: Partial<Record<ConditionalPresetAction, string>> = {
    'data-blue': '#638ec6',
    'data-green': '#63a95c',
    'data-red': '#c45a5a',
    'data-orange': '#d6a440',
    'data-purple': '#8a74b9',
    'data-teal': '#4ba1a8',
    'data-solid-blue': '#4472c4',
    'data-solid-green': '#70ad47',
    'data-solid-red': '#c00000',
    'data-solid-orange': '#ed7d31',
    'data-solid-purple': '#8064a2',
    'data-solid-gray': '#7f7f7f',
  };
  return map[action] ?? '#638ec6';
};

const dataBarGradient = (action: ConditionalPresetAction): boolean =>
  action === 'data-blue' ||
  action === 'data-green' ||
  action === 'data-red' ||
  action === 'data-orange' ||
  action === 'data-purple' ||
  action === 'data-teal';

const colorScaleStops = (
  action: ConditionalPresetAction,
): [string, string] | [string, string, string] => {
  const map: Partial<Record<ConditionalPresetAction, [string, string] | [string, string, string]>> =
    {
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

const iconSetThresholds = (icons: ConditionalIconSet): ConditionalScalePoint[] => {
  const slots = icons.endsWith('5') ? 5 : 3;
  return Array.from({ length: slots - 1 }, (_, index) => ({
    kind: 'percent',
    value: ((index + 1) * 100) / slots,
  }));
};

export function applyConditionalPresetAction(
  store: SpreadsheetStore,
  action: ConditionalPresetAction,
  range: ConditionalRule['range'] = store.getState().selection.range,
): boolean {
  if (action === 'clear-selection') {
    clearConditionalRulesInRange(store, range);
    return true;
  }
  if (action === 'clear-sheet') {
    clearConditionalRulesOnSheet(store, range.sheet);
    return true;
  }
  if (action === 'duplicates' || action === 'unique') {
    addConditionalRule(store, { kind: action, range, apply: highlightFill });
    return true;
  }
  if (
    action === 'top10' ||
    action === 'bottom10' ||
    action === 'top10-percent' ||
    action === 'bottom10-percent'
  ) {
    addConditionalRule(store, {
      kind: 'top-bottom',
      range,
      mode: action.startsWith('top') ? 'top' : 'bottom',
      n: 10,
      percent: action.endsWith('-percent'),
      apply: topBottomFill,
    });
    return true;
  }
  if (action === 'above-avg' || action === 'below-avg') {
    addConditionalRule(store, {
      kind: 'average',
      range,
      mode: action === 'above-avg' ? 'above' : 'below',
      apply: topBottomFill,
    });
    return true;
  }
  if (action.startsWith('data-')) {
    addConditionalRule(store, {
      kind: 'data-bar',
      range,
      color: dataBarColor(action),
      gradient: dataBarGradient(action),
      showValue: true,
    });
    return true;
  }
  if (action.startsWith('scale-')) {
    const stops = colorScaleStops(action);
    addConditionalRule(store, {
      kind: 'color-scale',
      range,
      stops,
      thresholds:
        stops.length === 2
          ? [{ kind: 'min' }, { kind: 'max' }]
          : [{ kind: 'min' }, { kind: 'percentile', value: 50 }, { kind: 'max' }],
    });
    return true;
  }
  if (action.startsWith('icons-')) {
    const icons = action.replace('icons-', '') as ConditionalIconSet;
    addConditionalRule(store, {
      kind: 'icon-set',
      range,
      icons,
      showValue: true,
      thresholds: iconSetThresholds(icons),
    });
    return true;
  }
  return false;
}
