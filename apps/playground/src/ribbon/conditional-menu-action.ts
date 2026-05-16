// Extracted from main.ts. Dispatches the ribbon's conditional formatting menu
// actions: opens the new-rule / manage dialogs, prompts for inline rule input,
// or applies preset actions on the store. State (workbook instance, locale,
// fills, prompt closures) is injected via a deps struct so the function stays
// host-agnostic.

import type {
  ConditionalDialogOpenOptions,
  ConditionalPresetAction,
  ConditionalRule,
  Range,
  SpreadsheetInstance,
} from '@libraz/formulon-cell';
import { applyConditionalPresetAction, recordConditionalRulesChange } from '@libraz/formulon-cell';

import { buildCfMenuText } from './menus/conditional.js';

// Pure helpers copied from main.ts so this module stands alone.

const DATE_PERIODS = [
  'yesterday',
  'today',
  'tomorrow',
  'last7',
  'last-week',
  'this-week',
  'next-week',
  'last-month',
  'this-month',
  'next-month',
] as const;

type DatePeriod = (typeof DATE_PERIODS)[number];

const isDatePeriod = (value: string): value is DatePeriod =>
  DATE_PERIODS.includes(value as DatePeriod);

const cfDatePeriodOptions = (
  ribbonLang: 'ja' | 'en',
): Array<{ value: DatePeriod; label: string }> =>
  ribbonLang === 'ja'
    ? [
        { value: 'yesterday', label: '昨日' },
        { value: 'today', label: '今日' },
        { value: 'tomorrow', label: '明日' },
        { value: 'last7', label: '過去 7 日間' },
        { value: 'last-week', label: '先週' },
        { value: 'this-week', label: '今週' },
        { value: 'next-week', label: '来週' },
        { value: 'last-month', label: '先月' },
        { value: 'this-month', label: '今月' },
        { value: 'next-month', label: '来月' },
      ]
    : [
        { value: 'yesterday', label: 'Yesterday' },
        { value: 'today', label: 'Today' },
        { value: 'tomorrow', label: 'Tomorrow' },
        { value: 'last7', label: 'In the last 7 days' },
        { value: 'last-week', label: 'Last week' },
        { value: 'this-week', label: 'This week' },
        { value: 'next-week', label: 'Next week' },
        { value: 'last-month', label: 'Last month' },
        { value: 'this-month', label: 'This month' },
        { value: 'next-month', label: 'Next month' },
      ];

const conditionalPresetActions = new Set<ConditionalPresetAction>([
  'clear-selection',
  'clear-sheet',
  'duplicates',
  'unique',
  'above-avg',
  'below-avg',
  'data-blue',
  'data-green',
  'data-red',
  'data-orange',
  'data-purple',
  'data-teal',
  'data-solid-blue',
  'data-solid-green',
  'data-solid-red',
  'data-solid-orange',
  'data-solid-purple',
  'data-solid-gray',
  'scale-gyr',
  'scale-ryg',
  'scale-gw',
  'scale-rw',
  'scale-bwr',
  'scale-rwb',
  'scale-gwg',
  'scale-ywg',
  'scale-rwr',
  'scale-bwb',
  'scale-yry',
  'scale-gyg',
  'icons-arrows3',
  'icons-arrows5',
  'icons-triangles3',
  'icons-traffic3',
  'icons-trafficRim3',
  'icons-symbols3',
  'icons-flags3',
  'icons-stars3',
  'icons-quarters5',
  'icons-ratings5',
  'icons-bars5',
  'icons-boxes5',
]);

const isConditionalPresetAction = (action: string): action is ConditionalPresetAction =>
  conditionalPresetActions.has(action as ConditionalPresetAction);

const conditionalRuleKindForPanel = (
  panel: string | undefined,
): ConditionalDialogOpenOptions['kind'] | undefined => {
  if (panel === 'dataBar') return 'data-bar';
  if (panel === 'colorScale') return 'color-scale';
  if (panel === 'iconSet') return 'icon-set';
  if (panel === 'topBottom') return 'top-bottom';
  if (panel === 'highlight') return 'cell-value';
  return undefined;
};

export type CfFillStyle = { readonly fill: string; readonly color: string };

export interface ConditionalMenuActionDeps {
  inst: SpreadsheetInstance | null;
  ribbonLang: 'ja' | 'en';
  range: Range | null;
  cfFill: CfFillStyle;
  cfTopFill: CfFillStyle;
  promptCfNumber: (
    title: string,
    initial?: number,
    options?: { min?: number; max?: number; step?: number },
  ) => Promise<number | null>;
  promptCfText: (title: string, label: string, initial?: string) => Promise<string | null>;
  showChoiceDialog: <T extends string>(spec: {
    title: string;
    label: string;
    options: ReadonlyArray<{ value: T; label: string }>;
    initial?: T;
    cancelLabel?: string;
  }) => Promise<T | null>;
  showMessage: (spec: { title: string; message: string }) => Promise<void>;
  refreshWorkbookCells: () => void;
  addConditionalRuleFromRibbon: (rule: ConditionalRule) => void;
}

export const applyConditionalMenuAction = async (
  deps: ConditionalMenuActionDeps,
  action: string,
  panel?: string,
): Promise<void> => {
  const {
    inst: i,
    range,
    ribbonLang,
    cfFill,
    cfTopFill,
    promptCfNumber,
    promptCfText,
    showChoiceDialog,
    showMessage,
    refreshWorkbookCells,
    addConditionalRuleFromRibbon,
  } = deps;
  if (!i || !range) return;
  const title = buildCfMenuText(ribbonLang);
  if (action === 'new-rule') {
    i.openConditionalDialog({ mode: 'new', kind: conditionalRuleKindForPanel(panel) });
    return;
  }
  if (action === 'manage') {
    i.openCfRulesDialog();
    return;
  }
  if (action === 'cell-gt' || action === 'cell-lt' || action === 'cell-eq') {
    const n = await promptCfNumber(
      action === 'cell-gt' ? title.greater : action === 'cell-lt' ? title.less : title.equal,
      0,
    );
    if (n === null) return;
    addConditionalRuleFromRibbon({
      kind: 'cell-value',
      range,
      op: action === 'cell-gt' ? '>' : action === 'cell-lt' ? '<' : '=',
      a: n,
      apply: cfFill,
    });
    return;
  }
  if (action === 'cell-between') {
    const a = await promptCfNumber(title.between, 0);
    if (a === null) return;
    const b = await promptCfNumber(title.between, 100);
    if (b === null) return;
    addConditionalRuleFromRibbon({
      kind: 'cell-value',
      range,
      op: 'between',
      a: Math.min(a, b),
      b: Math.max(a, b),
      apply: cfFill,
    });
    return;
  }
  if (action === 'text-contains') {
    const text = await promptCfText(title.text, title.textPrompt);
    if (text === null) return;
    addConditionalRuleFromRibbon({ kind: 'text-contains', range, text, apply: cfFill });
    return;
  }
  if (action === 'date-occurring') {
    const period = await showChoiceDialog<DatePeriod>({
      title: title.date,
      label: title.datePrompt,
      options: cfDatePeriodOptions(ribbonLang),
      initial: 'today',
      cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
    });
    if (period === null) return;
    if (!isDatePeriod(period)) {
      void showMessage({
        title: title.date,
        message:
          ribbonLang === 'ja'
            ? '指定できる日付条件を入力してください。'
            : 'Enter one of the supported date conditions.',
      });
      return;
    }
    addConditionalRuleFromRibbon({ kind: 'date-occurring', range, period, apply: cfFill });
    return;
  }
  if (
    action === 'top10' ||
    action === 'bottom10' ||
    action === 'top10-percent' ||
    action === 'bottom10-percent'
  ) {
    const isPercent = action.endsWith('-percent');
    const n = await promptCfNumber(
      action.startsWith('top')
        ? isPercent
          ? title.top10Percent
          : title.top10
        : isPercent
          ? title.bottom10Percent
          : title.bottom10,
      10,
      { min: 1, max: isPercent ? 100 : undefined, step: 1 },
    );
    if (n === null) return;
    addConditionalRuleFromRibbon({
      kind: 'top-bottom',
      range,
      mode: action.startsWith('top') ? 'top' : 'bottom',
      n: Math.max(1, Math.floor(n)),
      percent: isPercent,
      apply: cfTopFill,
    });
    return;
  }
  if (isConditionalPresetAction(action)) {
    let changed = false;
    recordConditionalRulesChange(i.history, i.store, () => {
      changed = applyConditionalPresetAction(i.store, action, range);
    });
    if (changed) refreshWorkbookCells();
  }
};
