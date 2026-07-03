// Extracted from main.ts. Dispatches the ribbon's conditional formatting menu
// actions: opens the new-rule / manage dialogs, prompts for inline rule input,
// or applies preset actions on the store. State (workbook instance, locale,
// fills, prompt closures) is injected via a deps struct so the function stays
// host-agnostic.

import type {
  CellFormat,
  ConditionalDialogOpenOptions,
  ConditionalPresetAction,
  ConditionalRule,
  Range,
  SpreadsheetInstance,
} from '../../index.js';
import { applyConditionalPresetAction, recordConditionalRulesChange } from '../../index.js';

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
  labels: Record<DatePeriod, string>,
): Array<{ value: DatePeriod; label: string }> =>
  DATE_PERIODS.map((value) => ({ value, label: labels[value] }));

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

export const SUPPORTED_CONDITIONAL_MENU_ACTIONS = new Set<string>([
  'new-rule',
  'manage',
  'cell-gt',
  'cell-lt',
  'cell-between',
  'cell-eq',
  'text-contains',
  'date-occurring',
  'top10',
  'bottom10',
  'top10-percent',
  'bottom10-percent',
  ...conditionalPresetActions,
]);

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
export type CfApplyStyle = Partial<
  Pick<CellFormat, 'fill' | 'color' | 'bold' | 'italic' | 'underline' | 'strike'>
>;
export type CfNumberDialogResult = {
  readonly values: readonly number[];
  readonly style: CfApplyStyle;
};
export type CfTextDialogResult = { readonly text: string; readonly style: CfApplyStyle };

export interface ConditionalMenuActionDeps {
  inst: SpreadsheetInstance | null;
  ribbonLang: 'ja' | 'en';
  range: Range | null;
  cfFill: CfFillStyle;
  promptCfNumber: (spec: {
    title: string;
    label: string;
    initial?: number;
    min?: number;
    max?: number;
    step?: number;
    secondLabel?: string;
    secondInitial?: number;
    initialStyle?: string;
  }) => Promise<CfNumberDialogResult | null>;
  promptCfText: (spec: {
    title: string;
    label: string;
    initial?: string;
    initialStyle?: string;
  }) => Promise<CfTextDialogResult | null>;
  showChoiceDialog: <T extends string>(spec: {
    title: string;
    label: string;
    options: ReadonlyArray<{ value: T; label: string }>;
    initial?: T;
    okLabel: string;
    cancelLabel: string;
  }) => Promise<T | null>;
  showMessage: (spec: { title: string; message: string; okLabel: string }) => Promise<void>;
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
    const dialogTitle =
      action === 'cell-gt' ? title.greater : action === 'cell-lt' ? title.less : title.equal;
    const result = await promptCfNumber({
      title: dialogTitle,
      label:
        action === 'cell-gt'
          ? title.greaterPrompt
          : action === 'cell-lt'
            ? title.lessPrompt
            : title.equalPrompt,
      initial: 0,
      initialStyle: 'light-red-dark-red',
    });
    if (result === null) return;
    const n = result.values[0] ?? 0;
    addConditionalRuleFromRibbon({
      kind: 'cell-value',
      range,
      op: action === 'cell-gt' ? '>' : action === 'cell-lt' ? '<' : '=',
      a: n,
      apply: result.style,
    });
    return;
  }
  if (action === 'cell-between') {
    const result = await promptCfNumber({
      title: title.between,
      label: title.betweenPrompt,
      initial: 0,
      secondLabel: title.betweenAndPrompt,
      secondInitial: 100,
      initialStyle: 'light-red-dark-red',
    });
    if (result === null) return;
    const a = result.values[0] ?? 0;
    const b = result.values[1] ?? a;
    addConditionalRuleFromRibbon({
      kind: 'cell-value',
      range,
      op: 'between',
      a: Math.min(a, b),
      b: Math.max(a, b),
      apply: result.style,
    });
    return;
  }
  if (action === 'text-contains') {
    const result = await promptCfText({
      title: title.text,
      label: title.textPrompt,
      initialStyle: 'light-red-dark-red',
    });
    if (result === null) return;
    addConditionalRuleFromRibbon({
      kind: 'text-contains',
      range,
      text: result.text,
      apply: result.style,
    });
    return;
  }
  if (action === 'date-occurring') {
    const period = await showChoiceDialog<DatePeriod>({
      title: title.date,
      label: title.datePrompt,
      options: cfDatePeriodOptions(title.datePeriods),
      initial: 'today',
      okLabel: title.ok,
      cancelLabel: title.cancel,
    });
    if (period === null) return;
    if (!isDatePeriod(period)) {
      void showMessage({
        title: title.date,
        message: title.dateUnsupported,
        okLabel: title.ok,
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
    const result = await promptCfNumber({
      title: action.startsWith('top')
        ? isPercent
          ? title.top10Percent
          : title.top10
        : isPercent
          ? title.bottom10Percent
          : title.bottom10,
      label: title.topBottomPrompt,
      initial: 10,
      min: 1,
      max: isPercent ? 100 : undefined,
      step: 1,
      initialStyle: 'green-dark-green',
    });
    if (result === null) return;
    const n = result.values[0] ?? 10;
    addConditionalRuleFromRibbon({
      kind: 'top-bottom',
      range,
      mode: action.startsWith('top') ? 'top' : 'bottom',
      n: Math.max(1, Math.floor(n)),
      percent: isPercent,
      apply: result.style,
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
