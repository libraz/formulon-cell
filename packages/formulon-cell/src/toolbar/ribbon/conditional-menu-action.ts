import type { Locale } from '../../i18n/strings.js';

// Dispatches the ribbon's conditional formatting menu actions: opens the
// new-rule / manage dialogs, prompts for inline rule input, or applies preset
// actions on the store. State (workbook instance, locale, fills, prompt
// closures) is injected via a deps struct so the function stays host-agnostic.

import type { ConditionalPresetAction } from '../../commands/conditional-format.js';
import {
  applyConditionalPresetAction,
  CONDITIONAL_PRESET_ACTIONS,
} from '../../commands/conditional-format.js';
import { recordConditionalRulesChange } from '../../commands/slice-history.js';
import type { Range } from '../../engine/types.js';
import type { ConditionalDialogOpenOptions } from '../../interact/conditional-dialog.js';
import type { SpreadsheetInstance } from '../../mount.js';
import type { CellFormat, ConditionalRule } from '../../store/store.js';

import { buildCfMenuText } from './menus/conditional.js';

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

const cfDatePeriodOptions = (
  labels: Record<DatePeriod, string>,
): Array<{ value: DatePeriod; label: string }> =>
  DATE_PERIODS.map((value) => ({ value, label: labels[value] }));

const conditionalPresetActions = new Set<string>(CONDITIONAL_PRESET_ACTIONS);

const isConditionalPresetAction = (action: string): action is ConditionalPresetAction =>
  conditionalPresetActions.has(action);

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

const isMacHost = (host: HTMLElement): boolean =>
  (host.closest<HTMLElement>('[data-fc-platform]') ?? host).dataset.fcPlatform === 'mac';

/** The full-dialog setup each Highlight Cells / Top-Bottom preset opens on the Mac profile. */
const macPresetDialogOptions = (action: string): ConditionalDialogOpenOptions | null => {
  switch (action) {
    case 'cell-gt':
      return { kind: 'cell-value', cellValueOp: '>' };
    case 'cell-lt':
      return { kind: 'cell-value', cellValueOp: '<' };
    case 'cell-eq':
      return { kind: 'cell-value', cellValueOp: '=' };
    case 'cell-between':
      return { kind: 'cell-value', cellValueOp: 'between' };
    case 'text-contains':
      return { kind: 'text-contains' };
    case 'date-occurring':
      return { kind: 'date-occurring', datePeriod: 'yesterday' };
    case 'duplicates':
      return { kind: 'duplicates' };
    case 'top10':
    case 'top10-percent':
    case 'bottom10':
    case 'bottom10-percent':
      return {
        kind: 'top-bottom',
        topBottomMode: action.startsWith('top') ? 'top' : 'bottom',
        topBottomPercent: action.endsWith('-percent'),
      };
    case 'above-avg':
      return { kind: 'average', averageMode: 'above' };
    case 'below-avg':
      return { kind: 'average', averageMode: 'below' };
    default:
      return null;
  }
};

export type CfFillStyle = { readonly fill: string; readonly color: string };
export type CfApplyStyle = Partial<
  Pick<CellFormat, 'fill' | 'color' | 'bold' | 'italic' | 'underline' | 'strike'>
>;
export type CfNumberDialogResult = {
  readonly values: readonly number[];
  /** Per field, the `=` reference the user entered, or `null` for the number in `values`. */
  readonly references?: readonly (string | null)[];
  readonly style: CfApplyStyle;
};
export type CfTextDialogResult = { readonly text: string; readonly style: CfApplyStyle };

export interface ConditionalMenuActionDeps {
  inst: SpreadsheetInstance | null;
  ribbonLang: Locale;
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
    /** The fields also take a cell reference. */
    allowReference?: boolean;
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
  const macDialog = macPresetDialogOptions(action);
  if (macDialog && isMacHost(i.host)) {
    i.openConditionalDialog({ mode: 'new', ...macDialog });
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
      allowReference: true,
    });
    if (result === null) return;
    const n = result.values[0] ?? 0;
    addConditionalRuleFromRibbon({
      kind: 'cell-value',
      range,
      op: action === 'cell-gt' ? '>' : action === 'cell-lt' ? '<' : '=',
      a: result.references?.[0] ?? n,
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
      allowReference: true,
    });
    if (result === null) return;
    const a = result.values[0] ?? 0;
    const b = result.values[1] ?? a;
    const refA = result.references?.[0] ?? null;
    const refB = result.references?.[1] ?? null;
    // Bounds are ordered only when both are numbers; a reference keeps its position.
    const ordered = refA === null && refB === null;
    addConditionalRuleFromRibbon({
      kind: 'cell-value',
      range,
      op: 'between',
      a: ordered ? Math.min(a, b) : (refA ?? a),
      b: ordered ? Math.max(a, b) : (refB ?? b),
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
