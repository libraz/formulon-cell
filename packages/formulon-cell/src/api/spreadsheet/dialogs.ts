import type { BuiltInCommandId } from '../../commands/built-in-command-ids.js';
import { canExecuteBuiltInRequest } from '../../commands/built-in-command-policy.js';
import type { Range } from '../../engine/types.js';
import type { FeatureId } from '../../extensions/features.js';
import {
  closeAutoCorrectDialog,
  openAutoCorrectDialog,
} from '../../interact/autocorrect-dialog.js';
import type { ConditionalDialogOpenOptions } from '../../interact/conditional-dialog.js';
import type { FxDialogOpenOptions } from '../../interact/fx-dialog.js';
import type { PasteSpecialOpenOptions } from '../../interact/paste-special.js';
import type { ChartSettingsTarget } from '../../interact/session-chart-settings.js';
import { closeSpellingDialog, openSpellingDialog } from '../../interact/spelling-dialog.js';
import type { SpreadsheetInstance as LegacyInstance } from '../../mount/types.js';
import { failureOfDecision, spreadsheetFailure } from '../errors.js';

/** What each dialog's `open` takes; `undefined` when it takes nothing. */
export interface DialogOptions {
  conditional: ConditionalDialogOpenOptions;
  iterative: undefined;
  externalLinks: undefined;
  cfRules: undefined;
  cellStylesGallery: undefined;
  evaluateFormula: undefined;
  functionArguments: { readonly name?: string } & FxDialogOpenOptions;
  hyperlink: undefined;
  comment: { readonly mode?: 'note' | 'thread' };
  findReplace: { readonly tab?: 'find' | 'replace' };
  pasteSpecial: PasteSpecialOpenOptions;
  insertCopiedCells: undefined;
  namedRange: undefined;
  defineName: undefined;
  pageSetup: { readonly tab?: 'page' | 'margins' | 'headerFooter' | 'sheet' };
  format: {
    readonly tab?: 'number' | 'align' | 'font' | 'border' | 'fill' | 'protection' | 'more';
  };
  dataValidation: undefined;
  goTo: undefined;
  goToSpecial: undefined;
  filterDropdown: { readonly range?: Range; readonly col?: number };
  watchWindow: undefined;
  quickAnalysis: undefined;
  workbookObjects: undefined;
  accessibilityChecker: undefined;
  chartSettings: { readonly tab?: ChartSettingsTarget };
  pivotFieldList: { readonly sheet?: number; readonly pivotIndex?: number };
  activePivotFieldList: undefined;
  pivotFieldSettings: undefined;
  pivotOptions: { readonly sheet?: number; readonly pivotIndex?: number };
  pivotTable: { readonly placement?: 'new' | 'existing' };
  spelling: undefined;
  autoCorrect: undefined;
}

/** A dialog `SpreadsheetInstance.dialogs` opens: the opener's name without `open` and `Dialog`. */
export type DialogId = keyof DialogOptions;

/** Opens and closes the host dialogs. */
export interface Dialogs {
  /** Opens a dialog on the current selection. Throws `feature-disabled` when its feature is off,
   *  `policy-denied` when the interaction policy blocks it, and `protected` when sheet or workbook
   *  protection does, deciding it as the ribbon decides the command that opens the same dialog. */
  open<K extends DialogId>(id: K, options?: DialogOptions[K]): void;
  /** Closes a dialog; does nothing when it is not open. */
  close(id: DialogId): void;
  /** Whether `open(id)` would open the dialog now. */
  canOpen(id: DialogId): boolean;
}

interface DialogRoute {
  /** Feature the dialog belongs to; `null` for the dialogs every mount carries. */
  readonly feature: FeatureId | null;
  /** The built-in command that opens it from the ribbon: opening is decided as that command is,
   *  from the instance API. `null` for a dialog the feature flag alone decides. */
  readonly commandId: BuiltInCommandId | null;
  /** Opens it; `false` when there was nothing to open it on (no chart, no pivot). */
  readonly open: (legacy: LegacyInstance, options: never) => unknown;
  readonly close?: (legacy: LegacyInstance) => void;
}

type Route<K extends DialogId> = Omit<DialogRoute, 'open'> & {
  readonly open: (legacy: LegacyInstance, options: DialogOptions[K] | undefined) => unknown;
};

const ROUTES: { readonly [K in DialogId]: Route<K> } = {
  spelling: {
    feature: null,
    commandId: 'spellingReview',
    open: (l) => openSpellingDialog(l),
    close: (l) => closeSpellingDialog(l),
  },
  autoCorrect: {
    feature: null,
    commandId: null,
    open: (l) => openAutoCorrectDialog(l),
    close: (l) => closeAutoCorrectDialog(l),
  },
  conditional: {
    feature: 'conditional',
    commandId: 'conditional',
    open: (l, o) => l.openConditionalDialog(o),
  },
  iterative: { feature: 'iterative', commandId: null, open: (l) => l.openIterativeDialog() },
  externalLinks: {
    feature: null,
    commandId: 'linksData',
    open: (l) => l.openExternalLinksDialog(),
  },
  cfRules: { feature: null, commandId: 'conditional', open: (l) => l.openCfRulesDialog() },
  cellStylesGallery: {
    feature: null,
    commandId: 'cellStyles',
    open: (l) => l.openCellStylesGallery(),
  },
  evaluateFormula: {
    feature: null,
    commandId: 'evaluateFormula',
    open: (l) => l.openEvaluateFormulaDialog(),
  },
  functionArguments: {
    feature: 'fxDialog',
    commandId: 'fx',
    open: (l, o) => {
      const { name, ...rest } = o ?? {};
      l.openFunctionArguments(name, rest);
    },
  },
  hyperlink: {
    feature: 'hyperlink',
    commandId: 'insertHyperlink',
    open: (l) => l.openHyperlinkDialog(),
  },
  comment: {
    feature: 'commentDialog',
    commandId: 'insertComment',
    open: (l, o) => l.openCommentDialog(o),
  },
  findReplace: {
    feature: 'findReplace',
    commandId: 'findHome',
    open: (l, o) => l.openFindReplace(o?.tab),
    close: (l) => l.closeFindReplace(),
  },
  pasteSpecial: {
    feature: 'pasteSpecial',
    commandId: 'pasteSpecial',
    open: (l, o) => l.openPasteSpecial(o),
  },
  insertCopiedCells: {
    feature: null,
    commandId: 'insertCopiedCells',
    open: (l) => l.openInsertCopiedCells(),
  },
  namedRange: {
    feature: 'namedRanges',
    commandId: 'namedRanges',
    open: (l) => l.openNamedRangeDialog(),
  },
  defineName: {
    feature: 'namedRanges',
    commandId: 'defineName',
    open: (l) => l.openDefineNameDialog(),
  },
  pageSetup: {
    feature: 'pageSetup',
    commandId: 'pageSetup',
    open: (l, o) => l.openPageSetup(o?.tab),
  },
  format: {
    feature: 'formatDialog',
    commandId: 'formatCells',
    open: (l, o) => l.openFormatDialog(o?.tab),
  },
  dataValidation: {
    feature: 'formatDialog',
    commandId: 'dataValidation',
    open: (l) => l.openDataValidationDialog(),
  },
  goTo: { feature: 'gotoSpecial', commandId: null, open: (l) => l.openGoTo() },
  goToSpecial: { feature: 'gotoSpecial', commandId: null, open: (l) => l.openGoToSpecial() },
  filterDropdown: {
    feature: null,
    commandId: 'filterByValue',

    open: (l, o) => l.openFilterDropdown(o?.range, o?.col),
  },
  watchWindow: {
    feature: null,
    commandId: 'watch',

    open: (l) => l.openWatchWindow(),
    close: (l) => l.closeWatchWindow(),
  },
  quickAnalysis: { feature: 'quickAnalysis', commandId: null, open: (l) => l.openQuickAnalysis() },
  workbookObjects: {
    feature: 'workbookObjects',
    commandId: 'workbookObjectsView',
    open: (l) => l.openWorkbookObjects(),
    close: (l) => l.closeWorkbookObjects(),
  },
  accessibilityChecker: {
    feature: null,
    commandId: 'accessibility',
    open: (l) => l.openAccessibilityChecker(),
  },
  chartSettings: {
    feature: 'charts',
    commandId: 'chartFormatSelection',
    open: (l, o) => l.openChartSettings(o?.tab),
  },
  pivotFieldList: {
    feature: 'workbookObjects',
    commandId: 'pivotFieldList',

    open: (l, o) =>
      o?.sheet !== undefined && o.pivotIndex !== undefined
        ? l.openPivotFieldList(o.sheet, o.pivotIndex)
        : l.openActivePivotFieldList(),
  },
  activePivotFieldList: {
    feature: 'workbookObjects',
    commandId: 'pivotFieldList',

    open: (l) => l.openActivePivotFieldList(),
  },
  pivotFieldSettings: {
    feature: 'workbookObjects',
    commandId: 'pivotFieldSettings',

    open: (l) => l.openPivotFieldSettings(),
  },
  pivotOptions: {
    feature: 'workbookObjects',
    commandId: 'pivotOptions',

    open: (l, o) => l.openPivotOptions(o?.sheet, o?.pivotIndex),
  },
  pivotTable: {
    feature: 'pivotTableDialog',
    commandId: 'pivotTableInsert',

    open: (l, o) => l.openPivotTableDialog(o),
  },
};

const routeOf = (id: DialogId): DialogRoute => {
  const route = (ROUTES as Readonly<Record<string, DialogRoute | undefined>>)[id];
  if (!route) throw new TypeError(`unknown dialog id ${String(id)}`);
  return route;
};

type DialogRefusal = 'feature-disabled' | 'policy-denied' | 'protected' | null;

const refusalOf = (legacy: LegacyInstance, route: DialogRoute): DialogRefusal => {
  if (route.feature !== null && legacy.features[route.feature] === undefined) {
    return 'feature-disabled';
  }
  if (route.commandId === null) return null;
  return failureOfDecision(canExecuteBuiltInRequest(legacy.store, route.commandId, 'instanceApi'));
};

/** The `dialogs` object over a mounted spreadsheet. */
export const createDialogs = (legacy: () => LegacyInstance): Dialogs => ({
  open(id, options) {
    const target = legacy();
    const route = routeOf(id);
    const refusal = refusalOf(target, route);
    if (refusal)
      throw spreadsheetFailure(refusal, `dialog ${id} cannot open`, { context: { dialog: id } });
    route.open(target, options as never);
  },
  close(id) {
    const target = legacy();
    const route = routeOf(id);
    if (route.close) {
      route.close(target);
      return;
    }
    const handle = route.feature ? target.features[route.feature] : undefined;
    (handle as { close?: () => void } | undefined)?.close?.();
  },
  canOpen(id) {
    return refusalOf(legacy(), routeOf(id)) === null;
  },
});
