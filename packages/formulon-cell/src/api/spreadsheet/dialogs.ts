import type { Range } from '../../engine/types.js';
import type { FeatureId } from '../../extensions/features.js';
import type { ConditionalDialogOpenOptions } from '../../interact/conditional-dialog.js';
import type { FxDialogOpenOptions } from '../../interact/fx-dialog.js';
import type { PasteSpecialOpenOptions } from '../../interact/paste-special.js';
import type { ChartSettingsTarget } from '../../interact/session-chart-settings.js';
import type { SpreadsheetInstance as LegacyInstance } from '../../mount/types.js';
import { spreadsheetFailure } from '../errors.js';

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
}

/** A dialog `SpreadsheetInstance.dialogs` opens: the opener's name without `open` and `Dialog`. */
export type DialogId = keyof DialogOptions;

/** Opens and closes the host dialogs. */
export interface Dialogs {
  /** Opens a dialog on the current selection. Throws `feature-disabled` when its feature is off,
   *  `policy-denied` when the interaction policy blocks it. */
  open<K extends DialogId>(id: K, options?: DialogOptions[K]): void;
  /** Closes a dialog; does nothing when it is not open. */
  close(id: DialogId): void;
  /** Whether `open(id)` would open the dialog now. */
  canOpen(id: DialogId): boolean;
}

interface DialogRoute {
  /** Feature the dialog belongs to; `null` for the dialogs every mount carries. */
  readonly feature: FeatureId | null;
  /** True for the routes that do not consult the interaction policy and stay shut under one. */
  readonly unpoliced?: boolean;
  /** Opens it; `false` when there was nothing to open it on (no chart, no pivot). */
  readonly open: (legacy: LegacyInstance, options: never) => unknown;
  readonly close?: (legacy: LegacyInstance) => void;
}

type Route<K extends DialogId> = Omit<DialogRoute, 'open'> & {
  readonly open: (legacy: LegacyInstance, options: DialogOptions[K] | undefined) => unknown;
};

const ROUTES: { readonly [K in DialogId]: Route<K> } = {
  conditional: { feature: 'conditional', open: (l, o) => l.openConditionalDialog(o) },
  iterative: { feature: 'iterative', open: (l) => l.openIterativeDialog() },
  externalLinks: { feature: null, open: (l) => l.openExternalLinksDialog() },
  cfRules: { feature: null, unpoliced: true, open: (l) => l.openCfRulesDialog() },
  cellStylesGallery: { feature: null, unpoliced: true, open: (l) => l.openCellStylesGallery() },
  evaluateFormula: { feature: null, open: (l) => l.openEvaluateFormulaDialog() },
  functionArguments: {
    feature: 'fxDialog',
    open: (l, o) => {
      const { name, ...rest } = o ?? {};
      l.openFunctionArguments(name, rest);
    },
  },
  hyperlink: { feature: 'hyperlink', open: (l) => l.openHyperlinkDialog() },
  comment: { feature: 'commentDialog', open: (l, o) => l.openCommentDialog(o) },
  findReplace: {
    feature: 'findReplace',
    open: (l, o) => l.openFindReplace(o?.tab),
    close: (l) => l.closeFindReplace(),
  },
  pasteSpecial: { feature: 'pasteSpecial', open: (l, o) => l.openPasteSpecial(o) },
  insertCopiedCells: { feature: null, open: (l) => l.openInsertCopiedCells() },
  namedRange: { feature: 'namedRanges', open: (l) => l.openNamedRangeDialog() },
  defineName: { feature: 'namedRanges', open: (l) => l.openDefineNameDialog() },
  pageSetup: { feature: 'pageSetup', open: (l, o) => l.openPageSetup(o?.tab) },
  format: { feature: 'formatDialog', open: (l, o) => l.openFormatDialog(o?.tab) },
  dataValidation: { feature: 'formatDialog', open: (l) => l.openDataValidationDialog() },
  goTo: { feature: 'gotoSpecial', open: (l) => l.openGoTo() },
  goToSpecial: { feature: 'gotoSpecial', open: (l) => l.openGoToSpecial() },
  filterDropdown: {
    feature: null,
    unpoliced: true,
    open: (l, o) => l.openFilterDropdown(o?.range, o?.col),
  },
  watchWindow: {
    feature: null,
    unpoliced: true,
    open: (l) => l.openWatchWindow(),
    close: (l) => l.closeWatchWindow(),
  },
  quickAnalysis: { feature: 'quickAnalysis', open: (l) => l.openQuickAnalysis() },
  workbookObjects: {
    feature: 'workbookObjects',
    open: (l) => l.openWorkbookObjects(),
    close: (l) => l.closeWorkbookObjects(),
  },
  accessibilityChecker: { feature: null, open: (l) => l.openAccessibilityChecker() },
  chartSettings: { feature: 'charts', open: (l, o) => l.openChartSettings(o?.tab) },
  pivotFieldList: {
    feature: 'workbookObjects',
    unpoliced: true,
    open: (l, o) =>
      o?.sheet !== undefined && o.pivotIndex !== undefined
        ? l.openPivotFieldList(o.sheet, o.pivotIndex)
        : l.openActivePivotFieldList(),
  },
  activePivotFieldList: {
    feature: 'workbookObjects',
    unpoliced: true,
    open: (l) => l.openActivePivotFieldList(),
  },
  pivotFieldSettings: {
    feature: 'workbookObjects',
    unpoliced: true,
    open: (l) => l.openPivotFieldSettings(),
  },
  pivotOptions: {
    feature: 'workbookObjects',
    unpoliced: true,
    open: (l, o) => l.openPivotOptions(o?.sheet, o?.pivotIndex),
  },
  pivotTable: {
    feature: 'pivotTableDialog',
    unpoliced: true,
    open: (l, o) => l.openPivotTableDialog(o),
  },
};

const routeOf = (id: DialogId): DialogRoute => {
  const route = (ROUTES as Readonly<Record<string, DialogRoute | undefined>>)[id];
  if (!route) throw new TypeError(`unknown dialog id ${String(id)}`);
  return route;
};

type DialogRefusal = 'feature-disabled' | 'policy-denied' | null;

const refusalOf = (legacy: LegacyInstance, route: DialogRoute): DialogRefusal => {
  if (route.feature !== null && legacy.features[route.feature] === undefined) {
    return 'feature-disabled';
  }
  if (route.unpoliced && legacy.commands.policy !== undefined) return 'policy-denied';
  return null;
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
