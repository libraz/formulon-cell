import { findPivotTableAtCell } from '../engine/passthrough-sync.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { ExtensionHandle } from '../extensions/index.js';
import type { FxDialogOpenOptions } from '../interact/fx-dialog.js';
import type { SpreadsheetStore } from '../store/store.js';
import type { AlwaysOnDialogs } from './always-on-dialogs.js';
import type { EngineBinding } from './engine-binding.js';
import type { HostFeatureState } from './host-features.js';
import type { SpreadsheetInstance } from './types.js';

export type InstanceDialogApi = Pick<
  SpreadsheetInstance,
  | 'openConditionalDialog'
  | 'openIterativeDialog'
  | 'openExternalLinksDialog'
  | 'openCfRulesDialog'
  | 'openCellStylesGallery'
  | 'openEvaluateFormulaDialog'
  | 'openFunctionArguments'
  | 'openHyperlinkDialog'
  | 'openCommentDialog'
  | 'openFindReplace'
  | 'closeFindReplace'
  | 'openPasteSpecial'
  | 'openNamedRangeDialog'
  | 'openDefineNameDialog'
  | 'openPageSetup'
  | 'openFormatDialog'
  | 'openDataValidationDialog'
  | 'openGoTo'
  | 'openGoToSpecial'
  | 'openFilterDropdown'
  | 'openWatchWindow'
  | 'closeWatchWindow'
  | 'toggleWatchWindow'
  | 'openQuickAnalysis'
  | 'openWorkbookObjects'
  | 'openPivotFieldList'
  | 'openActivePivotFieldList'
  | 'openPivotTableDialog'
  | 'addSlicer'
  | 'removeSlicer'
>;

export interface InstanceDialogApiDeps {
  featureState: HostFeatureState;
  alwaysOnDialogs: AlwaysOnDialogs;
  store: SpreadsheetStore;
  getWb: () => WorkbookHandle;
  getBinding: () => EngineBinding;
  getUserHandle: (id: string) => ExtensionHandle | undefined;
  /** True while an interaction policy blocks dialog routes. */
  isRestricted: () => boolean;
  ensureWatchWindow: () => void;
  refreshFeaturesView: () => void;
}

/** Instance methods that route to host dialogs, with user-extension overrides. */
export function createInstanceDialogApi(deps: InstanceDialogApiDeps): InstanceDialogApi {
  const {
    featureState,
    alwaysOnDialogs,
    store,
    getWb,
    getBinding,
    getUserHandle,
    isRestricted,
    ensureWatchWindow,
    refreshFeaturesView,
  } = deps;
  return {
    openConditionalDialog(options) {
      featureState.conditionalDialog?.open(options);
    },
    openIterativeDialog() {
      featureState.iterativeDialog?.open();
    },
    openExternalLinksDialog() {
      alwaysOnDialogs.openExternalLinks();
    },
    openCfRulesDialog() {
      if (isRestricted()) return;
      alwaysOnDialogs.openCfRules();
    },
    openCellStylesGallery() {
      if (isRestricted()) return;
      alwaysOnDialogs.openCellStyles();
    },
    openEvaluateFormulaDialog() {
      alwaysOnDialogs.openEvaluateFormula();
    },
    openFunctionArguments(seedName?: string, options?: FxDialogOpenOptions) {
      featureState.fxDialog?.open(seedName, options);
    },
    openHyperlinkDialog() {
      featureState.hyperlinkDialog?.open();
    },
    openCommentDialog() {
      featureState.commentDialog?.open();
    },
    openFindReplace(tab?: 'find' | 'replace') {
      getBinding().findReplace?.open(tab);
    },
    closeFindReplace() {
      getBinding().findReplace?.close();
    },
    openPasteSpecial(opts) {
      getBinding().pasteSpecialDialog?.open(opts);
    },
    openNamedRangeDialog() {
      featureState.namedRangeDialog?.open();
    },
    openDefineNameDialog() {
      featureState.namedRangeDialog?.openNew();
    },
    openPageSetup(tab) {
      featureState.pageSetupDialog?.open(tab);
    },
    openFormatDialog(tab) {
      featureState.formatDialog?.open(tab);
    },
    openDataValidationDialog() {
      featureState.formatDialog?.open('more', { mode: 'dataValidation', focus: 'validation' });
    },
    openGoTo() {
      featureState.goToDialog?.open('go-to');
    },
    openGoToSpecial() {
      featureState.goToDialog?.open('special');
    },
    openFilterDropdown(range, col) {
      if (isRestricted()) return;
      alwaysOnDialogs.openFilterAtHeader(range, col);
    },
    openWatchWindow() {
      if (isRestricted()) return;
      ensureWatchWindow();
      featureState.watchPanel?.open();
      refreshFeaturesView();
    },
    closeWatchWindow() {
      featureState.watchPanel?.close();
    },
    toggleWatchWindow() {
      if (isRestricted()) return;
      ensureWatchWindow();
      featureState.watchPanel?.toggle();
      refreshFeaturesView();
    },
    openQuickAnalysis() {
      const userQuick = getUserHandle('quickAnalysis') as
        | (ExtensionHandle & { open?: () => void })
        | undefined;
      if (userQuick?.open) {
        userQuick.open();
        return;
      }
      getBinding().quickAnalysis?.open();
    },
    openWorkbookObjects() {
      const userObjects = getUserHandle('workbookObjects') as
        | (ExtensionHandle & { open?: () => void })
        | undefined;
      if (userObjects?.open) {
        userObjects.open();
        return;
      }
      featureState.workbookObjects?.open();
    },
    openPivotFieldList(sheetIndex, pivotIndex) {
      if (isRestricted()) return false;
      const userObjects = getUserHandle('workbookObjects') as
        | (ExtensionHandle & {
            openPivotFieldList?: (sheetIndex: number, pivotIndex: number) => boolean;
          })
        | undefined;
      if (userObjects?.openPivotFieldList) {
        return userObjects.openPivotFieldList(sheetIndex, pivotIndex);
      }
      return featureState.workbookObjects?.openPivotFieldList(sheetIndex, pivotIndex) ?? false;
    },
    openActivePivotFieldList() {
      if (isRestricted()) return false;
      const pivot = findPivotTableAtCell(getWb(), store.getState().selection.active);
      if (!pivot) return false;
      const userObjects = getUserHandle('workbookObjects') as
        | (ExtensionHandle & {
            openPivotFieldList?: (sheetIndex: number, pivotIndex: number) => boolean;
          })
        | undefined;
      if (userObjects?.openPivotFieldList) {
        return userObjects.openPivotFieldList(pivot.sheetIndex, pivot.pivotIndex);
      }
      return (
        featureState.workbookObjects?.openPivotFieldList(pivot.sheetIndex, pivot.pivotIndex) ??
        false
      );
    },
    openPivotTableDialog(opts) {
      if (isRestricted()) return;
      const userPivot = getUserHandle('pivotTableDialog') as
        | (ExtensionHandle & { open?: (opts?: { placement?: 'new' | 'existing' }) => void })
        | undefined;
      if (userPivot?.open) {
        userPivot.open(opts);
        return;
      }
      featureState.pivotTableDialog?.open(opts);
    },
    addSlicer(input) {
      if (!featureState.slicer) {
        throw new Error('addSlicer: features.slicer is disabled');
      }
      return featureState.slicer.addSlicer(input);
    },
    removeSlicer(id) {
      featureState.slicer?.removeSlicer(id);
    },
  };
}
