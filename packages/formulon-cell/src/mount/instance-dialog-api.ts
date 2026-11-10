import { hyperlinkAt } from '../commands/hyperlinks.js';
import type { InteractionOperation } from '../commands/interaction-policy.js';
import { findPivotTableAtCell } from '../engine/passthrough-sync.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { ExtensionHandle } from '../extensions/index.js';
import {
  type DialogOpeners,
  INSTANCE_API_DEFINE_NAME_ENTRY,
  INSTANCE_API_FORMAT_ENTRY,
  INSTANCE_API_NAME_MANAGER_ENTRY,
  INSTANCE_API_VALIDATION_ENTRY,
  registerDialogOpeners,
} from '../interact/dialog-entry.js';
import type { FxDialogOpenOptions } from '../interact/fx-dialog.js';
import type { ChartSettingsTarget as ChartSettingsTab } from '../interact/session-chart-settings.js';
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
  | 'followActiveHyperlink'
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
  | 'closeWorkbookObjects'
  | 'openAccessibilityChecker'
  | 'openChartSettings'
  | 'openPivotFieldList'
  | 'openActivePivotFieldList'
  | 'openPivotOptions'
  | 'openPivotFieldSettings'
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
  /** Whether the host policy allows `operation`; `null` names a dialog that edits nothing. */
  permits: (operation: InteractionOperation | null) => boolean;
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
    permits,
    ensureWatchWindow,
    refreshFeaturesView,
  } = deps;
  // Internal UI callers reach these with their own entry; the public methods
  // below are the instance API by definition.
  const openFormatDialog: DialogOpeners['formatDialog'] = (tab, entry) => {
    featureState.formatDialog?.open(tab, { entry });
  };
  const openDataValidationDialog: DialogOpeners['dataValidationDialog'] = (entry) => {
    featureState.formatDialog?.open('more', { mode: 'dataValidation', focus: 'validation', entry });
  };
  const openNamedRangeDialog: DialogOpeners['namedRangeDialog'] = (entry) => {
    featureState.namedRangeDialog?.open(entry);
  };
  const openDefineNameDialog: DialogOpeners['defineNameDialog'] = (entry) => {
    featureState.namedRangeDialog?.openNew(entry);
  };
  registerDialogOpeners(store, {
    formatDialog: openFormatDialog,
    dataValidationDialog: openDataValidationDialog,
    namedRangeDialog: openNamedRangeDialog,
    defineNameDialog: openDefineNameDialog,
  });
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
      if (!permits('format')) return;
      alwaysOnDialogs.openCfRules();
    },
    openCellStylesGallery() {
      if (!permits('format')) return;
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
    followActiveHyperlink() {
      const state = store.getState();
      const target = hyperlinkAt(state, state.selection.active);
      return target !== null && getBinding().followHyperlink(target);
    },
    openCommentDialog(options) {
      featureState.commentDialog?.open(options);
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
      openNamedRangeDialog(INSTANCE_API_NAME_MANAGER_ENTRY);
    },
    openDefineNameDialog() {
      openDefineNameDialog(INSTANCE_API_DEFINE_NAME_ENTRY);
    },
    openPageSetup(tab) {
      featureState.pageSetupDialog?.open(tab);
    },
    openFormatDialog(tab) {
      openFormatDialog(tab, INSTANCE_API_FORMAT_ENTRY);
    },
    openDataValidationDialog() {
      openDataValidationDialog(INSTANCE_API_VALIDATION_ENTRY);
    },
    openGoTo() {
      featureState.goToDialog?.open('go-to');
    },
    openGoToSpecial() {
      featureState.goToDialog?.open('special');
    },
    openFilterDropdown(range, col) {
      if (!permits('filter')) return;
      alwaysOnDialogs.openFilterAtHeader(range, col);
    },
    openWatchWindow() {
      ensureWatchWindow();
      featureState.watchPanel?.open();
      refreshFeaturesView();
    },
    closeWatchWindow() {
      featureState.watchPanel?.close();
    },
    toggleWatchWindow() {
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
    openAccessibilityChecker() {
      if (!featureState.accessibilityPane) return false;
      featureState.accessibilityPane.open();
      return true;
    },
    closeWorkbookObjects() {
      const userObjects = getUserHandle('workbookObjects') as
        | (ExtensionHandle & { close?: () => void })
        | undefined;
      if (userObjects?.close) {
        userObjects.close();
        return;
      }
      featureState.workbookObjects?.close();
    },
    openChartSettings(tab = 'elements') {
      const userCharts = getUserHandle('charts') as
        | (ExtensionHandle & { openSettings?: (tab: ChartSettingsTab) => boolean })
        | undefined;
      if (userCharts?.openSettings) return userCharts.openSettings(tab);
      return featureState.sessionCharts?.openSettings(tab) ?? false;
    },
    openPivotFieldList(sheetIndex, pivotIndex) {
      if (!permits('table')) return false;
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
      if (!permits('table')) return false;
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
    openPivotFieldSettings() {
      if (!permits('table')) return false;
      const addr = store.getState().selection.active;
      const userObjects = getUserHandle('workbookObjects') as
        | (ExtensionHandle & { openPivotFieldSettings?: (addr: Addr) => boolean })
        | undefined;
      if (userObjects?.openPivotFieldSettings) return userObjects.openPivotFieldSettings(addr);
      return featureState.workbookObjects?.openPivotFieldSettings(addr) ?? false;
    },
    openPivotOptions(sheetIndex, pivotIndex) {
      if (!permits('table')) return false;
      const pivot =
        sheetIndex !== undefined && pivotIndex !== undefined
          ? { sheetIndex, pivotIndex }
          : findPivotTableAtCell(getWb(), store.getState().selection.active);
      if (!pivot) return false;
      const userObjects = getUserHandle('workbookObjects') as
        | (ExtensionHandle & {
            openPivotOptions?: (sheetIndex: number, pivotIndex: number) => boolean;
          })
        | undefined;
      if (userObjects?.openPivotOptions) {
        return userObjects.openPivotOptions(pivot.sheetIndex, pivot.pivotIndex);
      }
      return (
        featureState.workbookObjects?.openPivotOptions(pivot.sheetIndex, pivot.pivotIndex) ?? false
      );
    },
    openPivotTableDialog(opts) {
      if (!permits('table')) return;
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
