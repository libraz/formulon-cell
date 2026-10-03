import { hyperlinkAt } from '../../commands/hyperlinks.js';
import {
  createRibbonChartFromSelection,
  executeRibbonHyperlinkAction,
  executeRibbonPivotTableAction,
  mutators,
  type RibbonPivotTableAction,
  type SessionChartKind,
  type SpreadsheetInstance,
} from '../../index.js';
import { showSymbolDialog } from '../../toolbar/dialogs/symbol.js';
import type { DynamicDropdownsCtx } from '../../toolbar/ribbon/dynamic-dropdowns.js';
import { setMenuControlDisabled, showInstanceReport } from './menu-feedback.js';
import { normalizedSelectionRange } from './selection.js';

const buildLinksAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyLinksAction'] =>
  async (action) => {
    const linkAction =
      action === 'hyperlink'
        ? 'edit'
        : action === 'external' || action === 'open' || action === 'clear'
          ? action
          : null;
    if (!linkAction) return;
    const strings = instance.i18n.strings;
    const result = executeRibbonHyperlinkAction({
      store: instance.store,
      workbook: instance.workbook,
      history: instance.history,
      action: linkAction,
      strings: {
        linkOpen: strings.ribbonMenu.linkOpen,
        linkNoHyperlink: strings.ribbonMenu.linkNoHyperlink,
      },
    });
    if (result.kind === 'open-hyperlink-dialog') {
      instance.openHyperlinkDialog();
      return;
    }
    if (result.kind === 'open-external-dialog') {
      instance.openExternalLinksDialog();
      return;
    }
    if (result.kind === 'open-url') {
      window.open(result.url, '_blank', 'noopener,noreferrer');
      instance.host.focus();
      return;
    }
    if (result.kind === 'report') {
      await showInstanceReport(instance, result.report.title, result.report.items);
      return;
    }
    instance.host.focus();
  };

const updateLinksMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateLinksMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const hasHyperlink = hyperlinkAt(state, state.selection.active) !== null;
    const disabledReason = instance.i18n.strings.ribbonMenu.linkNoHyperlink;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-link-action]')) {
      const action = button.dataset.linkAction;
      const disabled = (action === 'open' || action === 'clear') && !hasHyperlink;
      setMenuControlDisabled(button, disabled, disabledReason);
    }
  };

const chartKindFromAction = (action: string): SessionChartKind => {
  if (
    action === 'bar' ||
    action === 'line' ||
    action === 'area' ||
    action === 'pie' ||
    action === 'scatter'
  ) {
    return action;
  }
  return 'column';
};

const buildChartAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['createChartFromSelection'] =>
  (kind) => {
    createRibbonChartFromSelection({
      store: instance.store,
      range: normalizedSelectionRange(instance),
      action: kind,
      history: instance.history,
    });
    instance.host.focus();
  };

const buildRecommendedChartAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['createRecommendedChartFromSelection'] =>
  async () => {
    const strings = instance.i18n.strings;
    await showInstanceReport(instance, strings.ribbonMenu.recommendedCharts, [
      {
        severity: 'info',
        label: strings.ribbon.chart,
        detail: strings.workbookObjects.compatibilityDetails.chartAuthoring,
      },
    ]);
    instance.host.focus();
  };

const buildPivotTableAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyPivotTableAction'] =>
  async (action) => {
    if (action === 'new-sheet' || action === 'existing-sheet') {
      (instance.openPivotTableDialog as (opts?: { placement?: 'new' | 'existing' }) => void)({
        placement: action === 'new-sheet' ? 'new' : 'existing',
      });
      return;
    }
    const pivotAction = (
      action === 'recommended' ||
      action === 'new-sheet' ||
      action === 'existing-sheet' ||
      action === 'refresh'
        ? action
        : 'dialog'
    ) as RibbonPivotTableAction;
    const strings = instance.i18n.strings;
    const result = executeRibbonPivotTableAction({
      store: instance.store,
      workbook: instance.workbook,
      action: pivotAction,
      history: instance.history,
      strings: {
        pivotTable: strings.ribbon.pivotTable,
        pivotTableNewSheet: strings.ribbonMenu.pivotTableNewSheet,
        pivotTableRefreshData: strings.ribbonMenu.pivotTableRefreshData,
        pivotTableRefreshUnavailable: strings.ribbonMenu.pivotTableRefreshUnavailable,
        recommendedPivotTables: strings.ribbonMenu.recommendedPivotTables,
        pivotAuthoringDetail: strings.workbookObjects.compatibilityDetails.pivotAuthoring,
        workbookStructureProtectedBlocked: strings.ribbonMenu.workbookStructureProtectedBlocked,
      },
    });
    if (result.kind === 'open-dialog') {
      instance.openPivotTableDialog();
      return;
    }
    if (result.kind === 'report') {
      await showInstanceReport(instance, result.report.title, result.report.items);
      return;
    }
    instance.host.focus();
  };

const insertSymbolAtActiveCell = (instance: SpreadsheetInstance, symbol: string): void => {
  const addr = instance.store.getState().selection.active;
  instance.history.begin();
  try {
    instance.workbook.setText(addr, symbol);
    mutators.replaceCells(instance.store, instance.workbook.cells(addr.sheet));
  } finally {
    instance.history.end();
  }
};

const buildSymbolAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applySymbolAction'] =>
  async (symbol) => {
    if (symbol === 'more') {
      const selected = await showSymbolDialog({
        text: instance.i18n.strings.ribbonMenu,
        okLabel: instance.i18n.strings.hyperlinkDialog.ok,
        cancelLabel: instance.i18n.strings.hyperlinkDialog.cancel,
      });
      if (selected) insertSymbolAtActiveCell(instance, selected);
      instance.host.focus();
      return;
    }
    insertSymbolAtActiveCell(instance, symbol);
    instance.host.focus();
  };

type InsertDropdownDefaults = Pick<
  DynamicDropdownsCtx,
  | 'applyPivotTableAction'
  | 'createRecommendedChartFromSelection'
  | 'createChartFromSelection'
  | 'chartKindFromAction'
  | 'applySymbolAction'
  | 'applyLinksAction'
  | 'updateLinksMenu'
>;

export function createInsertDropdownDefaults(
  instance: SpreadsheetInstance,
): InsertDropdownDefaults {
  return {
    applyPivotTableAction: buildPivotTableAction(instance),
    createRecommendedChartFromSelection: buildRecommendedChartAction(instance),
    createChartFromSelection: buildChartAction(instance),
    chartKindFromAction,
    applySymbolAction: buildSymbolAction(instance),
    applyLinksAction: buildLinksAction(instance),
    updateLinksMenu: updateLinksMenu(instance),
  };
}
