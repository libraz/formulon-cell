import type { ChartSpec } from '../../charts/types.js';
import { blockDataTableCells } from '../../commands/data-table-guard.js';
import { hyperlinkAt, isFollowableHyperlinkTarget } from '../../commands/hyperlinks.js';
import {
  executeRibbonPivotTableAction,
  type RibbonPivotTableAction,
} from '../../commands/pivot-table.js';
import {
  isCellWritable,
  runAfterAllowedRangeUnlock,
  warnProtected,
} from '../../commands/protection.js';
import { createRibbonChartFromSelection } from '../../commands/session-chart.js';
import { openHyperlinkDialogFrom, openPivotTableDialogFrom } from '../../interact/dialog-entry.js';
import { focusSheet } from '../../interact/dialog-focus.js';
import { executeRibbonHyperlinkAction } from '../../interact/ribbon-hyperlink.js';
import type { SpreadsheetInstance } from '../../mount.js';
import { rangeOfAddr } from '../../store/selection-geometry.js';
import { mutators, type SessionChartKind } from '../../store/store.js';
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
      context: {
        store: instance.store,
        history: instance.history,
        workbook: instance.workbook,
        origin: 'ribbon',
        commandId: 'hyperlinkInsert',
      },
      action: linkAction,
      strings: {
        linkOpen: strings.ribbonMenu.linkOpen,
        linkNoHyperlink: strings.ribbonMenu.linkNoHyperlink,
        linkUnsupportedTarget: strings.ribbonMenu.linkUnsupportedTarget,
      },
    });
    if (result.kind === 'open-hyperlink-dialog') {
      openHyperlinkDialogFrom(instance.store, { origin: 'ribbon', commandId: 'hyperlinkInsert' });
      return;
    }
    if (result.kind === 'open-external-dialog') {
      instance.openExternalLinksDialog();
      return;
    }
    if (result.kind === 'open-url') {
      instance.followActiveHyperlink();
      focusSheet(instance.host);
      return;
    }
    if (result.kind === 'report') {
      await showInstanceReport(instance, result.report.title, result.report.items);
      return;
    }
    focusSheet(instance.host);
  };

const updateLinksMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateLinksMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const target = hyperlinkAt(state, state.selection.active);
    const hasHyperlink = target !== null;
    const canOpenHyperlink = hasHyperlink && isFollowableHyperlinkTarget(target);
    const disabledReason = instance.i18n.strings.ribbonMenu.linkNoHyperlink;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-link-action]')) {
      const action = button.dataset.linkAction;
      const disabled =
        (action === 'open' && !canOpenHyperlink) || (action === 'clear' && !hasHyperlink);
      setMenuControlDisabled(
        button,
        disabled,
        action === 'open' && hasHyperlink && !canOpenHyperlink
          ? instance.i18n.strings.ribbonMenu.linkUnsupportedTarget
          : disabledReason,
      );
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

/** Insert a chart of the selection with settings beyond its kind, as the ribbon command
 *  `commandId`. */
export const createChartVariantFromSelection = (
  instance: SpreadsheetInstance,
  kind: SessionChartKind,
  settings?: Omit<ChartSpec, 'kind'>,
  commandId = 'chartInsert',
): void => {
  // The sheet takes focus first, so the chart the insert selects can take it over.
  focusSheet(instance.host);
  createRibbonChartFromSelection({
    store: instance.store,
    wb: instance.workbook,
    range: normalizedSelectionRange(instance),
    action: kind,
    settings,
    history: instance.history,
    authorization: { origin: 'ribbon', commandId },
  });
};

const buildChartAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['createChartFromSelection'] =>
  (kind) =>
    createChartVariantFromSelection(instance, kind);

const buildRecommendedChartAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['createRecommendedChartFromSelection'] =>
  async () => {
    // The sheet takes focus first, so the chart the insert selects can take it over.
    focusSheet(instance.host);
    createRibbonChartFromSelection({
      store: instance.store,
      wb: instance.workbook,
      range: normalizedSelectionRange(instance),
      action: 'recommended',
      history: instance.history,
      authorization: { origin: 'ribbon', commandId: 'chartInsert' },
    });
  };

const buildPivotTableAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyPivotTableAction'] =>
  async (action) => {
    if (action === 'new-sheet' || action === 'existing-sheet') {
      openPivotTableDialogFrom(instance.store, action === 'new-sheet' ? 'new' : 'existing', {
        origin: 'ribbon',
        commandId: 'pivotTableInsert',
      });
      return;
    }
    // The Mac desktop inserts its recommendation directly, without a dialog.
    const mac =
      (instance.host.closest<HTMLElement>('.fc-host') ?? instance.host).dataset.fcPlatform ===
      'mac';
    const pivotAction = (
      action === 'recommended' && mac
        ? 'recommended-insert'
        : action === 'recommended' ||
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
      origin: 'ribbon',
      commandId: pivotAction === 'refresh' ? 'pivotRefresh' : 'pivotTableInsert',
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
      openPivotTableDialogFrom(instance.store, undefined, {
        origin: 'ribbon',
        commandId: 'pivotTableInsert',
      });
      return;
    }
    if (result.kind === 'report') {
      await showInstanceReport(instance, result.report.title, result.report.items);
      return;
    }
    focusSheet(instance.host);
  };

const insertSymbolAtActiveCell = (instance: SpreadsheetInstance, symbol: string): void => {
  const addr = instance.store.getState().selection.active;
  const cell = rangeOfAddr(addr);
  runAfterAllowedRangeUnlock(instance.store, [cell], () => {
    const target = instance.store.getState().selection.active;
    if (!isCellWritable(instance.store.getState(), target)) {
      warnProtected(target);
      return;
    }
    if (blockDataTableCells(instance.workbook, [target])) return;
    instance.history.begin();
    try {
      instance.workbook.setText(target, symbol);
      mutators.replaceCells(instance.store, instance.workbook.cells(target.sheet));
    } finally {
      instance.history.end();
    }
  });
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
      focusSheet(instance.host);
      return;
    }
    insertSymbolAtActiveCell(instance, symbol);
    focusSheet(instance.host);
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
