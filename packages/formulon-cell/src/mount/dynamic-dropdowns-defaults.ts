// Default `DynamicDropdownsCtx` factory. Hosts that don't need a fully
// custom ribbon (React / Vue / quick embed) can call this and pass the
// result to `mountToolbar({ dynamicDropdowns })` — the toolbar's auto-wire
// then handles every menu-item click through these defaults. Hosts retain
// the ability to override any single handler via the `overrides` bag.
//
// Defaults are split into three buckets:
//   1. Pure instance handlers — derived entirely from the instance and the
//      already-exported command helpers (fill / clear / autosum / etc.).
//   2. Instance dispatch — forward to existing `instance.openX` methods.
//   3. Dialog / host-glue stubs — no-op (or browser fallback) when the host
//      doesn't supply a real implementation. Overriding via `overrides` lets
//      hosts plug in their own UI without re-wiring the click delegator.

// Engine and command helpers come through `../index.js`, the same entry
// point hosts use, so `SpreadsheetInstance`, `History` and `WorkbookHandle`
// keep a single type identity.
import {
  applyTextScriptToRange,
  buildRibbonAddInReport,
  formatA1Range,
  mutators,
  parseScriptCommand,
  type RibbonAddInAction,
  type RibbonPdfAction,
  resolveRibbonPdfAction,
  type SpreadsheetInstance,
} from '../index.js';
import { showScriptCommandDialog } from '../toolbar/dialogs/script-command.js';
import type { DynamicDropdownsCtx } from '../toolbar/ribbon/dynamic-dropdowns.js';
import { createCellsDropdownDefaults } from './dynamic-dropdown-defaults/cells.js';
import { createClipboardDropdownDefaults } from './dynamic-dropdown-defaults/clipboard.js';
import { createDataDropdownDefaults } from './dynamic-dropdown-defaults/data.js';
import { createEditingDropdownDefaults } from './dynamic-dropdown-defaults/editing.js';
import { createFormattingDropdownDefaults } from './dynamic-dropdown-defaults/formatting.js';
import { createFormulasDropdownDefaults } from './dynamic-dropdown-defaults/formulas.js';
import { createIllustrationDropdownDefaults } from './dynamic-dropdown-defaults/illustrations.js';
import { createInsertDropdownDefaults } from './dynamic-dropdown-defaults/insert.js';
import { showInstanceReport } from './dynamic-dropdown-defaults/menu-feedback.js';
import { createPageLayoutDropdownDefaults } from './dynamic-dropdown-defaults/page-layout.js';
import { createReviewDropdownDefaults } from './dynamic-dropdown-defaults/review.js';
import { normalizedSelectionRange } from './dynamic-dropdown-defaults/selection.js';
import { createStylesDropdownDefaults } from './dynamic-dropdown-defaults/styles.js';
import { createViewDropdownDefaults } from './dynamic-dropdown-defaults/view.js';

/** Options accepted alongside any host overrides. Lives separately from the
 *  partial-context bag so we can extend with cross-cutting knobs (e.g. a
 *  shared `focusSheet` closure) without polluting the dropdown ctx itself. */
export interface DefaultDynamicDropdownsOptions {
  /** Per-handler overrides. Merged on top of the defaults so the host only
   *  has to supply the ones that need a real dialog. Pass a getter when the
   *  overrides aren't ready at mount time (e.g. the playground builds its
   *  ctx after `mountToolbar` returns) — the ctx will lazily resolve each
   *  handler on every dispatch. */
  overrides?: Partial<DynamicDropdownsCtx> | (() => Partial<DynamicDropdownsCtx>);
  focusSheet?: () => void;
  projectFormatToolbar?: () => void;
  refreshCells?: () => void;
  renderSheetTabs?: () => void;
}

const noop = (): void => undefined;

const buildScriptAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyScriptAction'] =>
  async (action) => {
    const strings = instance.i18n.strings;
    const raw =
      action === 'custom'
        ? await showScriptCommandDialog({
            title: strings.ribbonMenu.scriptDialogTitle,
            label: strings.ribbonMenu.scriptDialogCommand,
            options: [
              { value: 'uppercase', label: strings.ribbonMenu.scriptCommandUppercase },
              { value: 'lowercase', label: strings.ribbonMenu.scriptCommandLowercase },
              { value: 'trim', label: strings.ribbonMenu.scriptCommandTrim },
              { value: 'clear', label: strings.ribbonMenu.scriptCommandClear },
            ],
            initial: 'uppercase',
            okLabel: strings.ribbonMenu.scriptDialogRun,
            cancelLabel: strings.hyperlinkDialog.cancel,
          })
        : action;
    if (raw === null) {
      instance.host.focus();
      return;
    }
    const command = parseScriptCommand(raw);
    if (!command) return;
    const range = normalizedSelectionRange(instance);
    instance.history.begin();
    let count = 0;
    try {
      count = applyTextScriptToRange(instance.store.getState(), instance.workbook, range, command);
      mutators.replaceCells(instance.store, instance.workbook.cells(range.sheet));
    } finally {
      instance.history.end();
    }
    await showInstanceReport(instance, strings.ribbonMenu.automationScriptsTitle, [
      {
        severity: 'info',
        label: strings.ribbonMenu.automationRunStatus.replace('{count}', String(count)),
        detail: strings.ribbonMenu.automationRunDetail
          .replace('{command}', command)
          .replace('{range}', formatA1Range(range))
          .replace('{count}', String(count)),
      },
    ]);
    instance.host.focus();
  };

const buildPdfAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyPdfAction'] =>
  async (action) => {
    const strings = instance.i18n.strings;
    const result = resolveRibbonPdfAction(action as RibbonPdfAction, {
      cellMenu: strings.ribbonMenu,
      pdfTitle: strings.ribbonMenu.pdfCreate,
    });
    if (result.kind === 'open-page-setup') {
      instance.openPageSetup();
      return;
    }
    instance.print('pdf');
    if (result.report) await showInstanceReport(instance, result.report.title, result.report.items);
    instance.host.focus();
  };

const buildAddInAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyAddInAction'] =>
  async (action) => {
    const strings = instance.i18n.strings;
    const report = buildRibbonAddInReport(action as RibbonAddInAction, {
      cellMenu: strings.ribbonMenu,
      addInDefaultTitle: strings.ribbon.addIn,
    });
    if (report) await showInstanceReport(instance, report.title, report.items);
    instance.host.focus();
  };

export function createDefaultDynamicDropdownsCtx(
  instance: SpreadsheetInstance,
  opts: DefaultDynamicDropdownsOptions = {},
): DynamicDropdownsCtx {
  const focusSheet = opts.focusSheet ?? ((): void => instance.host.focus());
  const illustrationDefaults = createIllustrationDropdownDefaults(instance);
  const clipboardDefaults = createClipboardDropdownDefaults(instance);
  const formattingDefaults = createFormattingDropdownDefaults(instance);
  const cellDefaults = createCellsDropdownDefaults(instance, opts);
  const editingDefaults = createEditingDropdownDefaults(instance);
  const formulaDefaults = createFormulasDropdownDefaults(instance);
  const styleDefaults = createStylesDropdownDefaults(instance);
  const insertDefaults = createInsertDropdownDefaults(instance);
  const dataDefaults = createDataDropdownDefaults(instance);
  const reviewDefaults = createReviewDropdownDefaults(instance);
  const pageLayoutDefaults = createPageLayoutDropdownDefaults(instance);
  const viewDefaults = createViewDropdownDefaults(instance);

  const base: DynamicDropdownsCtx = {
    getInst: () => instance,
    updateCalcOptionsMenu: formulaDefaults.updateCalcOptionsMenu,
    updateCellDeleteMenu: cellDefaults.updateCellDeleteMenu,
    updateCellInsertMenu: cellDefaults.updateCellInsertMenu,
    updateClearMenu: editingDefaults.updateClearMenu,
    updateClearArrowsMenu: formulaDefaults.updateClearArrowsMenu,
    updateCurrencyMenu: formattingDefaults.updateCurrencyMenu,
    updateDataValidationMenu: dataDefaults.updateDataValidationMenu,
    updateDefinedNamesMenu: formulaDefaults.updateDefinedNamesMenu,
    updateErrorCheckingMenu: formulaDefaults.updateErrorCheckingMenu,
    updateFormatCellsMenu: cellDefaults.updateFormatCellsMenu,
    updateLinksMenu: insertDefaults.updateLinksMenu,
    updatePageBreaksMenu: pageLayoutDefaults.updatePageBreaksMenu,
    updatePrintAreaMenu: pageLayoutDefaults.updatePrintAreaMenu,
    updateProtectMenu: reviewDefaults.updateProtectMenu,
    updatePageThemeMenu: pageLayoutDefaults.updatePageThemeMenu,
    updateReviewCommentsMenu: reviewDefaults.updateReviewCommentsMenu,
    updateSortMenu: editingDefaults.updateSortMenu,
    updateWatchMenu: formulaDefaults.updateWatchMenu,
    closeBorderMenu: noop,
    closeFreezeMenu: noop,
    closePrintAreaMenu: noop,
    closeSymbolMenu: noop,
    getConditionalMenu: () => document.getElementById('menu-conditional') as HTMLElement | null,
    focusSheet,

    // Pure / instance-derivable defaults.
    updateArrangeMenu: illustrationDefaults.updateArrangeMenu,
    applyCopyAction: clipboardDefaults.applyCopyAction,
    applyRibbonPasteAction: clipboardDefaults.applyRibbonPasteAction,
    updatePasteMenu: clipboardDefaults.updatePasteMenu,
    applyFillSeries: editingDefaults.applyFillSeries,
    updateFillMenu: editingDefaults.updateFillMenu,
    applyFillDirection: editingDefaults.applyFillDirection,
    applyClearAction: editingDefaults.applyClearAction,
    applyUnderlineAction: formattingDefaults.applyUnderlineAction,
    applyWrapAction: formattingDefaults.applyWrapAction,
    applyMergeAction: formattingDefaults.applyMergeAction,
    applyFreezeAction: viewDefaults.applyFreezeAction,
    updateFreezeMenu: viewDefaults.updateFreezeMenu,
    applyTextOrientationAction: formattingDefaults.applyTextOrientationAction,
    updateTextOrientationMenu: formattingDefaults.updateTextOrientationMenu,
    applyAutoSumFormula: formulaDefaults.applyAutoSumFormula,
    applyFormulaAuditAction: formulaDefaults.applyFormulaAuditAction,
    applyWatchAction: formulaDefaults.applyWatchAction,
    applyCalcOptionAction: formulaDefaults.applyCalcOptionAction,
    applyFindSelectAction: editingDefaults.applyFindSelectAction,
    applyDataValidationAction: dataDefaults.applyDataValidationAction,
    applyConditionalMenuAction: styleDefaults.applyConditionalMenuAction,
    applyUiTheme: pageLayoutDefaults.applyUiTheme,
    applySymbolAction: insertDefaults.applySymbolAction,

    // Dialog / host-glue — host opts in via `overrides`. Defaults are no-op
    // so the click delegator returns true (event consumed) and the open
    // menu still closes, instead of falling through to the legacy fallback.
    applyPivotTableAction: insertDefaults.applyPivotTableAction,
    applyDefinedNameAction: formulaDefaults.applyDefinedNameAction,
    applyLinksAction: insertDefaults.applyLinksAction,
    applyCellInsertAction: cellDefaults.applyCellInsertAction,
    applyCellDeleteAction: cellDefaults.applyCellDeleteAction,
    applyCellFormatAction: cellDefaults.applyCellFormatAction,
    applyPageBreakAction: pageLayoutDefaults.applyPageBreakAction,
    applySheetBackgroundAction: pageLayoutDefaults.applySheetBackgroundAction,
    applyPrintAreaAction: pageLayoutDefaults.applyPrintAreaAction,
    applyArrangeAction: illustrationDefaults.applyArrangeAction,
    applySortMenuAction: editingDefaults.applySortMenuAction,
    applyReviewCommentAction: reviewDefaults.applyReviewCommentAction,
    applyProtectAction: reviewDefaults.applyProtectAction,
    createRecommendedChartFromSelection: insertDefaults.createRecommendedChartFromSelection,
    createChartFromSelection: insertDefaults.createChartFromSelection,
    chartKindFromAction: insertDefaults.chartKindFromAction,
    insertPictureFromRibbon: illustrationDefaults.insertPictureFromRibbon,
    insertShapeFromRibbon: illustrationDefaults.insertShapeFromRibbon,
    insertScreenshotFromRibbon: illustrationDefaults.insertScreenshotFromRibbon,
    applyScriptAction: buildScriptAction(instance),
    applyPdfAction: buildPdfAction(instance),
    createTableFromSelection: styleDefaults.createTableFromSelection,
    openTableStyleFooterAction: styleDefaults.openTableStyleFooterAction,
    updateTableStylesMenu: styleDefaults.updateTableStylesMenu,
    applyPivotTableStyleFromRibbon: styleDefaults.applyPivotTableStyleFromRibbon,
    applyCellStyleFromRibbon: styleDefaults.applyCellStyleFromRibbon,
    updateCellStylesMenu: styleDefaults.updateCellStylesMenu,
    openCellStyleFooterAction: styleDefaults.openCellStyleFooterAction,
    applyCurrencyPreset: formattingDefaults.applyCurrencyPreset,
    openCurrencyFooterAction: formattingDefaults.openCurrencyFooterAction,
    splitTextToColumns: dataDefaults.splitTextToColumns,
    splitTextToColumnsCustom: dataDefaults.splitTextToColumnsCustom,
    applyAddInAction: buildAddInAction(instance),
  };

  // Object form: spread once at construction. Function form: build a live
  // ctx whose every property re-reads the latest override on each access so
  // hosts can swap handlers post-mount without re-creating the api.
  if (typeof opts.overrides !== 'function') {
    return { ...base, ...opts.overrides };
  }
  const getOverrides = opts.overrides;
  const ctx = {} as DynamicDropdownsCtx;
  for (const key of Object.keys(base) as (keyof DynamicDropdownsCtx)[]) {
    Object.defineProperty(ctx, key, {
      enumerable: true,
      get() {
        const override = (getOverrides() as Partial<DynamicDropdownsCtx>)[key];
        return override !== undefined ? override : base[key];
      },
    });
  }
  return ctx;
}
