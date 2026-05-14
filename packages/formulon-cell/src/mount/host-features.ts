import type { History } from '../commands/history.js';
import { setSheetZoom } from '../commands/structure.js';
import { tracePrecedents as tracePrecedentArrows } from '../commands/traces.js';
import { formatCellForEdit } from '../engine/edit-seed.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetEmitter } from '../events.js';
import type { ExtensionHandle, resolveFlags } from '../extensions/index.js';
import type { FormulaRegistry } from '../formula.js';
import type { Strings } from '../i18n/strings.js';
import { attachArgHelper } from '../interact/arg-helper.js';
import { attachAutocomplete } from '../interact/autocomplete.js';
import { attachCommentDialog } from '../interact/comment-dialog.js';
import { attachConditionalDialog } from '../interact/conditional-dialog.js';
import { attachErrorMenu, type ErrorMenuHandle } from '../interact/error-menu.js';
import { attachFormatDialog } from '../interact/format-dialog.js';
import { attachFormatPainter, type FormatPainterHandle } from '../interact/format-painter.js';
import { attachFxDialog, type FxDialogHandle } from '../interact/fx-dialog.js';
import { attachGoToDialog } from '../interact/goto-dialog.js';
import { attachHover } from '../interact/hover.js';
import { attachHyperlinkDialog } from '../interact/hyperlink-dialog.js';
import { attachIterativeDialog } from '../interact/iterative-dialog.js';
import { attachNamedRangeDialog } from '../interact/named-range-dialog.js';
import { attachPageSetupDialog } from '../interact/page-setup-dialog.js';
import { attachPivotTableDialog } from '../interact/pivot-table-dialog.js';
import { attachSessionCharts, type SessionChartsHandle } from '../interact/session-charts.js';
import { attachSlicer, type SlicerHandle } from '../interact/slicer.js';
import { attachStatusBar } from '../interact/status-bar.js';
import { attachViewToolbar, type ViewToolbarHandle } from '../interact/view-toolbar.js';
import { attachWatchPanel } from '../interact/watch-panel.js';
import { attachWheel } from '../interact/wheel.js';
import {
  attachWorkbookObjectsPanel,
  type WorkbookObjectsPanelHandle,
} from '../interact/workbook-objects.js';
import type { GridRenderer } from '../render/grid.js';
import type { SpreadsheetStore } from '../store/store.js';
import { mutators } from '../store/store.js';
import type { ChromeSlot } from './chrome.js';
import type { FormulaBarController } from './formula-bar.js';
import type { SheetTabsController } from './sheet-tabs-controller.js';

type FeatureFlags = ReturnType<typeof resolveFlags>;
export type AutocompleteHandle = ReturnType<typeof attachAutocomplete>;

export const HOST_TOGGLEABLE_IDS = [
  'formatDialog',
  'formatPainter',
  'hoverComment',
  'conditional',
  'iterative',
  'gotoSpecial',
  'fxDialog',
  'namedRanges',
  'pageSetup',
  'pivotTableDialog',
  'hyperlink',
  'commentDialog',
  'statusBar',
  'workbookObjects',
  'watchWindow',
  'slicer',
  'charts',
  'errorIndicators',
  'autocomplete',
  'wheel',
  'shortcuts',
  'formulaBar',
  'viewToolbar',
  'sheetTabs',
] as const;

export const HOST_FEATURE_USES_STRINGS: ReadonlySet<string> = new Set([
  'formatDialog',
  'conditional',
  'iterative',
  'gotoSpecial',
  'fxDialog',
  'namedRanges',
  'pageSetup',
  'pivotTableDialog',
  'hyperlink',
  'commentDialog',
  'statusBar',
  'workbookObjects',
  'viewToolbar',
  'watchWindow',
  'slicer',
  'charts',
  'errorIndicators',
]);

export const WB_TOGGLEABLE_IDS = [
  'shortcuts',
  'clipboard',
  'pasteSpecial',
  'quickAnalysis',
  'contextMenu',
  'findReplace',
  'validation',
] as const;

export interface HostFeatureState {
  formatDialog: ReturnType<typeof attachFormatDialog> | null;
  formatPainter: FormatPainterHandle | null;
  hover: ReturnType<typeof attachHover> | null;
  conditionalDialog: ReturnType<typeof attachConditionalDialog> | null;
  iterativeDialog: ReturnType<typeof attachIterativeDialog> | null;
  goToDialog: ReturnType<typeof attachGoToDialog> | null;
  fxDialog: FxDialogHandle | null;
  fxClickHandler: (() => void) | null;
  namedRangeDialog: ReturnType<typeof attachNamedRangeDialog> | null;
  pageSetupDialog: ReturnType<typeof attachPageSetupDialog> | null;
  pivotTableDialog: ReturnType<typeof attachPivotTableDialog> | null;
  hyperlinkDialog: ReturnType<typeof attachHyperlinkDialog> | null;
  commentDialog: ReturnType<typeof attachCommentDialog> | null;
  statusBar: ReturnType<typeof attachStatusBar> | null;
  watchPanel: ReturnType<typeof attachWatchPanel> | null;
  unsubWatchRecalc: () => void;
  unsubWatchWb: () => void;
  slicer: SlicerHandle | null;
  sessionCharts: SessionChartsHandle | null;
  viewToolbar: ViewToolbarHandle | null;
  workbookObjects: WorkbookObjectsPanelHandle | null;
  unsubSlicerRecalc: () => void;
  unsubSlicerWb: () => void;
  errorMenu: ErrorMenuHandle | null;
  detachWheel: () => void;
  fxAutocomplete: AutocompleteHandle;
  fxArgHelper: ReturnType<typeof attachArgHelper> | null;
  hostShortcutsAttached: boolean;
  canvasErrorClickAttached: boolean;
}

interface HostFeatureControllerInput {
  autocompleteStub: AutocompleteHandle;
  canvas: HTMLCanvasElement;
  emitter: SpreadsheetEmitter;
  featureRegistry: Map<string, ExtensionHandle>;
  flags: () => FeatureFlags;
  formulaRegistry: FormulaRegistry;
  fx: HTMLButtonElement;
  fxInput: HTMLTextAreaElement;
  getFormulaBar: () => FormulaBarController;
  getOnCanvasClick: () => (e: MouseEvent) => void;
  getOnHostKey: () => (e: KeyboardEvent) => void;
  getSheetTabs: () => SheetTabsController | null;
  grid: HTMLElement;
  history: History;
  host: HTMLElement;
  i18nLocale: () => string;
  refreshFeaturesView: () => void;
  renderer: GridRenderer;
  setChromeAttached: (slot: ChromeSlot, attached: boolean) => void;
  state: HostFeatureState;
  statusbar: HTMLElement;
  store: SpreadsheetStore;
  strings: () => Strings;
  viewbar: HTMLElement;
  watchDock: HTMLElement;
  wb: () => WorkbookHandle;
  wrapHandle: (raw: unknown, detach: () => void) => ExtensionHandle;
}

export function createAutocompleteStub(): AutocompleteHandle {
  return {
    isOpen: () => false,
    move: (_n: number) => {},
    acceptHighlighted: () => false,
    close: () => {},
    refresh: () => {},
    setLabels: () => {},
    detach: () => {},
  };
}

export function createHostFeatureState(autocompleteStub: AutocompleteHandle): HostFeatureState {
  return {
    formatDialog: null,
    formatPainter: null,
    hover: null,
    conditionalDialog: null,
    iterativeDialog: null,
    goToDialog: null,
    fxDialog: null,
    fxClickHandler: null,
    namedRangeDialog: null,
    pageSetupDialog: null,
    pivotTableDialog: null,
    hyperlinkDialog: null,
    commentDialog: null,
    statusBar: null,
    watchPanel: null,
    unsubWatchRecalc: (): void => {},
    unsubWatchWb: (): void => {},
    slicer: null,
    sessionCharts: null,
    viewToolbar: null,
    workbookObjects: null,
    unsubSlicerRecalc: (): void => {},
    unsubSlicerWb: (): void => {},
    errorMenu: null,
    detachWheel: (): void => {},
    fxAutocomplete: autocompleteStub,
    fxArgHelper: null,
    hostShortcutsAttached: false,
    canvasErrorClickAttached: false,
  };
}

export function createHostFeatureController(input: HostFeatureControllerInput): {
  attach: (id: string) => void;
  detach: (id: string) => void;
} {
  const s = input.state;
  const attach = (id: string): void => {
    const wb = input.wb();
    const strings = input.strings();
    switch (id) {
      case 'formatDialog':
        if (s.formatDialog) return;
        s.formatDialog = attachFormatDialog({
          host: input.host,
          store: input.store,
          strings,
          history: input.history,
          getWb: input.wb,
          getLocale: input.i18nLocale,
        });
        input.featureRegistry.set(
          'formatDialog',
          input.wrapHandle(s.formatDialog, () => s.formatDialog?.detach()),
        );
        break;
      case 'formatPainter':
        if (s.formatPainter) return;
        s.formatPainter = attachFormatPainter({
          host: input.host,
          store: input.store,
          history: input.history,
        });
        input.featureRegistry.set(
          'formatPainter',
          input.wrapHandle(s.formatPainter, () => s.formatPainter?.detach()),
        );
        break;
      case 'hoverComment':
        if (s.hover) return;
        s.hover = attachHover({ grid: input.grid, store: input.store });
        input.featureRegistry.set(
          'hoverComment',
          input.wrapHandle(s.hover, () => s.hover?.detach()),
        );
        break;
      case 'conditional':
        if (s.conditionalDialog) return;
        s.conditionalDialog = attachConditionalDialog({
          host: input.host,
          store: input.store,
          strings,
        });
        input.featureRegistry.set(
          'conditional',
          input.wrapHandle(s.conditionalDialog, () => s.conditionalDialog?.detach()),
        );
        break;
      case 'iterative':
        if (s.iterativeDialog) return;
        s.iterativeDialog = attachIterativeDialog({ host: input.host, getWb: input.wb, strings });
        input.featureRegistry.set(
          'iterative',
          input.wrapHandle(s.iterativeDialog, () => s.iterativeDialog?.detach()),
        );
        break;
      case 'gotoSpecial':
        if (s.goToDialog) return;
        s.goToDialog = attachGoToDialog({
          host: input.host,
          store: input.store,
          strings,
          getWb: input.wb,
        });
        input.featureRegistry.set(
          'gotoSpecial',
          input.wrapHandle(s.goToDialog, () => s.goToDialog?.detach()),
        );
        break;
      case 'fxDialog':
        if (s.fxDialog) return;
        s.fxDialog = attachFxDialog({
          host: input.host,
          store: input.store,
          strings,
          onInsert: (formula) => {
            input.fxInput.value = formula;
            input.fxInput.focus();
            input.getFormulaBar().commitFx('none');
          },
        });
        s.fxClickHandler = (): void => s.fxDialog?.open();
        input.fx.addEventListener('click', s.fxClickHandler);
        input.fx.disabled = false;
        input.fx.style.cursor = '';
        input.featureRegistry.set(
          'fxDialog',
          input.wrapHandle(s.fxDialog, () => s.fxDialog?.detach()),
        );
        break;
      case 'namedRanges':
        if (s.namedRangeDialog) return;
        s.namedRangeDialog = attachNamedRangeDialog({ host: input.host, wb, strings });
        input.featureRegistry.set(
          'namedRanges',
          input.wrapHandle(s.namedRangeDialog, () => s.namedRangeDialog?.detach()),
        );
        break;
      case 'pageSetup':
        if (s.pageSetupDialog) return;
        s.pageSetupDialog = attachPageSetupDialog({
          host: input.host,
          store: input.store,
          strings,
          history: input.history,
        });
        input.featureRegistry.set(
          'pageSetup',
          input.wrapHandle(s.pageSetupDialog, () => s.pageSetupDialog?.detach()),
        );
        break;
      case 'pivotTableDialog':
        if (s.pivotTableDialog) return;
        s.pivotTableDialog = attachPivotTableDialog({
          host: input.host,
          store: input.store,
          wb,
          strings,
          onAfterCreate: () => {
            mutators.replaceCells(input.store, wb.cells(input.store.getState().data.sheetIndex));
          },
          invalidate: () => input.renderer.invalidate(),
        });
        input.featureRegistry.set(
          'pivotTableDialog',
          input.wrapHandle(s.pivotTableDialog, () => s.pivotTableDialog?.detach()),
        );
        break;
      case 'hyperlink':
        if (s.hyperlinkDialog) return;
        s.hyperlinkDialog = attachHyperlinkDialog({
          host: input.host,
          store: input.store,
          strings,
          history: input.history,
          getWb: input.wb,
        });
        input.featureRegistry.set(
          'hyperlink',
          input.wrapHandle(s.hyperlinkDialog, () => s.hyperlinkDialog?.detach()),
        );
        break;
      case 'commentDialog':
        if (s.commentDialog) return;
        s.commentDialog = attachCommentDialog({
          host: input.host,
          store: input.store,
          strings,
          history: input.history,
          getWb: input.wb,
        });
        input.featureRegistry.set(
          'commentDialog',
          input.wrapHandle(s.commentDialog, () => s.commentDialog?.detach()),
        );
        break;
      case 'statusBar':
        if (s.statusBar) return;
        input.setChromeAttached('statusbar', true);
        s.statusBar = attachStatusBar({
          statusbar: input.statusbar,
          store: input.store,
          strings,
          getEngineLabel: () => (wb.isStub ? 'stub' : `formulon ${wb.version}`),
          getCalcMode: () => wb.calcMode(),
          onCycleCalcMode: () => {
            const cur = wb.calcMode();
            if (cur === null) return;
            const next = ((cur + 1) % 3) as 0 | 1 | 2;
            wb.setCalcMode(next);
            s.statusBar?.refresh();
          },
          onRecalc: () => {
            wb.recalc();
            mutators.replaceCells(input.store, wb.cells(input.store.getState().data.sheetIndex));
            input.renderer.invalidate();
          },
          onZoomChange: (zoom) => {
            setSheetZoom(input.store, zoom, wb);
            input.renderer.invalidate();
          },
        });
        input.featureRegistry.set(
          'statusBar',
          input.wrapHandle(s.statusBar, () => s.statusBar?.detach()),
        );
        break;
      case 'workbookObjects':
        if (s.workbookObjects) return;
        s.workbookObjects = attachWorkbookObjectsPanel({ host: input.host, wb, strings });
        input.featureRegistry.set(
          'workbookObjects',
          input.wrapHandle(s.workbookObjects, () => s.workbookObjects?.detach()),
        );
        break;
      case 'viewToolbar':
        if (s.viewToolbar) return;
        input.setChromeAttached('viewbar', true);
        s.viewToolbar = attachViewToolbar({
          toolbar: input.viewbar,
          store: input.store,
          wb,
          history: input.history,
          strings,
          onOpenObjects: input.flags().workbookObjects
            ? () => s.workbookObjects?.open()
            : undefined,
          onChanged: () => input.renderer.invalidate(),
        });
        input.featureRegistry.set(
          'viewToolbar',
          input.wrapHandle(s.viewToolbar, () => s.viewToolbar?.detach()),
        );
        break;
      case 'watchWindow':
        if (s.watchPanel) return;
        input.setChromeAttached('watchDock', true);
        s.watchPanel = attachWatchPanel({
          host: input.watchDock,
          store: input.store,
          getWb: input.wb,
          strings,
        });
        input.featureRegistry.set(
          'watchWindow',
          input.wrapHandle(s.watchPanel, () => s.watchPanel?.detach()),
        );
        s.unsubWatchRecalc = input.emitter.on('recalc', () => s.watchPanel?.refresh());
        s.unsubWatchWb = input.emitter.on('workbookChange', () => s.watchPanel?.refresh());
        break;
      case 'slicer':
        if (s.slicer) return;
        s.slicer = attachSlicer({
          host: input.host,
          store: input.store,
          getWb: input.wb,
          history: input.history,
          strings,
        });
        input.featureRegistry.set(
          'slicer',
          input.wrapHandle(s.slicer, () => s.slicer?.detach()),
        );
        s.unsubSlicerRecalc = input.emitter.on('recalc', () => s.slicer?.refresh());
        s.unsubSlicerWb = input.emitter.on('workbookChange', () => s.slicer?.refresh());
        break;
      case 'charts':
        if (s.sessionCharts) return;
        s.sessionCharts = attachSessionCharts({
          host: input.host,
          store: input.store,
          labels: strings.sessionCharts,
        });
        input.featureRegistry.set(
          'charts',
          input.wrapHandle(s.sessionCharts, () => s.sessionCharts?.detach()),
        );
        break;
      case 'errorIndicators':
        if (s.errorMenu) return;
        s.errorMenu = attachErrorMenu({
          host: input.host,
          store: input.store,
          getWb: input.wb,
          strings,
          onEditCell: (addr) => {
            mutators.setActive(input.store, addr);
            const cell = input.store
              .getState()
              .data.cells.get(`${addr.sheet}:${addr.row}:${addr.col}`);
            input.fxInput.value = formatCellForEdit(cell, wb, addr);
            input.fxInput.focus();
            input.fxInput.setSelectionRange(input.fxInput.value.length, input.fxInput.value.length);
          },
          onTraceError: (addr) => {
            mutators.setActive(input.store, addr);
            tracePrecedentArrows(input.store, wb, addr);
            input.renderer.invalidate();
          },
        });
        if (!s.canvasErrorClickAttached) {
          input.canvas.addEventListener('click', input.getOnCanvasClick());
          s.canvasErrorClickAttached = true;
        }
        input.featureRegistry.set(
          'errorIndicators',
          input.wrapHandle(s.errorMenu, () => s.errorMenu?.detach()),
        );
        break;
      case 'autocomplete':
        if (s.fxAutocomplete !== input.autocompleteStub) return;
        s.fxAutocomplete = attachAutocomplete({
          input: input.fxInput,
          onAfterInsert: () => input.getFormulaBar().syncFxRefs(),
          getTables: () => wb.getTables(),
          getCustomFunctions: () => input.formulaRegistry.list(),
          getFunctionNames: () => wb.functionNames(),
          labels: strings.autocomplete,
        });
        s.fxArgHelper = attachArgHelper({ input: input.fxInput, labels: strings.argHelper });
        break;
      case 'wheel':
        s.detachWheel();
        s.detachWheel = attachWheel({ grid: input.grid, store: input.store, wb });
        break;
      case 'shortcuts':
        if (s.hostShortcutsAttached) return;
        input.host.addEventListener('keydown', input.getOnHostKey());
        s.hostShortcutsAttached = true;
        break;
      case 'formulaBar':
        input.setChromeAttached('formulabar', true);
        break;
      case 'sheetTabs':
        input.setChromeAttached('sheetbar', true);
        input.getSheetTabs()?.update();
        break;
    }
    input.refreshFeaturesView();
  };

  const detach = (id: string): void => {
    switch (id) {
      case 'formatDialog':
        s.formatDialog?.detach();
        s.formatDialog = null;
        input.featureRegistry.delete('formatDialog');
        break;
      case 'formatPainter':
        s.formatPainter?.detach();
        s.formatPainter = null;
        input.featureRegistry.delete('formatPainter');
        break;
      case 'hoverComment':
        s.hover?.detach();
        s.hover = null;
        input.featureRegistry.delete('hoverComment');
        break;
      case 'conditional':
        s.conditionalDialog?.detach();
        s.conditionalDialog = null;
        input.featureRegistry.delete('conditional');
        break;
      case 'iterative':
        s.iterativeDialog?.detach();
        s.iterativeDialog = null;
        input.featureRegistry.delete('iterative');
        break;
      case 'gotoSpecial':
        s.goToDialog?.detach();
        s.goToDialog = null;
        input.featureRegistry.delete('gotoSpecial');
        break;
      case 'fxDialog':
        if (s.fxClickHandler) input.fx.removeEventListener('click', s.fxClickHandler);
        s.fxClickHandler = null;
        s.fxDialog?.detach();
        s.fxDialog = null;
        input.fx.disabled = true;
        input.fx.style.cursor = 'default';
        input.featureRegistry.delete('fxDialog');
        break;
      case 'namedRanges':
        s.namedRangeDialog?.detach();
        s.namedRangeDialog = null;
        input.featureRegistry.delete('namedRanges');
        break;
      case 'pageSetup':
        s.pageSetupDialog?.detach();
        s.pageSetupDialog = null;
        input.featureRegistry.delete('pageSetup');
        break;
      case 'pivotTableDialog':
        s.pivotTableDialog?.detach();
        s.pivotTableDialog = null;
        input.featureRegistry.delete('pivotTableDialog');
        break;
      case 'hyperlink':
        s.hyperlinkDialog?.detach();
        s.hyperlinkDialog = null;
        input.featureRegistry.delete('hyperlink');
        break;
      case 'commentDialog':
        s.commentDialog?.detach();
        s.commentDialog = null;
        input.featureRegistry.delete('commentDialog');
        break;
      case 'statusBar':
        s.statusBar?.detach();
        s.statusBar = null;
        input.featureRegistry.delete('statusBar');
        input.setChromeAttached('statusbar', false);
        break;
      case 'workbookObjects':
        s.workbookObjects?.detach();
        s.workbookObjects = null;
        input.featureRegistry.delete('workbookObjects');
        break;
      case 'viewToolbar':
        s.viewToolbar?.detach();
        s.viewToolbar = null;
        input.featureRegistry.delete('viewToolbar');
        input.setChromeAttached('viewbar', false);
        break;
      case 'watchWindow':
        s.unsubWatchRecalc();
        s.unsubWatchWb();
        s.unsubWatchRecalc = (): void => {};
        s.unsubWatchWb = (): void => {};
        s.watchPanel?.detach();
        s.watchPanel = null;
        input.featureRegistry.delete('watchWindow');
        input.setChromeAttached('watchDock', false);
        break;
      case 'slicer':
        s.unsubSlicerRecalc();
        s.unsubSlicerWb();
        s.unsubSlicerRecalc = (): void => {};
        s.unsubSlicerWb = (): void => {};
        s.slicer?.detach();
        s.slicer = null;
        input.featureRegistry.delete('slicer');
        break;
      case 'charts':
        s.sessionCharts?.detach();
        s.sessionCharts = null;
        input.featureRegistry.delete('charts');
        break;
      case 'errorIndicators':
        if (s.canvasErrorClickAttached)
          input.canvas.removeEventListener('click', input.getOnCanvasClick());
        s.canvasErrorClickAttached = false;
        s.errorMenu?.detach();
        s.errorMenu = null;
        input.featureRegistry.delete('errorIndicators');
        break;
      case 'autocomplete':
        s.fxAutocomplete.detach();
        s.fxAutocomplete = input.autocompleteStub;
        s.fxArgHelper?.detach();
        s.fxArgHelper = null;
        break;
      case 'wheel':
        s.detachWheel();
        s.detachWheel = (): void => {};
        break;
      case 'shortcuts':
        if (s.hostShortcutsAttached)
          input.host.removeEventListener('keydown', input.getOnHostKey());
        s.hostShortcutsAttached = false;
        break;
      case 'formulaBar':
        input.setChromeAttached('formulabar', false);
        break;
      case 'sheetTabs':
        input.setChromeAttached('sheetbar', false);
        break;
    }
    input.refreshFeaturesView();
  };

  return { attach, detach };
}
