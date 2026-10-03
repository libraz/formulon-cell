import { CellRegistry } from './cells.js';
import {
  insertCopiedBand,
  insertCopiedCellsFromTSV,
} from './commands/clipboard/insert-copied-cells.js';
import { History } from './commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from './commands/interaction-controller.js';
import { printSheet } from './commands/print.js';
import {
  isSheetProtected,
  setProtectedSheet,
  toggleProtectedSheet,
} from './commands/protection.js';
import {
  clearTraceArrows,
  traceDependents as traceDependentArrows,
  tracePrecedents as tracePrecedentArrows,
} from './commands/traces.js';
import {
  type SyncedConditionalRuleMap,
  syncTrackedConditionalRulesToEngine,
} from './engine/cf-writeback.js';
import { findPivotTableAtCell } from './engine/passthrough-sync.js';
import { WorkbookHandle } from './engine/workbook-handle.js';
import { SpreadsheetEmitter } from './events.js';
import {
  dedupeById,
  type Extension,
  type ExtensionContext,
  type ExtensionHandle,
  flattenExtensions,
  resolveSpreadsheetUiOptions,
  sortByPriority,
} from './extensions/index.js';
import { FormulaRegistry } from './formula.js';
import { createI18nController } from './i18n/controller.js';
import type { Strings } from './i18n/strings.js';
import { readClipboard } from './interact/context-menu-clipboard.js';
import type { FxDialogOpenOptions } from './interact/fx-dialog.js';
import { openInsertCopiedCellsDialog } from './interact/insert-copied-cells-dialog.js';
import { deactivateMacInk, disposeMacInk } from './interact/mac-ink.js';
import { attachNavigationPolicy, navigationBoundsFor } from './interact/navigation-policy.js';
import {
  disposeOverlayPortal,
  ensureOverlayPortal,
  setOverlayOptions,
  syncOverlayPortalTheme,
} from './interact/overlay-portal.js';
import { attachAlwaysOnDialogs } from './mount/always-on-dialogs.js';
import { createMountChrome } from './mount/chrome.js';
import { attachChromeSync, type ChromeSyncController } from './mount/chrome-sync.js';
import {
  attachEngineBinding,
  type EngineBinding,
  WB_REGISTRY_IDS,
} from './mount/engine-binding.js';
import { createErrorIndicatorClickHandler } from './mount/error-indicator-click.js';
import { resolveMountFlags } from './mount/feature-flags.js';
import { attachFormulaBarController } from './mount/formula-bar.js';
import { attachFormulaDraftMirror } from './mount/formula-draft-mirror.js';
import {
  applyPlatformLayoutDefaults,
  prepareMountHost,
  releaseMountHost,
  renderMountError,
} from './mount/host.js';
import {
  createAutocompleteStub,
  createHostFeatureController,
  createHostFeatureState,
  HOST_FEATURE_USES_STRINGS,
  HOST_TOGGLEABLE_IDS,
  WB_TOGGLEABLE_IDS,
} from './mount/host-features.js';
import { createHostShortcutHandler } from './mount/host-shortcuts.js';
import {
  dispatchWorkbookObjectSummaries,
  hydrateActiveSheetFromEngine,
  hydrateWorkbookMetadataFromEngine,
  validateViewportAgainstWorkbook,
} from './mount/hydration.js';
import { attachPivotFieldListFollow } from './mount/pivot-field-list-follow.js';
import { createPrinterProfiles } from './mount/printer-profiles.js';
import {
  attachSheetTabsController,
  type SheetTabsController,
} from './mount/sheet-tabs-controller.js';
import {
  type MountToolbarOptions,
  mountToolbar,
  type ToolbarInstance,
  type ToolbarInstanceRef,
} from './mount/toolbar.js';
import type { MountOptions, ScreenClipResult, SpreadsheetInstance } from './mount/types.js';
import { GridRenderer } from './render/grid.js';
import { isWholeColumnRange, isWholeRowRange } from './store/selection-geometry.js';
import { createSpreadsheetStore, mutators } from './store/store.js';
import { resolveTheme } from './theme/resolve.js';
import { disposeMacRibbonActions } from './toolbar/ribbon/mac/actions.js';

export type {
  MountToolbarOptions,
  RibbonDisplayMode,
  ToolbarInstance,
  ToolbarInstanceRef,
} from './mount/toolbar.js';
export {
  openHostMenuFirstDropdown,
  RIBBON_HOST_MENU_FIRST_COMMANDS,
} from './mount/toolbar.js';
export type {
  FunctionArgumentHelp,
  FunctionArgumentHelpProvider,
  FunctionCategory,
  FxDialogOpenOptions,
  MountOptions,
  ScreenClipCapture,
  ScreenClipCaptureResult,
  ScreenClipResult,
  SpreadsheetInstance,
} from './mount/types.js';

function normalizeScreenClipResult(result: string | ScreenClipResult | null | undefined) {
  if (!result) return null;
  if (typeof result === 'string') return result ? { src: result } : null;
  return result.src ? result : null;
}

/**
 * Mount a spreadsheet onto a DOM host. Returns an instance with imperative
 * controls. The host element is taken over — its existing children are
 * cleared. Idempotent dispose.
 */
export const Spreadsheet = {
  mountToolbar(
    host: HTMLElement,
    instance: ToolbarInstanceRef,
    opts: MountToolbarOptions,
  ): ToolbarInstance {
    return mountToolbar(host, instance, opts);
  },
  async mount(host: HTMLElement, opts: MountOptions = {}): Promise<SpreadsheetInstance> {
    if (!host) throw new Error('Spreadsheet.mount: host element required');

    // Reactive strings — extensions read `strings` (a `let` re-assigned by
    // the i18n subscription) so any future setStrings hook lands on a fresh
    // snapshot. v0.1 built-ins still snapshot at attach time; v0.2 will
    // wire setStrings hooks for live label updates.
    const i18n = createI18nController({ locale: opts.locale, overlay: opts.strings });
    let strings: Strings = i18n.strings;
    const printers = createPrinterProfiles({
      profiles: opts.printerProfiles,
      profileId: opts.printerProfileId,
      refreshHook: opts.refreshPrinterProfiles,
    });
    const captureScreenClipHook = opts.captureScreenClip;
    const getFunctionArgumentHelp = opts.getFunctionArgumentHelp;
    let uploadStatus = opts.uploadStatus ?? null;
    let macroRecording = opts.macroRecording ?? null;
    let ui = resolveSpreadsheetUiOptions(opts.ui);
    let requestedFeatures = { ...ui.features, ...opts.features };
    let contextMenuOptions = opts.contextMenu;
    const initialTheme = opts.theme ?? ui.theme;
    let flags = resolveMountFlags(requestedFeatures, opts.policy !== undefined, contextMenuOptions);
    const emitter = new SpreadsheetEmitter();
    const formulaRegistry = new FormulaRegistry();
    if (opts.functions) {
      for (const f of opts.functions) {
        formulaRegistry.register(f.name, f.impl, f.meta ?? {});
      }
    }

    const instanceId = prepareMountHost(host, strings, initialTheme, ui.platform);
    ensureOverlayPortal(host, opts.overlays);
    host.dataset.fcEngineState = 'loading';

    let sheetTabsController: SheetTabsController | null = null;

    // Track ownership before seeding — only owned (default-created) workbooks
    // should be touched by `seed`. Pre-loaded workbooks are the consumer's
    // data and must not be overwritten by the demo helper.
    let ownsWb = !opts.workbook;
    let wb: WorkbookHandle;
    try {
      wb = opts.workbook ?? (await WorkbookHandle.createDefault({ locale: i18n.locale }));
      if (opts.seed && ownsWb) opts.seed(wb);
      if (opts.viewport) {
        try {
          validateViewportAgainstWorkbook(opts.viewport, wb, opts.viewport.range?.sheet);
        } catch (err) {
          if (ownsWb) wb.dispose();
          throw err;
        }
      }
    } catch (err) {
      host.dataset.fcEngineState = 'error';
      disposeOverlayPortal(host);
      try {
        opts.onError?.(err);
      } catch (hookErr) {
        console.error('formulon-cell: mount error handler threw', hookErr);
      }
      if (opts.renderError !== false) renderMountError(host, err, strings.mountError);
      throw err;
    }

    const {
      formulabar,
      tag,
      fx,
      fxCancel,
      fxAccept,
      fxInput,
      viewbar,
      taskpaneDock,
      grid,
      canvas,
      a11y,
      a11yLive,
      statusbar,
      firstSheet,
      lastSheet,
      sheetTabs,
      addSheetBtn,
      sheetMenu,
      watchDock,
      refreshFormulaBarLabels,
      setChromeAttached,
    } = createMountChrome({
      host,
      getStrings: () => strings,
      flags,
      onSheetTabContextMenu: (idx, tab, x, y) => {
        sheetTabsController?.switchSheet(idx);
        sheetTabsController?.showMenu(idx, tab, x, y);
      },
    });

    const store = createSpreadsheetStore();
    if (opts.viewport?.range) mutators.setSheetIndex(store, opts.viewport.range.sheet);
    mutators.setTheme(store, initialTheme);
    applyPlatformLayoutDefaults(store, ui.platform);

    // Upload status / macro recording start off in the status-bar chooser
    // because most hosts never drive them. The first value a host reports
    // turns the item on; after that the chooser owns its visibility.
    const declaredStatusIndicators = new Set<'uploadStatus' | 'macroRecording'>();
    const declareStatusIndicator = (key: 'uploadStatus' | 'macroRecording'): void => {
      if (declaredStatusIndicators.has(key)) return;
      declaredStatusIndicators.add(key);
      mutators.setStatusOption(store, key, true);
    };
    if (uploadStatus !== null) declareStatusIndicator('uploadStatus');
    if (macroRecording !== null) declareStatusIndicator('macroRecording');

    // Unified undo/redo. Attach BEFORE seed-cell hydration so the seed itself
    // doesn't pollute the stack — but seed runs above on the wb. Clear the
    // stack after attach to drop any pre-attach entries (none expected, but
    // cheap insurance).
    const history = new History();
    wb.attachHistory(history);
    history.clear();

    hydrateActiveSheetFromEngine(wb, store);
    hydrateWorkbookMetadataFromEngine(wb, store);
    wb.attachStore(store);
    const navigation = attachNavigationPolicy(store, () => wb, opts.viewport);
    const commands = new InteractionController({
      store,
      getWb: () => wb,
      history,
      getBounds: () => navigationBoundsFor(store),
      onChanged: (result) => emitter.emit('changeBatch', result),
    });
    commands.setPolicy(opts.policy);
    const unregisterCommands = registerInteractionController(store, commands);
    dispatchPassthroughSummary();

    function dispatchPassthroughSummary(): void {
      // Surface preserved OOXML objects (charts/drawings/pivot parts) and
      // Spreadsheet Tables as host events so chrome (status bar, toast) can show a
      // read-only/editing-limited badge. Pivot layouts are rendered when the
      // engine exposes projection, but the object definition is still not
      // authorable from the UI.
      dispatchWorkbookObjectSummaries(host, wb);
    }

    function hydrateActiveSheet(): void {
      hydrateActiveSheetFromEngine(wb, store);
    }

    const cellRegistry = new CellRegistry();
    const renderer = new GridRenderer({
      host: grid,
      canvas,
      getState: () => store.getState(),
      getTheme: () => resolveTheme(host),
      onViewportSize: (rowCount, colCount, widthPx) =>
        mutators.setViewportSize(store, rowCount, colCount, widthPx),
      getWb: () => wb,
      getLocale: () => i18n.locale,
      getStrings: () => strings,
      getDisplay: (addr, value, formula, format) =>
        cellRegistry.resolveDisplay({ addr, value, formula, format }),
    });
    const unsubCellRegistry = cellRegistry.subscribe(() => renderer.invalidate());
    renderer.resize();
    sheetTabsController = attachSheetTabsController({
      addSheetBtn,
      firstSheet,
      getStrings: () => strings,
      getWb: () => wb,
      history,
      hydrateActiveSheet,
      invalidate: () => renderer.invalidate(),
      lastSheet,
      refreshStatusBar: () => featureState.statusBar?.refresh(),
      sheetMenu,
      sheetTabs,
      store,
    });
    sheetTabsController.update();

    // Always-on host features — not toggleable via `MountOptions.features`.
    const alwaysOnDialogs = attachAlwaysOnDialogs({
      host,
      store,
      history,
      getWb: () => wb,
      getStrings: () => strings,
      getLocale: () => i18n.locale,
      onConditionalRulesChanged: () => {
        syncSessionConditionalRules();
        renderer.invalidate();
      },
      openConditionalDialog: (options) => featureState.conditionalDialog?.open(options),
    });
    // Filter dropdown — opens when the pointer dispatches `fc:openfilter`
    // from a clicked column-filter chevron; no public toggle.
    interface OpenFilterDetail {
      range: import('./engine/types.js').Range;
      col: number;
      anchor: { x: number; y: number; h: number; clientX: number; clientY: number };
    }
    const onOpenFilter = (e: Event): void => {
      if (commands.policy !== undefined) return;
      const detail = (e as CustomEvent<OpenFilterDetail>).detail;
      if (!detail) return;
      // The dropdown is positioned with `position: fixed`, so it expects
      // viewport-relative coords. The pointer payload's `x/y` are host-relative;
      // use `clientX/clientY` instead. `- 4` matches the chevron offset.
      alwaysOnDialogs.openFilter(detail.range, detail.col, {
        x: detail.anchor.clientX,
        y: detail.anchor.clientY - 4,
        h: detail.anchor.h,
      });
    };
    host.addEventListener('fc:openfilter', onOpenFilter);

    const featureRegistry = new Map<string, ExtensionHandle>();
    const wrapHandle = (raw: unknown, detach: () => void): ExtensionHandle => {
      const h = (
        raw && typeof raw === 'object' ? (raw as Record<string, unknown>) : {}
      ) as ExtensionHandle;
      h.dispose = detach;
      return h;
    };

    const autocompleteStub = createAutocompleteStub();
    const featureState = createHostFeatureState(autocompleteStub);
    const captureScreenClip = async (): Promise<ScreenClipResult | null> =>
      normalizeScreenClipResult(await captureScreenClipHook?.());
    const pivotFieldListFollow = attachPivotFieldListFollow({
      store,
      getWb: () => wb,
      getWorkbookObjects: () => featureState.workbookObjects,
    });

    const syncBindingFeatures = (current: EngineBinding): void => {
      for (const id of WB_REGISTRY_IDS) featureRegistry.delete(id);
      if (current.clipboardH) {
        featureRegistry.set(
          'clipboard',
          wrapHandle(current.clipboardH, () => current.clipboardH?.detach()),
        );
      }
      if (current.pasteSpecialDialog) {
        featureRegistry.set(
          'pasteSpecial',
          wrapHandle(current.pasteSpecialDialog, () => current.pasteSpecialDialog?.detach()),
        );
      }
      if (current.quickAnalysis) {
        featureRegistry.set(
          'quickAnalysis',
          wrapHandle(current.quickAnalysis, () => current.quickAnalysis?.detach()),
        );
      }
      if (current.contextMenu) featureRegistry.set('contextMenu', current.contextMenu);
      if (current.findReplace) {
        featureRegistry.set(
          'findReplace',
          wrapHandle(current.findReplace, () => current.findReplace?.detach()),
        );
      }
      if (current.validation) {
        featureRegistry.set(
          'validation',
          wrapHandle(current.validation, () => current.validation?.detach()),
        );
      }
    };

    let chromeSync: ChromeSyncController | null = null;
    const updateChrome = (): void => chromeSync?.updateChrome();

    const bindEngine = (currentWb: WorkbookHandle): EngineBinding =>
      attachEngineBinding({
        contextMenuOptions,
        emitter,
        flags,
        formulaRegistry,
        getCommentDialog: () => featureState.commentDialog,
        getFormatDialog: () => featureState.formatDialog,
        getFormatPainter: () => featureState.formatPainter,
        getFormulaBarEditor: () => {
          if (host.dataset.fcPlatform !== 'mac') return null;
          if (featureState.macFormulaPalette?.isOpen()) {
            return featureState.macFormulaPalette.rangeInsertTarget();
          }
          return formulaBar;
        },
        getGoToDialog: () => featureState.goToDialog,
        getHyperlinkDialog: () => featureState.hyperlinkDialog,
        getLocale: () => i18n.locale,
        getNamedRangeDialog: () => featureState.namedRangeDialog,
        getPivotTableDialog: () => featureState.pivotTableDialog,
        getSessionCharts: () => featureState.sessionCharts,
        getSheetTabs: () => sheetTabsController,
        grid,
        history,
        host,
        renderer,
        store,
        strings,
        tag,
        updateChrome,
        wb: currentWb,
      });

    let binding = bindEngine(wb);
    syncBindingFeatures(binding);

    const onHostKey = createHostShortcutHandler({
      addSheet: () => sheetTabsController?.addSheet(),
      findReplace: () => binding.findReplace,
      formatDialog: () => featureState.formatDialog,
      formatPainter: () => featureState.formatPainter,
      getClipboardSnapshot: () => binding.clipboardH?.getSnapshot() ?? null,
      goToDialog: () => featureState.goToDialog,
      history,
      host,
      hostTag: tag,
      hyperlinkDialog: () => featureState.hyperlinkDialog,
      invalidate: () => renderer.invalidate(),
      locale: i18n.locale,
      namedRangeDialog: () => featureState.namedRangeDialog,
      pasteSpecialDialog: () => binding.pasteSpecialDialog,
      quickAnalysis: () => binding.quickAnalysis,
      store,
      strings: () => strings,
      wb: () => wb,
    });

    const onCanvasClick = createErrorIndicatorClickHandler({
      canvas,
      getErrorMenu: () => featureState.errorMenu,
    });

    const formulaBar = attachFormulaBarController({
      cancelBindingEditor: () => {
        if (binding.editor.isActive()) binding.editor.cancel();
      },
      formulabar,
      fxAccept,
      fxCancel,
      fxInput,
      getArgHelper: () => featureState.fxArgHelper,
      getAutocomplete: () => featureState.fxAutocomplete,
      getStrings: () => i18n.strings,
      host,
      onValidation: (outcome) => binding.validationAlert?.show(outcome),
      store,
      updateChrome,
      wb: () => wb,
    });

    chromeSync = attachChromeSync({
      a11y,
      a11yLive,
      emitter,
      fxInput,
      getFormulaEditing: () => formulaBar.isEditing(),
      getSheetTabs: () => sheetTabsController,
      getStrings: () => strings,
      getWb: () => wb,
      grid,
      host,
      invalidate: () => renderer.invalidate(),
      store,
      tag,
    });

    const formulaDraftMirror = attachFormulaDraftMirror({ grid, store });

    // Resize observer — we follow the host, not the window.
    const ro = new ResizeObserver(() => {
      renderer.resize();
      formulaDraftMirror.refresh();
    });
    ro.observe(grid);

    let disposed = false;

    // User extensions — additive on top of built-ins. Run after built-ins
    // and the engine binding so they can read other features via
    // `ctx.resolve()`.
    const userHandles = new Map<string, ExtensionHandle>();
    const refreshCells = (): void => {
      mutators.replaceCells(store, wb.cells(store.getState().data.sheetIndex));
    };
    const wbListeners = new Set<(next: WorkbookHandle) => void>();
    const ctx: ExtensionContext = {
      host,
      formulabar,
      viewbar,
      grid,
      statusbar,
      canvas,
      a11y: a11yLive,
      store,
      history,
      i18n,
      getWb: () => wb,
      refreshCells,
      invalidate: () => renderer.invalidate(),
      resolve: <T extends ExtensionHandle = ExtensionHandle>(id: string): T | undefined =>
        (featureRegistry.get(id) ?? userHandles.get(id)) as T | undefined,
      onWorkbookChange: (fn) => {
        wbListeners.add(fn);
        return () => {
          wbListeners.delete(fn);
        };
      },
    };

    const mountExtension = (ext: Extension): void => {
      if (userHandles.has(ext.id) || featureRegistry.has(ext.id)) {
        // last-wins via remove + re-add; users explicitly opt in
        userHandles.get(ext.id)?.dispose();
        userHandles.delete(ext.id);
      }
      const handle = ext.setup(ctx);
      if (handle) userHandles.set(ext.id, handle);
    };
    // Combined view exposed on `instance.features` — built-ins + user.
    const featuresView: Record<string, ExtensionHandle | undefined> = {};
    const refreshFeaturesView = (): void => {
      for (const k of Object.keys(featuresView)) delete featuresView[k];
      for (const [k, v] of featureRegistry) featuresView[k] = v;
      for (const [k, v] of userHandles) featuresView[k] = v;
    };
    let syncedSessionCfRules: SyncedConditionalRuleMap = new Map();
    const syncSessionConditionalRules = (): void => {
      const rules = store.getState().conditional.rules;
      for (let sheet = 0; sheet < wb.sheetCount; sheet += 1) {
        syncTrackedConditionalRulesToEngine(wb, rules, sheet, { tracked: syncedSessionCfRules });
      }
    };
    const hostFeatures = createHostFeatureController({
      autocompleteStub,
      canvas,
      emitter,
      featureRegistry,
      flags: () => flags,
      formulaRegistry,
      fx,
      fxInput,
      getFormulaBar: () => formulaBar,
      getInlineEditor: () => binding.editor,
      getOnCanvasClick: () => onCanvasClick,
      getOnHostKey: () => onHostKey,
      getPrintableBoundsForPageSetup: (setup, _sheet, _previous, selectedPrinterProfileId) =>
        printers.printableBounds(setup, selectedPrinterProfileId),
      getPrinterProfiles: printers.getProfiles,
      getPrinterProfileId: printers.getProfileId,
      setPrinterProfileId: printers.setProfileId,
      refreshPrinterProfiles: printers.refresh,
      getFunctionArgumentHelp,
      getUploadStatus: () => uploadStatus,
      getMacroRecording: () => macroRecording,
      onConditionalRulesChanged: syncSessionConditionalRules,
      getSheetTabs: () => sheetTabsController,
      grid,
      taskpaneDock,
      history,
      host,
      i18nLocale: () => i18n.locale,
      isMacPlatform: () => host.dataset.fcPlatform === 'mac',
      projectFormulaDraftMirror: formulaDraftMirror.project,
      refreshFeaturesView,
      renderer,
      setChromeAttached,
      state: featureState,
      statusbar,
      store,
      strings: () => strings,
      viewbar,
      watchDock,
      wb: () => wb,
      wrapHandle,
    });
    const attachHostFeature = hostFeatures.attach;
    const detachHostFeature = hostFeatures.detach;
    const closeMacPalette = (): void => {
      featureState.macFormulaPalette?.close();
    };
    const ensureWatchWindow = (): void => {
      if (!featureState.watchPanel) attachHostFeature('watchWindow');
    };

    // Initial host-feature attach — runs after every helper closure
    // (`onCanvasClick`, `onHostKey`, `syncFxRefs`, `commitFx`) is in
    // scope so the attacher bodies can resolve them at call time.
    for (const id of HOST_TOGGLEABLE_IDS) {
      if (flags[id as keyof typeof flags]) attachHostFeature(id);
    }

    if (opts.extensions) {
      const sorted = sortByPriority(dedupeById(flattenExtensions(opts.extensions)));
      for (const ext of sorted) mountExtension(ext);
    }
    refreshFeaturesView();

    // Locale change → push fresh strings everywhere. Built-ins that ship a
    // `setStrings` hook live-update labels in place; the rest are rebuilt by
    // detaching and re-attaching with the new dictionary in their closure.
    const unsubI18n = i18n.subscribe((next) => {
      strings = next;
      host.setAttribute('aria-label', strings.a11y.spreadsheet);
      tag.setAttribute('aria-label', strings.a11y.nameBox);
      refreshFormulaBarLabels();
      featureState.fxAutocomplete.setLabels(next.autocomplete);
      featureState.fxArgHelper?.setLabels(next.argHelper);
      sheetTabsController?.update();
      renderer.invalidate();
      alwaysOnDialogs.setStrings(next);

      // Toggleable host features: prefer setStrings when the handle exposes
      // it, otherwise fall back to detach+reattach.
      for (const id of HOST_TOGGLEABLE_IDS) {
        const handle = featureRegistry.get(id);
        if (!handle) continue;
        if (typeof handle.setStrings === 'function') {
          handle.setStrings(next);
        } else if (HOST_FEATURE_USES_STRINGS.has(id)) {
          detachHostFeature(id);
          attachHostFeature(id);
        }
      }

      // Engine-bound attaches (clipboard, paste-special, context-menu,
      // find-replace, validation) live inside `binding`. Rebuild it.
      binding.unbind();
      binding = bindEngine(wb);
      syncBindingFeatures(binding);
      featureState.viewToolbar?.bindWorkbook(wb);
      featureState.workbookObjects?.bindWorkbook(wb);
      featureState.pivotTableDialog?.bindWorkbook(wb);

      // User extensions opt-in via setStrings.
      for (const handle of userHandles.values()) handle.setStrings?.(next);

      emitter.emit('localeChange', { locale: i18n.locale, strings: next });
    });

    host.dataset.fcEngineState = wb.isStub ? 'ready-stub' : 'ready';

    let toolbarHandle: ToolbarInstance | null = null;
    let ribbonHost: HTMLElement | null = null;
    let requestedToolbar = opts.toolbar ?? (opts.ui ? ui.ribbon : false);

    const instance: SpreadsheetInstance = {
      host,
      get workbook() {
        return wb;
      },
      store,
      history,
      commands,
      applyChanges: (changes, options) => commands.applyChanges(changes, options),
      setPolicy(next) {
        if (disposed) return;
        closeMacPalette();
        binding.editor.cancel();
        formulaBar.cancelFx();
        commands.setPolicy(next);
        instance.setFeatures(requestedFeatures);
        instance.setToolbar(requestedToolbar);
      },
      setViewportOptions(next) {
        if (disposed) return;
        closeMacPalette();
        const targetSheet = next?.range?.sheet;
        if (targetSheet !== undefined && targetSheet !== store.getState().data.sheetIndex) {
          validateViewportAgainstWorkbook(next, wb, targetSheet);
          const previousState = store.getState();
          const previousOptions = navigation.options;
          binding.editor.cancel();
          formulaBar.cancelFx();
          wb.detachStore(store);
          try {
            navigation.setOptions(undefined);
            mutators.setSheetIndex(store, targetSheet);
            hydrateActiveSheet();
            navigation.setOptions(next);
          } catch (error) {
            store.setState(() => previousState);
            navigation.setOptions(previousOptions);
            throw error;
          } finally {
            wb.attachStore(store);
          }
          sheetTabsController?.update();
          updateChrome();
        } else {
          navigation.setOptions(next);
        }
        renderer.resize();
      },
      setContextMenu(next) {
        if (disposed) return;
        contextMenuOptions = next;
        instance.setFeatures(requestedFeatures);
        // Menu data may change without changing its feature flag.
        binding.unbind();
        binding = bindEngine(wb);
        syncBindingFeatures(binding);
        refreshFeaturesView();
      },
      setOverlayOptions(next) {
        if (disposed) return;
        setOverlayOptions(host, next);
      },
      setUi(next) {
        if (disposed) return;
        const nextUi = resolveSpreadsheetUiOptions(next);
        const platformChanged = nextUi.platform !== ui.platform;
        if (platformChanged) {
          closeMacPalette();
          if (featureState.fxDialog) detachHostFeature('fxDialog');
        }
        ui = nextUi;
        host.dataset.fcPlatform = ui.platform;
        applyPlatformLayoutDefaults(store, ui.platform);
        instance.setFeatures({ ...ui.features, ...opts.features });
        if (platformChanged && flags.fxDialog && !featureState.fxDialog) {
          attachHostFeature('fxDialog');
        }
        instance.setTheme(opts.theme ?? ui.theme);
        instance.setToolbar(opts.toolbar ?? (next ? ui.ribbon : false));
      },
      setToolbar(next) {
        if (disposed) return;
        requestedToolbar = next;
        disposeMacRibbonActions(instance);
        toolbarHandle?.dispose();
        toolbarHandle = null;
        ribbonHost?.remove();
        ribbonHost = null;
        if (!next) return;
        ribbonHost = host.ownerDocument.createElement('div');
        ribbonHost.className = 'fc-host__ribbon';
        ribbonHost.style.display = 'contents';
        host.insertBefore(ribbonHost, host.firstChild);
        const toolbarOpts = next === true ? {} : next;
        toolbarHandle = mountToolbar(ribbonHost, instance, {
          lang: i18n.locale === 'en' ? 'en' : 'ja',
          ...(commands.policy === undefined ? { dynamicDropdowns: true as const } : {}),
          ...toolbarOpts,
        });
      },
      i18n,
      features: featuresView,
      get toolbar() {
        return toolbarHandle;
      },
      get clipboard() {
        return binding.clipboardH;
      },
      get formatPainter() {
        return featureState.formatPainter ?? undefined;
      },
      get borderDraw() {
        return featureState.borderDraw ?? undefined;
      },
      formula: formulaRegistry,
      cells: cellRegistry,
      use(input) {
        const sorted = sortByPriority(dedupeById(flattenExtensions([input])));
        for (const ext of sorted) mountExtension(ext);
        refreshFeaturesView();
      },
      remove(id) {
        const handle = userHandles.get(id);
        if (!handle) return false;
        handle.dispose();
        userHandles.delete(id);
        refreshFeaturesView();
        return true;
      },
      setFeatures(next) {
        requestedFeatures = { ...next };
        const prevFlags = flags;
        const nextFlags = resolveMountFlags(
          next,
          commands.policy !== undefined,
          contextMenuOptions,
        );
        const shouldRebuildViewToolbarObjects =
          prevFlags.viewToolbar &&
          nextFlags.viewToolbar &&
          prevFlags.workbookObjects !== nextFlags.workbookObjects;
        flags = nextFlags;
        // Diff host-level features and dispatch attach/detach.
        for (const id of HOST_TOGGLEABLE_IDS) {
          const k = id as keyof typeof prevFlags;
          const was = prevFlags[k];
          const now = nextFlags[k];
          if (was === now) continue;
          if (was && !now) detachHostFeature(id);
          else if (!was && now) attachHostFeature(id);
        }
        // Wb-side rebuild only when a wb-bound feature flipped — keeps the
        // editor / pointer / undo state intact when only host-level flags
        // change.
        const wbChanged = WB_TOGGLEABLE_IDS.some(
          (id) =>
            prevFlags[id as keyof typeof prevFlags] !== nextFlags[id as keyof typeof nextFlags],
        );
        if (shouldRebuildViewToolbarObjects && featureState.viewToolbar) {
          detachHostFeature('viewToolbar');
          attachHostFeature('viewToolbar');
        }
        if (wbChanged) {
          binding.unbind();
          binding = bindEngine(wb);
          syncBindingFeatures(binding);
          featureState.viewToolbar?.bindWorkbook(wb);
          featureState.workbookObjects?.bindWorkbook(wb);
        }
        refreshFeaturesView();
      },
      setExtensions(next) {
        // Dispose all currently-mounted user extensions, then re-mount the
        // new list. Built-ins are untouched — use `setFeatures` for those.
        for (const handle of userHandles.values()) handle.dispose();
        userHandles.clear();
        if (next?.length) {
          const sorted = sortByPriority(dedupeById(flattenExtensions(next)));
          for (const ext of sorted) mountExtension(ext);
        }
        refreshFeaturesView();
      },
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
        if (commands.policy !== undefined) return;
        alwaysOnDialogs.openCfRules();
      },
      openCellStylesGallery() {
        if (commands.policy !== undefined) return;
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
        binding.findReplace?.open(tab);
      },
      closeFindReplace() {
        binding.findReplace?.close();
      },
      openPasteSpecial(opts) {
        binding.pasteSpecialDialog?.open(opts);
      },
      pasteSpecial(options, opts) {
        return binding.pasteSpecialDialog?.apply(options, opts) ?? false;
      },
      openInsertCopiedCells() {
        if (commands.policy !== undefined) return;

        // Whole-row/whole-column internal copies and cuts use the structural insert
        // command directly. This preserves source formats, merges, formulas,
        // and row/column dimensions; routing them through the TSV dialog
        // would reduce the payload to values and lose the band topology.
        const copied = binding.clipboardH?.getSnapshot() ?? null;
        const copiedLogical = copied?.logicalRange ?? copied?.range;
        const copiedWholeBand =
          copiedLogical !== undefined &&
          (isWholeRowRange(copiedLogical) || isWholeColumnRange(copiedLogical));
        if (copiedWholeBand && copied) {
          const target = store.getState().selection.range;
          const result = insertCopiedBand(store, wb, history, copied, target);
          if (result) {
            mutators.replaceCells(store, wb.cells(store.getState().data.sheetIndex));
            mutators.setRange(store, result.writtenRange);
            refreshCells();
            updateChrome();
            renderer.invalidate();
          }
          // A valid whole-band snapshot must not fall back to the direction
          // dialog when preflight rejects it: Excel leaves the sheet intact.
          return;
        }

        openInsertCopiedCellsDialog({
          host,
          strings: i18n.strings,
          onSubmit: (direction) => {
            void readClipboard().then((text) => {
              const snap = binding.clipboardH?.getSnapshot() ?? null;
              if (!text && !snap) return;
              const result = insertCopiedCellsFromTSV(store, wb, history, text, direction, snap);
              if (!result) return;
              // Marquee stays up, same as the context-menu variants.
              mutators.setRange(store, result.writtenRange);
              refreshCells();
              updateChrome();
              renderer.invalidate();
            });
          },
        });
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
      print(mode = 'print') {
        if (!ui.print) return;
        if (
          !commands.canExecute({
            operation: mode === 'pdf' ? 'export' : 'print',
            origin: 'instanceApi',
            effects: [{ kind: 'workbook' }],
          }).allowed
        )
          return;
        // The print command is wired through the same flag as the dialog —
        // when the feature is off, both call sites are no-ops. Skip if the
        // dialog never attached so consumers can rely on the gate.
        if (!featureState.pageSetupDialog) return;
        printSheet(
          wb,
          store,
          store.getState().data.sheetIndex,
          host,
          mode === 'pdf' ? strings.ribbon.pdf : strings.ribbon.print,
          mode,
          { printerProfiles: printers.getProfiles(), printerProfileId: printers.getProfileId() },
        );
      },
      setPrinterProfiles: printers.setProfiles,
      setPrinterProfileId: printers.setProfileId,
      refreshPrinterProfiles: printers.refresh,
      captureScreenClip,
      setUploadStatus(next) {
        uploadStatus = next;
        if (next !== null) declareStatusIndicator('uploadStatus');
        featureState.statusBar?.refresh();
      },
      setMacroRecording(next) {
        macroRecording = next;
        if (next !== null) declareStatusIndicator('macroRecording');
        featureState.statusBar?.refresh();
      },
      recalc() {
        wb.recalc();
        mutators.replaceCells(store, wb.cells(store.getState().data.sheetIndex));
        renderer.invalidate();
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
        if (commands.policy !== undefined) return;
        alwaysOnDialogs.openFilterAtHeader(range, col);
      },
      openWatchWindow() {
        if (commands.policy !== undefined) return;
        ensureWatchWindow();
        featureState.watchPanel?.open();
        refreshFeaturesView();
      },
      closeWatchWindow() {
        featureState.watchPanel?.close();
      },
      toggleWatchWindow() {
        if (commands.policy !== undefined) return;
        ensureWatchWindow();
        featureState.watchPanel?.toggle();
        refreshFeaturesView();
      },
      openQuickAnalysis() {
        const userQuick = userHandles.get('quickAnalysis') as
          | (ExtensionHandle & { open?: () => void })
          | undefined;
        if (userQuick?.open) {
          userQuick.open();
          return;
        }
        binding.quickAnalysis?.open();
      },
      openWorkbookObjects() {
        const userObjects = userHandles.get('workbookObjects') as
          | (ExtensionHandle & { open?: () => void })
          | undefined;
        if (userObjects?.open) {
          userObjects.open();
          return;
        }
        featureState.workbookObjects?.open();
      },
      openPivotFieldList(sheetIndex, pivotIndex) {
        if (commands.policy !== undefined) return false;
        const userObjects = userHandles.get('workbookObjects') as
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
        if (commands.policy !== undefined) return false;
        const pivot = findPivotTableAtCell(wb, store.getState().selection.active);
        if (!pivot) return false;
        const userObjects = userHandles.get('workbookObjects') as
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
        if (commands.policy !== undefined) return;
        const userPivot = userHandles.get('pivotTableDialog') as
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
      toggleSheetProtection() {
        if (commands.policy !== undefined) return;
        toggleProtectedSheet(store, store.getState().data.sheetIndex, { workbook: wb });
        renderer.invalidate();
      },
      setSheetProtected(
        on: boolean,
        password?: string,
        permissions?: import('./store/types.js').SheetProtectionPermissions,
      ) {
        if (commands.policy !== undefined) return;
        setProtectedSheet(store, store.getState().data.sheetIndex, on, {
          workbook: wb,
          password,
          permissions,
        });
        renderer.invalidate();
      },
      isSheetProtected() {
        return isSheetProtected(store.getState(), store.getState().data.sheetIndex);
      },
      tracePrecedents() {
        const count = tracePrecedentArrows(store, wb, store.getState().selection.active, history);
        renderer.invalidate();
        return count;
      },
      traceDependents() {
        const count = traceDependentArrows(store, wb, store.getState().selection.active, history);
        renderer.invalidate();
        return count;
      },
      clearTraces() {
        clearTraceArrows(store, history);
        renderer.invalidate();
      },
      setTheme(t) {
        host.dataset.fcTheme = t;
        syncOverlayPortalTheme(host);
        mutators.setTheme(store, t);
        renderer.invalidate();
        emitter.emit('themeChange', { theme: t });
      },
      undo() {
        // Batch the replay: a multi-cell entry restores cell by cell, and each
        // write would otherwise recalc on its own.
        const ok = wb.withBatchedRecalc(() => history.undo());
        if (ok) mutators.replaceCells(store, wb.cells(store.getState().data.sheetIndex));
        return ok;
      },
      redo() {
        const ok = wb.withBatchedRecalc(() => history.redo());
        if (ok) mutators.replaceCells(store, wb.cells(store.getState().data.sheetIndex));
        return ok;
      },
      async setWorkbook(next) {
        if (next === wb) return;
        // Validate before detaching the current workbook or clearing its history.
        if (navigation.options) {
          validateViewportAgainstWorkbook(
            navigation.options,
            next,
            navigation.options.range?.sheet ??
              Math.min(store.getState().data.sheetIndex, Math.max(0, next.sheetCount - 1)),
          );
        }
        // A Draw stroke belongs to the current workbook. Cancel it before
        // detaching the old engine so a late pointerup cannot commit points
        // into the newly bound workbook.
        deactivateMacInk(instance);
        closeMacPalette();
        binding.editor.cancel();
        formulaBar.cancelFx();
        wb.detachStore(store);
        binding.unbind();
        if (ownsWb) wb.dispose();
        wb = next;
        syncedSessionCfRules = new Map();
        ownsWb = true; // we now own the next handle and will dispose it
        wb.attachHistory(history);
        history.clear();
        const nextSheet =
          navigation.options?.range?.sheet ??
          Math.min(store.getState().data.sheetIndex, Math.max(0, wb.sheetCount - 1));
        mutators.setSheetIndex(store, nextSheet);
        mutators.clearIllustrations(store);
        store.setState((state) => ({
          ...state,
          format: { formats: new Map(), customCellStyles: [] },
          conditional: { rules: [] },
          ui: { ...state.ui, pendingFormat: null },
        }));
        hydrateActiveSheet();
        hydrateWorkbookMetadataFromEngine(wb, store);
        wb.attachStore(store);
        navigation.setOptions(navigation.options);
        dispatchPassthroughSummary();
        binding = bindEngine(wb);
        syncBindingFeatures(binding);
        featureState.viewToolbar?.bindWorkbook(wb);
        featureState.workbookObjects?.bindWorkbook(wb);
        featureState.namedRangeDialog?.bindWorkbook(wb);
        featureState.pivotTableDialog?.bindWorkbook(wb);
        featureState.statusBar?.refresh();
        sheetTabsController?.update();
        // Notify user extensions so they can rebind their wb references.
        for (const handle of [...userHandles.values()]) {
          try {
            handle.rebindWorkbook?.(wb);
          } catch (error) {
            console.warn('formulon-cell: extension workbook hook failed', error);
          }
        }
        for (const fn of [...wbListeners]) {
          try {
            fn(wb);
          } catch (error) {
            console.warn('formulon-cell: workbook hook failed', error);
          }
        }
        updateChrome();
        renderer.invalidate();
        emitter.emit('workbookChange', { workbook: wb });
      },
      on: (name, fn) => emitter.on(name, fn),
      off: (name, fn) => emitter.off(name, fn),
      dispose() {
        if (disposed) return;
        disposed = true;
        unregisterCommands();
        commands.dispose();
        navigation.dispose();
        closeMacPalette();
        wb.detachStore(store);
        disposeMacRibbonActions(instance);
        disposeMacInk(instance);
        toolbarHandle?.dispose();
        toolbarHandle = null;
        ribbonHost?.remove();
        ribbonHost = null;
        emitter.dispose();
        ro.disconnect();
        binding.unbind();
        for (const handle of userHandles.values()) handle.dispose();
        userHandles.clear();
        for (const id of HOST_TOGGLEABLE_IDS) detachHostFeature(id);
        formulaBar.detach();
        sheetTabsController?.detach();
        chromeSync?.detach();
        host.removeEventListener('fc:openfilter', onOpenFilter);
        alwaysOnDialogs.detach();
        unsubCellRegistry();
        formulaDraftMirror.detach();
        pivotFieldListFollow.detach();
        unsubI18n();
        i18n.dispose();
        renderer.dispose();
        if (ownsWb) wb.dispose();
        disposeOverlayPortal(host, instanceId);
        releaseMountHost(host, instanceId);
      },
    };

    // Single-call ribbon: mount the toolbar into a `display: contents` host at
    // the top of `.fc-host` so its shell participates in the host flex column
    // directly (grid fills the rest). The toolbar reads the same
    // `data-fc-theme`, so grid and toolbar theme together via the cascade.
    instance.setToolbar(requestedToolbar);

    return instance;
  },
};
