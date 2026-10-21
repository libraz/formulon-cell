import { CellRegistry } from './cells.js';
import { History } from './commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from './commands/interaction-controller.js';
import { PrintResourceLimitError, printSheet } from './commands/print.js';
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
import { rehydrateAxisLayout, syncActiveSheet } from './engine/layout-sync.js';
import { hydrateViewStateFromEngine } from './engine/view-state-sync.js';
import { WorkbookHandle } from './engine/workbook-handle.js';
import { SpreadsheetEmitter } from './events.js';
import { resolveSpreadsheetUiOptions } from './extensions/index.js';
import { FormulaRegistry } from './formula.js';
import { createI18nController } from './i18n/controller.js';
import { dictionaryLocaleFor, type Strings } from './i18n/strings.js';
import { deactivateMacInk, disposeMacInk } from './interact/mac-ink.js';
import { attachNavigationPolicy, navigationBoundsFor } from './interact/navigation-policy.js';
import {
  disposeOverlayPortal,
  ensureOverlayPortal,
  hasOpenLocaleDraft,
  setOverlayOptions,
  syncOverlayPortalTheme,
} from './interact/overlay-portal.js';
import { attachAlwaysOnDialogs } from './mount/always-on-dialogs.js';
import { createMountChrome } from './mount/chrome.js';
import { attachChromeSync, type ChromeSyncController } from './mount/chrome-sync.js';
import { attachEngineBinding, type EngineBinding } from './mount/engine-binding.js';
import { createErrorIndicatorClickHandler } from './mount/error-indicator-click.js';
import { createExtensionRegistry } from './mount/extension-registry.js';
import { resolveMountFlags } from './mount/feature-flags.js';
import { attachFormulaBarController } from './mount/formula-bar.js';
import { attachFormulaDraftMirror } from './mount/formula-draft-mirror.js';
import {
  applyPlatformChrome,
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
import { createInsertCopiedCellsOpener } from './mount/insert-copied-cells.js';
import { createInstanceDialogApi } from './mount/instance-dialog-api.js';
import { syncLocaleProfile } from './mount/locale-profile.js';
import { attachPivotFieldListFollow } from './mount/pivot-field-list-follow.js';
import { createPrinterProfiles } from './mount/printer-profiles.js';
import { createRibbonHost } from './mount/ribbon-host.js';
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
import { createSpreadsheetStore, mutators } from './store/store.js';
import { resolveTheme } from './theme/resolve.js';
import { showMessage } from './toolbar/dialogs/prompt.js';
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

// A workbook can be mounted by more than one host at a time.  The engine
// exposes one history slot per workbook. The latest live mount owns that slot;
// teardown restores the preceding owner without detaching another live mount.
const workbookHistoryOwners = new WeakMap<WorkbookHandle, History[]>();

function attachWorkbookHistory(wb: WorkbookHandle, history: History): void {
  const owners = workbookHistoryOwners.get(wb) ?? [];
  owners.push(history);
  workbookHistoryOwners.set(wb, owners);
  wb.attachHistory(history);
}

function detachWorkbookHistory(wb: WorkbookHandle, history: History): void {
  const owners = workbookHistoryOwners.get(wb);
  if (!owners) return;
  const index = owners.indexOf(history);
  if (index < 0) return;
  const wasLatest = index === owners.length - 1;
  owners.splice(index, 1);
  if (owners.length === 0) workbookHistoryOwners.delete(wb);
  if (wasLatest) wb.attachHistory(owners.at(-1) ?? null);
}

function throwCleanupErrors(errors: unknown[]): void {
  if (errors.length === 0) return;
  if (errors.length === 1) throw errors[0];
  throw new AggregateError(errors, 'formulon-cell: one or more cleanup operations failed');
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

    const previousHostLanguage = host.getAttribute('lang');
    const languageTag = (): string => {
      try {
        return (
          Intl.getCanonicalLocales(i18n.locale.replaceAll('_', '-'))[0] ??
          dictionaryLocaleFor(i18n.locale)
        );
      } catch {
        return dictionaryLocaleFor(i18n.locale);
      }
    };
    const instanceId = prepareMountHost(host, strings, initialTheme, ui.platform);
    host.lang = languageTag();
    const overlayPortal = ensureOverlayPortal(host, opts.overlays);
    host.dataset.fcEngineState = 'loading';

    let sheetTabsController: SheetTabsController | null = null;

    // Track ownership before seeding — only owned (default-created) workbooks
    // should be touched by `seed`. Pre-loaded workbooks are the consumer's
    // data and must not be overwritten by the demo helper.
    let ownsWb = !opts.workbook;
    let wb!: WorkbookHandle;
    let workbookDisposed = false;
    const disposeOwnedWorkbook = (): void => {
      if (!ownsWb || workbookDisposed || !wb) return;
      workbookDisposed = true;
      wb.dispose();
    };
    try {
      wb =
        opts.workbook ??
        (await WorkbookHandle.createDefault({ locale: i18n.locale, platform: ui.platform }));
      wb.geometryPlatform = ui.platform;
      // The engine profile is workbook state and is not persisted in files.
      // Select it before a seed writes formulas so seeded values use the same
      // host semantics as the first hydrated paint.
      syncLocaleProfile(wb, i18n.locale, ui.platform, { throwOnFailure: true });
      if (opts.seed && ownsWb) opts.seed(wb);
      if (opts.viewport) {
        validateViewportAgainstWorkbook(opts.viewport, wb, opts.viewport.range?.sheet);
      }
    } catch (err) {
      // The default workbook is owned before the mount promise resolves.  A
      // seed/profile/viewport failure must release it even though no instance
      // exists for the caller to dispose.  External workbooks stay borrowed.
      try {
        disposeOwnedWorkbook();
      } catch (cleanupError) {
        console.error(
          'formulon-cell: failed to dispose workbook after mount failure',
          cleanupError,
        );
      }
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
      localeSwitcher,
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
    applyPlatformChrome(store, ui.platform);

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
    attachWorkbookHistory(wb, history);
    history.clear();

    hydrateViewStateFromEngine(wb, store, { followActiveSheet: !opts.viewport?.range });
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
      isRestricted: () => commands.policy !== undefined,
    });

    const refreshCells = (): void => {
      const sheetIndex = store.getState().data.sheetIndex;
      // A host can change the store's active sheet while a workbook has no
      // matching sheet yet (for example during a platform/workbook transition).
      // Profile refresh must not turn that transient state into an engine
      // sheet-index exception.
      if (sheetIndex < 0 || sheetIndex >= wb.sheetCount) return;
      mutators.replaceCells(store, wb.cells(sheetIndex));
    };
    const extensionRegistry = createExtensionRegistry({
      context: {
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
      },
    });
    const { featureRegistry, featuresView, wrapHandle, syncBindingFeatures, refreshFeaturesView } =
      extensionRegistry;

    const autocompleteStub = createAutocompleteStub();
    const featureState = createHostFeatureState(autocompleteStub);
    const captureScreenClip = async (): Promise<ScreenClipResult | null> =>
      normalizeScreenClipResult(await captureScreenClipHook?.());
    const pivotFieldListFollow = attachPivotFieldListFollow({
      store,
      getWb: () => wb,
      getWorkbookObjects: () => featureState.workbookObjects,
    });

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
        getWorkbookObjects: () => featureState.workbookObjects,
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
    const rebindEngine = (): void => {
      binding.unbind();
      binding = bindEngine(wb);
      syncBindingFeatures(binding);
    };

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
      getLocale: () => i18n.locale,
      namedRangeDialog: () => featureState.namedRangeDialog,
      pasteSpecialDialog: () => binding.pasteSpecialDialog,
      quickAnalysis: () => binding.quickAnalysis,
      store,
      strings: () => strings,
      wb: () => wb,
    });

    const onCanvasClick = createErrorIndicatorClickHandler({
      grid,
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
      getHistory: () => history,
      getSheetTabs: () => sheetTabsController,
      getStrings: () => strings,
      getWb: () => wb,
      grid,
      host,
      invalidate: () => renderer.invalidate(),
      onAlert: (outcome) => binding.validationAlert?.show(outcome),
      store,
      tag,
    });

    const formulaDraftMirror = attachFormulaDraftMirror({ grid, store, wb });

    // Resize observer — we follow the host, not the window.
    const ro = new ResizeObserver(() => {
      renderer.resize();
      formulaDraftMirror.refresh();
    });
    ro.observe(grid);

    let disposed = false;

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
      getWorkbookName: opts.getWorkbookName,
      localeSwitcher,
      setLocale: (locale) => instance.setLocale(locale),
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
    const discardMacPalette = (): void => {
      featureState.macFormulaPalette?.discard();
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

    refreshFeaturesView();

    // Locale change → push fresh strings everywhere. Built-ins that ship a
    // `setStrings` hook live-update labels in place; the rest are rebuilt by
    // detaching and re-attaching with the new dictionary in their closure.
    // i18n.setLocale updates its current value before notifying us, so settle
    // an active edit while the workbook still has the old profile. If a live
    // profile write fails, restore the previous locale instead of leaving the
    // UI and parser on different languages.
    let appliedLocale = i18n.locale;
    let rollingBackLocale = false;
    const rollbackLocale = (previous: string, reason: unknown): void => {
      console.warn('formulon-cell: locale profile switch rejected', reason);
      if (rollingBackLocale || i18n.locale === previous) return;
      rollingBackLocale = true;
      try {
        i18n.setLocale(previous);
      } finally {
        rollingBackLocale = false;
      }
    };
    const settleEditsBeforeLocaleSwitch = (): boolean => {
      if (binding.editor.isActive()) {
        binding.editor.commit('none');
        if (binding.editor.isActive()) return false;
      }
      if (formulaBar.isEditing()) {
        if (!formulaBar.commitFx('none') || formulaBar.isEditing()) return false;
      }
      return true;
    };
    let deferredStrings: Strings | null = null;
    const applyLocaleStrings = (next: Strings): void => {
      const localeChanged = i18n.locale !== appliedLocale;
      let profileChanged = false;
      // Keep the dialog and all its inputs intact, including during the
      // synchronous callback used to restore the controller's old locale.
      if (hasOpenLocaleDraft(host)) {
        if (localeChanged && !rollingBackLocale) {
          rollbackLocale(appliedLocale, new Error('an active dialog draft is open'));
        } else if (!rollingBackLocale) {
          // Overlay updates do not change the language. Refresh safe chrome
          // now and rebuild hidden factories once the last draft closes.
          deferredStrings = next;
          strings = next;
          host.setAttribute('aria-label', strings.a11y.spreadsheet);
          tag.setAttribute('aria-label', strings.a11y.nameBox);
          refreshFormulaBarLabels();
        }
        return;
      }
      deferredStrings = null;
      // The Mac formula palette owns an external formula draft while it is
      // open. Its public handle can discard that lease, but cannot restore
      // the user's raw text after the locale callback has already run. Keep
      // the pane and reject the switch so the caller can retry after closing
      // or committing it.
      if (localeChanged && !rollingBackLocale && featureState.macFormulaPalette?.isOpen()) {
        rollbackLocale(appliedLocale, new Error('an active Mac formula palette edit is open'));
        return;
      }
      // A rollback callback must refresh the old labels without discarding
      // the palette that caused the rejection.
      if (!(rollingBackLocale && featureState.macFormulaPalette?.isOpen())) {
        discardMacPalette();
      }
      if (localeChanged && !rollingBackLocale) {
        if (!settleEditsBeforeLocaleSwitch()) {
          rollbackLocale(appliedLocale, new Error('an active formula edit could not be committed'));
          return;
        }
        try {
          profileChanged = syncLocaleProfile(wb, i18n.locale, ui.platform, {
            throwOnFailure: true,
          });
        } catch (error) {
          rollbackLocale(appliedLocale, error);
          return;
        }
        appliedLocale = i18n.locale;
      }
      strings = next;
      host.lang = languageTag();
      overlayPortal.lang = host.lang;
      host.setAttribute('aria-label', strings.a11y.spreadsheet);
      tag.setAttribute('aria-label', strings.a11y.nameBox);
      refreshFormulaBarLabels();
      featureState.fxAutocomplete.setLabels(next.autocomplete);
      featureState.fxArgHelper?.setLabels(next.argHelper);
      sheetTabsController?.update();
      renderer.invalidate();
      alwaysOnDialogs.setStrings(next);
      if (profileChanged) refreshCells();

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
      rebindEngine();

      // User extensions opt-in via setStrings.
      extensionRegistry.setUserStrings(next);

      emitter.emit('localeChange', { locale: i18n.locale, strings: next });
    };
    const unsubI18n = i18n.subscribe(applyLocaleStrings);
    const draftObserver = new MutationObserver(() => {
      if (disposed || deferredStrings === null || hasOpenLocaleDraft(host)) return;
      applyLocaleStrings(deferredStrings);
    });
    const draftObservation = {
      subtree: true,
      childList: true,
      attributes: true,
      attributeFilter: ['hidden', 'aria-hidden', 'style', 'class', 'open'],
    };
    draftObserver.observe(host, draftObservation);
    draftObserver.observe(overlayPortal, draftObservation);

    host.dataset.fcEngineState = wb.isStub ? 'ready-stub' : 'ready';

    const ribbon = createRibbonHost({
      host,
      getInstance: () => instance,
      getLocale: () => i18n.locale,
      isRestricted: () => commands.policy !== undefined,
    });
    let requestedToolbar = opts.toolbar ?? (opts.ui ? ui.ribbon : false);

    const openInsertCopiedCells = createInsertCopiedCellsOpener({
      host,
      store,
      history,
      getWb: () => wb,
      getStrings: () => i18n.strings,
      isRestricted: () => commands.policy !== undefined,
      getClipboardSnapshot: () => binding.clipboardH?.getSnapshot() ?? null,
      refreshCells,
      updateChrome,
      invalidate: () => renderer.invalidate(),
    });

    const instanceDialogApi = createInstanceDialogApi({
      featureState,
      alwaysOnDialogs,
      store,
      getWb: () => wb,
      getBinding: () => binding,
      getUserHandle: extensionRegistry.getUserHandle,
      isRestricted: () => commands.policy !== undefined,
      ensureWatchWindow,
      refreshFeaturesView,
    });

    const instance: SpreadsheetInstance = {
      host,
      get workbook() {
        return wb;
      },
      store,
      history,
      commands,
      applyChanges: (changes, options) => commands.applyChanges(changes, options),
      get locale() {
        return i18n.locale;
      },
      setLocale(next) {
        if (disposed) return false;
        const previous = i18n.locale;
        if (next === previous) return true;
        if (hasOpenLocaleDraft(host)) return false;
        try {
          i18n.setLocale(next);
        } catch (error) {
          // A subscriber can still fail after the profile transaction (for
          // example while rebinding a host feature). Keep the public locale
          // and engine profile aligned before reporting rejection.
          rollbackLocale(previous, error);
          return false;
        }
        return i18n.locale === next;
      },
      setPolicy(next) {
        if (disposed) return;
        discardMacPalette();
        binding.editor.cancel();
        formulaBar.discardFx();
        commands.setPolicy(next);
        instance.setFeatures(requestedFeatures);
        instance.setToolbar(requestedToolbar);
      },
      setViewportOptions(next) {
        if (disposed) return;
        discardMacPalette();
        const targetSheet = next?.range?.sheet;
        if (targetSheet !== undefined && targetSheet !== store.getState().data.sheetIndex) {
          validateViewportAgainstWorkbook(next, wb, targetSheet);
          const previousState = store.getState();
          const previousOptions = navigation.options;
          binding.editor.cancel();
          formulaBar.discardFx();
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
        rebindEngine();
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
        let profileChanged = false;
        if (platformChanged) {
          // Keep the old UI platform attached when the profile-capable engine
          // rejects the corresponding platform profile.
          profileChanged = syncLocaleProfile(wb, i18n.locale, nextUi.platform, {
            throwOnFailure: true,
          });
        }
        if (platformChanged) {
          discardMacPalette();
          if (featureState.fxDialog) detachHostFeature('fxDialog');
        }
        ui = nextUi;
        host.dataset.fcPlatform = ui.platform;
        if (platformChanged) wb.geometryPlatform = ui.platform;
        applyPlatformChrome(store, ui.platform);
        if (platformChanged) {
          const sheetIndex = store.getState().data.sheetIndex;
          if (sheetIndex >= 0 && sheetIndex < wb.sheetCount) {
            rehydrateAxisLayout(wb, store, sheetIndex);
          }
          // History snapshots hold column px projected by the previous
          // platform's width model; replaying them would write wrong widths.
          history.clear();
        }
        if (profileChanged) refreshCells();
        if (platformChanged) renderer.invalidate();
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
        ribbon.setToolbar(next);
      },
      i18n,
      features: featuresView,
      get toolbar() {
        return ribbon.getToolbar();
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
      use: extensionRegistry.use,
      remove: extensionRegistry.remove,
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
        if (wbChanged) rebindEngine();
        refreshFeaturesView();
      },
      setExtensions: extensionRegistry.setExtensions,
      ...instanceDialogApi,
      pasteSpecial(options, opts) {
        return binding.pasteSpecialDialog?.apply(options, opts) ?? false;
      },
      openInsertCopiedCells,
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
        try {
          printSheet(
            wb,
            store,
            store.getState().data.sheetIndex,
            host,
            mode === 'pdf' ? strings.ribbon.pdf : strings.ribbon.print,
            mode,
            { printerProfiles: printers.getProfiles(), printerProfileId: printers.getProfileId() },
          );
        } catch (error) {
          if (!(error instanceof PrintResourceLimitError)) throw error;
          void showMessage({
            host,
            title: strings.ribbon.print,
            message: strings.pageSetup.printRangeTooLarge,
            okLabel: strings.formatDialog.ok,
          });
        }
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
        if (disposed) throw new Error('formulon-cell: spreadsheet is disposed');
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
        // Configure the incoming workbook while the old one is still live.
        // A rejected profile leaves the current workbook and its store
        // untouched because detachment happens only after this call.
        syncLocaleProfile(next, i18n.locale, ui.platform, { throwOnFailure: true });
        next.geometryPlatform = ui.platform;
        discardMacPalette();
        // A Draw stroke belongs to the current workbook. Cancel it before
        // detaching the old engine so a late pointerup cannot commit points
        // into the newly bound workbook.
        deactivateMacInk(instance);
        binding.editor.cancel();
        formulaBar.discardFx();
        wb.detachStore(store);
        detachWorkbookHistory(wb, history);
        binding.unbind();
        disposeOwnedWorkbook();
        wb = next;
        syncedSessionCfRules = new Map();
        ownsWb = true; // we now own the next handle and will dispose it
        workbookDisposed = false;
        attachWorkbookHistory(wb, history);
        history.clear();
        const nextSheet =
          navigation.options?.range?.sheet ??
          Math.min(store.getState().data.sheetIndex, Math.max(0, wb.sheetCount - 1));
        mutators.setSheetIndex(store, nextSheet);
        // Per-sheet selection and scroll belong to the previous workbook; the
        // incoming one brings its own saved view.
        store.setState((state) => ({ ...state, ui: { ...state.ui, sheetViewMemos: undefined } }));
        hydrateViewStateFromEngine(wb, store, { followActiveSheet: !navigation.options?.range });
        syncActiveSheet(wb, store.getState().data.sheetIndex);
        mutators.clearIllustrations(store);
        store.setState((state) => ({
          ...state,
          threadedComments: { bySheet: new Map(), persons: [] },
        }));
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
        hostFeatures.bindWorkbook(wb);
        featureState.statusBar?.refresh();
        sheetTabsController?.update();
        extensionRegistry.notifyWorkbookChange(wb);
        updateChrome();
        renderer.invalidate();
        emitter.emit('workbookChange', { workbook: wb });
      },
      on: (name, fn) => emitter.on(name, fn),
      off: (name, fn) => emitter.off(name, fn),
      dispose() {
        if (disposed) return;
        disposed = true;
        const errors: unknown[] = [];
        const cleanup = (fn: () => void): void => {
          try {
            fn();
          } catch (error) {
            errors.push(error);
          }
        };

        // Every teardown is independent.  In particular, a user extension is
        // allowed to throw without preventing the engine binding, workbook,
        // portal, or host from being released.
        cleanup(discardMacPalette);
        cleanup(formulaBar.discardFx);
        cleanup(unregisterCommands);
        cleanup(() => commands.dispose());
        cleanup(() => navigation.dispose());
        cleanup(() => wb.detachStore(store));
        cleanup(() => detachWorkbookHistory(wb, history));
        cleanup(() => disposeMacRibbonActions(instance));
        cleanup(() => disposeMacInk(instance));
        cleanup(() => ribbon.dispose());
        cleanup(() => emitter.dispose());
        cleanup(() => ro.disconnect());
        cleanup(() => binding.unbind());
        cleanup(() => extensionRegistry.dispose());
        for (const id of HOST_TOGGLEABLE_IDS) cleanup(() => detachHostFeature(id));
        cleanup(() => formulaBar.detach());
        cleanup(() => sheetTabsController?.detach());
        cleanup(() => chromeSync?.detach());
        cleanup(() => alwaysOnDialogs.detach());
        cleanup(unsubCellRegistry);
        cleanup(() => formulaDraftMirror.detach());
        cleanup(() => pivotFieldListFollow.detach());
        cleanup(unsubI18n);
        cleanup(() => draftObserver.disconnect());
        cleanup(() => i18n.dispose());
        cleanup(() => renderer.dispose());
        cleanup(disposeOwnedWorkbook);
        cleanup(() => disposeOverlayPortal(host, instanceId));
        cleanup(() => {
          if (host.dataset.fcInstId === instanceId) {
            if (previousHostLanguage === null) host.removeAttribute('lang');
            else host.setAttribute('lang', previousHostLanguage);
          }
        });
        cleanup(() => releaseMountHost(host, instanceId));
        throwCleanupErrors(errors);
      },
    };

    try {
      // Install the locale transaction before user extensions can subscribe.
      if (opts.extensions) extensionRegistry.mountAll(opts.extensions);
      refreshFeaturesView();

      // Single-call ribbon: mount the toolbar into a `display: contents` host at
      // the top of `.fc-host` so its shell participates in the host flex column
      // directly (grid fills the rest). The toolbar reads the same
      // `data-fc-theme`, so grid and toolbar theme together via the cascade.
      instance.setToolbar(requestedToolbar);
    } catch (error) {
      // At this point the instance is fully live. A setup failure must still
      // release every resource, while the caller receives the original setup
      // error rather than a cleanup exception.
      try {
        instance.dispose();
      } catch (cleanupError) {
        console.error('formulon-cell: mount cleanup after setup failure threw', cleanupError);
      }
      try {
        opts.onError?.(error);
      } catch (hookError) {
        console.error('formulon-cell: mount error handler threw', hookError);
      }
      if (opts.renderError !== false) renderMountError(host, error, strings.mountError);
      throw error;
    }

    return instance;
  },
};
