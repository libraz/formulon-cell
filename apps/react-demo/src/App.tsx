import {
  type CellEvent,
  type HostProviders,
  type SpreadsheetInstance,
  type ThemeName,
  Workbook,
} from '@libraz/formulon-cell';
import type { FeatureFlags, FeatureId } from '@libraz/formulon-cell/extensions';
import type { RibbonHooks, RibbonTab, Toolbar } from '@libraz/formulon-cell/toolbar';
import { Spreadsheet, SpreadsheetToolbar } from '@libraz/formulon-cell-react';
import { type ReactElement, useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  adoptDemoWorkbook,
  buildDemoCommands,
  buildDemoReviewDialog,
  buildDemoSearchItems,
  type ChangeLogEntry,
  composeDemoUiOptions,
  createDemoStrings,
  createInitialDemoWorkbook,
  DEMO_FORMULA_HINTS,
  DEMO_MAC_RIBBON_TABS,
  DEMO_PRINTER_PROFILE_ID,
  DEMO_PRINTER_PROFILES,
  DEMO_RIBBON_TABS,
  type DemoBackstageAction,
  type DemoF6Toolbar,
  type DemoLocale,
  type DemoPlatform,
  type DemoReviewDialogState,
  demoChangeLogEntry,
  demoCommandText,
  demoFunctionArgumentHelp,
  demoUiStringsForLocale,
  FORMATTERS,
  formatLoadError,
  installDemoF6Navigation,
  nextDemoFeatureOverrides,
  normalizeDemoLocale,
  openDemoWorkbookFile,
  type PresetKey,
  pushDemoChangeLog,
  refreshDemoPrinterProfiles,
  resolveInitialLocale,
  resolveInitialPlatform,
  runDemoBackstageAction,
  saveDemoWorkbookAsToDownload,
  saveDemoWorkbookToDownload,
} from '../../demo-shared/index.js';
import { DemoBackstage } from './DemoBackstage.js';
import { DemoOptionsPanel } from './DemoOptionsPanel.js';
import { DemoReviewDialog } from './DemoReviewDialog.js';
import { DemoTitleBar } from './DemoTitleBar.js';

const UI = createDemoStrings('React');

/** The ribbon as F6 navigation reads it: the wrapper element around it and its tab focus. */
const f6ToolbarOf = (toolbar: Toolbar | null, host: HTMLElement | null): DemoF6Toolbar | null =>
  host ? { host, focusActiveTab: () => toolbar?.focusActiveTab() ?? false } : null;

export const App = (): ReactElement => {
  const [theme, setTheme] = useState<ThemeName>('paper');
  const [locale, setLocale] = useState<DemoLocale>(() => resolveInitialLocale());
  const [platform] = useState<DemoPlatform>(() => resolveInitialPlatform());
  const [workbook, setWorkbook] = useState<Workbook | null>(null);
  const [instance, setInstance] = useState<SpreadsheetInstance | null>(null);
  const [log, setLog] = useState<ChangeLogEntry[]>([]);
  const [formatters, setFormatters] = useState({ uppercase: false, arrows: false });
  const [preset, setPreset] = useState<PresetKey>('full');
  const [overrides, setOverrides] = useState<FeatureFlags>({});
  const [showRibbon, setShowRibbon] = useState(true);
  const [showPanel, setShowPanel] = useState(false);
  const [ribbonTab, setRibbonTab] = useState<RibbonTab>('home');
  const [backstageAction, setBackstageAction] = useState<DemoBackstageAction>('info');
  const [loadError, setLoadError] = useState<string | null>(null);
  const [reviewDialog, setReviewDialog] = useState<DemoReviewDialogState | null>(null);
  const [uploadStatus, setUploadStatus] = useState<'saved' | 'saving' | 'error' | null>(null);
  // Workbook display name. Untitled until the user opens or saves a file —
  // mirrors the spreadsheet titlebar convention. Stripping the extension
  // keeps it tidy in the chrome while preserving the user's filename for
  // re-saves.
  const [bookName, setBookName] = useState('Book1');
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const quickAccessRef = useRef<HTMLDivElement | null>(null);
  const ribbonHostRef = useRef<HTMLDivElement | null>(null);
  const toolbarRef = useRef<Toolbar | null>(null);
  const workbookOwnerRef = useRef<Workbook | null>(null);
  const instanceRef = useRef<SpreadsheetInstance | null>(null);
  const aliveRef = useRef(true);

  const resolvedUi = useMemo(
    () => composeDemoUiOptions({ preset, overrides, showRibbon, platform }),
    [overrides, platform, preset, showRibbon],
  );
  const ribbonTabs = platform === 'mac' ? DEMO_MAC_RIBBON_TABS : DEMO_RIBBON_TABS;
  const features = resolvedUi.features;
  const ui = demoUiStringsForLocale(UI, locale);
  const commandText = useMemo(() => demoCommandText(locale), [locale]);
  const closeReviewDialog = useCallback(() => setReviewDialog(null), []);
  const host = useMemo<HostProviders>(
    () => ({
      workbookName: () => bookName,
      printerProfiles: DEMO_PRINTER_PROFILES,
      printerProfileId: DEMO_PRINTER_PROFILE_ID,
      refreshPrinterProfiles: refreshDemoPrinterProfiles,
      functionArgumentHelp: demoFunctionArgumentHelp,
    }),
    [bookName],
  );

  const onCellChange = useCallback((live: SpreadsheetInstance, e: CellEvent) => {
    // The id is drawn once, outside the updater React may call twice.
    const entry = demoChangeLogEntry(live.workbook, e);
    setLog((prev) => pushDemoChangeLog(prev, entry));
  }, []);

  useEffect(() => {
    let alive = true;
    aliveRef.current = true;
    void createInitialDemoWorkbook()
      .then((wb) => {
        if (!alive || !aliveRef.current) {
          // The demo created this workbook, so a late completion must
          // release it instead of leaving the engine workbook orphaned.
          wb.dispose();
          return;
        }
        workbookOwnerRef.current = wb;
        setLoadError(null);
        setWorkbook(wb);
      })
      .catch((err: unknown) => {
        if (!alive) return;
        setLoadError(formatLoadError(err));
      });
    return () => {
      alive = false;
      aliveRef.current = false;
      // Dispose the instance first, then the workbook the demo owns.
      try {
        instanceRef.current?.dispose();
      } catch (error) {
        console.error('formulon demo: spreadsheet cleanup failed', error);
      } finally {
        instanceRef.current = null;
        workbookOwnerRef.current?.dispose();
        workbookOwnerRef.current = null;
      }
    };
  }, []);

  useEffect(() => {
    if (!instance) return undefined;
    return instance.on('cell', (e) => onCellChange(instance, e));
  }, [instance, onCellChange]);

  useEffect(() => {
    if (!instance) return undefined;
    const disposers: (() => void)[] = [];
    if (formatters.uppercase) disposers.push(instance.cells.add(FORMATTERS.uppercaseA));
    if (formatters.arrows) disposers.push(instance.cells.add(FORMATTERS.arrowNegatives));
    return () => {
      for (const d of disposers) d();
    };
  }, [instance, formatters.uppercase, formatters.arrows]);

  useEffect(() => {
    document.documentElement.lang = locale;
  }, [locale]);

  useEffect(
    () =>
      installDemoF6Navigation({
        getQuickAccess: () => quickAccessRef.current,
        getToolbar: () => f6ToolbarOf(toolbarRef.current, ribbonHostRef.current),
        getInstance: () => instance,
      }),
    [instance],
  );

  // Expose the live instance on `window.__fcInst` so cross-demo E2E scenarios
  // can drive imperative paths (named-range, paste-special, etc.) without
  // depending on demo-specific UI.
  useEffect(() => {
    (window as unknown as { __fcInst?: SpreadsheetInstance | null }).__fcInst = instance;
    return () => {
      delete (window as unknown as { __fcInst?: SpreadsheetInstance | null }).__fcInst;
    };
  }, [instance]);

  const showRibbonNotice = useCallback(
    (title: string, detail: string) => {
      setReviewDialog(buildDemoReviewDialog(title, commandText.ribbonCommand, detail));
    },
    [commandText.ribbonCommand],
  );

  const ribbonHooks = useMemo<RibbonHooks>(
    () => ({
      review: {
        translate: () =>
          showRibbonNotice(commandText.translate, commandText.translationUnavailable),
      },
      drawing: {
        setInkMode: (mode) =>
          showRibbonNotice(
            commandText.draw,
            mode === 'pen' ? commandText.inkNotPersisted : commandText.selectInkFirst,
          ),
      },
      automation: {
        addInManager: () => showRibbonNotice(commandText.addIns, commandText.addInsHostCallbacks),
      },
    }),
    [commandText, showRibbonNotice],
  );

  const onSave = useCallback(() => {
    void saveDemoWorkbookToDownload({ instance, bookName, setUploadStatus });
  }, [bookName, instance]);

  const onSaveAs = useCallback(() => {
    void saveDemoWorkbookAsToDownload({ instance, bookName, setUploadStatus });
  }, [bookName, instance]);

  const onSpreadsheetError = useCallback((error: unknown) => {
    console.error('formulon demo: spreadsheet error', error);
  }, []);

  const onNewWorkbook = useCallback(async () => {
    const wb = await Workbook.createDefault({ locale, platform });
    const previous = workbookOwnerRef.current;
    const live = instanceRef.current ?? instance;
    try {
      // Swapping here rather than through the prop reports a refused swap to this flow; the
      // prop then already matches the instance.
      if (!(await adoptDemoWorkbook(live, wb, () => aliveRef.current))) return;
    } catch (error) {
      if (!aliveRef.current) return;
      setReviewDialog(
        buildDemoReviewDialog(commandText.openFailed, commandText.workbook, formatLoadError(error)),
      );
      return;
    }
    workbookOwnerRef.current = wb;
    setWorkbook(wb);
    previous?.dispose();
    setBookName('Book1');
    setLog([]);
    setRibbonTab('home');
  }, [commandText.openFailed, commandText.workbook, instance, locale, platform]);

  const onOpen = useCallback(
    async (file: File) => {
      const live = instance;
      if (!live || !aliveRef.current) return;
      const previous = workbookOwnerRef.current;
      try {
        const { workbook: next, name } = await openDemoWorkbookFile(file, { locale, platform });
        const current = () => aliveRef.current && instanceRef.current === live;
        if (!(await adoptDemoWorkbook(live, next, current))) return;
        workbookOwnerRef.current = next;
        setWorkbook(next);
        // The instance no longer shows the previous workbook; release the one the demo owned.
        previous?.dispose();
        setBookName(name);
        setLoadError(null);
      } catch (err) {
        if (!aliveRef.current) return;
        setReviewDialog(
          buildDemoReviewDialog(commandText.openFailed, commandText.workbook, formatLoadError(err)),
        );
      }
    },
    [commandText.openFailed, commandText.workbook, instance, locale, platform],
  );

  const runBackstageAction = useCallback(
    (action: DemoBackstageAction): void => {
      if (action === 'info' || action === 'print') {
        setBackstageAction(action);
        return;
      }
      runDemoBackstageAction({
        action,
        instance,
        ui,
        newWorkbook: onNewWorkbook,
        openWorkbook: () => fileInputRef.current?.click(),
        saveWorkbook: onSave,
        saveWorkbookAs: onSaveAs,
        showNotice: showRibbonNotice,
        toggleOptions: () => setShowPanel((v) => !v),
        closeBackstage: () => {
          toolbarRef.current?.update({ backstageOpen: false });
          setBackstageAction('info');
          setRibbonTab('home');
        },
      });
    },
    [instance, onNewWorkbook, onSave, onSaveAs, showRibbonNotice, ui],
  );

  const onFormatterChange = useCallback((key: 'uppercase' | 'arrows', checked: boolean) => {
    setFormatters((f) => ({ ...f, [key]: checked }));
  }, []);

  const onPresetChange = useCallback(
    (next: PresetKey) => {
      if (next === preset) return;
      setPreset(next);
      setOverrides({});
    },
    [preset],
  );

  const onFeatureToggle = useCallback(
    (id: FeatureId) => {
      setOverrides(nextDemoFeatureOverrides({ preset, overrides, features, id }));
    },
    [features, overrides, preset],
  );

  const commands = useMemo(
    () =>
      buildDemoCommands({
        commandText,
        instance,
        openWorkbook: () => fileInputRef.current?.click(),
        saveWorkbook: onSave,
        setRibbonTab,
        togglePanel: () => setShowPanel((v) => !v),
        setTheme,
        setLocale: (next) => {
          const resolved = normalizeDemoLocale(next);
          if (resolved) setLocale(resolved);
        },
      }),
    [commandText, instance, onSave],
  );
  const applyToolbarCommand = useCallback((commandId: string): boolean => {
    const toolbar = toolbarRef.current;
    if (!toolbar) return false;
    try {
      toolbar.execute(commandId);
      return true;
    } catch {
      // A refused command did not run; the caller falls back to showing its tab.
      return false;
    }
  }, []);
  const searchItems = useMemo(
    () => buildDemoSearchItems(commands, locale, setRibbonTab, applyToolbarCommand, ribbonTabs),
    [applyToolbarCommand, commands, locale, ribbonTabs],
  );

  const openBackstage = useCallback(() => {
    setRibbonTab('file');
    toolbarRef.current?.update({ backstageOpen: true });
  }, []);

  if (!workbook) {
    return (
      <div className="demo demo--loading">
        {loadError ? (
          <div className="demo__load-error" role="alert">
            <strong>{ui.engineUnavailable}</strong>
            <span>{ui.engineSetup}</span>
            <code>{loadError}</code>
          </div>
        ) : (
          ui.loadingEngine
        )}
      </div>
    );
  }

  return (
    <div className="demo" data-fc-theme={theme} data-fc-platform={platform}>
      <DemoTitleBar
        ui={ui}
        platform={platform}
        bookName={bookName}
        instance={instance}
        showPanel={showPanel}
        searchItems={searchItems}
        quickAccessRef={quickAccessRef}
        onSave={onSave}
        onSaveAs={onSaveAs}
        onRename={setBookName}
        onNewWorkbook={() => void onNewWorkbook()}
        applyCommand={applyToolbarCommand}
        onOpenBackstage={openBackstage}
        onTogglePanel={() => setShowPanel((v) => !v)}
        onTabChange={setRibbonTab}
      />
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xlsm,.xlsb"
        hidden
        onChange={(ev) => {
          const f = ev.target.files?.[0];
          if (f) void onOpen(f);
          ev.target.value = '';
        }}
      />

      <main className={`demo__body${showPanel ? ' demo__body--panel' : ''}`}>
        <div className="fc-tb__sheet-col">
          {resolvedUi.ribbon ? (
            <div ref={ribbonHostRef} style={{ display: 'contents' }}>
              <SpreadsheetToolbar
                spreadsheet={instance}
                activeTab={ribbonTab}
                onTabChange={setRibbonTab}
                ribbonTabs={ribbonTabs}
                hooks={ribbonHooks}
                onReady={(toolbar) => {
                  toolbarRef.current = toolbar;
                  (window as unknown as { __fcToolbar?: Toolbar | null }).__fcToolbar = toolbar;
                }}
              />
            </div>
          ) : null}
          <Spreadsheet
            className="demo__sheet"
            ui={resolvedUi}
            workbook={workbook}
            formulaHints={DEMO_FORMULA_HINTS}
            theme={theme}
            locale={locale}
            host={host}
            uploadStatus={uploadStatus}
            onReady={(next) => {
              instanceRef.current = next;
              setInstance(next);
            }}
            onError={onSpreadsheetError}
            onLocaleChange={(event) => {
              const resolved = normalizeDemoLocale(event.locale);
              if (resolved) setLocale(resolved);
            }}
          />
          {ribbonTab === 'file' ? (
            <DemoBackstage
              ui={ui}
              instance={instance}
              bookName={bookName}
              action={backstageAction}
              onAction={runBackstageAction}
            />
          ) : null}
        </div>
        <DemoOptionsPanel
          ui={ui}
          commandText={commandText}
          spreadsheet={instance}
          hidden={!showPanel}
          theme={theme}
          locale={locale}
          preset={preset}
          features={features}
          ribbon={resolvedUi.ribbon}
          formatters={formatters}
          log={log}
          onThemeChange={setTheme}
          onLocaleChange={(next) => {
            const resolved = normalizeDemoLocale(next);
            if (resolved) setLocale(resolved);
          }}
          onPresetChange={onPresetChange}
          onFeatureToggle={onFeatureToggle}
          onRibbonChange={setShowRibbon}
          onFormatterChange={onFormatterChange}
        />
      </main>
      {reviewDialog ? (
        <DemoReviewDialog dialog={reviewDialog} ui={ui} onClose={closeReviewDialog} />
      ) : null}
    </div>
  );
};
