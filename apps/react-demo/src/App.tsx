import {
  type CellChangeEvent,
  type FeatureFlags,
  type FeatureId,
  type parseScriptCommand,
  type ScriptCommand,
  type SpreadsheetInstance,
  type ThemeName,
  type ToolbarInstance,
  WorkbookHandle,
} from '@libraz/formulon-cell';
import { type RibbonTab, Spreadsheet, SpreadsheetToolbar } from '@libraz/formulon-cell-react';
import { type ReactElement, useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  buildDemoCommands,
  buildDemoReviewDialog,
  buildDemoReviewFindings,
  buildDemoSearchItems,
  type ChangeLogEntry,
  composeDemoUiOptions,
  createDemoStrings,
  createInitialDemoWorkbook,
  DEMO_FUNCTIONS,
  DEMO_MAC_RIBBON_TABS,
  DEMO_PRINTER_PROFILE_ID,
  DEMO_PRINTER_PROFILES,
  DEMO_RIBBON_TABS,
  type DemoBackstageAction,
  type DemoPlatform,
  type DemoReviewDialogState,
  demoCommandText,
  demoFunctionArgumentHelp,
  FORMATTERS,
  formatLoadError,
  installDemoF6Navigation,
  installDemoScriptMenu,
  nextDemoFeatureOverrides,
  openDemoWorkbookFile,
  type PresetKey,
  pushDemoChangeLog,
  refreshDemoPrinterProfiles,
  reportDemoScriptRun,
  resolveInitialLocale,
  resolveInitialPlatform,
  runDemoBackstageAction,
  saveDemoWorkbookToDownload,
} from '../../demo-shared/index.js';
import { DemoBackstage } from './DemoBackstage.js';
import { DemoOptionsPanel } from './DemoOptionsPanel.js';
import { DemoReviewDialog } from './DemoReviewDialog.js';
import { DemoScriptDialog } from './DemoScriptDialog.js';
import { DemoTitleBar } from './DemoTitleBar.js';

const UI = createDemoStrings('React');

export const App = (): ReactElement => {
  const [theme, setTheme] = useState<ThemeName>('paper');
  const [locale, setLocale] = useState<string>(() => resolveInitialLocale());
  const [platform] = useState<DemoPlatform>(() => resolveInitialPlatform());
  const [workbook, setWorkbook] = useState<WorkbookHandle | null>(null);
  const [instance, setInstance] = useState<SpreadsheetInstance | null>(null);
  const [log, setLog] = useState<ChangeLogEntry[]>([]);
  const [formatters, setFormatters] = useState({ uppercase: true, arrows: true });
  const [preset, setPreset] = useState<PresetKey>('full');
  const [overrides, setOverrides] = useState<FeatureFlags>({});
  const [showRibbon, setShowRibbon] = useState(true);
  const [showPanel, setShowPanel] = useState(false);
  const [ribbonTab, setRibbonTab] = useState<RibbonTab>('home');
  const [backstageAction, setBackstageAction] = useState<DemoBackstageAction>('info');
  const [loadError, setLoadError] = useState<string | null>(null);
  const [reviewDialog, setReviewDialog] = useState<DemoReviewDialogState | null>(null);
  const [scriptOpen, setScriptOpen] = useState(false);
  const [uploadStatus, setUploadStatus] = useState<'saved' | 'saving' | 'error' | null>(null);
  // Workbook display name. Untitled until the user opens or saves a file —
  // mirrors the spreadsheet titlebar convention. Stripping the extension
  // keeps it tidy in the chrome while preserving the user's filename for
  // re-saves.
  const [bookName, setBookName] = useState('Book1');
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const quickAccessRef = useRef<HTMLDivElement | null>(null);
  const toolbarRef = useRef<ToolbarInstance | null>(null);

  const resolvedUi = useMemo(
    () => composeDemoUiOptions({ preset, overrides, showRibbon, theme, platform }),
    [overrides, platform, preset, showRibbon, theme],
  );
  const ribbonTabs = platform === 'mac' ? DEMO_MAC_RIBBON_TABS : DEMO_RIBBON_TABS;
  const features = resolvedUi.features;
  const ui = UI[locale === 'ja' ? 'ja' : 'en'];
  const commandText = useMemo(() => demoCommandText(locale), [locale]);
  const closeReviewDialog = useCallback(() => setReviewDialog(null), []);
  const closeScriptDialog = useCallback(() => setScriptOpen(false), []);

  useEffect(() => {
    let alive = true;
    void createInitialDemoWorkbook()
      .then((wb) => {
        if (!alive) return;
        setLoadError(null);
        setWorkbook(wb);
      })
      .catch((err: unknown) => {
        if (!alive) return;
        setLoadError(formatLoadError(err));
      });
    return () => {
      alive = false;
    };
  }, []);

  useEffect(() => {
    if (!instance) return undefined;
    const disposers: (() => void)[] = [];
    if (formatters.uppercase) {
      disposers.push(instance.cells.registerFormatter(FORMATTERS.uppercaseA));
    }
    if (formatters.arrows) {
      disposers.push(instance.cells.registerFormatter(FORMATTERS.arrowNegatives));
    }
    return () => {
      for (const d of disposers) d();
    };
  }, [instance, formatters.uppercase, formatters.arrows]);

  useEffect(() => {
    document.documentElement.lang = locale === 'ja' ? 'ja' : 'en';
  }, [locale]);

  useEffect(
    () =>
      installDemoF6Navigation({
        getQuickAccess: () => quickAccessRef.current,
        getToolbar: () => toolbarRef.current,
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

  const onCellChange = useCallback((e: CellChangeEvent) => {
    setLog((prev) => pushDemoChangeLog(prev, e));
  }, []);

  const onSpellingReview = useCallback(() => {
    if (!instance) return;
    setReviewDialog(buildDemoReviewFindings('spelling', instance, locale, commandText));
  }, [commandText, instance, locale]);

  const onAccessibilityCheck = useCallback(() => {
    if (!instance) return;
    setReviewDialog(buildDemoReviewFindings('accessibility', instance, locale, commandText));
  }, [commandText, instance, locale]);

  const onRunScript = useCallback(() => {
    if (!instance) return;
    setScriptOpen(true);
  }, [instance]);

  const showRibbonNotice = useCallback(
    (title: string, detail: string) => {
      setReviewDialog(buildDemoReviewDialog(title, commandText.ribbonCommand, detail));
    },
    [commandText.ribbonCommand],
  );

  const applyParsedScript = useCallback(
    (command: ReturnType<typeof parseScriptCommand>) => {
      if (!instance || !command) return;
      setReviewDialog(reportDemoScriptRun(instance, command, commandText));
    },
    [commandText, instance],
  );

  const onScriptSubmit = useCallback(
    (command: ScriptCommand) => {
      setScriptOpen(false);
      applyParsedScript(command);
    },
    [applyParsedScript],
  );

  // Runs the built-in script commands from `#menu-script`. The toolbar owns the
  // rest of that click: it closes the menu, moves focus back to the Script
  // command, and routes `custom` to the `onRunScript` prop — so this handler
  // neither reopens that dialog nor touches focus, which would pull focus back
  // out of the dialog the action just opened.
  useEffect(() => installDemoScriptMenu(applyParsedScript), [applyParsedScript]);

  const onSave = useCallback(() => {
    saveDemoWorkbookToDownload({ instance, bookName, setUploadStatus });
  }, [bookName, instance]);

  const onNewWorkbook = useCallback(async () => {
    const wb = await WorkbookHandle.createDefault();
    setWorkbook(wb);
    await instance?.setWorkbook(wb);
    setBookName('Book1');
    setLog([]);
    setRibbonTab('home');
  }, [instance]);

  const onOpen = useCallback(
    async (file: File) => {
      if (!instance) return;
      try {
        setBookName(await openDemoWorkbookFile(instance, file));
        setLoadError(null);
      } catch (err) {
        setReviewDialog(
          buildDemoReviewDialog(commandText.openFailed, commandText.workbook, formatLoadError(err)),
        );
      }
    },
    [commandText.openFailed, commandText.workbook, instance],
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
        showNotice: showRibbonNotice,
        toggleOptions: () => setShowPanel((v) => !v),
        closeBackstage: () => {
          toolbarRef.current?.setBackstageOpen(false);
          setBackstageAction('info');
          setRibbonTab('home');
        },
      });
    },
    [instance, onNewWorkbook, onSave, showRibbonNotice, ui],
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
        setLocale,
      }),
    [commandText, instance, onSave],
  );
  const searchItems = useMemo(
    () =>
      buildDemoSearchItems(
        commands,
        locale,
        setRibbonTab,
        (commandId) => toolbarRef.current?.applyCommand(commandId) ?? false,
        ribbonTabs,
      ),
    [commands, locale, ribbonTabs],
  );

  const openBackstage = useCallback(() => {
    setRibbonTab('file');
    toolbarRef.current?.setBackstageOpen(true);
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
        locale={locale}
        bookName={bookName}
        instance={instance}
        showPanel={showPanel}
        searchItems={searchItems}
        quickAccessRef={quickAccessRef}
        onSave={onSave}
        onOpenBackstage={openBackstage}
        onTogglePanel={() => setShowPanel((v) => !v)}
        onTabChange={setRibbonTab}
      />
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xlsm"
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
            <SpreadsheetToolbar
              instance={instance}
              activeTab={ribbonTab}
              onTabChange={setRibbonTab}
              locale={locale}
              ribbonTabs={ribbonTabs}
              onSpellingReview={onSpellingReview}
              onAccessibilityCheck={onAccessibilityCheck}
              onRunScript={onRunScript}
              onDrawPen={() => showRibbonNotice(commandText.draw, commandText.inkNotPersisted)}
              onDrawEraser={() => showRibbonNotice(commandText.draw, commandText.selectInkFirst)}
              onTranslate={() =>
                showRibbonNotice(commandText.translate, commandText.translationUnavailable)
              }
              onAddIn={() => showRibbonNotice(commandText.addIns, commandText.addInsHostCallbacks)}
              onToolbarReady={(toolbar) => {
                if (toolbar) toolbar.host.dataset.fcPlatform = platform;
                toolbarRef.current = toolbar;
                (window as unknown as { __fcToolbar?: ToolbarInstance | null }).__fcToolbar =
                  toolbar;
              }}
            />
          ) : null}
          <Spreadsheet
            className="demo__sheet"
            ui={resolvedUi}
            toolbar={false}
            workbook={workbook}
            theme={theme}
            locale={locale}
            features={features}
            functions={DEMO_FUNCTIONS}
            printerProfiles={DEMO_PRINTER_PROFILES}
            printerProfileId={DEMO_PRINTER_PROFILE_ID}
            refreshPrinterProfiles={refreshDemoPrinterProfiles}
            getFunctionArgumentHelp={demoFunctionArgumentHelp}
            uploadStatus={uploadStatus}
            macroRecording={scriptOpen}
            onReady={setInstance}
            onCellChange={onCellChange}
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
          instance={instance}
          hidden={!showPanel}
          theme={theme}
          locale={locale}
          preset={preset}
          features={features}
          ribbon={resolvedUi.ribbon}
          formatters={formatters}
          log={log}
          onThemeChange={setTheme}
          onLocaleChange={setLocale}
          onPresetChange={onPresetChange}
          onFeatureToggle={onFeatureToggle}
          onRibbonChange={setShowRibbon}
          onFormatterChange={onFormatterChange}
        />
      </main>
      {reviewDialog ? (
        <DemoReviewDialog dialog={reviewDialog} ui={ui} onClose={closeReviewDialog} />
      ) : null}
      {scriptOpen ? (
        <DemoScriptDialog
          ui={ui}
          commandText={commandText}
          onSubmit={onScriptSubmit}
          onClose={closeScriptDialog}
        />
      ) : null}
    </div>
  );
};
