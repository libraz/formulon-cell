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
  type DemoSearchItem,
  type DemoSearchUsagePrior,
  demoCommandText,
  demoFunctionArgumentHelp,
  demoSearchOptionId,
  FORMATTERS,
  formatLoadError,
  installDemoF6Navigation,
  installDemoScriptMenu,
  installDemoSearchShortcut,
  loadDemoSearchUsagePrior,
  nextDemoFeatureOverrides,
  openDemoWorkbookFile,
  type PresetKey,
  pushDemoChangeLog,
  queryDemoSearchItems,
  recordDemoSearchUsage,
  refreshDemoPrinterProfiles,
  reportDemoScriptRun,
  resolveDemoSearchKey,
  resolveInitialLocale,
  resolveInitialPlatform,
  runDemoBackstageAction,
  saveDemoSearchUsagePrior,
  saveDemoWorkbookToDownload,
} from '../../demo-shared/index.js';
import { DemoBackstage } from './DemoBackstage.js';
import { DemoIcon } from './DemoIcon.js';
import { DemoOptionsPanel } from './DemoOptionsPanel.js';
import { DemoReviewDialog } from './DemoReviewDialog.js';
import { DemoScriptDialog } from './DemoScriptDialog.js';

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
  const [searchQuery, setSearchQuery] = useState('');
  const [searchOpen, setSearchOpen] = useState(false);
  const [searchActiveIndex, setSearchActiveIndex] = useState(-1);
  const [searchUsagePrior, setSearchUsagePrior] = useState<DemoSearchUsagePrior>(() =>
    loadDemoSearchUsagePrior(),
  );
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
  const searchInputRef = useRef<HTMLInputElement | null>(null);
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
    instance?.i18n.setLocale(locale);
    document.documentElement.lang = locale === 'ja' ? 'ja' : 'en';
  }, [instance, locale]);

  useEffect(() => installDemoSearchShortcut(() => searchInputRef.current, platform), [platform]);

  useEffect(() => saveDemoSearchUsagePrior(searchUsagePrior), [searchUsagePrior]);
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

  const filteredCommands = useMemo(() => {
    return queryDemoSearchItems(searchItems, searchQuery, 8, searchUsagePrior);
  }, [searchItems, searchQuery, searchUsagePrior]);

  const runCommand = useCallback((cmd: DemoSearchItem) => {
    setSearchUsagePrior((prior) => recordDemoSearchUsage(prior, cmd));
    if (cmd.tab) setRibbonTab(cmd.tab);
    cmd.run();
    setSearchQuery('');
    setSearchOpen(false);
    setSearchActiveIndex(-1);
  }, []);

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
      <header className="demo__head">
        <div className="fc-tb__titlebar">
          <div
            ref={quickAccessRef}
            className="demo__quick"
            role="toolbar"
            aria-label={ui.quickAccessToolbar}
          >
            {platform === 'mac' ? (
              <button
                type="button"
                className="demo__brand-mark"
                aria-label={ui.file}
                title={ui.file}
                onClick={openBackstage}
              >
                <DemoIcon name="app" />
              </button>
            ) : (
              <span className="demo__brand-mark" aria-hidden="true">
                <DemoIcon name="app" />
              </span>
            )}
            <button
              type="button"
              className="demo__title-icon"
              aria-label={ui.save}
              onClick={onSave}
            >
              <DemoIcon name="save" />
            </button>
            <button
              type="button"
              className="demo__title-icon"
              aria-label={ui.undo}
              onClick={() => instance?.undo()}
            >
              <DemoIcon name="undo" />
            </button>
            <button
              type="button"
              className="demo__title-icon"
              aria-label={ui.redo}
              onClick={() => instance?.redo()}
            >
              <DemoIcon name="redo" />
            </button>
          </div>
          <div className="fc-tb__title">
            <strong>{bookName}</strong>
            <span>{ui.saved}</span>
          </div>
          <div className="fc-tb__search">
            <DemoIcon name="search" />
            <input
              ref={searchInputRef}
              type="search"
              role="combobox"
              placeholder={ui.search}
              aria-label={
                platform === 'mac' ? (locale === 'ja' ? '検索' : 'Search') : ui.searchCommands
              }
              aria-controls="demo-search-results"
              aria-expanded={searchOpen}
              aria-activedescendant={
                searchOpen && searchActiveIndex >= 0
                  ? demoSearchOptionId(searchActiveIndex)
                  : undefined
              }
              value={searchQuery}
              onFocus={() => {
                setSearchOpen(true);
                setSearchActiveIndex(-1);
              }}
              onChange={(e) => {
                const input = e.currentTarget;
                setSearchQuery(input.value);
                // Escape clears a search input natively, and that clear lands
                // as a change on an input we just blurred. Only a change the
                // user typed reopens the list.
                setSearchOpen(document.activeElement === input);
                setSearchActiveIndex(-1);
              }}
              onKeyDown={(e) => {
                const action = resolveDemoSearchKey(
                  e.key,
                  searchActiveIndex,
                  filteredCommands.length,
                );
                if (action?.kind === 'close') {
                  setSearchOpen(false);
                  setSearchActiveIndex(-1);
                  e.currentTarget.blur();
                } else if (action?.kind === 'move') {
                  e.preventDefault();
                  setSearchOpen(true);
                  setSearchActiveIndex(action.index);
                } else if (action?.kind === 'run') {
                  e.preventDefault();
                  const command = filteredCommands[action.index];
                  if (command) runCommand(command);
                }
              }}
              onBlur={() => setSearchOpen(false)}
            />
            {searchOpen ? (
              <div id="demo-search-results" className="fc-tb__command-menu" role="listbox">
                {filteredCommands.length === 0 ? (
                  <div className="fc-tb__command-empty">{ui.noCommands}</div>
                ) : (
                  filteredCommands.map((cmd, index) => (
                    <button
                      key={cmd.id}
                      id={demoSearchOptionId(index)}
                      type="button"
                      role="option"
                      aria-selected={index === searchActiveIndex}
                      aria-disabled={cmd.disabled ? 'true' : undefined}
                      data-disabled-reason={cmd.disabledReason}
                      className={`fc-tb__command-item${
                        index === searchActiveIndex ? ' fc-tb__command-item--active' : ''
                      }${cmd.disabled ? ' fc-tb__command-item--disabled' : ''}`}
                      onMouseDown={(e) => e.preventDefault()}
                      onMouseEnter={() => setSearchActiveIndex(index)}
                      onClick={() => runCommand(cmd)}
                    >
                      <strong>{cmd.label}</strong>
                      <span>{cmd.hint}</span>
                    </button>
                  ))
                )}
              </div>
            ) : null}
          </div>
          <div className="demo__account">
            <button type="button" className="demo__share">
              {ui.share}
            </button>
            <button
              type="button"
              className={`demo__share${showPanel ? ' demo__share--active' : ''}`}
              onClick={() => setShowPanel((v) => !v)}
              aria-pressed={showPanel}
            >
              {ui.demoPane}
            </button>
            <span className="demo__avatar" role="img" aria-label={ui.signedInUser}>
              FC
            </span>
          </div>
        </div>
      </header>
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
