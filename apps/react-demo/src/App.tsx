import {
  type CellChangeEvent,
  type CellValue,
  type FeatureFlags,
  type FeatureId,
  type parseScriptCommand,
  type ScriptCommand,
  type SpreadsheetInstance,
  type ThemeName,
  type ToolbarInstance,
  WorkbookHandle,
} from '@libraz/formulon-cell';
import {
  type RibbonTab,
  Spreadsheet,
  SpreadsheetToolbar,
  useSelection,
} from '@libraz/formulon-cell-react';
import { type ReactElement, useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  buildDemoBackstageCards,
  buildDemoBackstageNav,
  buildDemoCommands,
  buildDemoPrintPreviewModel,
  buildDemoReviewDialog,
  buildDemoReviewFindings,
  buildDemoSearchItems,
  type ChangeLogEntry,
  composeDemoUiOptions,
  createDemoStrings,
  createInitialDemoWorkbook,
  DEMO_FUNCTIONS,
  DEMO_MAC_RIBBON_TABS,
  DEMO_PRINT_PREVIEW_LINES,
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
  demoSelectionLabel,
  evaluateDemoProbe,
  FEATURE_GROUPS,
  FORMATTERS,
  formatLoadError,
  installDemoF6Navigation,
  installDemoScriptMenu,
  installDemoSearchShortcut,
  isDemoBackstageActionDisabled,
  isDemoFeatureOn,
  LOCALES,
  loadDemoSearchUsagePrior,
  nextDemoFeatureOverrides,
  openDemoWorkbookFile,
  PRESETS,
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
  THEMES,
} from '../../demo-shared/index.js';
import { DemoIcon } from './DemoIcon.js';
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
  const [probe, setProbe] = useState<{ name: string; result: string } | null>(null);
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

  const selection = useSelection(instance);
  const selectionLabel = useMemo(() => demoSelectionLabel(selection), [selection]);

  const runProbe = useCallback(
    (name: string, args: CellValue[]) => {
      if (!instance) return;
      setProbe(evaluateDemoProbe(instance, name, args));
    },
    [instance],
  );

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

  const backstageNav = useMemo(
    () => buildDemoBackstageNav(ui, backstageAction),
    [backstageAction, ui],
  );
  const backstageCards = useMemo(() => buildDemoBackstageCards(ui), [ui]);
  const printPreview = useMemo(() => {
    void backstageAction;
    return buildDemoPrintPreviewModel(ui, instance, bookName);
  }, [backstageAction, bookName, instance, ui]);
  const backstageActionDisabled = useCallback(
    (action: DemoBackstageAction): boolean => isDemoBackstageActionDisabled(action, instance),
    [instance],
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
            <div className="fc-tb__backstage" role="dialog" aria-label={ui.file}>
              <nav className="fc-tb__backstage-nav" aria-label={ui.file}>
                <strong>{ui.file}</strong>
                {backstageNav.map((item) => (
                  <button
                    key={item.action}
                    type="button"
                    className={`fc-tb__backstage-navitem${
                      item.active ? ' fc-tb__backstage-navitem--active' : ''
                    }`}
                    onClick={() => runBackstageAction(item.action)}
                    disabled={backstageActionDisabled(item.action)}
                  >
                    {item.label}
                  </button>
                ))}
              </nav>
              <div className="fc-tb__backstage-main">
                <div className="fc-tb__backstage-title">
                  <span className="fc-tb__backstage-xl" aria-hidden="true">
                    <DemoIcon name="app" />
                  </span>
                  <div>
                    <h1>{bookName}</h1>
                    <p>{ui.backstageSub}</p>
                  </div>
                </div>
                {backstageAction === 'print' ? (
                  <div className="fc-tb__print-preview" data-demo-print-preview>
                    <section className="fc-tb__print-settings" aria-label={ui.printSettings}>
                      <h2>{printPreview.title}</h2>
                      <p>{printPreview.subtitle}</p>
                      <button
                        type="button"
                        className="fc-tb__print-action fc-tb__print-action--primary"
                        onClick={() => instance?.print('print')}
                        disabled={!instance}
                      >
                        {printPreview.printLabel}
                      </button>
                      <button
                        type="button"
                        className="fc-tb__print-action"
                        onClick={() => instance?.print('pdf')}
                        disabled={!instance}
                      >
                        {printPreview.pdfLabel}
                      </button>
                      <button
                        type="button"
                        className="fc-tb__print-action"
                        onClick={() => instance?.openPageSetup()}
                        disabled={!instance}
                      >
                        {printPreview.pageSetupLabel}
                      </button>
                      <dl className="fc-tb__print-meta">
                        {printPreview.settings.map((row) => (
                          <div key={row.label}>
                            <dt>{row.label}</dt>
                            <dd>{row.value}</dd>
                          </div>
                        ))}
                      </dl>
                    </section>
                    <section className="fc-tb__print-paper" aria-label={printPreview.previewTitle}>
                      {printPreview.previewHtml ? (
                        <iframe
                          className="fc-tb__print-frame"
                          title={printPreview.previewTitle}
                          sandbox=""
                          srcDoc={printPreview.previewHtml}
                        />
                      ) : (
                        <div className="fc-tb__print-page">
                          <strong>{printPreview.previewTitle}</strong>
                          <div aria-hidden="true" className="fc-tb__print-sheet-lines">
                            {DEMO_PRINT_PREVIEW_LINES.map((line) => (
                              <span key={line} />
                            ))}
                          </div>
                        </div>
                      )}
                      <p>{printPreview.previewHint}</p>
                    </section>
                  </div>
                ) : (
                  <div className="fc-tb__backstage-grid">
                    {backstageCards.map((item) => (
                      <button
                        key={item.action}
                        type="button"
                        className="fc-tb__backstage-card"
                        onClick={() => runBackstageAction(item.action)}
                        disabled={backstageActionDisabled(item.action)}
                      >
                        <strong>{item.label}</strong>
                        <span>{item.desc}</span>
                      </button>
                    ))}
                  </div>
                )}
              </div>
            </div>
          ) : null}
        </div>
        <aside className="demo__panel" aria-label={ui.optionsPanel} hidden={!showPanel}>
          <section className="demo__card">
            <h2>{ui.demoChrome}</h2>
            <div className="demo__controls demo__controls--panel">
              <div className="demo__seg" role="group" aria-label={ui.theme}>
                {THEMES.map((t) => (
                  <button
                    key={t.value}
                    type="button"
                    className={`demo__seg-btn${t.value === theme ? ' demo__seg-btn--active' : ''}`}
                    onClick={() => setTheme(t.value)}
                    aria-pressed={t.value === theme}
                  >
                    {ui.themeLabels[t.value] ?? t.label}
                  </button>
                ))}
              </div>
              <div className="demo__seg" role="group" aria-label={ui.locale}>
                {LOCALES.map((l) => (
                  <button
                    key={l.value}
                    type="button"
                    className={`demo__seg-btn${l.value === locale ? ' demo__seg-btn--active' : ''}`}
                    onClick={() => setLocale(l.value)}
                    aria-pressed={l.value === locale}
                  >
                    {l.label}
                  </button>
                ))}
              </div>
            </div>
          </section>

          <section className="demo__card">
            <h2>{ui.preset}</h2>
            <p className="demo__hint">{ui.presetHint}</p>
            <div className="demo__preset">
              {PRESETS.map((p) => (
                <button
                  key={p.value}
                  type="button"
                  className={`demo__preset-btn${
                    p.value === preset ? ' demo__preset-btn--active' : ''
                  }`}
                  onClick={() => onPresetChange(p.value)}
                  aria-pressed={p.value === preset}
                >
                  <span className="demo__preset-name">{ui.presets[p.value]?.label ?? p.label}</span>
                  <span className="demo__preset-hint">{ui.presets[p.value]?.hint ?? p.hint}</span>
                </button>
              ))}
            </div>
          </section>

          <section className="demo__card">
            <h2>{ui.features}</h2>
            <p className="demo__hint">{ui.featuresHint}</p>
            {FEATURE_GROUPS.map((group) => (
              <div key={group.title} className="demo__feat-group">
                <h3 className="demo__feat-title">
                  {ui.featureGroupLabels[group.title] ?? group.title}
                </h3>
                <div className="demo__feat-grid">
                  {group.features.map((f) => {
                    const enabled = isDemoFeatureOn(features, f.id);
                    return (
                      <label key={f.id} className={`demo__feat${enabled ? ' demo__feat--on' : ''}`}>
                        <input
                          type="checkbox"
                          checked={enabled}
                          onChange={() => onFeatureToggle(f.id)}
                        />
                        <span>{ui.featureLabels[f.id] ?? f.label}</span>
                      </label>
                    );
                  })}
                  {group.title === 'Chrome' ? (
                    <label className={`demo__feat${resolvedUi.ribbon ? ' demo__feat--on' : ''}`}>
                      <input
                        type="checkbox"
                        checked={resolvedUi.ribbon}
                        onChange={(e) => setShowRibbon(e.target.checked)}
                      />
                      <span>{ui.spreadsheetRibbon}</span>
                    </label>
                  ) : null}
                </div>
              </div>
            ))}
          </section>

          <section className="demo__card">
            <h2>{commandText.selection}</h2>
            <p className="demo__mono">{selectionLabel}</p>
          </section>

          <section className="demo__card">
            <h2>{ui.cellRenderers}</h2>
            <p className="demo__hint">{ui.cellRenderersHint}</p>
            <label className="fc-tb__check">
              <input
                type="checkbox"
                checked={formatters.uppercase}
                onChange={(e) => setFormatters((f) => ({ ...f, uppercase: e.target.checked }))}
              />
              {ui.uppercaseColumnA}
            </label>
            <label className="fc-tb__check">
              <input
                type="checkbox"
                checked={formatters.arrows}
                onChange={(e) => setFormatters((f) => ({ ...f, arrows: e.target.checked }))}
              />
              {ui.arrowPrefixNegatives}
            </label>
          </section>

          <section className="demo__card">
            <h2>{ui.customFunctions}</h2>
            <p className="demo__hint">{ui.customFunctionsHint}</p>
            <div className="demo__probe">
              <button
                type="button"
                className="fc-tb__btn fc-tb__btn--ghost"
                onClick={() => runProbe('GREET', [{ kind: 'text', value: 'Workbook' }])}
                disabled={!instance}
              >
                GREET("Workbook")
              </button>
              <button
                type="button"
                className="fc-tb__btn fc-tb__btn--ghost"
                onClick={() => runProbe('FAHRENHEIT', [{ kind: 'number', value: 100 }])}
                disabled={!instance}
              >
                FAHRENHEIT(100)
              </button>
              {probe ? (
                <p className="demo__probe-out">
                  → <code>{probe.result}</code>
                </p>
              ) : null}
            </div>
          </section>

          <section className="demo__card demo__card--log">
            <h2>{ui.cellChangeLog}</h2>
            <p className="demo__hint">{ui.cellChangeLogHint}</p>
            {log.length === 0 ? (
              <p className="fc-tb__empty">{ui.editCellToSeeEvents}</p>
            ) : (
              <ul className="demo__log">
                {log.map((entry) => (
                  <li key={entry.id}>
                    <span className="demo__log-cell">{entry.cell}</span>
                    <span className="demo__log-arrow">→</span>
                    <span className="demo__mono">{entry.preview}</span>
                  </li>
                ))}
              </ul>
            )}
          </section>
        </aside>
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
