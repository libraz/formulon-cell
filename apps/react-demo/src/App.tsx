import {
  analyzeAccessibilityCells,
  analyzeSpellingCells,
  applyTextScript,
  type CellChangeEvent,
  EXCEL365_STANDARD_RIBBON_TABS,
  type FeatureFlags,
  type FeatureId,
  mutators,
  parseScriptCommand,
  presets,
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
import {
  type ReactElement,
  type RefObject,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import {
  activateDemoModal,
  buildDemoBackstageCards,
  buildDemoBackstageNav,
  buildDemoCommands,
  buildDemoPrintPreviewModel,
  buildDemoSearchItems,
  composeDemoUiOptions,
  createDemoStrings,
  DEMO_FUNCTIONS,
  DEMO_ICONS,
  DEMO_PRINT_PREVIEW_LINES,
  type DemoBackstageAction,
  type DemoIconName,
  type DemoSearchItem,
  type DemoSearchUsagePrior,
  demoColLabel,
  demoCommandText,
  demoSearchOptionId,
  FEATURE_GROUPS,
  FORMATTERS,
  formatLoadError,
  installDemoF6Navigation,
  installDemoSearchShortcut,
  isDemoBackstageActionDisabled,
  LOCALES,
  loadDemoSearchUsagePrior,
  nextDemoSearchIndex,
  PRESETS,
  type PresetKey,
  previewCellChange,
  queryDemoSearchItems,
  recordDemoSearchUsage,
  resolveInitialLocale,
  reviewCellsForInstance,
  runDemoBackstageAction,
  saveDemoSearchUsagePrior,
  seedDemoWorkbook,
  THEMES,
} from '../../demo-shared/index.js';

const UI = createDemoStrings('React');

const colLabel = demoColLabel;

interface ChangeLogEntry {
  readonly id: number;
  readonly cell: string;
  readonly preview: string;
}

interface ReviewDialogState {
  readonly title: string;
  readonly items: readonly { label: string; detail: string }[];
}

let changeId = 0;

// Modal focus trap + Esc-to-close. `activateDemoModal` lives in demo-shared
// and is shared with the Vue demo; this hook adapts it to React's effect
// model by attaching on mount/open and detaching on unmount/close.
const useDemoModalFocus = (
  rootRef: RefObject<HTMLElement | null>,
  open: boolean,
  onClose: () => void,
): void => {
  useEffect(() => {
    if (!open) return;
    const root = rootRef.current;
    if (!root) return;
    return activateDemoModal(root, onClose);
  }, [rootRef, open, onClose]);
};

const previewValue = previewCellChange;
const seed = seedDemoWorkbook;

const DemoIcon = ({ name }: { name: DemoIconName }): ReactElement => (
  <svg
    className="demo__rb-icon"
    viewBox="0 0 20 20"
    fill="none"
    stroke="currentColor"
    strokeWidth="1.45"
    strokeLinecap="round"
    strokeLinejoin="round"
    aria-hidden="true"
  >
    {DEMO_ICONS[name].map((path) => (
      <path key={path} d={path} />
    ))}
  </svg>
);

export const App = (): ReactElement => {
  const [theme, setTheme] = useState<ThemeName>('paper');
  const [locale, setLocale] = useState<string>(() => resolveInitialLocale());
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
  const [reviewDialog, setReviewDialog] = useState<ReviewDialogState | null>(null);
  const [scriptOpen, setScriptOpen] = useState(false);
  const [scriptCommand, setScriptCommand] = useState('uppercase');
  const [scriptError, setScriptError] = useState<string | null>(null);
  // Workbook display name. Untitled until the user opens or saves a file —
  // mirrors the spreadsheet titlebar convention. Stripping the extension
  // keeps it tidy in the chrome while preserving the user's filename for
  // re-saves.
  const [bookName, setBookName] = useState('Book1');
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const searchInputRef = useRef<HTMLInputElement | null>(null);
  const quickAccessRef = useRef<HTMLDivElement | null>(null);
  const toolbarRef = useRef<ToolbarInstance | null>(null);
  const reviewModalRef = useRef<HTMLDivElement | null>(null);
  const scriptModalRef = useRef<HTMLDivElement | null>(null);

  const resolvedUi = useMemo(
    () => composeDemoUiOptions({ preset, overrides, showRibbon, theme }),
    [overrides, preset, showRibbon, theme],
  );
  const features = resolvedUi.features;
  const ui = UI[locale === 'ja' ? 'ja' : 'en'];
  const commandText = useMemo(() => demoCommandText(locale), [locale]);
  const closeReviewDialog = useCallback(() => setReviewDialog(null), []);
  const closeScriptDialog = useCallback(() => setScriptOpen(false), []);

  useDemoModalFocus(reviewModalRef, !!reviewDialog, closeReviewDialog);
  useDemoModalFocus(scriptModalRef, scriptOpen, closeScriptDialog);

  useEffect(() => {
    let alive = true;
    void WorkbookHandle.createDefault()
      .then((wb) => {
        if (!alive) return;
        // Core only auto-seeds when it owns the workbook (no `workbook` prop).
        // The demo passes a pre-built handle, so seed by hand here. `?fixture=empty`
        // (used by E2E specs that need a deterministic blank workbook) skips this.
        const fx = new URLSearchParams(window.location.search).get('fixture');
        if (fx !== 'empty') seed(wb);
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

  useEffect(() => installDemoSearchShortcut(() => searchInputRef.current), []);

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
    const cell = `${colLabel(e.addr.col)}${e.addr.row + 1}`;
    setLog((prev) => [{ id: ++changeId, cell, preview: previewValue(e) }, ...prev].slice(0, 8));
  }, []);

  const selection = useSelection(instance);
  const selectionLabel = useMemo(() => {
    const { active, range } = selection;
    if (range.r0 === range.r1 && range.c0 === range.c1) {
      return `${colLabel(active.col)}${active.row + 1}`;
    }
    const tl = `${colLabel(range.c0)}${range.r0 + 1}`;
    const br = `${colLabel(range.c1)}${range.r1 + 1}`;
    const cells = (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);
    return `${tl}:${br} · ${cells}`;
  }, [selection]);

  const runProbe = useCallback(
    (name: string, args: unknown[]) => {
      if (!instance) return;
      try {
        const out = instance.formula.evaluate(name, args as never);
        const display =
          out.kind === 'number'
            ? out.value.toString()
            : out.kind === 'text'
              ? out.value
              : JSON.stringify(out);
        setProbe({ name, result: display });
      } catch (err) {
        setProbe({ name, result: err instanceof Error ? err.message : String(err) });
      }
    },
    [instance],
  );

  const onSpellingReview = useCallback(() => {
    if (!instance) return;
    setReviewDialog({
      title: 'Spelling Review',
      items: analyzeSpellingCells(reviewCellsForInstance(instance)),
    });
  }, [instance]);

  const onAccessibilityCheck = useCallback(() => {
    if (!instance) return;
    setReviewDialog({
      title: commandText.accessibilityCheck,
      items: analyzeAccessibilityCells(reviewCellsForInstance(instance)),
    });
  }, [commandText.accessibilityCheck, instance]);

  const onRunScript = useCallback(() => {
    if (!instance) return;
    setScriptCommand('uppercase');
    setScriptError(null);
    setScriptOpen(true);
  }, [instance]);

  const showRibbonNotice = useCallback(
    (title: string, detail: string) => {
      setReviewDialog({ title, items: [{ label: commandText.ribbonCommand, detail }] });
    },
    [commandText.ribbonCommand],
  );

  const applyParsedScript = useCallback(
    (command: ReturnType<typeof parseScriptCommand>) => {
      if (!instance || !command) return;
      const range = instance.store.getState().selection.range;
      let changed = 0;
      instance.history.begin();
      try {
        for (let row = range.r0; row <= range.r1; row += 1) {
          for (let col = range.c0; col <= range.c1; col += 1) {
            const addr = { sheet: range.sheet, row, col };
            const value = instance.workbook.getValue(addr);
            if (command === 'clear') {
              if (value.kind !== 'blank' || instance.workbook.cellFormula(addr)) {
                instance.workbook.setBlank(addr);
                changed += 1;
              }
              continue;
            }
            if (value.kind === 'text') {
              const next = applyTextScript(value.value, command);
              if (next !== value.value) {
                instance.workbook.setText(addr, next);
                changed += 1;
              }
            }
          }
        }
      } finally {
        instance.history.end();
      }
      mutators.replaceCells(instance.store, instance.workbook.cells(range.sheet));
      setReviewDialog({
        title: commandText.script,
        items: [
          {
            label: commandText.selection,
            detail: commandText.cellsUpdated.replace('{count}', String(changed)),
          },
        ],
      });
    },
    [commandText.cellsUpdated, commandText.script, commandText.selection, instance],
  );

  const applyScriptCommand = useCallback(() => {
    const command = parseScriptCommand(scriptCommand);
    if (!command) {
      setScriptError(commandText.scriptCommandError);
      return;
    }
    setScriptOpen(false);
    applyParsedScript(command);
  }, [applyParsedScript, commandText.scriptCommandError, scriptCommand]);

  // Wire script-menu items. mountToolbar opens `#menu-script` for the Script
  // split button but doesn't dispatch its `[data-script-action]` children;
  // we handle them here so the demo's dialog/quick actions stay in App scope.
  useEffect(() => {
    const onMenuClick = (e: MouseEvent): void => {
      const target = e.target;
      if (!(target instanceof Element)) return;
      const btn = target.closest<HTMLButtonElement>('[data-script-action]');
      if (!btn) return;
      const menu = btn.closest<HTMLDivElement>('#menu-script');
      if (!menu) return;
      const action = btn.dataset.scriptAction ?? '';
      menu.hidden = true;
      const opener = menu.previousElementSibling;
      if (opener instanceof HTMLElement) {
        opener.setAttribute('aria-expanded', 'false');
        // Focus the ribbon command so `activateDemoModal` captures it as the
        // restore target; without this, Esc would try to refocus the now-
        // hidden menu item and fail.
        opener.focus({ preventScroll: true });
      }
      if (action === 'custom') {
        if (!instance) return;
        setScriptCommand('uppercase');
        setScriptError(null);
        setScriptOpen(true);
        return;
      }
      const command = parseScriptCommand(action);
      if (command) applyParsedScript(command);
    };
    document.addEventListener('click', onMenuClick);
    return () => document.removeEventListener('click', onMenuClick);
  }, [applyParsedScript, instance]);

  const onSave = useCallback(() => {
    if (!instance) return;
    const bytes = instance.workbook.save();
    const blob = new Blob([bytes as BlobPart], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${bookName}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1_000);
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
        const buf = await file.arrayBuffer();
        const next = await WorkbookHandle.loadBytes(new Uint8Array(buf));
        await instance.setWorkbook(next);
        setLoadError(null);
        setBookName(file.name.replace(/\.(xlsx|xlsm)$/i, ''));
      } catch (err) {
        setReviewDialog({
          title: commandText.openFailed,
          items: [{ label: commandText.workbook, detail: formatLoadError(err) }],
        });
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
      // Compute the next override map. If toggling back to the preset's
      // default, drop the override so the preset's value wins.
      const presetFlags = presets[preset]();
      const defaultOff = id === 'watchWindow' || id === 'slicer';
      const presetDefault = defaultOff ? presetFlags[id] === true : presetFlags[id] !== false;
      const currentVal = defaultOff ? features[id] === true : features[id] !== false;
      const nextVal = !currentVal;
      const nextOverrides = { ...overrides };
      if (nextVal === presetDefault) {
        delete nextOverrides[id];
      } else {
        nextOverrides[id] = nextVal;
      }
      setOverrides(nextOverrides);
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
      ),
    [commands, locale],
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
          'Loading engine…'
        )}
      </div>
    );
  }

  return (
    <div className="demo" data-theme={theme}>
      <header className="demo__head">
        <div className="demo__titlebar">
          <div
            ref={quickAccessRef}
            className="demo__quick"
            role="toolbar"
            aria-label="Quick access toolbar"
          >
            <span className="demo__brand-mark" aria-hidden="true">
              <DemoIcon name="app" />
            </span>
            <button type="button" className="demo__title-icon" aria-label="Save" onClick={onSave}>
              <DemoIcon name="save" />
            </button>
            <button
              type="button"
              className="demo__title-icon"
              aria-label="Undo"
              onClick={() => instance?.undo()}
            >
              <DemoIcon name="undo" />
            </button>
            <button
              type="button"
              className="demo__title-icon"
              aria-label="Redo"
              onClick={() => instance?.redo()}
            >
              <DemoIcon name="redo" />
            </button>
          </div>
          <div className="demo__title">
            <strong>{bookName}</strong>
            <span>{ui.saved}</span>
          </div>
          <div className="demo__search">
            <DemoIcon name="search" />
            <input
              ref={searchInputRef}
              type="search"
              role="combobox"
              placeholder={ui.search}
              aria-label={ui.searchCommands}
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
                setSearchQuery(e.currentTarget.value);
                setSearchOpen(true);
                setSearchActiveIndex(-1);
              }}
              onKeyDown={(e) => {
                if (e.key === 'Escape') {
                  setSearchOpen(false);
                  setSearchActiveIndex(-1);
                  e.currentTarget.blur();
                }
                if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
                  e.preventDefault();
                  setSearchOpen(true);
                  setSearchActiveIndex((current) =>
                    nextDemoSearchIndex(
                      current,
                      filteredCommands.length,
                      e.key === 'ArrowDown' ? 'next' : 'previous',
                    ),
                  );
                }
                if (e.key === 'Enter' && filteredCommands.length > 0) {
                  e.preventDefault();
                  const index = nextDemoSearchIndex(
                    searchActiveIndex,
                    filteredCommands.length,
                    'first',
                  );
                  const command = filteredCommands[index];
                  if (command) runCommand(command);
                }
              }}
              onBlur={() => setSearchOpen(false)}
            />
            {searchOpen ? (
              <div id="demo-search-results" className="demo__command-menu" role="listbox">
                {filteredCommands.length === 0 ? (
                  <div className="demo__command-empty">{ui.noCommands}</div>
                ) : (
                  filteredCommands.map((cmd, index) => (
                    <button
                      key={cmd.id}
                      id={demoSearchOptionId(index)}
                      type="button"
                      role="option"
                      aria-selected={index === searchActiveIndex}
                      className={`demo__command-item${
                        index === searchActiveIndex ? ' demo__command-item--active' : ''
                      }`}
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
        <div className="demo__sheet-col">
          {resolvedUi.ribbon ? (
            <SpreadsheetToolbar
              instance={instance}
              activeTab={ribbonTab}
              onTabChange={setRibbonTab}
              locale={locale}
              ribbonTabs={EXCEL365_STANDARD_RIBBON_TABS}
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
                toolbarRef.current = toolbar;
              }}
            />
          ) : null}
          <Spreadsheet
            className="demo__sheet"
            workbook={workbook}
            theme={theme}
            locale={locale}
            features={features}
            functions={DEMO_FUNCTIONS}
            onReady={setInstance}
            onCellChange={onCellChange}
          />
          {ribbonTab === 'file' ? (
            <div className="demo__backstage" role="dialog" aria-label={ui.file}>
              <nav className="demo__backstage-nav" aria-label={ui.file}>
                <strong>{ui.file}</strong>
                {backstageNav.map((item) => (
                  <button
                    key={item.action}
                    type="button"
                    className={`demo__backstage-navitem${
                      item.active ? ' demo__backstage-navitem--active' : ''
                    }`}
                    onClick={() => runBackstageAction(item.action)}
                    disabled={backstageActionDisabled(item.action)}
                  >
                    {item.label}
                  </button>
                ))}
              </nav>
              <div className="demo__backstage-main">
                <div className="demo__backstage-title">
                  <span className="demo__backstage-xl" aria-hidden="true">
                    <DemoIcon name="app" />
                  </span>
                  <div>
                    <h1>{bookName}</h1>
                    <p>{ui.backstageSub}</p>
                  </div>
                </div>
                {backstageAction === 'print' ? (
                  <div className="demo__print-preview" data-demo-print-preview>
                    <section className="demo__print-settings" aria-label={ui.printSettings}>
                      <h2>{printPreview.title}</h2>
                      <p>{printPreview.subtitle}</p>
                      <button
                        type="button"
                        className="demo__print-action demo__print-action--primary"
                        onClick={() => instance?.print('print')}
                        disabled={!instance}
                      >
                        {printPreview.printLabel}
                      </button>
                      <button
                        type="button"
                        className="demo__print-action"
                        onClick={() => instance?.print('pdf')}
                        disabled={!instance}
                      >
                        {printPreview.pdfLabel}
                      </button>
                      <button
                        type="button"
                        className="demo__print-action"
                        onClick={() => instance?.openPageSetup()}
                        disabled={!instance}
                      >
                        {printPreview.pageSetupLabel}
                      </button>
                      <dl className="demo__print-meta">
                        {printPreview.settings.map((row) => (
                          <div key={row.label}>
                            <dt>{row.label}</dt>
                            <dd>{row.value}</dd>
                          </div>
                        ))}
                      </dl>
                    </section>
                    <section className="demo__print-paper" aria-label={printPreview.previewTitle}>
                      {printPreview.previewHtml ? (
                        <iframe
                          className="demo__print-frame"
                          title={printPreview.previewTitle}
                          sandbox=""
                          srcDoc={printPreview.previewHtml}
                        />
                      ) : (
                        <div className="demo__print-page">
                          <strong>{printPreview.previewTitle}</strong>
                          <div aria-hidden="true" className="demo__print-sheet-lines">
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
                  <div className="demo__backstage-grid">
                    {backstageCards.map((item) => (
                      <button
                        key={item.action}
                        type="button"
                        className="demo__backstage-card"
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
        <aside className="demo__panel" aria-label="Options panel" hidden={!showPanel}>
          <section className="demo__card">
            <h2>Demo chrome</h2>
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
                    {t.label}
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
            <h2>Preset</h2>
            <p className="demo__hint">
              Toggle entire feature bundles, or override individual flags below. Changes flow
              through <code>inst.setFeatures()</code> live — edits survive.
            </p>
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
                  <span className="demo__preset-name">{p.label}</span>
                  <span className="demo__preset-hint">{p.hint}</span>
                </button>
              ))}
            </div>
          </section>

          <section className="demo__card">
            <h2>Features</h2>
            <p className="demo__hint">
              Live-toggle individual <code>FeatureFlags</code>. Disabled flags skip their
              <code>attach*</code> in <code>mount.ts</code>.
            </p>
            {FEATURE_GROUPS.map((group) => (
              <div key={group.title} className="demo__feat-group">
                <h3 className="demo__feat-title">{group.title}</h3>
                <div className="demo__feat-grid">
                  {group.features.map((f) => {
                    // `watchWindow` and `slicer` ship default-off; everything else is opt-out.
                    const defaultOff = f.id === 'watchWindow' || f.id === 'slicer';
                    const enabled = defaultOff ? features[f.id] === true : features[f.id] !== false;
                    return (
                      <label key={f.id} className={`demo__feat${enabled ? ' demo__feat--on' : ''}`}>
                        <input
                          type="checkbox"
                          checked={enabled}
                          onChange={() => onFeatureToggle(f.id)}
                        />
                        <span>{f.label}</span>
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
                      <span>Spreadsheet ribbon</span>
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
            <h2>Cell renderers</h2>
            <p className="demo__hint">
              Wired via <code>inst.cells.registerFormatter</code>.
            </p>
            <label className="demo__check">
              <input
                type="checkbox"
                checked={formatters.uppercase}
                onChange={(e) => setFormatters((f) => ({ ...f, uppercase: e.target.checked }))}
              />
              Uppercase column A
            </label>
            <label className="demo__check">
              <input
                type="checkbox"
                checked={formatters.arrows}
                onChange={(e) => setFormatters((f) => ({ ...f, arrows: e.target.checked }))}
              />
              Arrow-prefix negatives
            </label>
          </section>

          <section className="demo__card">
            <h2>Custom functions</h2>
            <p className="demo__hint">
              Registered via the <code>functions</code> prop. Probe the host-side registry directly:
            </p>
            <div className="demo__probe">
              <button
                type="button"
                className="demo__btn demo__btn--ghost"
                onClick={() => runProbe('GREET', [{ kind: 'text', value: 'React' }])}
                disabled={!instance}
              >
                GREET("React")
              </button>
              <button
                type="button"
                className="demo__btn demo__btn--ghost"
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
            <h2>Cell change log</h2>
            <p className="demo__hint">
              Mirrors <code>onCellChange</code> events into React state.
            </p>
            {log.length === 0 ? (
              <p className="demo__empty">Edit a cell to see events stream in.</p>
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
        <div
          ref={reviewModalRef}
          className="demo__modal"
          role="dialog"
          aria-modal="true"
          aria-label={reviewDialog.title}
        >
          <section className="demo__modal-panel">
            <header className="demo__modal-header">
              <h2>{reviewDialog.title}</h2>
              <button
                type="button"
                className="demo__modal-x"
                aria-label="Close"
                onClick={closeReviewDialog}
              >
                ×
              </button>
            </header>
            <div className="demo__modal-body">
              {reviewDialog.items.length === 0 ? (
                <p className="demo__modal-empty">No issues found.</p>
              ) : (
                <ul className="demo__modal-list">
                  {reviewDialog.items.map((item) => (
                    <li key={`${item.label}-${item.detail}`}>
                      <strong>{item.label}</strong>
                      <span>{item.detail}</span>
                    </li>
                  ))}
                </ul>
              )}
            </div>
            <footer className="demo__modal-footer">
              <button type="button" className="demo__btn" onClick={closeReviewDialog}>
                OK
              </button>
            </footer>
          </section>
        </div>
      ) : null}
      {scriptOpen ? (
        <div
          ref={scriptModalRef}
          className="demo__modal"
          role="dialog"
          aria-modal="true"
          aria-label="Script"
        >
          <form
            className="demo__modal-panel demo__modal-panel--narrow"
            onSubmit={(ev) => {
              ev.preventDefault();
              applyScriptCommand();
            }}
          >
            <header className="demo__modal-header">
              <h2>Script</h2>
              <button
                type="button"
                className="demo__modal-x"
                aria-label="Close"
                onClick={closeScriptDialog}
              >
                ×
              </button>
            </header>
            <div className="demo__modal-body">
              <label className="demo__modal-field">
                <span>Command</span>
                <input
                  value={scriptCommand}
                  onChange={(ev) => {
                    setScriptCommand(ev.target.value);
                    setScriptError(null);
                  }}
                />
              </label>
              {scriptError ? <p className="demo__modal-error">{scriptError}</p> : null}
            </div>
            <footer className="demo__modal-footer">
              <button type="button" className="demo__btn" onClick={closeScriptDialog}>
                Cancel
              </button>
              <button type="submit" className="demo__btn demo__btn--active">
                Run
              </button>
            </footer>
          </form>
        </div>
      ) : null}
    </div>
  );
};
