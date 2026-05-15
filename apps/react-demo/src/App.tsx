import {
  analyzeAccessibilityCells,
  analyzeSpellingCells,
  applyTextScript,
  type CellChangeEvent,
  type FeatureFlags,
  type FeatureId,
  mutators,
  parseScriptCommand,
  presets,
  type ReviewCell,
  type SpreadsheetInstance,
  type ThemeName,
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
  createDemoStrings,
  DEMO_FUNCTIONS,
  FEATURE_GROUPS,
  FORMATTERS,
  formatLoadError,
  LOCALES,
  PRESETS,
  type PresetKey,
  THEMES,
} from '../../demo-shared/index.js';

const UI = createDemoStrings('React');

const colLabel = (n: number): string => {
  let out = '';
  let v = n;
  do {
    out = String.fromCharCode(65 + (v % 26)) + out;
    v = Math.floor(v / 26) - 1;
  } while (v >= 0);
  return out;
};

interface ChangeLogEntry {
  readonly id: number;
  readonly cell: string;
  readonly preview: string;
}

interface CommandItem {
  readonly id: string;
  readonly label: string;
  readonly hint: string;
  readonly tab?: RibbonTab;
  readonly run: () => void;
}

interface ReviewDialogState {
  readonly title: string;
  readonly items: readonly { label: string; detail: string }[];
}

let changeId = 0;

const FOCUSABLE_MODAL_SELECTOR = [
  'button',
  'input',
  'select',
  'textarea',
  'a[href]',
  '[tabindex]:not([tabindex="-1"])',
].join(',');

const focusableModalItems = (root: HTMLElement): HTMLElement[] =>
  Array.from(root.querySelectorAll<HTMLElement>(FOCUSABLE_MODAL_SELECTOR)).filter((el) => {
    if (el.closest('[hidden],[aria-hidden="true"]')) return false;
    if ('disabled' in el && (el as HTMLButtonElement | HTMLInputElement).disabled) return false;
    return el.tabIndex >= 0;
  });

const useDemoModalFocus = (
  rootRef: RefObject<HTMLElement | null>,
  open: boolean,
  onClose: () => void,
): void => {
  useEffect(() => {
    if (!open) return;
    const root = rootRef.current;
    if (!root) return;
    const restoreFocusEl =
      document.activeElement instanceof HTMLElement ? document.activeElement : null;
    const focusFirst = window.requestAnimationFrame(() => {
      (focusableModalItems(root)[0] ?? root).focus({ preventScroll: true });
    });
    const onKeyDown = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') {
        event.preventDefault();
        onClose();
        return;
      }
      if (event.key !== 'Tab') return;
      const items = focusableModalItems(root);
      if (items.length === 0) {
        event.preventDefault();
        root.focus({ preventScroll: true });
        return;
      }
      const first = items[0];
      const last = items[items.length - 1];
      if (event.shiftKey && document.activeElement === first) {
        event.preventDefault();
        last?.focus({ preventScroll: true });
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault();
        first?.focus({ preventScroll: true });
      }
    };
    root.addEventListener('keydown', onKeyDown);
    return () => {
      window.cancelAnimationFrame(focusFirst);
      root.removeEventListener('keydown', onKeyDown);
      if (
        restoreFocusEl &&
        (root.contains(document.activeElement) || document.activeElement === document.body)
      ) {
        restoreFocusEl.focus({ preventScroll: true });
      }
    };
  }, [rootRef, open, onClose]);
};

const previewValue = (e: CellChangeEvent): string => {
  if (e.formula) return e.formula;
  switch (e.value.kind) {
    case 'number':
      return String(e.value.value);
    case 'text':
      return JSON.stringify(e.value.value);
    case 'bool':
      return String(e.value.value);
    case 'error':
      return `#${e.value.code}`;
    case 'blank':
      return '∅';
    default:
      return '?';
  }
};

// Demo seed — only runs once on the initial blank workbook (core gates
// `seed` on `ownsWb`, so re-mounts and Open xlsx don't re-trigger it).
const seed = (wb: WorkbookHandle): void => {
  wb.setText({ sheet: 0, row: 0, col: 0 }, 'item');
  wb.setText({ sheet: 0, row: 0, col: 1 }, 'celsius');
  wb.setText({ sheet: 0, row: 0, col: 2 }, 'fahrenheit');
  wb.setText({ sheet: 0, row: 0, col: 3 }, 'greeting');
  const rows: [string, number][] = [
    ['London', 8],
    ['Tokyo', 22],
    ['Reykjavík', -3],
    ['Cairo', 31],
  ];
  rows.forEach(([city, c], i) => {
    const r = i + 1;
    wb.setText({ sheet: 0, row: r, col: 0 }, city);
    wb.setNumber({ sheet: 0, row: r, col: 1 }, c);
    wb.setFormula({ sheet: 0, row: r, col: 2 }, `=B${r + 1}*1.8+32`);
    wb.setFormula({ sheet: 0, row: r, col: 3 }, `=A${r + 1}&" ☼"`);
  });
  wb.recalc();
};

// Combine a preset's flags with explicit overrides. Overrides win — that
// way the user can pick "Full" then individually disable e.g. context
// menu without losing the rest of the preset.
const composeFeatures = (preset: PresetKey, overrides: FeatureFlags): FeatureFlags => ({
  ...presets[preset](),
  ...overrides,
});

const reviewCellsForInstance = (inst: SpreadsheetInstance): ReviewCell[] => {
  const sheet = inst.store.getState().data.sheetIndex;
  return Array.from(inst.workbook.cells(sheet), (entry) => ({
    label: `${colLabel(entry.addr.col)}${entry.addr.row + 1}`,
    value:
      entry.value.kind === 'text'
        ? { kind: 'text' as const, value: entry.value.value }
        : entry.value.kind === 'error'
          ? { kind: 'error' as const, text: entry.value.text }
          : entry.value.kind === 'number'
            ? { kind: 'number' as const }
            : entry.value.kind === 'bool'
              ? { kind: 'bool' as const }
              : { kind: 'blank' as const },
    formula: entry.formula,
  }));
};

export const App = (): ReactElement => {
  const [theme, setTheme] = useState<ThemeName>('paper');
  const [locale, setLocale] = useState<string>('en');
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
  const [searchQuery, setSearchQuery] = useState('');
  const [searchOpen, setSearchOpen] = useState(false);
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
  const reviewModalRef = useRef<HTMLDivElement | null>(null);
  const scriptModalRef = useRef<HTMLDivElement | null>(null);

  const features = useMemo(() => composeFeatures(preset, overrides), [preset, overrides]);
  const ui = UI[locale === 'ja' ? 'ja' : 'en'];
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
  }, [instance, locale]);

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
      title: 'Accessibility Check',
      items: analyzeAccessibilityCells(reviewCellsForInstance(instance)),
    });
  }, [instance]);

  const onRunScript = useCallback(() => {
    if (!instance) return;
    setScriptCommand('uppercase');
    setScriptError(null);
    setScriptOpen(true);
  }, [instance]);

  const showRibbonNotice = useCallback((title: string, detail: string) => {
    setReviewDialog({ title, items: [{ label: 'Ribbon command', detail }] });
  }, []);

  const applyScriptCommand = useCallback(() => {
    if (!instance) return;
    const command = parseScriptCommand(scriptCommand);
    if (!command) {
      setScriptError('Use one of: uppercase, lowercase, trim, clear.');
      return;
    }
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
    setScriptOpen(false);
    setReviewDialog({
      title: 'Script',
      items: [{ label: 'Selection', detail: `${changed} cells updated.` }],
    });
  }, [instance, scriptCommand]);

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
          title: 'Open failed',
          items: [{ label: 'Workbook', detail: formatLoadError(err) }],
        });
      }
    },
    [instance],
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

  const commands = useMemo<CommandItem[]>(
    () => [
      {
        id: 'open',
        label: 'Open',
        hint: 'Open an xlsx or xlsm workbook',
        tab: 'file',
        run: () => fileInputRef.current?.click(),
      },
      {
        id: 'save',
        label: 'Save',
        hint: 'Download the workbook as xlsx',
        tab: 'file',
        run: onSave,
      },
      {
        id: 'page-setup',
        label: 'Page Setup',
        hint: 'Open page setup',
        tab: 'file',
        run: () => instance?.openPageSetup(),
      },
      {
        id: 'print',
        label: 'Print',
        hint: 'Open browser print dialog',
        tab: 'file',
        run: () => instance?.print(),
      },
      {
        id: 'format-cells',
        label: 'Format Cells',
        hint: 'Open the format dialog',
        tab: 'home',
        run: () => instance?.openFormatDialog(),
      },
      {
        id: 'conditional',
        label: 'Conditional Formatting',
        hint: 'Create or edit conditional formatting',
        tab: 'insert',
        run: () => instance?.openConditionalDialog(),
      },
      {
        id: 'cell-styles',
        label: 'Cell Styles',
        hint: 'Open the style gallery',
        tab: 'insert',
        run: () => instance?.openCellStylesGallery(),
      },
      {
        id: 'name-manager',
        label: 'Name Manager',
        hint: 'Inspect named ranges',
        tab: 'insert',
        run: () => instance?.openNamedRangeDialog(),
      },
      {
        id: 'insert-function',
        label: 'Insert Function',
        hint: 'Open function arguments',
        tab: 'formulas',
        run: () => instance?.openFunctionArguments(),
      },
      {
        id: 'trace-precedents',
        label: 'Trace Precedents',
        hint: 'Show precedent arrows',
        tab: 'formulas',
        run: () => instance?.tracePrecedents(),
      },
      {
        id: 'watch-window',
        label: 'Watch Window',
        hint: 'Toggle Watch Window',
        tab: 'formulas',
        run: () => instance?.toggleWatchWindow(),
      },
      {
        id: 'filter',
        label: 'Filter',
        hint: 'Show the Data tab filter tools',
        tab: 'data',
        run: () => setRibbonTab('data'),
      },
      {
        id: 'sort',
        label: 'Sort',
        hint: 'Show sort buttons',
        tab: 'data',
        run: () => setRibbonTab('data'),
      },
      {
        id: 'freeze-panes',
        label: 'Freeze Panes',
        hint: 'Show Freeze Panes',
        tab: 'view',
        run: () => setRibbonTab('view'),
      },
      {
        id: 'protect-sheet',
        label: 'Protect Sheet',
        hint: 'Toggle sheet protection from View',
        tab: 'view',
        run: () => instance?.toggleSheetProtection(),
      },
      {
        id: 'options-pane',
        label: 'Options',
        hint: 'Show or hide the integration panel',
        run: () => setShowPanel((v) => !v),
      },
      {
        id: 'theme-light',
        label: 'Light Theme',
        hint: 'Switch to light workbook theme',
        run: () => setTheme('paper'),
      },
      {
        id: 'theme-dark',
        label: 'Dark Theme',
        hint: 'Switch to dark workbook theme',
        run: () => setTheme('ink'),
      },
      {
        id: 'locale-ja',
        label: 'Japanese Locale',
        hint: 'Switch labels to JA',
        run: () => setLocale('ja'),
      },
      {
        id: 'locale-en',
        label: 'English Locale',
        hint: 'Switch labels to EN',
        run: () => setLocale('en'),
      },
    ],
    [instance, onSave],
  );

  const filteredCommands = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    if (!q) return commands.slice(0, 8);
    return commands
      .filter((cmd) => `${cmd.label} ${cmd.hint}`.toLowerCase().includes(q))
      .slice(0, 8);
  }, [commands, searchQuery]);

  const runCommand = useCallback((cmd: CommandItem) => {
    if (cmd.tab) setRibbonTab(cmd.tab);
    cmd.run();
    setSearchQuery('');
    setSearchOpen(false);
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
          <div className="demo__quick" role="toolbar" aria-label="Quick access toolbar">
            <span className="demo__brand-mark">⊞</span>
            <button type="button" className="demo__title-icon" aria-label="Save" onClick={onSave}>
              💾
            </button>
            <button
              type="button"
              className="demo__title-icon"
              aria-label="Undo"
              onClick={() => instance?.undo()}
            >
              ↶
            </button>
            <button
              type="button"
              className="demo__title-icon"
              aria-label="Redo"
              onClick={() => instance?.redo()}
            >
              ↷
            </button>
          </div>
          <div className="demo__title">
            <strong>{bookName}</strong>
            <span>{ui.saved}</span>
          </div>
          <div className="demo__search">
            <span aria-hidden="true">⌕</span>
            <input
              type="search"
              placeholder={ui.search}
              aria-label="Search commands"
              value={searchQuery}
              onFocus={() => setSearchOpen(true)}
              onChange={(e) => {
                setSearchQuery(e.currentTarget.value);
                setSearchOpen(true);
              }}
              onKeyDown={(e) => {
                if (e.key === 'Escape') {
                  setSearchOpen(false);
                  e.currentTarget.blur();
                }
                if (e.key === 'Enter' && filteredCommands[0]) {
                  e.preventDefault();
                  runCommand(filteredCommands[0]);
                }
              }}
              onBlur={() => setSearchOpen(false)}
            />
            {searchOpen ? (
              <div className="demo__command-menu">
                {filteredCommands.length === 0 ? (
                  <div className="demo__command-empty">{ui.noCommands}</div>
                ) : (
                  filteredCommands.map((cmd) => (
                    <button
                      key={cmd.id}
                      type="button"
                      className="demo__command-item"
                      onMouseDown={(e) => e.preventDefault()}
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
            <span className="demo__avatar" role="img" aria-label="Signed in user">
              FC
            </span>
          </div>
        </div>
        <div className="demo__commandbar">
          <div className="demo__brand">
            <strong>formulon-cell</strong>
            <span className="demo__brand-sep">·</span>
            <span className="demo__brand-tag">{ui.workbook}</span>
          </div>
          <div className="demo__controls">
            <div className="demo__seg" role="group" aria-label="Theme">
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
            <div className="demo__seg" role="group" aria-label="Locale">
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
            <button
              type="button"
              className={`demo__btn${showPanel ? ' demo__btn--active' : ''}`}
              onClick={() => setShowPanel((v) => !v)}
              aria-pressed={showPanel}
            >
              {ui.demoPane}
            </button>
            <button
              type="button"
              className="demo__btn"
              onClick={() => fileInputRef.current?.click()}
            >
              {ui.open}
            </button>
            <button type="button" className="demo__btn" onClick={onSave} disabled={!instance}>
              {ui.save}
            </button>
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
          </div>
        </div>
      </header>

      <main className={`demo__body${showPanel ? ' demo__body--panel' : ''}`}>
        <div className="demo__sheet-col">
          {showRibbon ? (
            <SpreadsheetToolbar
              instance={instance}
              activeTab={ribbonTab}
              onTabChange={setRibbonTab}
              locale={locale}
              onSpellingReview={onSpellingReview}
              onAccessibilityCheck={onAccessibilityCheck}
              onRunScript={onRunScript}
              onDrawPen={() =>
                showRibbonNotice('Draw', 'Ink strokes are not persisted in this demo workbook.')
              }
              onDrawEraser={() =>
                showRibbonNotice('Draw', 'Select an ink stroke first to use the eraser.')
              }
              onTranslate={() =>
                showRibbonNotice('Translate', 'No translation service is connected in this demo.')
              }
              onAddIn={() =>
                showRibbonNotice(
                  'Add-ins',
                  'Office add-ins are represented by host callbacks here.',
                )
              }
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
                <button
                  type="button"
                  className="demo__backstage-navitem demo__backstage-navitem--active"
                >
                  {ui.info}
                </button>
                <button
                  type="button"
                  className="demo__backstage-navitem"
                  onClick={() => fileInputRef.current?.click()}
                >
                  {ui.openTitle}
                </button>
                <button type="button" className="demo__backstage-navitem" onClick={onSave}>
                  {ui.save}
                </button>
                <button
                  type="button"
                  className="demo__backstage-navitem"
                  onClick={() => instance?.print()}
                  disabled={!instance}
                >
                  {ui.print}
                </button>
                <button
                  type="button"
                  className="demo__backstage-navitem"
                  onClick={() => instance?.openPageSetup()}
                  disabled={!instance}
                >
                  {ui.pageSetup}
                </button>
                <button
                  type="button"
                  className="demo__backstage-navitem"
                  onClick={() => setRibbonTab('home')}
                >
                  {ui.close}
                </button>
              </nav>
              <div className="demo__backstage-main">
                <div className="demo__backstage-title">
                  <span className="demo__backstage-xl">⊞</span>
                  <div>
                    <h1>{bookName}</h1>
                    <p>{ui.backstageSub}</p>
                  </div>
                </div>
                <div className="demo__backstage-grid">
                  <button
                    type="button"
                    className="demo__backstage-card"
                    onClick={() => fileInputRef.current?.click()}
                  >
                    <strong>{ui.openTitle}</strong>
                    <span>{ui.openDesc}</span>
                  </button>
                  <button type="button" className="demo__backstage-card" onClick={onSave}>
                    <strong>{ui.saveCopy}</strong>
                    <span>{ui.saveDesc}</span>
                  </button>
                  <button
                    type="button"
                    className="demo__backstage-card"
                    onClick={() => instance?.print()}
                    disabled={!instance}
                  >
                    <strong>{ui.print}</strong>
                    <span>{ui.printDesc}</span>
                  </button>
                  <button
                    type="button"
                    className="demo__backstage-card"
                    onClick={() => instance?.openPageSetup()}
                    disabled={!instance}
                  >
                    <strong>{ui.pageSetup}</strong>
                    <span>{ui.pageSetupDesc}</span>
                  </button>
                  <button
                    type="button"
                    className="demo__backstage-card"
                    onClick={() => instance?.openExternalLinksDialog()}
                    disabled={!instance}
                  >
                    <strong>{ui.editLinks}</strong>
                    <span>{ui.linksDesc}</span>
                  </button>
                  <button
                    type="button"
                    className="demo__backstage-card"
                    onClick={() => setShowPanel((v) => !v)}
                  >
                    <strong>{ui.options}</strong>
                    <span>{ui.optionsDesc}</span>
                  </button>
                </div>
              </div>
            </div>
          ) : null}
        </div>
        <aside className="demo__panel" aria-label="Options panel" hidden={!showPanel}>
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
                    <label className={`demo__feat${showRibbon ? ' demo__feat--on' : ''}`}>
                      <input
                        type="checkbox"
                        checked={showRibbon}
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
            <h2>Selection</h2>
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
