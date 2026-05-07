import {
  type CellChangeEvent,
  type CellRenderInput,
  type CellValue,
  type FeatureFlags,
  type FeatureId,
  presets,
  type SpreadsheetInstance,
  type ThemeName,
  WorkbookHandle,
} from '@libraz/formulon-cell';
import { Spreadsheet, useSelection } from '@libraz/formulon-cell-react';
import { type ReactElement, useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { type RibbonTab, Toolbar } from './Toolbar';

const THEMES: { value: ThemeName; label: string }[] = [
  { value: 'paper', label: 'Light' },
  { value: 'ink', label: 'Dark' },
  { value: 'contrast', label: 'Contrast' },
];
const LOCALES = [
  { value: 'en', label: 'EN' },
  { value: 'ja', label: 'JA' },
];

const UI = {
  en: {
    saved: 'Saved to this device',
    search: 'Search',
    share: 'Share',
    workbook: 'React workbook',
    demoPane: 'Demo pane',
    open: 'Open xlsx…',
    save: 'Save',
    file: 'File',
    info: 'Info',
    print: 'Print',
    pageSetup: 'Page Setup',
    close: 'Close',
    backstageTitle: 'Book1',
    backstageSub: 'React workbook · Excel 365 layout mode',
    openTitle: 'Open',
    openDesc: 'Load an .xlsx or .xlsm workbook from this device.',
    saveCopy: 'Save a Copy',
    saveDesc: 'Download the current workbook as an .xlsx file.',
    printDesc: 'Use the browser print dialog or save as PDF.',
    pageSetupDesc: 'Set orientation, margins, paper size, headers, and print titles.',
    editLinks: 'Edit Links',
    linksDesc: 'Inspect external workbook references carried by the file.',
    options: 'Options',
    optionsDesc: 'Show the demo integration panel and feature toggles.',
    noCommands: 'No commands found',
  },
  ja: {
    saved: 'このデバイスに保存済み',
    search: '検索',
    share: '共有',
    workbook: 'React ブック',
    demoPane: 'デモ ペイン',
    open: 'xlsx を開く…',
    save: '保存',
    file: 'ファイル',
    info: '情報',
    print: '印刷',
    pageSetup: 'ページ設定',
    close: '閉じる',
    backstageTitle: 'Book1',
    backstageSub: 'React ブック · Excel 365 レイアウト',
    openTitle: '開く',
    openDesc: '.xlsx または .xlsm ブックをこのデバイスから読み込みます。',
    saveCopy: 'コピーを保存',
    saveDesc: '現在のブックを .xlsx ファイルとしてダウンロードします。',
    printDesc: 'ブラウザーの印刷ダイアログ、または PDF 保存を使用します。',
    pageSetupDesc: '用紙方向、余白、用紙サイズ、ヘッダー、印刷タイトルを設定します。',
    editLinks: 'リンクの編集',
    linksDesc: 'ファイルに含まれる外部ブック参照を確認します。',
    options: 'オプション',
    optionsDesc: 'デモ統合パネルと機能トグルを表示します。',
    noCommands: 'コマンドが見つかりません',
  },
} as const;

type PresetKey = 'minimal' | 'standard' | 'excel';
const PRESETS: { value: PresetKey; label: string; hint: string }[] = [
  { value: 'minimal', label: 'Minimal', hint: 'grid + formula bar only' },
  { value: 'standard', label: 'Standard', hint: 'menus, find/replace, painter' },
  { value: 'excel', label: 'Excel', hint: 'full Excel 365 chrome' },
];

// Feature flags grouped semantically for the panel. Order matches the
// columns of `ALL_FEATURE_IDS` in core.
const FEATURE_GROUPS: { title: string; features: { id: FeatureId; label: string }[] }[] = [
  {
    title: 'Chrome',
    features: [
      { id: 'formulaBar', label: 'Formula bar' },
      { id: 'sheetTabs', label: 'Sheet tabs' },
      { id: 'statusBar', label: 'Status bar' },
      { id: 'contextMenu', label: 'Context menu' },
      { id: 'watchWindow', label: 'Watch window' },
    ],
  },
  {
    title: 'Editing',
    features: [
      { id: 'clipboard', label: 'Clipboard' },
      { id: 'pasteSpecial', label: 'Paste special' },
      { id: 'formatPainter', label: 'Format painter' },
      { id: 'autocomplete', label: 'Autocomplete' },
      { id: 'shortcuts', label: 'Shortcuts' },
      { id: 'wheel', label: 'Wheel scroll' },
    ],
  },
  {
    title: 'Dialogs & overlays',
    features: [
      { id: 'findReplace', label: 'Find & replace' },
      { id: 'formatDialog', label: 'Format dialog' },
      { id: 'fxDialog', label: 'Function dialog' },
      { id: 'conditional', label: 'Conditional formatting' },
      { id: 'namedRanges', label: 'Named ranges' },
      { id: 'hyperlink', label: 'Hyperlink' },
      { id: 'validation', label: 'Data validation' },
      { id: 'hoverComment', label: 'Hover comment' },
      { id: 'errorIndicators', label: 'Error indicators' },
    ],
  },
];

const colLabel = (n: number): string => {
  let out = '';
  let v = n;
  do {
    out = String.fromCharCode(65 + (v % 26)) + out;
    v = Math.floor(v / 26) - 1;
  } while (v >= 0);
  return out;
};

const DEMO_FUNCTIONS = [
  {
    name: 'GREET',
    impl: (...args: CellValue[]) => {
      const v = args[0];
      const who = v?.kind === 'text' ? v.value : 'World';
      return `Hello, ${who}!`;
    },
    meta: { description: 'Friendly greeting', args: ['name'], returnType: 'text' as const },
  },
  {
    name: 'FAHRENHEIT',
    impl: (...args: CellValue[]) => {
      const v = args[0];
      const c = v?.kind === 'number' ? v.value : Number.NaN;
      return Number.isFinite(c) ? c * 1.8 + 32 : null;
    },
    meta: {
      description: 'Celsius to Fahrenheit',
      args: ['celsius'],
      returnType: 'number' as const,
    },
  },
];

const FORMATTERS = {
  uppercaseA: {
    id: 'demo:uppercaseA',
    match: (i: CellRenderInput) => i.addr.col === 0 && i.value.kind === 'text',
    format: (i: CellRenderInput) => (i.value.kind === 'text' ? i.value.value.toUpperCase() : null),
  },
  arrowNegatives: {
    id: 'demo:arrowNegatives',
    match: (i: CellRenderInput) => i.value.kind === 'number' && i.value.value < 0,
    format: (i: CellRenderInput) =>
      i.value.kind === 'number' ? `↓ ${Math.abs(i.value.value).toFixed(2)}` : null,
  },
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

let changeId = 0;

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
// way the user can pick "Excel" then individually disable e.g. context
// menu without losing the rest of the preset.
const composeFeatures = (preset: PresetKey, overrides: FeatureFlags): FeatureFlags => ({
  ...presets[preset](),
  ...overrides,
});

export const App = (): ReactElement => {
  const [theme, setTheme] = useState<ThemeName>('paper');
  const [locale, setLocale] = useState<string>('en');
  const [workbook, setWorkbook] = useState<WorkbookHandle | null>(null);
  const [instance, setInstance] = useState<SpreadsheetInstance | null>(null);
  const [log, setLog] = useState<ChangeLogEntry[]>([]);
  const [formatters, setFormatters] = useState({ uppercase: true, arrows: true });
  const [probe, setProbe] = useState<{ name: string; result: string } | null>(null);
  const [preset, setPreset] = useState<PresetKey>('excel');
  const [overrides, setOverrides] = useState<FeatureFlags>({});
  const [showRibbon, setShowRibbon] = useState(true);
  const [showPanel, setShowPanel] = useState(false);
  const [ribbonTab, setRibbonTab] = useState<RibbonTab>('home');
  const [searchQuery, setSearchQuery] = useState('');
  const [searchOpen, setSearchOpen] = useState(false);
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  const features = useMemo(() => composeFeatures(preset, overrides), [preset, overrides]);
  const ui = UI[locale === 'ja' ? 'ja' : 'en'];

  useEffect(() => {
    let alive = true;
    void WorkbookHandle.createDefault().then((wb) => {
      if (!alive) return;
      // Core only auto-seeds when it owns the workbook (no `workbook` prop).
      // The demo passes a pre-built handle, so seed by hand here.
      seed(wb);
      setWorkbook(wb);
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

  const onSave = useCallback(() => {
    if (!instance) return;
    const bytes = instance.workbook.save();
    const blob = new Blob([bytes as BlobPart], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'react-demo.xlsx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1_000);
  }, [instance]);

  const onOpen = useCallback(
    async (file: File) => {
      if (!instance) return;
      const buf = await file.arrayBuffer();
      const next = await WorkbookHandle.loadBytes(new Uint8Array(buf));
      await instance.setWorkbook(next);
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
      const presetDefault =
        id === 'watchWindow' ? presetFlags[id] === true : presetFlags[id] !== false;
      const currentVal = id === 'watchWindow' ? features[id] === true : features[id] !== false;
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
        id: 'demo-pane',
        label: 'Demo Pane',
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
    return <div className="demo demo--loading">Loading engine…</div>;
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
            <strong>Book1</strong>
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
            <Toolbar
              instance={instance}
              activeTab={ribbonTab}
              onTabChange={setRibbonTab}
              locale={locale}
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
                    <h1>Book1</h1>
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
        <aside className="demo__panel" aria-label="Demo panel" hidden={!showPanel}>
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
                    // `watchWindow` ships default-off; everything else is opt-out.
                    const enabled =
                      f.id === 'watchWindow' ? features[f.id] === true : features[f.id] !== false;
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
                      <span>Excel ribbon</span>
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
    </div>
  );
};
