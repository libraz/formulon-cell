/**
 * Shared configuration for the React and Vue demo apps.
 *
 * The two demos historically maintained their own byte-identical copies of
 * `THEMES`, `LOCALES`, `PRESETS`, `FEATURE_GROUPS`, `DEMO_FUNCTIONS`,
 * `FORMATTERS`, the `formatLoadError` helper, and the `UI` strings table —
 * the only divergence being the framework label ("React" vs "Vue") used in
 * the workbook title and backstage subtitle. That layout invited drift bugs
 * (a translation added to one demo would silently miss the other) and
 * doubled review surface for every UI tweak. This module is the single
 * source of truth so the two demos can stay aligned.
 *
 * Shared helpers keep the React and Vue demos on the same integration line.
 */

import type {
  CellChangeEvent,
  CellRenderInput,
  CellValue,
  FeatureFlags,
  FeatureId,
  ReviewCell,
  SpreadsheetInstance,
  SpreadsheetUiOptions,
  ThemeName,
  WorkbookHandle,
} from '@libraz/formulon-cell';
import { resolveSpreadsheetUiOptions } from '@libraz/formulon-cell';

export type DemoFramework = 'React' | 'Vue';

export const THEMES: { value: ThemeName; label: string }[] = [
  { value: 'paper', label: 'Light' },
  { value: 'ink', label: 'Dark' },
  { value: 'contrast', label: 'Contrast' },
];

export const LOCALES = [
  { value: 'en', label: 'EN' },
  { value: 'ja', label: 'JA' },
] as const;

export type DemoLocale = (typeof LOCALES)[number]['value'];

const isDemoLocale = (value: string | null | undefined): value is DemoLocale =>
  value === 'en' || value === 'ja';

export const resolveInitialLocale = (
  search = globalThis.location?.search ?? '',
  languages = globalThis.navigator?.languages ?? [globalThis.navigator?.language ?? ''],
): DemoLocale => {
  const param = new URLSearchParams(search).get('locale');
  if (isDemoLocale(param)) return param;
  const first = languages.find((lang): lang is string => typeof lang === 'string' && lang !== '');
  if (first?.toLowerCase().startsWith('en')) return 'en';
  return 'ja';
};

export type PresetKey = 'minimal' | 'standard' | 'full';
export const PRESETS: { value: PresetKey; label: string; hint: string }[] = [
  { value: 'minimal', label: 'Minimal', hint: 'bare spreadsheet chrome' },
  { value: 'standard', label: 'Standard', hint: 'lightweight editing chrome' },
  { value: 'full', label: 'Full', hint: 'complete spreadsheet chrome' },
];

export const composeDemoUiOptions = (input: {
  preset: PresetKey;
  overrides: FeatureFlags;
  showRibbon: boolean;
  theme: ThemeName;
}): ReturnType<typeof resolveSpreadsheetUiOptions> => {
  const profile: SpreadsheetUiOptions['profile'] =
    input.preset === 'full' ? 'excel365' : input.preset;
  return resolveSpreadsheetUiOptions({
    profile,
    theme: input.theme,
    features: { ribbon: input.showRibbon },
    advancedFeatures: input.overrides,
  });
};

export const FEATURE_GROUPS: {
  title: string;
  features: { id: FeatureId; label: string }[];
}[] = [
  {
    title: 'Chrome',
    features: [
      { id: 'formulaBar', label: 'Formula bar' },
      { id: 'viewToolbar', label: 'View toolbar' },
      { id: 'sheetTabs', label: 'Sheet tabs' },
      { id: 'statusBar', label: 'Status bar' },
      { id: 'workbookObjects', label: 'Workbook objects' },
      { id: 'contextMenu', label: 'Context menu' },
      { id: 'charts', label: 'Charts' },
      { id: 'watchWindow', label: 'Watch window' },
      { id: 'slicer', label: 'Slicer' },
    ],
  },
  {
    title: 'Editing',
    features: [
      { id: 'clipboard', label: 'Clipboard' },
      { id: 'pasteSpecial', label: 'Paste special' },
      { id: 'quickAnalysis', label: 'Quick Analysis' },
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
      { id: 'gotoSpecial', label: 'Go To Special' },
      { id: 'formatDialog', label: 'Format dialog' },
      { id: 'fxDialog', label: 'Function dialog' },
      { id: 'pageSetup', label: 'Page setup' },
      { id: 'iterative', label: 'Iterative calc' },
      { id: 'conditional', label: 'Conditional formatting' },
      { id: 'namedRanges', label: 'Named ranges' },
      { id: 'hyperlink', label: 'Hyperlink' },
      { id: 'commentDialog', label: 'Comment popover' },
      { id: 'pivotTableDialog', label: 'PivotTable dialog' },
      { id: 'validation', label: 'Data validation' },
      { id: 'hoverComment', label: 'Hover comment' },
      { id: 'errorIndicators', label: 'Error indicators' },
    ],
  },
];

export const formatLoadError = (err: unknown): string =>
  err instanceof Error ? err.message : String(err);

export interface DemoUiStrings {
  saved: string;
  search: string;
  searchCommands: string;
  share: string;
  workbook: string;
  demoPane: string;
  open: string;
  save: string;
  file: string;
  info: string;
  print: string;
  pageSetup: string;
  theme: string;
  locale: string;
  signedInUser: string;
  close: string;
  backstageSub: string;
  openTitle: string;
  openDesc: string;
  saveCopy: string;
  saveDesc: string;
  printDesc: string;
  pageSetupDesc: string;
  editLinks: string;
  linksDesc: string;
  options: string;
  optionsDesc: string;
  noCommands: string;
  engineUnavailable: string;
  engineSetup: string;
}

export interface DemoCommandStrings {
  ribbonCommand: string;
  selection: string;
  workbook: string;
  cellsUpdated: string;
  draw: string;
  translate: string;
  addIns: string;
  openFailed: string;
  script: string;
  scriptCommandError: string;
  accessibilityCheck: string;
  inkNotPersisted: string;
  selectInkFirst: string;
  translationUnavailable: string;
  addInsHostCallbacks: string;
  commands: Record<
    | 'open'
    | 'save'
    | 'pageSetup'
    | 'print'
    | 'formatCells'
    | 'conditionalFormatting'
    | 'cellStyles'
    | 'nameManager'
    | 'insertFunction'
    | 'tracePrecedents'
    | 'watchWindow'
    | 'filter'
    | 'sort'
    | 'freezePanes'
    | 'protectSheet'
    | 'options'
    | 'lightTheme'
    | 'darkTheme'
    | 'japaneseLocale'
    | 'englishLocale',
    { label: string; hint: string }
  >;
}

export interface DemoStrings {
  en: DemoUiStrings;
  ja: DemoUiStrings;
}

const JA_FRAMEWORK: Record<DemoFramework, string> = {
  React: 'React',
  Vue: 'Vue',
};

/** Build the UI string table for a demo. The English/Japanese tables are
 *  identical between React and Vue except for the framework label embedded
 *  in `workbook` and `backstageSub`. */
export function createDemoStrings(framework: DemoFramework): DemoStrings {
  const ja = JA_FRAMEWORK[framework];
  return {
    en: {
      saved: 'Saved to this device',
      search: 'Search',
      searchCommands: 'Search commands',
      share: 'Share',
      workbook: `${framework} workbook`,
      demoPane: 'Options',
      open: 'Open xlsx…',
      save: 'Save',
      file: 'File',
      info: 'Info',
      print: 'Print',
      pageSetup: 'Page Setup',
      theme: 'Theme',
      locale: 'Locale',
      signedInUser: 'Signed in user',
      close: 'Close',
      backstageSub: `${framework} workbook · full spreadsheet layout`,
      openTitle: 'Open',
      openDesc: 'Load an .xlsx or .xlsm workbook from this device.',
      saveCopy: 'Save a Copy',
      saveDesc: 'Download the current workbook as an .xlsx file.',
      printDesc: 'Use the browser print dialog or save as PDF.',
      pageSetupDesc: 'Set orientation, margins, paper size, headers, and print titles.',
      editLinks: 'Edit Links',
      linksDesc: 'Inspect external workbook references carried by the file.',
      options: 'Options',
      optionsDesc: 'Show the integration panel and feature toggles.',
      noCommands: 'No commands found',
      engineUnavailable: 'Spreadsheet engine unavailable',
      engineSetup:
        'Serve this demo with COOP: same-origin and COEP: require-corp so SharedArrayBuffer is available.',
    },
    ja: {
      saved: 'このデバイスに保存済み',
      search: '検索',
      searchCommands: 'コマンドの検索',
      share: '共有',
      workbook: `${ja} ブック`,
      demoPane: 'オプション',
      open: 'xlsx を開く…',
      save: '保存',
      file: 'ファイル',
      info: '情報',
      print: '印刷',
      pageSetup: 'ページ設定',
      theme: 'テーマ',
      locale: '表示言語',
      signedInUser: 'サインイン中のユーザー',
      close: '閉じる',
      backstageSub: `${ja} ブック · スプレッドシート レイアウト`,
      openTitle: '開く',
      openDesc: '.xlsx または .xlsm ブックをこのデバイスから読み込みます。',
      saveCopy: 'コピーを保存',
      saveDesc: '現在のブックを .xlsx ファイルとしてダウンロードします。',
      printDesc: 'ブラウザーの印刷ダイアログ、または PDF 保存を使用します。',
      pageSetupDesc: '用紙方向、余白、用紙サイズ、ヘッダー、印刷タイトルを設定します。',
      editLinks: 'リンクの編集',
      linksDesc: 'ファイルに含まれる外部ブック参照を確認します。',
      options: 'オプション',
      optionsDesc: '統合パネルと機能トグルを表示します。',
      noCommands: 'コマンドが見つかりません',
      engineUnavailable: 'スプレッドシートエンジンを起動できません',
      engineSetup:
        'SharedArrayBuffer を有効にするため、COOP: same-origin と COEP: require-corp 付きで配信してください。',
    },
  };
}

export const demoCommandText = (locale: string): DemoCommandStrings =>
  locale === 'ja'
    ? {
        ribbonCommand: 'リボン コマンド',
        selection: '選択範囲',
        workbook: 'ブック',
        cellsUpdated: '{count} 個のセルを更新しました。',
        draw: '描画',
        translate: '翻訳',
        addIns: 'アドイン',
        openFailed: 'ファイルを開けませんでした',
        script: 'スクリプト',
        scriptCommandError: '次のいずれかを使用してください: uppercase, lowercase, trim, clear.',
        accessibilityCheck: 'アクセシビリティ チェック',
        inkNotPersisted: 'このデモ ブックではインク ストロークは保存されません。',
        selectInkFirst: '消しゴムを使うには、先にインク ストロークを選択してください。',
        translationUnavailable: 'このデモには翻訳サービスが接続されていません。',
        addInsHostCallbacks: 'ここでは Office アドインをホスト コールバックで表しています。',
        commands: {
          open: { label: '開く', hint: 'xlsx または xlsm ブックを開きます' },
          save: { label: '保存', hint: 'ブックを xlsx としてダウンロードします' },
          pageSetup: { label: 'ページ設定', hint: 'ページ設定を開きます' },
          print: { label: '印刷', hint: 'ブラウザーの印刷ダイアログを開きます' },
          formatCells: { label: 'セルの書式設定', hint: '書式設定ダイアログを開きます' },
          conditionalFormatting: {
            label: '条件付き書式',
            hint: '条件付き書式を作成または編集します',
          },
          cellStyles: { label: 'セルのスタイル', hint: 'スタイル ギャラリーを開きます' },
          nameManager: { label: '名前の管理', hint: '名前付き範囲を確認します' },
          insertFunction: { label: '関数の挿入', hint: '関数の引数を開きます' },
          tracePrecedents: { label: '参照元のトレース', hint: '参照元矢印を表示します' },
          watchWindow: { label: 'ウォッチ ウィンドウ', hint: 'ウォッチ ウィンドウを切り替えます' },
          filter: { label: 'フィルター', hint: 'データ タブのフィルター ツールを表示します' },
          sort: { label: '並べ替え', hint: '並べ替えボタンを表示します' },
          freezePanes: { label: 'ウィンドウ枠の固定', hint: 'ウィンドウ枠の固定を表示します' },
          protectSheet: { label: 'シートの保護', hint: '表示タブからシート保護を切り替えます' },
          options: { label: 'オプション', hint: '統合パネルの表示を切り替えます' },
          lightTheme: { label: 'ライト テーマ', hint: 'ブックをライト テーマに切り替えます' },
          darkTheme: { label: 'ダーク テーマ', hint: 'ブックをダーク テーマに切り替えます' },
          japaneseLocale: { label: '日本語表示', hint: 'ラベルを日本語に切り替えます' },
          englishLocale: { label: '英語表示', hint: 'ラベルを英語に切り替えます' },
        },
      }
    : {
        ribbonCommand: 'Ribbon command',
        selection: 'Selection',
        workbook: 'Workbook',
        cellsUpdated: '{count} cells updated.',
        draw: 'Draw',
        translate: 'Translate',
        addIns: 'Add-ins',
        openFailed: 'Open failed',
        script: 'Script',
        scriptCommandError: 'Use one of: uppercase, lowercase, trim, clear.',
        accessibilityCheck: 'Accessibility Check',
        inkNotPersisted: 'Ink strokes are not persisted in this demo workbook.',
        selectInkFirst: 'Select an ink stroke first to use the eraser.',
        translationUnavailable: 'No translation service is connected in this demo.',
        addInsHostCallbacks: 'Office add-ins are represented by host callbacks here.',
        commands: {
          open: { label: 'Open', hint: 'Open an xlsx or xlsm workbook' },
          save: { label: 'Save', hint: 'Download the workbook as xlsx' },
          pageSetup: { label: 'Page Setup', hint: 'Open page setup' },
          print: { label: 'Print', hint: 'Open browser print dialog' },
          formatCells: { label: 'Format Cells', hint: 'Open the format dialog' },
          conditionalFormatting: {
            label: 'Conditional Formatting',
            hint: 'Create or edit conditional formatting',
          },
          cellStyles: { label: 'Cell Styles', hint: 'Open the style gallery' },
          nameManager: { label: 'Name Manager', hint: 'Inspect named ranges' },
          insertFunction: { label: 'Insert Function', hint: 'Open function arguments' },
          tracePrecedents: { label: 'Trace Precedents', hint: 'Show precedent arrows' },
          watchWindow: { label: 'Watch Window', hint: 'Toggle Watch Window' },
          filter: { label: 'Filter', hint: 'Show the Data tab filter tools' },
          sort: { label: 'Sort', hint: 'Show sort buttons' },
          freezePanes: { label: 'Freeze Panes', hint: 'Show Freeze Panes' },
          protectSheet: { label: 'Protect Sheet', hint: 'Toggle sheet protection from View' },
          options: { label: 'Options', hint: 'Show or hide the integration panel' },
          lightTheme: { label: 'Light Theme', hint: 'Switch to light workbook theme' },
          darkTheme: { label: 'Dark Theme', hint: 'Switch to dark workbook theme' },
          japaneseLocale: { label: 'Japanese Locale', hint: 'Switch labels to JA' },
          englishLocale: { label: 'English Locale', hint: 'Switch labels to EN' },
        },
      };

/** Sample user-defined functions registered into both demos so testers can
 *  verify host function injection without re-deriving the boilerplate. */
export const DEMO_FUNCTIONS = [
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

/** Custom cell formatters demonstrating the formatter registry. */
export const FORMATTERS = {
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

// ─── Demo runtime helpers ──────────────────────────────────────────────
// These were previously duplicated byte-for-byte between react-demo and
// vue-demo. Each one is framework-agnostic — the React app wraps the
// modal focus helper with `useEffect`, Vue calls it directly.

export const demoColLabel = (n: number): string => {
  let out = '';
  let v = n;
  do {
    out = String.fromCharCode(65 + (v % 26)) + out;
    v = Math.floor(v / 26) - 1;
  } while (v >= 0);
  return out;
};

const FOCUSABLE_DEMO_MODAL_SELECTOR = [
  'button',
  'input',
  'select',
  'textarea',
  'a[href]',
  '[tabindex]:not([tabindex="-1"])',
].join(',');

const focusableDemoModalItems = (root: HTMLElement): HTMLElement[] =>
  Array.from(root.querySelectorAll<HTMLElement>(FOCUSABLE_DEMO_MODAL_SELECTOR)).filter((el) => {
    if (el.closest('[hidden],[aria-hidden="true"]')) return false;
    if ('disabled' in el && (el as HTMLButtonElement | HTMLInputElement).disabled) return false;
    return el.tabIndex >= 0;
  });

/** Wires Tab/Shift+Tab focus trap + Escape close for a demo modal and
 *  returns a teardown callback. Used directly by Vue; React wraps this
 *  in a `useEffect` to attach/detach with the modal's open state. */
export const activateDemoModal = (root: HTMLElement, onClose: () => void): (() => void) => {
  const restoreFocusEl =
    document.activeElement instanceof HTMLElement ? document.activeElement : null;
  const focusFirst = window.requestAnimationFrame(() => {
    (focusableDemoModalItems(root)[0] ?? root).focus({ preventScroll: true });
  });
  const onKeyDown = (event: KeyboardEvent): void => {
    if (event.key === 'Escape') {
      event.preventDefault();
      onClose();
      return;
    }
    if (event.key !== 'Tab') return;
    const items = focusableDemoModalItems(root);
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
};

/** Formats a cell change event for the demo's change log (favours the
 *  raw formula text when present so users see what they typed). */
export const previewCellChange = (e: CellChangeEvent): string => {
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

/** Demo seed — only runs once on the initial blank workbook. Core gates
 *  `seed` on `ownsWb`, so re-mounts and Open xlsx don't re-trigger it. */
export const seedDemoWorkbook = (wb: WorkbookHandle): void => {
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

/** Projects a workbook's cells for the demo review dialog. */
export const reviewCellsForInstance = (inst: SpreadsheetInstance): ReviewCell[] => {
  const sheet = inst.store.getState().data.sheetIndex;
  return Array.from(inst.workbook.cells(sheet), (entry) => ({
    label: `${demoColLabel(entry.addr.col)}${entry.addr.row + 1}`,
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
