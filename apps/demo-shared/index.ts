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
 * The vanilla `playground` demo uses a different chrome model and does not
 * consume this file — that's intentional.
 */
import type { CellRenderInput, CellValue, FeatureId, ThemeName } from '@libraz/formulon-cell';

export type DemoFramework = 'React' | 'Vue';

export const THEMES: { value: ThemeName; label: string }[] = [
  { value: 'paper', label: 'Light' },
  { value: 'ink', label: 'Dark' },
  { value: 'contrast', label: 'Contrast' },
];

export const LOCALES = [
  { value: 'en', label: 'EN' },
  { value: 'ja', label: 'JA' },
];

export type PresetKey = 'minimal' | 'standard' | 'full';
export const PRESETS: { value: PresetKey; label: string; hint: string }[] = [
  { value: 'minimal', label: 'Minimal', hint: 'bare spreadsheet chrome' },
  { value: 'standard', label: 'Standard', hint: 'lightweight editing chrome' },
  { value: 'full', label: 'Full', hint: 'complete spreadsheet chrome' },
];

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
  share: string;
  workbook: string;
  demoPane: string;
  open: string;
  save: string;
  file: string;
  info: string;
  print: string;
  pageSetup: string;
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
      share: 'Share',
      workbook: `${framework} workbook`,
      demoPane: 'Options',
      open: 'Open xlsx…',
      save: 'Save',
      file: 'File',
      info: 'Info',
      print: 'Print',
      pageSetup: 'Page Setup',
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
      share: '共有',
      workbook: `${ja} ブック`,
      demoPane: 'オプション',
      open: 'xlsx を開く…',
      save: '保存',
      file: 'ファイル',
      info: '情報',
      print: '印刷',
      pageSetup: 'ページ設定',
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
