import type { FeatureId, ThemeName } from '@libraz/formulon-cell';
import type { DemoFramework, PresetKey } from './index.js';

export interface DemoUiStrings {
  saved: string;
  quickAccessToolbar: string;
  search: string;
  searchCommands: string;
  share: string;
  workbook: string;
  demoPane: string;
  demoChrome: string;
  optionsPanel: string;
  open: string;
  save: string;
  undo: string;
  redo: string;
  file: string;
  info: string;
  print: string;
  pageSetup: string;
  theme: string;
  themeLabels: Partial<Record<ThemeName, string>>;
  locale: string;
  signedInUser: string;
  close: string;
  ok: string;
  cancel: string;
  run: string;
  command: string;
  noIssuesFound: string;
  preset: string;
  presetHint: string;
  presets: Record<PresetKey, { label: string; hint: string }>;
  features: string;
  featuresHint: string;
  featureGroupLabels: Record<string, string>;
  featureLabels: Partial<Record<FeatureId, string>>;
  spreadsheetRibbon: string;
  cellRenderers: string;
  cellRenderersHint: string;
  uppercaseColumnA: string;
  arrowPrefixNegatives: string;
  customFunctions: string;
  customFunctionsHint: string;
  cellChangeLog: string;
  cellChangeLogHint: string;
  editCellToSeeEvents: string;
  backstageSub: string;
  newWorkbook: string;
  newWorkbookDesc: string;
  openTitle: string;
  openDesc: string;
  saveCopy: string;
  saveDesc: string;
  saveAsDesc: string;
  printDesc: string;
  printPreviewTitle: string;
  printNow: string;
  printToPdf: string;
  printSettings: string;
  printPreviewSheet: string;
  printPreviewOrientation: string;
  printPreviewOrientPortrait: string;
  printPreviewOrientLandscape: string;
  printPreviewPaper: string;
  printPreviewPrinter: string;
  printPreviewPrinterMargins: string;
  printPreviewMargins: string;
  printPreviewScale: string;
  printPreviewArea: string;
  printPreviewNoArea: string;
  printPreviewPage: string;
  printPreviewHint: string;
  printPreviewUnavailable: string;
  pageSetupDesc: string;
  editLinks: string;
  linksDesc: string;
  export: string;
  exportDesc: string;
  shareDesc: string;
  options: string;
  optionsDesc: string;
  noCommands: string;
  loadingEngine: string;
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
  spellingReview: string;
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

/** Build the UI string table for a demo. The visible workbook chrome is kept
 *  framework-neutral so the React and Vue wrappers share the same Excel-style
 *  surface. */
export function createDemoStrings(_framework: DemoFramework): DemoStrings {
  return {
    en: {
      saved: 'Saved to this device',
      quickAccessToolbar: 'Quick Access Toolbar',
      search: 'Search',
      searchCommands: 'Search commands',
      share: 'Share',
      workbook: 'Workbook',
      demoPane: 'Options',
      demoChrome: 'Demo chrome',
      optionsPanel: 'Options panel',
      open: 'Open xlsx…',
      save: 'Save',
      undo: 'Undo',
      redo: 'Redo',
      file: 'File',
      info: 'Info',
      print: 'Print',
      pageSetup: 'Page Setup',
      theme: 'Theme',
      themeLabels: {
        paper: 'Light',
        ink: 'Dark',
        contrast: 'Contrast',
      },
      locale: 'Locale',
      signedInUser: 'Signed in user',
      close: 'Close',
      ok: 'OK',
      cancel: 'Cancel',
      run: 'Run',
      command: 'Command',
      noIssuesFound: 'No issues found.',
      preset: 'Preset',
      presetHint:
        'Toggle entire feature bundles, or override individual flags below. Changes apply live.',
      presets: {
        minimal: { label: 'Minimal', hint: 'bare spreadsheet chrome' },
        standard: { label: 'Standard', hint: 'lightweight editing chrome' },
        full: { label: 'Full', hint: 'complete spreadsheet chrome' },
      },
      features: 'Features',
      featuresHint: 'Live-toggle individual feature flags.',
      featureGroupLabels: {
        Chrome: 'Chrome',
        Editing: 'Editing',
        'Dialogs & overlays': 'Dialogs & overlays',
      },
      featureLabels: {
        formulaBar: 'Formula bar',
        viewToolbar: 'View toolbar',
        sheetTabs: 'Sheet tabs',
        statusBar: 'Status bar',
        workbookObjects: 'Workbook objects',
        contextMenu: 'Context menu',
        charts: 'Charts',
        watchWindow: 'Watch window',
        slicer: 'Slicer',
        clipboard: 'Clipboard',
        pasteSpecial: 'Paste special',
        quickAnalysis: 'Quick Analysis',
        formatPainter: 'Format painter',
        autocomplete: 'Autocomplete',
        shortcuts: 'Shortcuts',
        wheel: 'Wheel scroll',
        findReplace: 'Find & replace',
        gotoSpecial: 'Go To Special',
        formatDialog: 'Format dialog',
        fxDialog: 'Function dialog',
        pageSetup: 'Page setup',
        iterative: 'Iterative calc',
        conditional: 'Conditional formatting',
        namedRanges: 'Named ranges',
        hyperlink: 'Hyperlink',
        commentDialog: 'Comment popover',
        pivotTableDialog: 'PivotTable dialog',
        validation: 'Data validation',
        hoverComment: 'Hover comment',
        errorIndicators: 'Error indicators',
      },
      spreadsheetRibbon: 'Spreadsheet ribbon',
      cellRenderers: 'Cell renderers',
      cellRenderersHint: 'Wired through the host formatter registry.',
      uppercaseColumnA: 'Uppercase column A',
      arrowPrefixNegatives: 'Arrow-prefix negatives',
      customFunctions: 'Custom functions',
      customFunctionsHint: 'Probe the host-side function registry directly.',
      cellChangeLog: 'Cell change log',
      cellChangeLogHint: 'Mirrors cell edits into the demo log.',
      editCellToSeeEvents: 'Edit a cell to see events stream in.',
      backstageSub: 'Workbook · spreadsheet layout',
      newWorkbook: 'New',
      newWorkbookDesc: 'Start from a blank workbook in this demo session.',
      openTitle: 'Open',
      openDesc: 'Load an .xlsx or .xlsm workbook from this device.',
      saveCopy: 'Save As',
      saveDesc: 'Download the current workbook as an .xlsx file.',
      saveAsDesc: 'Download a separate copy with the current workbook name.',
      printDesc: 'Use the browser print dialog or save as PDF.',
      printPreviewTitle: 'Print',
      printNow: 'Print',
      printToPdf: 'Export to PDF',
      printSettings: 'Settings',
      printPreviewSheet: 'Active sheet',
      printPreviewOrientation: 'Orientation',
      printPreviewOrientPortrait: 'Portrait Orientation',
      printPreviewOrientLandscape: 'Landscape Orientation',
      printPreviewPaper: 'Paper size',
      printPreviewPrinter: 'Printer',
      printPreviewPrinterMargins: 'Minimum margins',
      printPreviewMargins: 'Margins',
      printPreviewScale: 'Scaling',
      printPreviewArea: 'Print area',
      printPreviewNoArea: 'No print area set',
      printPreviewPage: 'Page',
      printPreviewHint: 'Preview reflects the active sheet page setup.',
      printPreviewUnavailable: 'Open a workbook to preview print settings.',
      pageSetupDesc: 'Set orientation, margins, paper size, headers, and print titles.',
      editLinks: 'Edit Links',
      linksDesc: 'Inspect external workbook references carried by the file.',
      export: 'Export',
      exportDesc: 'Use the browser print flow to export as PDF.',
      shareDesc: 'Show the sharing status for this host-driven workbook.',
      options: 'Options',
      optionsDesc: 'Show the integration panel and feature toggles.',
      noCommands: 'No commands found',
      loadingEngine: 'Loading engine...',
      engineUnavailable: 'Spreadsheet engine unavailable',
      engineSetup: 'Check WebAssembly support and that the engine WASM asset can be loaded.',
    },
    ja: {
      saved: 'このデバイスに保存済み',
      quickAccessToolbar: 'クイック アクセス ツール バー',
      search: '検索',
      searchCommands: 'コマンドの検索',
      share: '共有',
      workbook: 'ブック',
      demoPane: 'オプション',
      demoChrome: 'デモ表示',
      optionsPanel: 'オプション パネル',
      open: 'xlsx を開く…',
      save: '保存',
      undo: '元に戻す',
      redo: 'やり直し',
      file: 'ファイル',
      info: '情報',
      print: '印刷',
      pageSetup: 'ページ設定',
      theme: 'テーマ',
      themeLabels: {
        paper: 'ライト',
        ink: 'ダーク',
        contrast: 'コントラスト',
      },
      locale: '表示言語',
      signedInUser: 'サインイン中のユーザー',
      close: '閉じる',
      ok: 'OK',
      cancel: 'キャンセル',
      run: '実行',
      command: 'コマンド',
      noIssuesFound: '問題は見つかりませんでした。',
      preset: 'プリセット',
      presetHint:
        '機能セット全体を切り替えるか、下の個別フラグで上書きします。変更はすぐに反映されます。',
      presets: {
        minimal: { label: '最小', hint: '最小限のスプレッドシート表示' },
        standard: { label: '標準', hint: '軽量な編集用表示' },
        full: { label: 'フル', hint: '完全なスプレッドシート表示' },
      },
      features: '機能',
      featuresHint: '個別の機能フラグをライブ切り替えします。',
      featureGroupLabels: {
        Chrome: '表示',
        Editing: '編集',
        'Dialogs & overlays': 'ダイアログとオーバーレイ',
      },
      featureLabels: {
        formulaBar: '数式バー',
        viewToolbar: '表示ツール バー',
        sheetTabs: 'シート タブ',
        statusBar: 'ステータス バー',
        workbookObjects: 'ブック オブジェクト',
        contextMenu: 'コンテキスト メニュー',
        charts: 'グラフ',
        watchWindow: 'ウォッチ ウィンドウ',
        slicer: 'スライサー',
        clipboard: 'クリップボード',
        pasteSpecial: '形式を選択して貼り付け',
        quickAnalysis: 'クイック分析',
        formatPainter: '書式のコピー/貼り付け',
        autocomplete: 'オートコンプリート',
        shortcuts: 'ショートカット',
        wheel: 'ホイール スクロール',
        findReplace: '検索と置換',
        gotoSpecial: 'ジャンプ',
        formatDialog: 'セルの書式設定',
        fxDialog: '関数ダイアログ',
        pageSetup: 'ページ設定',
        iterative: '反復計算',
        conditional: '条件付き書式',
        namedRanges: '名前付き範囲',
        hyperlink: 'ハイパーリンク',
        commentDialog: 'コメント ポップアップ',
        pivotTableDialog: 'ピボットテーブル ダイアログ',
        validation: 'データの入力規則',
        hoverComment: 'ホバー コメント',
        errorIndicators: 'エラー インジケーター',
      },
      spreadsheetRibbon: 'スプレッドシート リボン',
      cellRenderers: 'セル レンダラー',
      cellRenderersHint: 'ホスト側のフォーマッター登録を通じて適用されます。',
      uppercaseColumnA: '列 A を大文字にする',
      arrowPrefixNegatives: '負の値に矢印を付ける',
      customFunctions: 'カスタム関数',
      customFunctionsHint: 'ホスト側の関数レジストリを直接確認します。',
      cellChangeLog: 'セル変更ログ',
      cellChangeLogHint: 'セル編集をデモログに反映します。',
      editCellToSeeEvents: 'セルを編集するとイベントが表示されます。',
      backstageSub: 'ブック · スプレッドシート レイアウト',
      newWorkbook: '新規',
      newWorkbookDesc: 'このデモセッションで空のブックを開始します。',
      openTitle: '開く',
      openDesc: '.xlsx または .xlsm ブックをこのデバイスから読み込みます。',
      saveCopy: '名前を付けて保存',
      saveDesc: '現在のブックを .xlsx ファイルとしてダウンロードします。',
      saveAsDesc: '現在のブック名で別コピーをダウンロードします。',
      printDesc: 'ブラウザーの印刷ダイアログ、または PDF 保存を使用します。',
      printPreviewTitle: '印刷',
      printNow: '印刷',
      printToPdf: 'PDF にエクスポート',
      printSettings: '設定',
      printPreviewSheet: 'アクティブ シート',
      printPreviewOrientation: '印刷の向き',
      printPreviewOrientPortrait: '縦方向',
      printPreviewOrientLandscape: '横方向',
      printPreviewPaper: '用紙サイズ',
      printPreviewPrinter: 'プリンター',
      printPreviewPrinterMargins: '最小余白',
      printPreviewMargins: '余白',
      printPreviewScale: '拡大縮小',
      printPreviewArea: '印刷範囲',
      printPreviewNoArea: '印刷範囲なし',
      printPreviewPage: 'ページ',
      printPreviewHint: 'プレビューはアクティブ シートのページ設定を反映します。',
      printPreviewUnavailable: 'ブックを開くと印刷設定をプレビューできます。',
      pageSetupDesc: '用紙方向、余白、用紙サイズ、ヘッダー、印刷タイトルを設定します。',
      editLinks: 'リンクの編集',
      linksDesc: 'ファイルに含まれる外部ブック参照を確認します。',
      export: 'エクスポート',
      exportDesc: 'ブラウザーの印刷フローを使って PDF として出力します。',
      shareDesc: 'このホスト管理ブックの共有状態を表示します。',
      options: 'オプション',
      optionsDesc: '統合パネルと機能トグルを表示します。',
      noCommands: 'コマンドが見つかりません',
      loadingEngine: 'エンジンを読み込んでいます...',
      engineUnavailable: 'スプレッドシートエンジンを起動できません',
      engineSetup:
        'WebAssemblyの対応状況と、エンジンのWASMアセットを読み込めるかを確認してください。',
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
        spellingReview: 'スペル チェック',
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
        spellingReview: 'Spelling Review',
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
