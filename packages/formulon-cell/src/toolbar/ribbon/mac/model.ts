import { canExecuteBuiltIn } from '../../../commands/built-in-command-policy.js';
import { cellValueIsFormulaError } from '../../../commands/error-indicators.js';
import {
  allFunctionNames,
  type CatalogFunctionCategory,
  supportedFunctionNames,
} from '../../../commands/function-categories.js';
import { addrKey } from '../../../engine/address.js';
import { getMacInk } from '../../../interact/mac-ink.js';
import type { SpreadsheetInstance } from '../../../mount/types.js';
import type { ExcelRibbonIconName } from '../../excel-ribbon-icons.js';
import { projectDisabledState } from '../../menu-a11y.js';
import { projectActiveState } from '../../ribbon-active-state.js';
import type {
  RibbonCommand,
  RibbonGroupModel,
  RibbonOption,
  RibbonTabModel,
  ToolbarLang,
} from '../../ribbon-model.js';
import { MAC_AUTOMATION_GALLERY, MAC_AUTOMATION_GALLERY_TITLE } from './automation-gallery.js';
import { toolbarLangForLocale } from './locale.js';

/** A leaf rendered in one of the Mac Office-style command menus. */
export interface MacRibbonMenuItem {
  id: string;
  label: string;
  labelJa?: string;
  disabled?: boolean;
  disabledReason?: string;
  icon?: string;
}

type Label = { ja: string; en: string };

const text = (lang: ToolbarLang, label: Label): string => label[lang];

const OFFICE_REQUIRED: Label = {
  ja: 'クラウドサービスへの接続が必要です。',
  en: 'Requires a cloud service connection.',
};
const NOT_IMPLEMENTED: Label = {
  ja: 'この機能は未実装です。',
  en: 'This feature is not implemented yet.',
};
const EMBEDDED_UNSUPPORTED: Label = {
  ja: 'この埋め込み環境では未対応です。',
  en: 'This feature is not available in the embedded workbook.',
};
const WORKBOOK_THEME_UNSUPPORTED: Label = {
  ja: 'ブックのテーマはこの埋め込み環境では変更できません。',
  en: 'Workbook themes cannot be changed in the embedded workbook.',
};
const THREADED_COMMENTS_UNSUPPORTED: Label = {
  ja: 'スレッド形式のコメントは未対応です。セルのメモを使用してください。',
  en: 'Threaded comments are unavailable; use cell notes instead.',
};

const command = (
  id: string,
  label: Label,
  icon?: string,
  options: {
    kind?: RibbonCommand['kind'];
    disabled?: boolean;
    disabledReason?: Label;
    className?: string;
    layout?: RibbonCommand['layout'];
  } = {},
): RibbonCommand => ({
  id,
  title: text('en', label),
  label: text('en', label),
  icon,
  kind: options.kind ?? 'wide',
  layout: options.layout,
  className: options.className,
  disabled: options.disabled,
  disabledReason: options.disabledReason ? text('en', options.disabledReason) : undefined,
});

const group = (title: Label, commands: RibbonCommand[], variant = 'tiles'): RibbonGroupModel => ({
  title: text('en', title),
  commands,
  variant,
});

const menuCommand = (id: string, label: Label, icon?: string): RibbonCommand =>
  command(id, label, icon, { kind: 'wide' });

const compactCommand = (id: string, label: Label, icon?: string): RibbonCommand =>
  command(id, label, icon, { kind: 'wide', layout: 'stacked' });

const iconCommand = (id: string, label: Label, icon?: string): RibbonCommand =>
  command(id, label, icon, { kind: 'button', className: 'fc-tb__rb--mac-icon-only' });

const disabledIconCommand = (
  id: string,
  label: Label,
  icon: string | undefined,
  reason: Label,
): RibbonCommand =>
  command(id, label, icon, {
    kind: 'button',
    className: 'fc-tb__rb--mac-icon-only',
    disabled: true,
    disabledReason: reason,
  });

const disabledCompactCommand = (
  id: string,
  label: Label,
  icon: string | undefined,
  reason: Label,
): RibbonCommand =>
  command(id, label, icon, {
    kind: 'wide',
    layout: 'stacked',
    disabled: true,
    disabledReason: reason,
  });

const checkboxCommand = (id: string, label: Label, icon?: string): RibbonCommand =>
  command(id, label, icon, {
    kind: 'wide',
    layout: 'stacked',
    className: 'fc-tb__rb--mac-checkbox',
  });

const selectCommand = (
  id: string,
  label: Label,
  options: readonly RibbonOption[],
  className = 'fc-tb__rb-select--border',
): RibbonCommand => ({
  id,
  title: label.en,
  label: label.en,
  kind: 'select',
  options,
  className,
});

const disabledCommand = (id: string, label: Label, icon: string | undefined, reason: Label) =>
  command(id, label, icon, { disabled: true, disabledReason: reason, kind: 'wide' });

const macTab = (
  id: RibbonTabModel['id'],
  label: Label,
  groups: RibbonGroupModel[],
): RibbonTabModel => ({
  id,
  label: label.en,
  groups,
});

const labels: Record<string, Label> = {
  home: { ja: 'ホーム', en: 'Home' },
  addins: { ja: 'アドイン', en: 'Add-ins' },
  pivotTable: { ja: 'ピボットテーブル', en: 'PivotTable' },
  recommendedPivotTable: { ja: 'おすすめピボットテーブル', en: 'Recommended PivotTables' },
  table: { ja: 'テーブル', en: 'Table' },
  forms: { ja: 'フォーム', en: 'Forms' },
  pictureFromFile: { ja: '画像から', en: 'Picture from File' },
  pictureFromOnline: { ja: 'オンライン画像', en: 'Online Pictures' },
  photo: { ja: '写真', en: 'Photos' },
  shapes: { ja: '図形', en: 'Shapes' },
  icons: { ja: 'アイコン', en: 'Icons' },
  threeDModel: { ja: '3D モデル', en: '3D Models' },
  smartArt: { ja: 'SmartArt', en: 'SmartArt' },
  screenshot: { ja: 'スクリーンショット', en: 'Screenshot' },
  checkBox: { ja: 'チェック ボックス', en: 'Checkbox' },
  recommendedChart: { ja: 'おすすめグラフ', en: 'Recommended Charts' },
  chartColumn: { ja: '縦棒', en: 'Column' },
  chartBar: { ja: '横棒', en: 'Bar' },
  chartLine: { ja: '折れ線', en: 'Line' },
  chartArea: { ja: '面', en: 'Area' },
  chartPie: { ja: '円', en: 'Pie' },
  chartHierarchy: { ja: '階層', en: 'Hierarchy' },
  chartStatistical: { ja: '統計', en: 'Statistical' },
  chartScatter: { ja: '散布図', en: 'Scatter' },
  chartWaterfall: { ja: 'ウォーターフォール', en: 'Waterfall' },
  chartCombo: { ja: '組み合わせ', en: 'Combo' },
  chartMap: { ja: 'マップ', en: 'Map' },
  pivotChart: { ja: 'ピボットグラフ', en: 'PivotChart' },
  sparkline: { ja: 'スパークライン', en: 'Sparklines' },
  slicer: { ja: 'スライサー', en: 'Slicer' },
  timeline: { ja: 'タイムライン', en: 'Timeline' },
  insert: { ja: '挿入', en: 'Insert' },
  formulas: { ja: '数式', en: 'Formulas' },
  data: { ja: 'データ', en: 'Data' },
  review: { ja: '校閲', en: 'Review' },
  view: { ja: '表示', en: 'View' },
  automate: { ja: '自動化', en: 'Automate' },
  sortFilter: { ja: '並べ替えとフィルター', en: 'Sort & Filter' },
  sheetViews: { ja: 'シート ビュー', en: 'Sheet Views' },
  shapeLine: { ja: '直線', en: 'Line' },
  shapeRectangle: { ja: '四角形', en: 'Rectangle' },
  shapeRoundedRectangle: { ja: '角丸四角形', en: 'Rounded Rectangle' },
  shapeOval: { ja: '楕円', en: 'Oval' },
  shapeTriangle: { ja: '三角形', en: 'Triangle' },
  shapeDiamond: { ja: 'ひし形', en: 'Diamond' },
  shapeArrow: { ja: '矢印', en: 'Arrow' },
  shapeCallout: { ja: '吹き出し', en: 'Callout' },
  gridlines: { ja: '枠線', en: 'Gridlines' },
  headings: { ja: '見出し', en: 'Headings' },
  formulaBar: { ja: '数式バー', en: 'Formula Bar' },
  zeros: { ja: 'ゼロ値', en: 'Zero Values' },
  link: { ja: 'リンク', en: 'Link' },
  newComment: { ja: '新しいコメント', en: 'New Comment' },
  textBox: { ja: 'テキスト ボックス', en: 'Text Box' },
  headerFooter: { ja: 'ヘッダーとフッター', en: 'Header & Footer' },
  wordArt: { ja: 'ワードアート', en: 'WordArt' },
  object: { ja: 'オブジェクト', en: 'Object' },
  equation: { ja: '数式', en: 'Equation' },
  symbol: { ja: '記号と特殊文字', en: 'Symbols' },
  draw: { ja: '描画', en: 'Draw' },
  toggle: { ja: '描画', en: 'Draw' },
  eraser: { ja: '消しゴム', en: 'Eraser' },
  lasso: { ja: 'なげわ選択', en: 'Lasso Select' },
  penBlack: { ja: 'ペン（黒）', en: 'Black Pen' },
  penRed: { ja: 'ペン（赤）', en: 'Red Pen' },
  pencil: { ja: '鉛筆書き（灰）', en: 'Pencil' },
  highlighter: { ja: '蛍光ペン（黄）', en: 'Highlighter' },
  add: { ja: '追加', en: 'Add Pen' },
  trackpad: { ja: 'トラックパッドで描画', en: 'Draw with Trackpad' },
  theme: { ja: 'テーマ', en: 'Themes' },
  color: { ja: '色', en: 'Colors' },
  font: { ja: 'フォント', en: 'Fonts' },
  margins: { ja: '余白', en: 'Margins' },
  orientation: { ja: '印刷の向き', en: 'Orientation' },
  size: { ja: 'サイズ', en: 'Size' },
  printArea: { ja: '印刷範囲', en: 'Print Area' },
  pageBreaks: { ja: '改ページ', en: 'Breaks' },
  background: { ja: '背景', en: 'Background' },
  printTitles: { ja: '印刷タイトル', en: 'Print Titles' },
  pageSetup: { ja: 'ページ設定', en: 'Page Setup' },
  width: { ja: '幅', en: 'Width' },
  height: { ja: '高さ', en: 'Height' },
  scaleWidth: { ja: '幅', en: 'Width' },
  scaleHeight: { ja: '高さ', en: 'Height' },
  showGridlines: { ja: '表示（枠線）', en: 'View Gridlines' },
  printGridlines: { ja: '印刷（枠線）', en: 'Print Gridlines' },
  showHeadings: { ja: '表示（見出し）', en: 'View Headings' },
  printHeadings: { ja: '印刷（見出し）', en: 'Print Headings' },
  insertFunction: { ja: '関数の挿入', en: 'Insert Function' },
  autoSum: { ja: 'オートSUM', en: 'AutoSum' },
  recent: { ja: '最近使ったもの', en: 'Recently Used' },
  financial: { ja: '財務', en: 'Financial' },
  logical: { ja: '論理', en: 'Logical' },
  text: { ja: '文字列操作', en: 'Text' },
  dateTime: { ja: '日付/時刻', en: 'Date & Time' },
  lookup: { ja: '検索/行列', en: 'Lookup & Reference' },
  math: { ja: '数学/三角', en: 'Math & Trig' },
  more: { ja: 'その他の関数', en: 'More Functions' },
  statistical: { ja: '統計', en: 'Statistical' },
  dynamicArray: { ja: '動的配列', en: 'Dynamic Array' },
  allFunctions: { ja: 'すべての関数', en: 'All Functions' },
  python: { ja: 'Python の挿入', en: 'Insert Python' },
  reset: { ja: 'リセット', en: 'Reset' },
  editor: { ja: 'エディター', en: 'Editor' },
  initialize: { ja: '初期化', en: 'Initialize' },
  namesManager: { ja: 'ネームマネージャー', en: 'Name Manager' },
  defineName: { ja: '名前の定義', en: 'Define Name' },
  useInFormula: { ja: '数式で使用', en: 'Use in Formula' },
  createFromSelection: { ja: '選択範囲から作成', en: 'Create from Selection' },
  precedents: { ja: '参照元のトレース', en: 'Trace Precedents' },
  dependents: { ja: '参照先のトレース', en: 'Trace Dependents' },
  removeArrows: { ja: 'トレース矢印の削除', en: 'Remove Arrows' },
  showFormulas: { ja: '数式の表示', en: 'Show Formulas' },
  errorCheck: { ja: 'エラーチェック', en: 'Error Checking' },
  watch: { ja: 'ウォッチウィンドウ', en: 'Watch Window' },
  calcOptions: { ja: '計算方法の設定', en: 'Calculation Options' },
  recalc: { ja: '再計算実行', en: 'Calculate Now' },
  sheetRecalc: { ja: 'シート再計算', en: 'Calculate Sheet' },
  powerQuery: { ja: 'データファイル指定', en: 'Get Data' },
  pictureFromData: { ja: '画像から', en: 'From Picture' },
  refreshAll: { ja: 'すべて更新', en: 'Refresh All' },
  queriesConnections: { ja: 'クエリと接続', en: 'Queries & Connections' },
  properties: { ja: 'プロパティ', en: 'Properties' },
  workbookLinks: { ja: 'ブックのリンク', en: 'Workbook Links' },
  dataTypes: { ja: 'データの種類', en: 'Data Types' },
  sortAsc: { ja: '昇順', en: 'Sort Ascending' },
  sortDesc: { ja: '降順', en: 'Sort Descending' },
  sortCustom: { ja: '並べ替え', en: 'Custom Sort' },
  filter: { ja: 'フィルター', en: 'Filter' },
  clear: { ja: 'クリア', en: 'Clear' },
  reapply: { ja: '再適用', en: 'Reapply' },
  advancedFilter: { ja: '詳細設定', en: 'Advanced' },
  textToColumns: { ja: '区切り位置', en: 'Text to Columns' },
  flashFill: { ja: 'フラッシュフィル', en: 'Flash Fill' },
  removeDuplicates: { ja: '重複削除', en: 'Remove Duplicates' },
  validation: { ja: '入力規則', en: 'Data Validation' },
  consolidate: { ja: '統合', en: 'Consolidate' },
  whatIf: { ja: 'What-If 分析', en: 'What-If Analysis' },
  group: { ja: 'グループ化', en: 'Group' },
  ungroup: { ja: 'グループ解除', en: 'Ungroup' },
  subtotal: { ja: '小計', en: 'Subtotal' },
  showDetail: { ja: '詳細データ表示', en: 'Show Detail' },
  hideDetail: { ja: '詳細を表示しない', en: 'Hide Detail' },
  analysis: { ja: '分析ツール', en: 'Analysis ToolPak' },
  spelling: { ja: 'スペルチェック', en: 'Spelling' },
  thesaurus: { ja: '類義語辞典', en: 'Thesaurus' },
  stats: { ja: 'ブック統計', en: 'Workbook Statistics' },
  accessibility: { ja: 'アクセシビリティチェック', en: 'Check Accessibility' },
  translate: { ja: '翻訳', en: 'Translate' },
  showChanges: { ja: '変更内容を表示', en: 'Show Changes' },
  deleteComment: { ja: '削除', en: 'Delete Comment' },
  previousComment: { ja: '前', en: 'Previous Comment' },
  nextComment: { ja: '次', en: 'Next Comment' },
  showComments: { ja: 'コメントの表示', en: 'Show Comments' },
  notes: { ja: 'メモ', en: 'Notes' },
  protectSheet: { ja: 'シート保護', en: 'Protect Sheet' },
  protectWorkbook: { ja: 'ブック保護', en: 'Protect Workbook' },
  alwaysReadOnly: { ja: '常に読み取り専用で開く', en: 'Always Open Read-Only' },
  hideInk: { ja: 'インクを非表示にする', en: 'Hide Ink' },
  sheetView: { ja: 'シートビュー', en: 'Sheet View' },
  sheetViewSelect: { ja: '現在のビュー', en: 'Current View' },
  sheetViewSave: { ja: 'ビューを保存', en: 'Save View' },
  sheetViewDelete: { ja: 'ビューを削除', en: 'Delete View' },
  standard: { ja: '標準', en: 'Normal' },
  pageBreakPreview: { ja: '改ページプレビュー', en: 'Page Break Preview' },
  pageLayout: { ja: 'ページレイアウト', en: 'Page Layout' },
  customViews: { ja: 'ユーザー設定のビュー', en: 'Custom Views' },
  show: { ja: '表示', en: 'Show' },
  zoom: { ja: 'ズーム', en: 'Zoom' },
  zoomSelection: { ja: '選択範囲に合わせて拡大/縮小', en: 'Zoom to Selection' },
  newWindow: { ja: '新しいウィンドウ', en: 'New Window' },
  arrange: { ja: '整列', en: 'Arrange All' },
  freeze: { ja: 'ウィンドウ枠の固定', en: 'Freeze Panes' },
  split: { ja: '分割', en: 'Split' },
  hide: { ja: '非表示', en: 'Hide' },
  unhide: { ja: '再表示', en: 'Unhide' },
  sideBySide: { ja: '並べて比較', en: 'View Side by Side' },
  syncScroll: { ja: '同時にスクロール', en: 'Synchronous Scrolling' },
  resetWindow: { ja: 'ウィンドウ位置のリセット', en: 'Reset Window Position' },
  switchWindow: { ja: 'ウィンドウの切り替え', en: 'Switch Windows' },
  macros: { ja: 'マクロ', en: 'Macros' },
  newScript: { ja: '新しいスクリプト', en: 'New Script' },
  showScripts: { ja: 'スクリプトの表示', en: 'View Scripts' },
  gallery: MAC_AUTOMATION_GALLERY_TITLE,
};

const t = (key: string): Label => labels[key] ?? { ja: key, en: key };

const GROUP_LABELS: Readonly<Record<string, Label>> = {
  Tables: { ja: 'テーブル', en: 'Tables' },
  Forms: { ja: 'フォーム', en: 'Forms' },
  Pictures: { ja: '画像', en: 'Pictures' },
  Illustrations: { ja: 'イラスト', en: 'Illustrations' },
  Controls: { ja: 'コントロール', en: 'Controls' },
  Charts: { ja: 'グラフ', en: 'Charts' },
  Sparklines: { ja: 'スパークライン', en: 'Sparklines' },
  Filters: { ja: 'フィルター', en: 'Filters' },
  Links: { ja: 'リンク', en: 'Links' },
  Comments: { ja: 'コメント', en: 'Comments' },
  Text: { ja: 'テキスト', en: 'Text' },
  Symbols: { ja: '記号', en: 'Symbols' },
  Tools: { ja: 'ツール', en: 'Tools' },
  Pens: { ja: 'ペン', en: 'Pens' },
  Draw: { ja: '描画', en: 'Draw' },
  Themes: { ja: 'テーマ', en: 'Themes' },
  'Page Setup': { ja: 'ページ設定', en: 'Page Setup' },
  'Scale to Fit': { ja: '拡大縮小印刷', en: 'Scale to Fit' },
  'Sheet Options': { ja: 'シートのオプション', en: 'Sheet Options' },
  'Function Library': { ja: '関数ライブラリ', en: 'Function Library' },
  Python: { ja: 'Python', en: 'Python' },
  'Defined Names': { ja: '定義された名前', en: 'Defined Names' },
  'Formula Auditing': { ja: '数式の監査', en: 'Formula Auditing' },
  Calculation: { ja: '計算方法', en: 'Calculation' },
  'Sheet Views': { ja: 'シート ビュー', en: 'Sheet Views' },
  'Get & Transform Data': { ja: 'データの取得と変換', en: 'Get & Transform Data' },
  'Queries & Connections': { ja: 'クエリと接続', en: 'Queries & Connections' },
  'Data Types': { ja: 'データの種類', en: 'Data Types' },
  'Sort & Filter': { ja: '並べ替えとフィルター', en: 'Sort & Filter' },
  'Data Tools': { ja: 'データ ツール', en: 'Data Tools' },
  Forecast: { ja: '予測', en: 'Forecast' },
  Outline: { ja: 'アウトライン', en: 'Outline' },
  Analysis: { ja: '分析', en: 'Analysis' },
  Proofing: { ja: '文章校正', en: 'Proofing' },
  Accessibility: { ja: 'アクセシビリティ', en: 'Accessibility' },
  Language: { ja: '言語', en: 'Language' },
  Changes: { ja: '変更', en: 'Changes' },
  Options: { ja: 'オプション', en: 'Options' },
  Notes: { ja: 'メモ', en: 'Notes' },
  Protect: { ja: '保護', en: 'Protect' },
  'Workbook Views': { ja: 'ブックの表示', en: 'Workbook Views' },
  Show: { ja: '表示/非表示', en: 'Show' },
  Zoom: { ja: 'ズーム', en: 'Zoom' },
  Window: { ja: 'ウィンドウ', en: 'Window' },
  Macros: { ja: 'マクロ', en: 'Macros' },
  Automation: { ja: '自動化', en: 'Automation' },
  Gallery: { ja: 'ギャラリー', en: 'Gallery' },
};

export const macRibbonDisabledReason = (
  reason: string | undefined,
  lang: ToolbarLang,
): string | undefined => {
  if (!reason) return undefined;
  if (reason === OFFICE_REQUIRED.en) return OFFICE_REQUIRED[lang];
  if (reason === NOT_IMPLEMENTED.en) return NOT_IMPLEMENTED[lang];
  if (reason === EMBEDDED_UNSUPPORTED.en) return EMBEDDED_UNSUPPORTED[lang];
  if (reason === WORKBOOK_THEME_UNSUPPORTED.en) return WORKBOOK_THEME_UNSUPPORTED[lang];
  if (reason === THREADED_COMMENTS_UNSUPPORTED.en) return THREADED_COMMENTS_UNSUPPORTED[lang];
  return reason;
};

/** Localized label lookup used by both the model and the static Mac menus. */
export const macRibbonLabelForCommand = (id: string, lang: ToolbarLang): string => {
  const key = id.split('.').at(-1) ?? id;
  const label = MENU_LABELS[id] ?? labels[key];
  return label ? label[lang] : id;
};

/** Root command ids whose primary face opens a Mac menu. The mount can use
 * this explicit list to apply menu-first behavior without changing the
 * generic ribbon activation manifest. */
export const MAC_RIBBON_MENU_COMMANDS = [
  'mac.insert.pivotTable',
  'mac.insert.shapes',
  'mac.insert.recommendedChart',
  'mac.insert.chartColumn',
  'mac.insert.chartBar',
  'mac.insert.chartLine',
  'mac.insert.chartArea',
  'mac.insert.chartPie',
  'mac.insert.chartScatter',
  'mac.draw.eraser',
  'mac.draw.add',
  'mac.page.margins',
  'mac.page.orientation',
  'mac.page.size',
  'mac.page.printArea',
  'mac.page.pageBreaks',
  'mac.formulas.autoSum',
  'mac.formulas.recent',
  'mac.formulas.financial',
  'mac.formulas.logical',
  'mac.formulas.text',
  'mac.formulas.dateTime',
  'mac.formulas.lookup',
  'mac.formulas.math',
  'mac.formulas.more',
  'mac.formulas.calcOptions',
  'mac.formulas.createFromSelection',
  'mac.formulas.removeArrows',
  'mac.formulas.errorCheck',
  'mac.data.dataTypes',
  'mac.data.sortFilter',
  'mac.data.validation',
  'mac.data.whatIf',
  'mac.review.notes',
  'mac.view.show',
  'mac.view.freeze',
  'mac.automate.gallery',
] as const;

export const MAC_RIBBON_MENU_COMMAND_SET: ReadonlySet<string> = new Set(MAC_RIBBON_MENU_COMMANDS);

const MENU_LABELS: Readonly<Record<string, Label>> = {
  'mac.formulas.removeArrows.all': { ja: '矢印の削除', en: 'Remove Arrows' },
  'mac.formulas.removeArrows.precedents': {
    ja: '参照元の矢印の削除',
    en: 'Remove Precedent Arrows',
  },
  'mac.formulas.removeArrows.dependents': {
    ja: '参照先の矢印の削除',
    en: 'Remove Dependent Arrows',
  },
  'mac.formulas.errorCheck.run': { ja: 'エラー チェック...', en: 'Error Checking...' },
  'mac.formulas.errorCheck.trace': { ja: 'エラー トレース', en: 'Trace Error' },
  'mac.formulas.errorCheck.ignore': { ja: 'エラーを無視する', en: 'Ignore Error' },
  'mac.data.validation.settings': { ja: 'データの入力規則...', en: 'Data Validation...' },
  'mac.data.validation.circleInvalid': { ja: '無効データのマーク', en: 'Circle Invalid Data' },
  'mac.data.validation.clearCircles': {
    ja: '入力規則マークのクリア',
    en: 'Clear Validation Circles',
  },
  'mac.data.validation.clearRules': { ja: '入力規則のクリア', en: 'Clear Validation' },
  'mac.page.margins.normal': { ja: '標準', en: 'Normal' },
  'mac.page.margins.wide': { ja: '広い', en: 'Wide' },
  'mac.page.margins.narrow': { ja: '狭い', en: 'Narrow' },
  'mac.page.margins.custom': { ja: 'ユーザー設定の余白', en: 'Custom Margins' },
  'mac.page.orientation.portrait': { ja: '縦', en: 'Portrait' },
  'mac.page.orientation.landscape': { ja: '横', en: 'Landscape' },
  'mac.page.printArea.set': { ja: '印刷範囲の設定', en: 'Set Print Area' },
  'mac.page.printArea.add': { ja: '印刷範囲に追加', en: 'Add to Print Area' },
  'mac.page.printArea.clear': { ja: '印刷範囲のクリア', en: 'Clear Print Area' },
  'mac.formulas.createNames.topRow': { ja: '上端行', en: 'Top row' },
  'mac.formulas.createNames.bottomRow': { ja: '下端行', en: 'Bottom row' },
  'mac.formulas.createNames.leftColumn': { ja: '左端列', en: 'Left column' },
  'mac.formulas.createNames.rightColumn': { ja: '右端列', en: 'Right column' },
  'mac.data.dataTypes.stocks': { ja: '株式', en: 'Stocks' },
  'mac.data.dataTypes.geography': { ja: '地理', en: 'Geography' },
  'mac.data.goalSeek': { ja: 'ゴール シーク', en: 'Goal Seek' },
  'mac.data.scenarioManager': { ja: 'シナリオ マネージャー', en: 'Scenario Manager' },
  'mac.data.dataTable': { ja: 'データ テーブル', en: 'Data Table' },
  'mac.review.newNote': { ja: '新しいメモ', en: 'New Note' },
  'mac.review.nextNote': { ja: '次のメモ', en: 'Next Note' },
  'mac.review.previousNote': { ja: '前のメモ', en: 'Previous Note' },
  'mac.review.showNotes': { ja: 'すべてのメモを表示', en: 'Show Notes' },
  'mac.view.freeze.off': { ja: 'ウィンドウ枠固定の解除', en: 'Unfreeze Panes' },
  'mac.view.freeze.firstRow': { ja: '先頭行の固定', en: 'Freeze Top Row' },
  'mac.view.freeze.firstColumn': { ja: '先頭列の固定', en: 'Freeze First Column' },
  ...Object.fromEntries(
    MAC_AUTOMATION_GALLERY.map((script) => [
      script.commandId,
      { ja: script.labelJa, en: script.label },
    ]),
  ),
};

const menu = (items: readonly MacRibbonMenuItem[]): readonly MacRibbonMenuItem[] =>
  items.map((item) => {
    const label = MENU_LABELS[item.id] ?? labels[item.id.split('.').at(-1) ?? ''];
    return label ? { ...item, label: label.en, labelJa: label.ja } : item;
  });
const leaf = (
  id: string,
  key: string,
  disabled = false,
  disabledReason?: Label,
): MacRibbonMenuItem => ({
  id,
  label: labels[key]?.en ?? key,
  labelJa: labels[key]?.ja,
  disabled,
  disabledReason: disabledReason?.en,
});

const functionMenu = (category: CatalogFunctionCategory): readonly MacRibbonMenuItem[] =>
  menu(
    supportedFunctionNames(category).map((name) => ({
      id: `mac.function.${name}`,
      label: name,
    })),
  );

type FormulaFamily =
  | 'recent'
  | 'financial'
  | 'logical'
  | 'text'
  | 'dateTime'
  | 'lookup'
  | 'math'
  | 'more';

const FORMULA_FAMILY_KEYS = [
  'recent',
  'financial',
  'logical',
  'text',
  'dateTime',
  'lookup',
  'math',
  'more',
] as const satisfies readonly FormulaFamily[];

const FORMULA_FAMILY_ICONS = {
  recent: 'functionRecent',
  financial: 'functionFinancial',
  logical: 'functionLogical',
  text: 'functionText',
  dateTime: 'functionDateTime',
  lookup: 'functionLookup',
  math: 'functionMath',
  more: 'functionMore',
} as const satisfies Record<FormulaFamily, ExcelRibbonIconName>;

/** Menu leaves are data commands so the host's normal ribbon click delegation
 *  and restricted-interaction guard can handle them exactly like top-level
 *  buttons. */
export const MAC_RIBBON_MENU_ITEMS: Readonly<Record<string, readonly MacRibbonMenuItem[]>> = {
  'mac.insert.pivotTable': menu([
    leaf('mac.insert.pivotTable', 'pivotTable'),
    leaf('mac.insert.recommendedPivotTable', 'recommendedPivotTable'),
  ]),
  'mac.insert.shapes': menu([
    ...['Line', 'Arrow', 'Rectangle', 'RoundedRectangle', 'Oval', 'Triangle', 'Diamond'].map(
      (kind) => ({
        ...leaf(`mac.insert.shape${kind}`, `shape${kind}`),
        icon: `shape${kind}`,
      }),
    ),
    { ...leaf('mac.insert.shapeCallout', 'shapeCallout', true, NOT_IMPLEMENTED), icon: 'shapes' },
  ]),
  'mac.insert.recommendedChart': menu([
    leaf('mac.insert.chartColumn', 'chartColumn'),
    leaf('mac.insert.chartLine', 'chartLine'),
    leaf('mac.insert.chartPie', 'chartPie'),
    leaf('mac.insert.chartHierarchy', 'chartHierarchy', true, NOT_IMPLEMENTED),
    leaf('mac.insert.chartStatistical', 'chartStatistical', true, NOT_IMPLEMENTED),
    leaf('mac.insert.chartScatter', 'chartScatter'),
    leaf('mac.insert.chartWaterfall', 'chartWaterfall', true, NOT_IMPLEMENTED),
    leaf('mac.insert.chartCombo', 'chartCombo', true, NOT_IMPLEMENTED),
    leaf('mac.insert.chartMap', 'chartMap', true, OFFICE_REQUIRED),
    leaf('mac.insert.pivotChart', 'pivotChart', true, NOT_IMPLEMENTED),
  ]),
  'mac.insert.chartColumn': menu([
    leaf('mac.insert.chartColumn', 'chartColumn'),
    leaf('mac.insert.chartBar', 'chartBar'),
  ]),
  'mac.insert.chartBar': menu([leaf('mac.insert.chartBar', 'chartBar')]),
  'mac.insert.chartLine': menu([
    leaf('mac.insert.chartLine', 'chartLine'),
    leaf('mac.insert.chartArea', 'chartArea'),
  ]),
  'mac.insert.chartArea': menu([leaf('mac.insert.chartArea', 'chartArea')]),
  'mac.insert.chartPie': menu([leaf('mac.insert.chartPie', 'chartPie')]),
  'mac.insert.chartScatter': menu([leaf('mac.insert.chartScatter', 'chartScatter')]),
  'mac.draw.eraser': menu([leaf('mac.draw.eraser', 'eraser'), leaf('mac.draw.toggle', 'draw')]),
  'mac.draw.add': menu([
    leaf('mac.draw.penBlack', 'penBlack'),
    leaf('mac.draw.penRed', 'penRed'),
    leaf('mac.draw.pencil', 'pencil'),
    leaf('mac.draw.highlighter', 'highlighter'),
  ]),
  'mac.page.margins': menu([
    { id: 'mac.page.margins.normal', label: 'Normal' },
    { id: 'mac.page.margins.wide', label: 'Wide' },
    { id: 'mac.page.margins.narrow', label: 'Narrow' },
    { id: 'mac.page.margins.custom', label: 'Custom Margins' },
  ]),
  'mac.page.orientation': menu([
    { id: 'mac.page.orientation.portrait', label: 'Portrait' },
    { id: 'mac.page.orientation.landscape', label: 'Landscape' },
  ]),
  'mac.page.size': menu([
    { id: 'mac.page.size.a4', label: 'A4' },
    { id: 'mac.page.size.a3', label: 'A3' },
    { id: 'mac.page.size.letter', label: 'Letter' },
    { id: 'mac.page.size.legal', label: 'Legal' },
  ]),
  'mac.page.printArea': menu([
    { id: 'mac.page.printArea.set', label: 'Set Print Area' },
    { id: 'mac.page.printArea.add', label: 'Add to Print Area' },
    { id: 'mac.page.printArea.clear', label: 'Clear Print Area' },
  ]),
  'mac.page.pageBreaks': menu([
    { id: 'mac.page.break.insert', label: 'Insert Page Break', labelJa: '改ページを挿入' },
    { id: 'mac.page.break.remove', label: 'Remove Page Break', labelJa: '改ページを削除' },
    {
      id: 'mac.page.break.reset',
      label: 'Reset All Page Breaks',
      labelJa: 'すべての改ページをリセット',
    },
  ]),
  'mac.formulas.autoSum': menu([
    { id: 'mac.autosum.SUM', label: 'Sum', labelJa: '合計' },
    { id: 'mac.autosum.AVERAGE', label: 'Average', labelJa: '平均' },
    { id: 'mac.autosum.COUNT', label: 'Count Numbers', labelJa: '数値の個数' },
    { id: 'mac.autosum.MAX', label: 'Max', labelJa: '最大値' },
    { id: 'mac.autosum.MIN', label: 'Min', labelJa: '最小値' },
    { id: 'mac.formulas.category.all', label: 'More Functions...', labelJa: 'その他の関数...' },
  ]),
  'mac.formulas.recent': menu([
    { id: 'mac.formulas.category.recent', label: 'More Functions...', labelJa: 'その他の関数...' },
  ]),
  'mac.formulas.financial': functionMenu('financial'),
  'mac.formulas.logical': functionMenu('logical'),
  'mac.formulas.text': functionMenu('text'),
  'mac.formulas.dateTime': functionMenu('datetime'),
  'mac.formulas.lookup': functionMenu('lookup'),
  'mac.formulas.math': functionMenu('math'),
  'mac.formulas.more': menu([]),
  'mac.formulas.calcOptions': menu([
    { id: 'mac.formulas.calc.auto', label: 'Automatic', labelJa: '自動' },
    {
      id: 'mac.formulas.calc.autoNoTable',
      label: 'Automatic Except for Data Tables',
      labelJa: 'データ テーブル以外は自動',
    },
    { id: 'mac.formulas.calc.manual', label: 'Manual', labelJa: '手動' },
    {
      id: 'mac.formulas.calc.iterative',
      label: 'Enable Iterative Calculation',
      labelJa: '反復計算を有効にする',
    },
  ]),
  'mac.formulas.createFromSelection': menu([
    { id: 'mac.formulas.createNames.topRow', label: 'Top row' },
    { id: 'mac.formulas.createNames.bottomRow', label: 'Bottom row' },
    { id: 'mac.formulas.createNames.leftColumn', label: 'Left column' },
    { id: 'mac.formulas.createNames.rightColumn', label: 'Right column' },
  ]),
  'mac.data.dataTypes': menu([
    leaf('mac.data.dataTypes.stocks', 'dataTypes', true, OFFICE_REQUIRED),
    leaf('mac.data.dataTypes.geography', 'dataTypes', true, OFFICE_REQUIRED),
  ]),
  'mac.data.sortFilter': menu([
    leaf('mac.data.sortAsc', 'sortAsc'),
    leaf('mac.data.sortDesc', 'sortDesc'),
    leaf('mac.data.sortCustom', 'sortCustom'),
    leaf('mac.data.filter', 'filter'),
    leaf('mac.data.clear', 'clear'),
    leaf('mac.data.reapply', 'reapply'),
    leaf('mac.data.advancedFilter', 'advancedFilter'),
  ]),
  'mac.formulas.removeArrows': menu([
    { id: 'mac.formulas.removeArrows.all', label: 'Remove Arrows' },
    { id: 'mac.formulas.removeArrows.precedents', label: 'Remove Precedent Arrows' },
    { id: 'mac.formulas.removeArrows.dependents', label: 'Remove Dependent Arrows' },
  ]),
  'mac.formulas.errorCheck': menu([
    { id: 'mac.formulas.errorCheck.run', label: 'Error Checking...' },
    { id: 'mac.formulas.errorCheck.trace', label: 'Trace Error' },
    { id: 'mac.formulas.errorCheck.ignore', label: 'Ignore Error' },
  ]),
  'mac.data.validation': menu([
    { id: 'mac.data.validation.settings', label: 'Data Validation...' },
    { id: 'mac.data.validation.circleInvalid', label: 'Circle Invalid Data' },
    { id: 'mac.data.validation.clearCircles', label: 'Clear Validation Circles' },
    { id: 'mac.data.validation.clearRules', label: 'Clear Validation' },
  ]),
  'mac.data.whatIf': menu([
    { id: 'mac.data.goalSeek', label: 'Goal Seek' },
    {
      id: 'mac.data.scenarioManager',
      label: 'Scenario Manager',
      disabled: true,
      disabledReason: NOT_IMPLEMENTED.en,
    },
    {
      id: 'mac.data.dataTable',
      label: 'Data Table',
      disabled: true,
      disabledReason: EMBEDDED_UNSUPPORTED.en,
    },
  ]),
  'mac.review.notes': menu([
    { id: 'mac.review.newNote', label: 'New Note' },
    { id: 'mac.review.nextNote', label: 'Next Note' },
    { id: 'mac.review.previousNote', label: 'Previous Note' },
    { id: 'mac.review.showNotes', label: 'Show Notes' },
  ]),
  'mac.view.show': menu([
    { id: 'mac.view.gridlines', label: 'Gridlines' },
    { id: 'mac.view.headings', label: 'Headings' },
    { id: 'mac.view.formulaBar', label: 'Formula Bar' },
    { id: 'mac.view.zeros', label: 'Zero Values' },
  ]),
  'mac.view.freeze': menu([
    { id: 'mac.view.freeze', label: 'Freeze Panes' },
    { id: 'mac.view.freeze.off', label: 'Unfreeze Panes' },
    { id: 'mac.view.freeze.firstRow', label: 'Freeze Top Row' },
    { id: 'mac.view.freeze.firstColumn', label: 'Freeze First Column' },
  ]),
  'mac.automate.gallery': menu(
    MAC_AUTOMATION_GALLERY.map((script) => ({ id: script.commandId, label: script.label })),
  ),
};

/** DOM id of the Mac Formulas "More Functions" menu. */
export const MAC_FORMULAS_MORE_MENU_ID = 'menu-mac-formulas-more';

export const macRibbonMenuIdForCommand = (commandId: string): string | null => {
  if (!MAC_RIBBON_MENU_COMMAND_SET.has(commandId)) return null;
  const ids: Readonly<Record<string, string>> = {
    'mac.insert.pivotTable': 'menu-mac-insert-pivot',
    'mac.insert.shapes': 'menu-mac-insert-shapes',
    'mac.insert.recommendedChart': 'menu-mac-insert-chart',
    'mac.insert.chartColumn': 'menu-mac-insert-chart-column',
    'mac.insert.chartBar': 'menu-mac-insert-chart-bar',
    'mac.insert.chartLine': 'menu-mac-insert-chart-line',
    'mac.insert.chartArea': 'menu-mac-insert-chart-area',
    'mac.insert.chartPie': 'menu-mac-insert-chart-pie',
    'mac.insert.chartScatter': 'menu-mac-insert-chart-scatter',
    'mac.draw.eraser': 'menu-mac-draw-eraser',
    'mac.draw.add': 'menu-mac-draw-add',
    'mac.page.margins': 'menu-mac-page-margins',
    'mac.page.orientation': 'menu-mac-page-orientation',
    'mac.page.size': 'menu-mac-page-size',
    'mac.page.printArea': 'menu-mac-page-print-area',
    'mac.page.pageBreaks': 'menu-mac-page-breaks',
    'mac.formulas.autoSum': 'menu-mac-formulas-autosum',
    'mac.formulas.recent': 'menu-mac-formulas-recent',
    'mac.formulas.financial': 'menu-mac-formulas-financial',
    'mac.formulas.logical': 'menu-mac-formulas-logical',
    'mac.formulas.text': 'menu-mac-formulas-text',
    'mac.formulas.dateTime': 'menu-mac-formulas-date-time',
    'mac.formulas.lookup': 'menu-mac-formulas-lookup',
    'mac.formulas.math': 'menu-mac-formulas-math',
    'mac.formulas.more': MAC_FORMULAS_MORE_MENU_ID,
    'mac.formulas.calcOptions': 'menu-mac-formulas-calc-options',
    'mac.formulas.createFromSelection': 'menu-mac-formulas-create-names',
    'mac.formulas.removeArrows': 'menu-mac-formulas-remove-arrows',
    'mac.formulas.errorCheck': 'menu-mac-formulas-error-check',
    'mac.data.dataTypes': 'menu-mac-data-types',
    'mac.data.sortFilter': 'menu-mac-data-sort-filter',
    'mac.data.validation': 'menu-mac-data-validation',
    'mac.data.whatIf': 'menu-mac-data-what-if',
    'mac.review.notes': 'menu-mac-review-notes',
    'mac.view.show': 'menu-mac-view-show',
    'mac.view.freeze': 'menu-mac-view-freeze',
    'mac.automate.gallery': 'menu-mac-automate-gallery',
  };
  return ids[commandId] ?? null;
};

const insertTab = (): RibbonTabModel =>
  macTab('insert', t('insert'), [
    group({ ja: 'テーブル', en: 'Tables' }, [
      menuCommand('mac.insert.pivotTable', t('pivotTable'), 'pivotTable'),
      menuCommand('mac.insert.recommendedPivotTable', t('recommendedPivotTable'), 'pivotTable'),
      menuCommand('mac.insert.table', t('table'), 'table'),
    ]),
    group({ ja: 'フォーム', en: 'Forms' }, [
      disabledCommand('mac.insert.forms', t('forms'), 'form', OFFICE_REQUIRED),
    ]),
    group(
      { ja: '画像', en: 'Pictures' },
      [
        disabledCompactCommand(
          'mac.insert.pictureFromData',
          t('pictureFromData'),
          'picture',
          OFFICE_REQUIRED,
        ),
      ],
      'compact',
    ),
    group(
      { ja: 'イラスト', en: 'Illustrations' },
      [
        compactCommand('mac.insert.photo', t('photo'), 'picture'),
        compactCommand('mac.insert.shapes', t('shapes'), 'shapes'),
        disabledCompactCommand('mac.insert.icons', t('icons'), 'icons', NOT_IMPLEMENTED),
        disabledCompactCommand(
          'mac.insert.threeDModel',
          t('threeDModel'),
          'threeD',
          NOT_IMPLEMENTED,
        ),
        disabledCompactCommand('mac.insert.smartArt', t('smartArt'), 'smartArt', NOT_IMPLEMENTED),
        compactCommand('mac.insert.screenshot', t('screenshot'), 'screenshot'),
      ],
      'compact',
    ),
    group({ ja: 'コントロール', en: 'Controls' }, [
      disabledCommand('mac.insert.checkBox', t('checkBox'), 'checkBox', NOT_IMPLEMENTED),
    ]),
    group(
      { ja: 'グラフ', en: 'Charts' },
      [
        iconCommand('mac.insert.recommendedChart', t('recommendedChart'), 'chartRecommended'),
        iconCommand('mac.insert.chartColumn', t('chartColumn'), 'chartColumn'),
        iconCommand('mac.insert.chartLine', t('chartLine'), 'chartLine'),
        iconCommand('mac.insert.chartPie', t('chartPie'), 'chartPie'),
        disabledIconCommand(
          'mac.insert.chartHierarchy',
          t('chartHierarchy'),
          'chartHierarchy',
          NOT_IMPLEMENTED,
        ),
        disabledIconCommand(
          'mac.insert.chartStatistical',
          t('chartStatistical'),
          'chartStatistical',
          NOT_IMPLEMENTED,
        ),
        iconCommand('mac.insert.chartScatter', t('chartScatter'), 'chartScatter'),
        disabledIconCommand(
          'mac.insert.chartWaterfall',
          t('chartWaterfall'),
          'chartWaterfall',
          NOT_IMPLEMENTED,
        ),
        disabledIconCommand(
          'mac.insert.chartCombo',
          t('chartCombo'),
          'chartCombo',
          NOT_IMPLEMENTED,
        ),
        disabledIconCommand('mac.insert.chartMap', t('chartMap'), 'chartMap', OFFICE_REQUIRED),
        disabledIconCommand(
          'mac.insert.pivotChart',
          t('pivotChart'),
          'pivotTable',
          NOT_IMPLEMENTED,
        ),
      ],
      'compact-icons',
    ),
    group({ ja: 'スパークライン', en: 'Sparklines' }, [
      menuCommand('mac.insert.sparkline', t('sparkline'), 'sparkline'),
    ]),
    group({ ja: 'フィルター', en: 'Filters' }, [
      menuCommand('mac.insert.slicer', t('slicer'), 'slicer'),
      disabledCommand('mac.insert.timeline', t('timeline'), 'timeline', NOT_IMPLEMENTED),
    ]),
    group({ ja: 'リンク', en: 'Links' }, [menuCommand('mac.insert.link', t('link'), 'link')]),
    group({ ja: 'コメント', en: 'Comments' }, [
      disabledCommand(
        'mac.insert.newComment',
        t('newComment'),
        'commentAdd',
        THREADED_COMMENTS_UNSUPPORTED,
      ),
    ]),
    group({ ja: 'テキスト', en: 'Text' }, [
      disabledCommand('mac.insert.textBox', t('textBox'), 'textBox', EMBEDDED_UNSUPPORTED),
      menuCommand('mac.insert.headerFooter', t('headerFooter'), 'headerFooter'),
      disabledCommand('mac.insert.wordArt', t('wordArt'), 'wordArt', NOT_IMPLEMENTED),
      disabledCommand('mac.insert.object', t('object'), 'object', NOT_IMPLEMENTED),
    ]),
    group({ ja: '記号', en: 'Symbols' }, [
      disabledCommand('mac.insert.equation', t('equation'), 'function', NOT_IMPLEMENTED),
      menuCommand('mac.insert.symbol', t('symbol'), 'symbol'),
    ]),
  ]);

const drawTab = (): RibbonTabModel =>
  macTab('draw', t('draw'), [
    group(
      { ja: 'ツール', en: 'Tools' },
      [
        compactCommand('mac.draw.toggle', t('draw'), 'pen'),
        compactCommand('mac.draw.eraser', t('eraser'), 'eraser'),
        disabledCompactCommand('mac.draw.lasso', t('lasso'), 'lasso', EMBEDDED_UNSUPPORTED),
      ],
      'compact',
    ),
    group(
      { ja: 'ペン', en: 'Pens' },
      [
        iconCommand('mac.draw.penBlack', t('penBlack'), 'penBlack'),
        iconCommand('mac.draw.penRed', t('penRed'), 'penRed'),
        iconCommand('mac.draw.pencil', t('pencil'), 'pencil'),
        iconCommand('mac.draw.highlighter', t('highlighter'), 'highlighter'),
        iconCommand('mac.draw.add', t('add'), 'plus'),
      ],
      'pens',
    ),
    group(
      { ja: '描画', en: 'Draw' },
      [compactCommand('mac.draw.trackpad', t('trackpad'), 'trackpad')],
      'compact',
    ),
  ]);

const pageLayoutTab = (): RibbonTabModel =>
  macTab('pageLayout', { ja: 'ページ レイアウト', en: 'Page Layout' }, [
    group(
      { ja: 'テーマ', en: 'Themes' },
      [
        disabledCommand('mac.page.theme', t('theme'), 'pageTheme', WORKBOOK_THEME_UNSUPPORTED),
        disabledCommand('mac.page.color', t('color'), 'pageTheme', WORKBOOK_THEME_UNSUPPORTED),
        disabledCommand('mac.page.font', t('font'), 'font', WORKBOOK_THEME_UNSUPPORTED),
      ],
      'compact',
    ),
    group(
      { ja: 'ページ設定', en: 'Page Setup' },
      [
        compactCommand('mac.page.margins', t('margins'), 'pageSetup'),
        compactCommand('mac.page.orientation', t('orientation'), 'pageSetup'),
        compactCommand('mac.page.size', t('size'), 'pageSetup'),
        compactCommand('mac.page.printArea', t('printArea'), 'printArea'),
        compactCommand('mac.page.pageBreaks', t('pageBreaks'), 'pageBreaks'),
        compactCommand('mac.page.background', t('background'), 'sheetBackground'),
        compactCommand('mac.page.printTitles', t('printTitles'), 'printTitles'),
        compactCommand('mac.page.pageSetup', t('pageSetup'), 'pageSetup'),
      ],
      'compact',
    ),
    group({ ja: '拡大縮小印刷', en: 'Scale to Fit' }, [
      selectCommand('scaleWidth', t('scaleWidth'), [
        { value: '0', label: 'Automatic' },
        { value: '1', label: '1 page' },
        { value: '2', label: '2 pages' },
        { value: '3', label: '3 pages' },
        { value: 'custom', label: 'Custom' },
      ]),
      selectCommand('scaleHeight', t('scaleHeight'), [
        { value: '0', label: 'Automatic' },
        { value: '1', label: '1 page' },
        { value: '2', label: '2 pages' },
        { value: '3', label: '3 pages' },
        { value: 'custom', label: 'Custom' },
      ]),
    ]),
    group(
      { ja: 'シートのオプション', en: 'Sheet Options' },
      [
        checkboxCommand('mac.page.showGridlines', t('showGridlines')),
        checkboxCommand('mac.page.printGridlines', t('printGridlines')),
        checkboxCommand('mac.page.showHeadings', t('showHeadings')),
        checkboxCommand('mac.page.printHeadings', t('printHeadings')),
      ],
      'checks',
    ),
  ]);

const formulasTab = (): RibbonTabModel =>
  macTab('formulas', t('formulas'), [
    group(
      { ja: '関数ライブラリ', en: 'Function Library' },
      [
        menuCommand('mac.formulas.insertFunction', t('insertFunction'), 'function'),
        menuCommand('mac.formulas.autoSum', t('autoSum'), 'autosum'),
        ...FORMULA_FAMILY_KEYS.map((key) =>
          menuCommand(`mac.formulas.${key}`, t(key), FORMULA_FAMILY_ICONS[key]),
        ),
      ],
      'function-library',
    ),
    group(
      { ja: 'Python', en: 'Python' },
      [
        disabledCommand('mac.formulas.python', t('python'), 'python', OFFICE_REQUIRED),
        disabledCommand('mac.formulas.reset', t('reset'), 'clear', OFFICE_REQUIRED),
        disabledCommand('mac.formulas.editor', t('editor'), 'script', OFFICE_REQUIRED),
        disabledCommand('mac.formulas.initialize', t('initialize'), 'script', OFFICE_REQUIRED),
      ],
      'compact',
    ),
    group(
      { ja: '定義された名前', en: 'Defined Names' },
      [
        menuCommand('mac.formulas.namesManager', t('namesManager'), 'names'),
        menuCommand('mac.formulas.defineName', t('defineName'), 'names'),
        menuCommand('mac.formulas.useInFormula', t('useInFormula'), 'names'),
        menuCommand('mac.formulas.createFromSelection', t('createFromSelection'), 'names'),
      ],
      'compact',
    ),
    group(
      { ja: '数式の監査', en: 'Formula Auditing' },
      [
        menuCommand('mac.formulas.precedents', t('precedents'), 'trace'),
        menuCommand('mac.formulas.dependents', t('dependents'), 'dependents'),
        menuCommand('mac.formulas.removeArrows', t('removeArrows'), 'clearArrows'),
        menuCommand('mac.formulas.showFormulas', t('showFormulas'), 'function'),
        menuCommand('mac.formulas.errorCheck', t('errorCheck'), 'errorChecking'),
        menuCommand('mac.formulas.watch', t('watch'), 'watch'),
      ],
      'compact',
    ),
    group(
      { ja: '計算方法', en: 'Calculation' },
      [
        menuCommand('mac.formulas.calcOptions', t('calcOptions'), 'calcOptions'),
        menuCommand('mac.formulas.recalc', t('recalc'), 'autosum'),
        menuCommand('mac.formulas.sheetRecalc', t('sheetRecalc'), 'autosum'),
      ],
      'compact',
    ),
  ]);

const dataTab = (): RibbonTabModel =>
  macTab('data', t('data'), [
    group({ ja: 'データの取得と変換', en: 'Get & Transform Data' }, [
      disabledCommand('mac.data.powerQuery', t('powerQuery'), 'data', NOT_IMPLEMENTED),
      disabledCommand('mac.data.pictureFromData', t('pictureFromData'), 'picture', OFFICE_REQUIRED),
    ]),
    group(
      { ja: 'クエリと接続', en: 'Queries & Connections' },
      [
        disabledCompactCommand('mac.data.refreshAll', t('refreshAll'), 'refresh', NOT_IMPLEMENTED),
        disabledCompactCommand(
          'mac.data.queriesConnections',
          t('queriesConnections'),
          'link',
          NOT_IMPLEMENTED,
        ),
        disabledCompactCommand('mac.data.properties', t('properties'), 'options', NOT_IMPLEMENTED),
        compactCommand('mac.data.workbookLinks', t('workbookLinks'), 'link'),
      ],
      'compact',
    ),
    group(
      { ja: 'データの種類', en: 'Data Types' },
      [menuCommand('mac.data.dataTypes', t('dataTypes'), 'data')],
      'compact',
    ),
    group(
      { ja: '並べ替えとフィルター', en: 'Sort & Filter' },
      [
        compactCommand('mac.data.sortAsc', t('sortAsc'), 'sortAsc'),
        compactCommand('mac.data.sortDesc', t('sortDesc'), 'sortDesc'),
        compactCommand('mac.data.sortCustom', t('sortCustom'), 'sortAsc'),
        checkboxCommand('mac.data.filter', t('filter'), 'filter'),
        compactCommand('mac.data.clear', t('clear'), 'clear'),
        compactCommand('mac.data.reapply', t('reapply'), 'refresh'),
        compactCommand('mac.data.advancedFilter', t('advancedFilter'), 'filter'),
      ],
      'compact',
    ),
    group(
      { ja: 'データ ツール', en: 'Data Tools' },
      [
        compactCommand('mac.data.textToColumns', t('textToColumns'), 'textToColumns'),
        compactCommand('mac.data.flashFill', t('flashFill'), 'flashFill'),
        compactCommand('mac.data.removeDuplicates', t('removeDuplicates'), 'removeDuplicates'),
        compactCommand('mac.data.validation', t('validation'), 'dataValidation'),
        compactCommand('mac.data.consolidate', t('consolidate'), 'merge'),
      ],
      'compact',
    ),
    group(
      { ja: '予測', en: 'Forecast' },
      [menuCommand('mac.data.whatIf', t('whatIf'), 'options')],
      'compact',
    ),
    group(
      { ja: 'アウトライン', en: 'Outline' },
      [
        compactCommand('mac.data.group', t('group'), 'outlineGroup'),
        compactCommand('mac.data.ungroup', t('ungroup'), 'outlineUngroup'),
        compactCommand('mac.data.subtotal', t('subtotal'), 'outlineGroup'),
        compactCommand('mac.data.showDetail', t('showDetail'), 'outlineShow'),
        compactCommand('mac.data.hideDetail', t('hideDetail'), 'outlineHide'),
      ],
      'compact',
    ),
    group(
      { ja: '分析', en: 'Analysis' },
      [disabledCompactCommand('mac.data.analysis', t('analysis'), 'chart', NOT_IMPLEMENTED)],
      'compact',
    ),
  ]);

const reviewTab = (): RibbonTabModel =>
  macTab('review', t('review'), [
    group(
      { ja: '文章校正', en: 'Proofing' },
      [
        compactCommand('mac.review.spelling', t('spelling'), 'spelling'),
        disabledCompactCommand(
          'mac.review.thesaurus',
          t('thesaurus'),
          'spelling',
          EMBEDDED_UNSUPPORTED,
        ),
        compactCommand('mac.review.stats', t('stats'), 'options'),
      ],
      'compact',
    ),
    group(
      { ja: 'アクセシビリティ', en: 'Accessibility' },
      [compactCommand('mac.review.accessibility', t('accessibility'), 'accessibility')],
      'compact',
    ),
    group(
      { ja: '言語', en: 'Language' },
      [
        disabledCompactCommand(
          'mac.review.translate',
          t('translate'),
          'translate',
          NOT_IMPLEMENTED,
        ),
      ],
      'compact',
    ),
    group(
      { ja: '変更', en: 'Changes' },
      [
        disabledCompactCommand(
          'mac.review.showChanges',
          t('showChanges'),
          'options',
          EMBEDDED_UNSUPPORTED,
        ),
      ],
      'compact',
    ),
    group(
      { ja: 'コメント', en: 'Comments' },
      [
        disabledCompactCommand(
          'mac.review.newComment',
          t('newComment'),
          'commentAdd',
          THREADED_COMMENTS_UNSUPPORTED,
        ),
        disabledCompactCommand(
          'mac.review.deleteComment',
          t('deleteComment'),
          'clear',
          THREADED_COMMENTS_UNSUPPORTED,
        ),
        disabledCompactCommand(
          'mac.review.previousComment',
          t('previousComment'),
          'goTo',
          THREADED_COMMENTS_UNSUPPORTED,
        ),
        disabledCompactCommand(
          'mac.review.nextComment',
          t('nextComment'),
          'goTo',
          THREADED_COMMENTS_UNSUPPORTED,
        ),
        disabledCompactCommand(
          'mac.review.showComments',
          t('showComments'),
          'commentAdd',
          THREADED_COMMENTS_UNSUPPORTED,
        ),
      ],
      'compact',
    ),
    group(
      { ja: 'メモ', en: 'Notes' },
      [compactCommand('mac.review.notes', t('notes'), 'commentAdd')],
      'compact',
    ),
    group(
      { ja: '保護', en: 'Protect' },
      [
        compactCommand('mac.review.protectSheet', t('protectSheet'), 'protect'),
        compactCommand('mac.review.protectWorkbook', t('protectWorkbook'), 'protect'),
      ],
      'compact',
    ),
    group(
      { ja: 'オプション', en: 'Options' },
      [
        disabledCompactCommand(
          'mac.review.alwaysReadOnly',
          t('alwaysReadOnly'),
          'protect',
          NOT_IMPLEMENTED,
        ),
        disabledCompactCommand('mac.review.hideInk', t('hideInk'), 'pen', EMBEDDED_UNSUPPORTED),
      ],
      'compact',
    ),
  ]);

const viewTab = (): RibbonTabModel =>
  macTab('view', t('view'), [
    group(
      { ja: 'シート ビュー', en: 'Sheet Views' },
      [
        selectCommand('sheetViewSelect', t('sheetViewSelect'), [
          { value: 'current', label: 'Current View' },
        ]),
        compactCommand('mac.view.sheetViewSave', t('sheetViewSave'), 'options'),
        compactCommand('mac.view.sheetViewDelete', t('sheetViewDelete'), 'clear'),
      ],
      'compact',
    ),
    group(
      { ja: 'ブックの表示', en: 'Workbook Views' },
      [
        compactCommand('mac.view.standard', t('standard'), 'table'),
        compactCommand('mac.view.pageBreakPreview', t('pageBreakPreview'), 'table'),
        compactCommand('mac.view.pageLayout', t('pageLayout'), 'page'),
        disabledCompactCommand(
          'mac.view.customViews',
          t('customViews'),
          'options',
          NOT_IMPLEMENTED,
        ),
      ],
      'compact',
    ),
    group(
      { ja: '表示/非表示', en: 'Show' },
      [compactCommand('mac.view.show', t('show'), 'options')],
      'compact',
    ),
    group(
      { ja: 'ズーム', en: 'Zoom' },
      [
        compactCommand('mac.view.zoom', t('zoom'), 'zoom'),
        compactCommand('mac.view.zoom100', { ja: '100%', en: '100%' }, 'zoom'),
        compactCommand('mac.view.zoomSelection', t('zoomSelection'), 'zoom'),
      ],
      'compact',
    ),
    group(
      { ja: 'ウィンドウ', en: 'Window' },
      [
        disabledCompactCommand(
          'mac.view.newWindow',
          t('newWindow'),
          'window',
          EMBEDDED_UNSUPPORTED,
        ),
        disabledCompactCommand('mac.view.arrange', t('arrange'), 'arrange', EMBEDDED_UNSUPPORTED),
        compactCommand('mac.view.freeze', t('freeze'), 'freeze'),
        disabledCompactCommand('mac.view.split', t('split'), 'split', EMBEDDED_UNSUPPORTED),
        disabledCompactCommand('mac.view.hide', t('hide'), 'hide', EMBEDDED_UNSUPPORTED),
        disabledCompactCommand('mac.view.unhide', t('unhide'), 'show', EMBEDDED_UNSUPPORTED),
        disabledCompactCommand(
          'mac.view.sideBySide',
          t('sideBySide'),
          'window',
          EMBEDDED_UNSUPPORTED,
        ),
        disabledCompactCommand(
          'mac.view.syncScroll',
          t('syncScroll'),
          'window',
          EMBEDDED_UNSUPPORTED,
        ),
        disabledCompactCommand(
          'mac.view.resetWindow',
          t('resetWindow'),
          'window',
          EMBEDDED_UNSUPPORTED,
        ),
        disabledCompactCommand(
          'mac.view.switchWindow',
          t('switchWindow'),
          'window',
          EMBEDDED_UNSUPPORTED,
        ),
      ],
      'compact',
    ),
    group(
      { ja: 'マクロ', en: 'Macros' },
      [disabledCompactCommand('mac.view.macros', t('macros'), 'script', NOT_IMPLEMENTED)],
      'compact',
    ),
  ]);

const automateTab = (): RibbonTabModel =>
  macTab('automate', t('automate'), [
    group(
      { ja: '自動化', en: 'Automation' },
      [
        disabledCompactCommand('mac.automate.newScript', t('newScript'), 'script', OFFICE_REQUIRED),
        compactCommand('mac.automate.showScripts', t('showScripts'), 'script'),
        menuCommand('mac.automate.gallery', t('gallery'), 'script'),
      ],
      'compact',
    ),
    group(
      { ja: 'ギャラリー', en: 'Gallery' },
      [
        compactCommand('mac.automate.allRowsColumns', t('allRowsColumns'), 'show'),
        compactCommand('mac.automate.freezeSelection', t('freezeSelection'), 'freeze'),
        compactCommand('mac.automate.makeSubtable', t('makeSubtable'), 'table'),
        compactCommand('mac.automate.removeHyperlinks', t('removeHyperlinks'), 'link'),
        compactCommand('mac.automate.countEmptyRows', t('countEmptyRows'), 'count'),
        compactCommand('mac.automate.tableToJson', t('tableToJson'), 'code'),
        compactCommand('mac.automate.newPivotTable', t('newPivotTable'), 'pivotTable'),
      ],
      'gallery',
    ),
  ]);

const MAC_HOME_ICON_ONLY_COMMAND_IDS = new Set(['autosum', 'fillHome', 'clearFormat']);

const macHomeCommand = (candidate: RibbonCommand): RibbonCommand => {
  const commandWithMacTitle = {
    ...candidate,
    title: candidate.title.replace(/\bCtrl\+/g, 'Cmd+'),
  };
  if (!MAC_HOME_ICON_ONLY_COMMAND_IDS.has(candidate.id)) return commandWithMacTitle;
  return {
    ...commandWithMacTitle,
    kind: 'button',
    layout: 'stacked',
    className: [candidate.className, 'fc-tb__rb--mac-icon-only'].filter(Boolean).join(' '),
  };
};

const macHomeTab = (home: RibbonTabModel): RibbonTabModel => ({
  ...home,
  groups: [
    ...home.groups.map((current) => {
      if (current.variant !== 'font') {
        return {
          ...current,
          commands: current.commands.map(macHomeCommand),
        };
      }
      const phonetic = current.commands.find((candidate) => candidate.id === 'editPhonetic');
      return {
        ...current,
        commands: [
          ...current.commands
            .filter((candidate) => candidate.id !== 'strike' && candidate.id !== 'editPhonetic')
            .map(macHomeCommand),
          ...(phonetic ? [phonetic] : []),
        ],
      };
    }),
    group({ ja: 'アドイン', en: 'Add-ins' }, [
      disabledCommand('mac.home.addins', t('addins'), 'addIn', OFFICE_REQUIRED),
    ]),
  ],
});

/** Builds the Office 365 Mac command surface. The shared Home model is passed
 *  in by ribbon-model.ts so its established formatting controls and active
 *  state projection remain identical across profiles. */
export const buildMacRibbonModel = (lang: ToolbarLang, home: RibbonTabModel): RibbonTabModel[] => {
  const localizeSelectOptions = (
    commandId: string,
    options: readonly RibbonOption[] | undefined,
  ): readonly RibbonOption[] | undefined => {
    if (!options) return undefined;
    if (
      commandId !== 'scaleWidth' &&
      commandId !== 'scaleHeight' &&
      commandId !== 'sheetViewSelect'
    ) {
      return options;
    }
    const pageLabel = (value: string): string => {
      if (lang === 'en') {
        if (value === '0') return 'Automatic';
        if (value === 'custom') return 'Custom';
        return `${value} ${value === '1' ? 'page' : 'pages'}`;
      }
      if (value === '0') return '自動';
      if (value === 'custom') return 'ユーザー設定';
      return `${value} ページ`;
    };
    return options.map((option) => ({
      ...option,
      label:
        commandId === 'sheetViewSelect' && option.value === 'current'
          ? lang === 'ja'
            ? '現在のビュー'
            : 'Current View'
          : pageLabel(option.value),
    }));
  };
  const localize = (tab: RibbonTabModel): RibbonTabModel => ({
    ...tab,
    label: labels[tab.id]?.[lang] ?? tab.label,
    groups: tab.groups.map((g) => ({
      ...g,
      title: GROUP_LABELS[g.title]?.[lang] ?? g.title,
      commands: g.commands.map((c) => ({
        ...c,
        title:
          macRibbonLabelForCommand(c.id, lang) === c.id
            ? c.title
            : macRibbonLabelForCommand(c.id, lang),
        label:
          macRibbonLabelForCommand(c.id, lang) === c.id
            ? c.label
            : macRibbonLabelForCommand(c.id, lang),
        disabledReason: macRibbonDisabledReason(c.disabledReason, lang),
        options: localizeSelectOptions(c.id, c.options),
      })),
    })),
  });
  return [
    macHomeTab(home),
    insertTab(),
    drawTab(),
    pageLayoutTab(),
    formulasTab(),
    dataTab(),
    reviewTab(),
    viewTab(),
    automateTab(),
  ].map(localize);
};

const macCatalogHome: RibbonTabModel = { id: 'home', label: 'Home', groups: [] };
const macCatalog = buildMacRibbonModel('en', macCatalogHome).filter((tab) => tab.id !== 'home');
const macMenuLeafCommandIds = Object.values(MAC_RIBBON_MENU_ITEMS).flatMap((items) =>
  items.map((item) => item.id),
);

export const MAC_RIBBON_COMMAND_IDS: readonly string[] = Object.freeze(
  Array.from(
    new Set([
      ...macCatalog.flatMap((tab) => tab.groups.flatMap((g) => g.commands.map((c) => c.id))),
      ...macMenuLeafCommandIds,
      ...MAC_RIBBON_MENU_COMMANDS,
      'mac.home.addins',
      ...allFunctionNames().map((name) => `mac.function.${name}`),
    ]),
  ),
);

export const MAC_RIBBON_DISABLED_COMMAND_IDS: ReadonlySet<string> = new Set([
  'mac.home.addins',
  ...macCatalog.flatMap((tab) =>
    tab.groups.flatMap((g) => g.commands.filter((c) => c.disabled).map((c) => c.id)),
  ),
  ...Object.values(MAC_RIBBON_MENU_ITEMS).flatMap((items) =>
    items.filter((item) => item.disabled).map((item) => item.id),
  ),
]);

const MAC_ACTIVE_STATE_KEYS: Readonly<Record<string, keyof ReturnType<typeof projectActiveState>>> =
  {
    'mac.page.showGridlines': 'gridlinesVisible',
    'mac.page.printGridlines': 'printGridlines',
    'mac.page.showHeadings': 'headingsVisible',
    'mac.page.printHeadings': 'printHeadings',
    'mac.view.standard': 'workbookView',
    'mac.view.pageBreakPreview': 'workbookView',
    'mac.view.pageLayout': 'workbookView',
    'mac.view.gridlines': 'gridlinesVisible',
    'mac.view.headings': 'headingsVisible',
    'mac.view.zeros': 'zerosVisible',
    'mac.formulas.showFormulas': 'formulasVisible',
    'mac.data.filter': 'filterOn',
    'mac.review.protectSheet': 'protected',
    'mac.view.freeze': 'frozen',
  };

const MAC_POLICY_PROJECTED_COMMANDS = [
  'mac.data.sortAsc',
  'mac.data.sortDesc',
  'mac.data.sortCustom',
  'mac.data.filter',
  'mac.data.clear',
  'mac.data.reapply',
  'mac.data.advancedFilter',
  'mac.formulas.removeArrows.all',
  'mac.formulas.removeArrows.precedents',
  'mac.formulas.removeArrows.dependents',
  'mac.formulas.errorCheck.run',
  'mac.formulas.errorCheck.trace',
  'mac.formulas.errorCheck.ignore',
  'mac.data.validation.settings',
  'mac.data.validation.circleInvalid',
  'mac.data.validation.clearCircles',
  'mac.data.validation.clearRules',
] as const;

const FILTER_RANGE_REQUIRED: Label = {
  ja: '適用中のフィルター範囲はありません。',
  en: 'There is no active filter range.',
};
const FILTER_CRITERIA_REQUIRED: Label = {
  ja: '再適用できるフィルター条件はありません。',
  en: 'There are no filter criteria to reapply.',
};

const hasValidationInSelection = (instance: SpreadsheetInstance): boolean => {
  const { format, selection } = instance.store.getState();
  const { sheet, r0, r1, c0, c1 } = selection.range;
  const area = (r1 - r0 + 1) * (c1 - c0 + 1);
  if (area <= format.formats.size) {
    for (let row = r0; row <= r1; row += 1) {
      for (let col = c0; col <= c1; col += 1) {
        if (format.formats.get(addrKey({ sheet, row, col }))?.validation) return true;
      }
    }
    return false;
  }
  for (const [key, cellFormat] of format.formats) {
    if (!cellFormat.validation) continue;
    const [keySheet, row, col] = key.split(':').map(Number);
    if (
      keySheet === sheet &&
      row !== undefined &&
      col !== undefined &&
      row >= r0 &&
      row <= r1 &&
      col >= c0 &&
      col <= c1
    ) {
      return true;
    }
  }
  return false;
};

const contextualDisabledReason = (instance: SpreadsheetInstance, id: string): string | null => {
  const state = instance.store.getState();
  const strings = instance.i18n.strings.ribbonMenu;
  const lang = toolbarLangForLocale(instance.i18n.locale);
  if (id === 'mac.data.clear' && !state.ui.filterRange) return text(lang, FILTER_RANGE_REQUIRED);
  if (id === 'mac.data.reapply' && state.ui.filterCriteria.length === 0) {
    return text(lang, FILTER_CRITERIA_REQUIRED);
  }

  if (id.startsWith('mac.formulas.removeArrows.')) {
    const traces = state.traces.items;
    const hasPrecedents = traces.some((trace) => trace.kind === 'precedent');
    const hasDependents = traces.some((trace) => trace.kind === 'dependent');
    if (id === 'mac.formulas.removeArrows.all' && !hasPrecedents && !hasDependents) {
      return strings.removeArrowsRequiresAny;
    }
    if (id === 'mac.formulas.removeArrows.precedents' && !hasPrecedents) {
      return strings.removePrecedentArrowsRequiresAny;
    }
    if (id === 'mac.formulas.removeArrows.dependents' && !hasDependents) {
      return strings.removeDependentArrowsRequiresAny;
    }
  }

  if (id === 'mac.formulas.errorCheck.trace' || id === 'mac.formulas.errorCheck.ignore') {
    const active = state.selection.active;
    const cell = state.data.cells.get(addrKey(active));
    if (!cell?.formula || !cellValueIsFormulaError(cell.value)) {
      return strings.traceErrorRequiresFormulaError;
    }
  }

  if (id === 'mac.data.validation.circleInvalid' || id === 'mac.data.validation.clearRules') {
    if (!hasValidationInSelection(instance)) return strings.validationRequiresRules;
  }
  if (
    id === 'mac.data.validation.clearCircles' &&
    state.errorIndicators.validationCircles.size === 0
  ) {
    return strings.validationClearCirclesRequiresAny;
  }
  return null;
};

/** Projects Mac alias buttons onto the same store-derived state as the
 *  generic ribbon. Hosts call this after an instance mutation; it is kept
 *  separate from the render pass so active controls do not require a full
 *  menu rebuild. */
export const projectMacRibbonState = (
  host: HTMLElement,
  instance: SpreadsheetInstance | null,
): void => {
  if (!instance) return;
  const lang = toolbarLangForLocale(instance.i18n.locale);
  const active = projectActiveState(instance);
  const ink = getMacInk(instance);
  const inkTool = ink?.getTool();
  const drawingStates: Readonly<Record<string, boolean>> = {
    'mac.draw.toggle': ink?.isActive() ?? false,
    'mac.draw.eraser': inkTool === 'eraser',
    'mac.draw.penBlack': inkTool === 'pen-black',
    'mac.draw.penRed': inkTool === 'pen-red',
    'mac.draw.pencil': inkTool === 'pencil',
    'mac.draw.highlighter': inkTool === 'highlighter',
    'mac.draw.trackpad': ink?.isActive() === true && ink.getTrackpadMode(),
  };
  for (const [command, pressed] of Object.entries(drawingStates)) {
    for (const button of host.querySelectorAll<HTMLButtonElement>(
      `[data-ribbon-command="${command}"]`,
    )) {
      button.classList.toggle('fc-tb__rb--active', pressed);
      button.setAttribute('aria-pressed', String(pressed));
    }
  }

  for (const [command, key] of Object.entries(MAC_ACTIVE_STATE_KEYS)) {
    let pressed = Boolean(active[key]);
    if (
      command === 'mac.view.standard' ||
      command === 'mac.view.pageLayout' ||
      command === 'mac.view.pageBreakPreview'
    ) {
      const view = command.endsWith('standard')
        ? 'normal'
        : command.endsWith('pageLayout')
          ? 'pageLayout'
          : 'pageBreakPreview';
      pressed = active.workbookView === view;
    }
    for (const button of host.querySelectorAll<HTMLButtonElement>(
      `[data-ribbon-command="${command}"]`,
    )) {
      button.classList.toggle('fc-tb__rb--active', pressed);
      button.setAttribute('aria-pressed', pressed ? 'true' : 'false');
    }
  }

  for (const id of MAC_POLICY_PROJECTED_COMMANDS) {
    const decision = canExecuteBuiltIn(instance.store, id, 'ribbon');
    const contextualReason = contextualDisabledReason(instance, id);
    const disabled = !decision.allowed || contextualReason !== null;
    const reason = decision.allowed
      ? contextualReason
      : instance.i18n.strings.backstage.commandUnavailable;
    for (const button of host.querySelectorAll<HTMLButtonElement>(
      `[data-ribbon-command="${id}"]`,
    )) {
      const titlePrefix =
        button.getAttribute('aria-label')?.trim() || macRibbonLabelForCommand(id, lang);
      projectDisabledState(button, disabled, reason, {
        datasetKey: 'ribbonDisabledReason',
        titlePrefix,
      });
    }
  }
};
