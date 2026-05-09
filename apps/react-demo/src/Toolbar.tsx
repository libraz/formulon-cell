import {
  Accessibility24Regular,
  Add24Regular,
  AlignCenterHorizontal24Regular,
  AlignLeft24Regular,
  AlignRight24Regular,
  AlignSpaceEvenlyVertical24Regular,
  AppsAddIn24Regular,
  ArrowRedo24Regular,
  ArrowSortDown24Regular,
  ArrowSortUp24Regular,
  ArrowTrendingLines24Regular,
  ArrowUndo24Regular,
  Autosum24Regular,
  Book24Regular,
  BorderAll24Regular,
  Calculator24Regular,
  ChartMultiple24Regular,
  Checkmark16Regular,
  ChevronDown12Regular,
  ClipboardPaste24Regular,
  Code24Regular,
  Comma24Regular,
  Comment24Regular,
  CommentAdd24Regular,
  CommentMultiple24Regular,
  Copy24Regular,
  CurrencyDollarEuro24Regular,
  Cut24Regular,
  DecimalArrowLeft24Regular,
  DecimalArrowRight24Regular,
  Delete24Regular,
  DocumentMargins24Regular,
  DocumentOnePage24Regular,
  DocumentPdf24Regular,
  Eraser24Regular,
  Filter24Regular,
  FontDecrease24Regular,
  FontIncrease24Regular,
  Highlight24Regular,
  Link24Regular,
  LockClosed24Regular,
  Merge24Regular,
  Orientation24Regular,
  PaintBrush24Regular,
  Pen24Regular,
  Print24Regular,
  Search24Regular,
  SearchSquare24Regular,
  Settings24Regular,
  Shield24Regular,
  TableDeleteColumn24Regular,
  TableDeleteRow24Regular,
  TableDismiss24Regular,
  TableFreezeColumnAndRow24Regular,
  TableInsertColumn24Regular,
  TableInsertRow24Regular,
  TableLightning24Regular,
  TableSettings24Regular,
  TableSimple24Regular,
  TagPercent24Regular,
  TextAlignCenter24Regular,
  TextBold24Regular,
  TextBulletListSquare24Regular,
  TextColor24Regular,
  TextItalic24Regular,
  TextProofingTools24Regular,
  TextStrikethrough24Regular,
  TextUnderline24Regular,
  TextWrap24Regular,
  Translate24Regular,
  Window24Regular,
  ZoomFit24Regular,
} from '@fluentui/react-icons';
import {
  applyMerge,
  applyUnmerge,
  autoSum,
  bumpDecimals,
  type CellBorderStyle,
  clearFilter,
  clearFormat,
  commentAt,
  cycleBorders,
  cycleCurrency,
  cyclePercent,
  deleteCols,
  deleteRows,
  formatAsTable,
  hiddenInSelection,
  hideCols,
  hideRows,
  insertCols,
  insertRows,
  type MarginPreset,
  marginPresetOf,
  mutators,
  type PageOrientation,
  type PaperSize,
  pageSetupForSheet,
  recordFormatChange,
  recordPageSetupChange,
  removeDuplicates,
  type SpreadsheetInstance,
  setAlign,
  setAutoFilter,
  setBorderPreset,
  setFillColor,
  setFont,
  setFontColor,
  setFreezePanes,
  setMarginPreset,
  setNumFmt,
  setPageOrientation,
  setPaperSize,
  setSheetZoom,
  setVAlign,
  showCols,
  showRows,
  sortRange,
  toggleBold,
  toggleItalic,
  toggleStrike,
  toggleUnderline,
  toggleWrap,
} from '@libraz/formulon-cell';
import { type ReactElement, useCallback, useEffect, useRef, useState } from 'react';

interface DropdownOption<V extends string | number> {
  value: V;
  label: string;
}

interface DropdownProps<V extends string | number> {
  title: string;
  value: V;
  options: readonly DropdownOption<V>[];
  onChange: (value: V) => void;
  disabled?: boolean;
  className?: string;
  /** Optional override of what's shown in the closed display. Defaults to the
   *  option label matching `value`, or the raw `value` if no label matches.
   *  Used for the font-name dropdown so unknown faces still render. */
  display?: string;
}

function Dropdown<V extends string | number>({
  title,
  value,
  options,
  onChange,
  disabled,
  className,
  display,
}: DropdownProps<V>): ReactElement {
  const [open, setOpen] = useState(false);
  const wrapRef = useRef<HTMLDivElement | null>(null);
  const listRef = useRef<HTMLDivElement | null>(null);
  const matched = options.find((o) => o.value === value);
  const shown = display ?? matched?.label ?? String(value);

  useEffect(() => {
    if (!open) return;
    const onDocDown = (e: MouseEvent): void => {
      const node = wrapRef.current;
      if (!node) return;
      if (e.target instanceof Node && node.contains(e.target)) return;
      setOpen(false);
    };
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') {
        e.preventDefault();
        setOpen(false);
      }
    };
    document.addEventListener('mousedown', onDocDown, true);
    document.addEventListener('keydown', onKey, true);
    return () => {
      document.removeEventListener('mousedown', onDocDown, true);
      document.removeEventListener('keydown', onKey, true);
    };
  }, [open]);

  // Scroll the active row into view when opening, so long lists (font sizes)
  // don't strand the user at the top.
  useEffect(() => {
    if (!open) return;
    const list = listRef.current;
    if (!list) return;
    const sel = list.querySelector<HTMLElement>('[aria-selected="true"]');
    sel?.scrollIntoView({ block: 'nearest' });
  }, [open]);

  return (
    <div
      ref={wrapRef}
      className={`demo__rb-dd${className ? ` ${className}` : ''}${
        open ? ' demo__rb-dd--open' : ''
      }`}
    >
      <button
        type="button"
        className="demo__rb-dd__btn"
        title={title}
        aria-label={title}
        aria-haspopup="listbox"
        aria-expanded={open}
        disabled={disabled}
        onClick={() => setOpen((o) => !o)}
        onKeyDown={(e) => {
          if (e.key === 'ArrowDown' || e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            setOpen(true);
          }
        }}
      >
        <span className="demo__rb-dd__value">{shown}</span>
        <ChevronDown12Regular className="demo__rb-dd__chev" />
      </button>
      {open ? (
        <div
          ref={listRef}
          className="demo__rb-dd__list"
          role="listbox"
          aria-label={title}
          tabIndex={-1}
        >
          {options.map((o) => {
            const selected = o.value === value;
            return (
              <button
                key={o.value}
                type="button"
                role="option"
                aria-selected={selected}
                className={`demo__rb-dd__opt${selected ? ' demo__rb-dd__opt--selected' : ''}`}
                onClick={() => {
                  onChange(o.value);
                  setOpen(false);
                }}
              >
                <span className="demo__rb-dd__check" aria-hidden="true">
                  {selected ? <Checkmark16Regular /> : null}
                </span>
                <span className="demo__rb-dd__label">{o.label}</span>
              </button>
            );
          })}
        </div>
      ) : null}
    </div>
  );
}

interface Props {
  instance: SpreadsheetInstance | null;
  activeTab: RibbonTab;
  onTabChange: (tab: RibbonTab) => void;
  locale: string;
}

export type RibbonTab =
  | 'file'
  | 'home'
  | 'insert'
  | 'draw'
  | 'pageLayout'
  | 'formulas'
  | 'data'
  | 'review'
  | 'view'
  | 'automate'
  | 'acrobat';

const RIBBON_TAB_LABELS: Record<RibbonTab, { en: string; ja: string }> = {
  file: { en: 'File', ja: 'ファイル' },
  home: { en: 'Home', ja: 'ホーム' },
  insert: { en: 'Insert', ja: '挿入' },
  draw: { en: 'Draw', ja: '描画' },
  pageLayout: { en: 'Page Layout', ja: 'ページ レイアウト' },
  formulas: { en: 'Formulas', ja: '数式' },
  data: { en: 'Data', ja: 'データ' },
  review: { en: 'Review', ja: '校閲' },
  view: { en: 'View', ja: '表示' },
  automate: { en: 'Automate', ja: '自動化' },
  acrobat: { en: 'Acrobat', ja: 'Acrobat' },
};

interface ActiveState {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  alignLeft: boolean;
  alignCenter: boolean;
  alignRight: boolean;
  currency: boolean;
  percent: boolean;
  frozen: boolean;
  filterOn: boolean;
  rowsHidden: boolean;
  colsHidden: boolean;
  protected: boolean;
  zoom: number;
  fontFamily: string;
  fontSize: number;
  fontColor: string;
  fillColor: string;
  formatPainterArmed: boolean;
  hasComment: boolean;
  pageOrientation: PageOrientation;
  paperSize: PaperSize;
  /** Closest named preset for the active sheet's margins, or `null` when
   *  the user has set custom values via the Page Setup dialog. */
  marginPreset: MarginPreset | null;
}

const EMPTY: ActiveState = {
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  alignLeft: false,
  alignCenter: false,
  alignRight: false,
  currency: false,
  percent: false,
  frozen: false,
  filterOn: false,
  rowsHidden: false,
  colsHidden: false,
  protected: false,
  zoom: 1,
  fontFamily: 'Aptos',
  fontSize: 11,
  fontColor: '#201f1e',
  fillColor: '#ffffff',
  formatPainterArmed: false,
  hasComment: false,
  pageOrientation: 'portrait',
  paperSize: 'A4',
  marginPreset: 'normal',
};

const FONT_FAMILIES = ['Aptos', 'Calibri', 'Arial', 'Segoe UI', 'Times New Roman', 'Consolas'];
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36];

type BorderPreset =
  | 'none'
  | 'outline'
  | 'all'
  | 'top'
  | 'bottom'
  | 'left'
  | 'right'
  | 'doubleBottom';

const BORDER_STYLES: { value: CellBorderStyle; label: string }[] = [
  { value: 'thin', label: 'Thin' },
  { value: 'medium', label: 'Medium' },
  { value: 'thick', label: 'Thick' },
  { value: 'dashed', label: 'Dashed' },
  { value: 'dotted', label: 'Dotted' },
  { value: 'double', label: 'Double' },
];

const BORDER_PRESETS: { value: BorderPreset; label: string }[] = [
  { value: 'none', label: 'No Border' },
  { value: 'outline', label: 'Outside Borders' },
  { value: 'all', label: 'All Borders' },
  { value: 'top', label: 'Top Border' },
  { value: 'bottom', label: 'Bottom Border' },
  { value: 'left', label: 'Left Border' },
  { value: 'right', label: 'Right Border' },
  { value: 'doubleBottom', label: 'Double Bottom' },
];

const project = (inst: SpreadsheetInstance): ActiveState => {
  const s = inst.store.getState();
  const a = s.selection.active;
  const r = s.selection.range;
  const f = s.format.formats.get(`${a.sheet}:${a.row}:${a.col}`);
  const setup = pageSetupForSheet(s, s.data.sheetIndex);
  return {
    bold: !!f?.bold,
    italic: !!f?.italic,
    underline: !!f?.underline,
    strike: !!f?.strike,
    alignLeft: f?.align === 'left',
    alignCenter: f?.align === 'center',
    alignRight: f?.align === 'right',
    currency: f?.numFmt?.kind === 'currency',
    percent: f?.numFmt?.kind === 'percent',
    frozen: s.layout.freezeRows > 0 || s.layout.freezeCols > 0,
    filterOn: s.ui.filterRange != null,
    rowsHidden: hiddenInSelection(s.layout, 'row', r.r0, r.r1).length > 0,
    colsHidden: hiddenInSelection(s.layout, 'col', r.c0, r.c1).length > 0,
    protected: inst.isSheetProtected(),
    zoom: s.viewport.zoom,
    fontFamily: f?.fontFamily ?? 'Aptos',
    fontSize: f?.fontSize ?? 11,
    fontColor: f?.color ?? '#201f1e',
    fillColor: f?.fill ?? '#ffffff',
    formatPainterArmed: !!inst.formatPainter?.isActive(),
    hasComment: commentAt(s, a) != null,
    pageOrientation: setup.orientation,
    paperSize: setup.paperSize,
    marginPreset: marginPresetOf(setup.margins),
  };
};

type IconName =
  | 'accessibility'
  | 'add'
  | 'paste'
  | 'cut'
  | 'copy'
  | 'paint'
  | 'undo'
  | 'redo'
  | 'fontGrow'
  | 'fontShrink'
  | 'bold'
  | 'italic'
  | 'underline'
  | 'strike'
  | 'fontColor'
  | 'fillColor'
  | 'top'
  | 'middle'
  | 'currency'
  | 'percent'
  | 'comma'
  | 'decDown'
  | 'decUp'
  | 'autosum'
  | 'alignLeft'
  | 'alignCenter'
  | 'alignRight'
  | 'borders'
  | 'merge'
  | 'wrap'
  | 'freeze'
  | 'insertRows'
  | 'deleteRows'
  | 'insertCols'
  | 'deleteCols'
  | 'filter'
  | 'sortAsc'
  | 'sortDesc'
  | 'table'
  | 'tableStyle'
  | 'conditional'
  | 'formatCells'
  | 'removeDuplicates'
  | 'link'
  | 'pen'
  | 'eraser'
  | 'page'
  | 'margins'
  | 'orientation'
  | 'scale'
  | 'print'
  | 'function'
  | 'names'
  | 'trace'
  | 'dependents'
  | 'clearArrows'
  | 'options'
  | 'watch'
  | 'comment'
  | 'commentAdd'
  | 'commentMultiple'
  | 'protect'
  | 'zoom'
  | 'script'
  | 'addIn'
  | 'pdf'
  | 'goTo'
  | 'find'
  | 'findSelect'
  | 'translate'
  | 'spelling'
  | 'chart'
  | 'clear';

const Icon = ({ name }: { name: IconName }): ReactElement => {
  const common = { className: 'demo__rb-icon', 'aria-hidden': true };
  switch (name) {
    case 'accessibility':
      return <Accessibility24Regular {...common} />;
    case 'add':
      return <Add24Regular {...common} />;
    case 'paste':
      return <ClipboardPaste24Regular {...common} />;
    case 'cut':
      return <Cut24Regular {...common} />;
    case 'copy':
      return <Copy24Regular {...common} />;
    case 'paint':
      return <PaintBrush24Regular {...common} />;
    case 'undo':
      return <ArrowUndo24Regular {...common} />;
    case 'redo':
      return <ArrowRedo24Regular {...common} />;
    case 'fontGrow':
      return <FontIncrease24Regular {...common} />;
    case 'fontShrink':
      return <FontDecrease24Regular {...common} />;
    case 'bold':
      return <TextBold24Regular {...common} />;
    case 'italic':
      return <TextItalic24Regular {...common} />;
    case 'underline':
      return <TextUnderline24Regular {...common} />;
    case 'strike':
      return <TextStrikethrough24Regular {...common} />;
    case 'fontColor':
      return <TextColor24Regular {...common} />;
    case 'fillColor':
      return <Highlight24Regular {...common} />;
    case 'top':
      return <AlignSpaceEvenlyVertical24Regular {...common} />;
    case 'middle':
      return <TextAlignCenter24Regular {...common} />;
    case 'currency':
      return <CurrencyDollarEuro24Regular {...common} />;
    case 'percent':
      return <TagPercent24Regular {...common} />;
    case 'comma':
      return <Comma24Regular {...common} />;
    case 'decDown':
      return <DecimalArrowLeft24Regular {...common} />;
    case 'decUp':
      return <DecimalArrowRight24Regular {...common} />;
    case 'autosum':
      return <Autosum24Regular {...common} />;
    case 'alignLeft':
      return <AlignLeft24Regular {...common} />;
    case 'alignCenter':
      return <AlignCenterHorizontal24Regular {...common} />;
    case 'alignRight':
      return <AlignRight24Regular {...common} />;
    case 'borders':
      return <BorderAll24Regular {...common} />;
    case 'merge':
      return <Merge24Regular {...common} />;
    case 'wrap':
      return <TextWrap24Regular {...common} />;
    case 'freeze':
      return <TableFreezeColumnAndRow24Regular {...common} />;
    case 'insertRows':
      return <TableInsertRow24Regular {...common} />;
    case 'deleteRows':
      return <TableDeleteRow24Regular {...common} />;
    case 'insertCols':
      return <TableInsertColumn24Regular {...common} />;
    case 'deleteCols':
      return <TableDeleteColumn24Regular {...common} />;
    case 'filter':
      return <Filter24Regular {...common} />;
    case 'sortAsc':
      return <ArrowSortUp24Regular {...common} />;
    case 'sortDesc':
      return <ArrowSortDown24Regular {...common} />;
    case 'table':
      return <TableSimple24Regular {...common} />;
    case 'tableStyle':
      return <TableSettings24Regular {...common} />;
    case 'conditional':
      return <TableLightning24Regular {...common} />;
    case 'formatCells':
      return <TableSettings24Regular {...common} />;
    case 'removeDuplicates':
      return <TableDismiss24Regular {...common} />;
    case 'link':
      return <Link24Regular {...common} />;
    case 'pen':
      return <Pen24Regular {...common} />;
    case 'eraser':
      return <Eraser24Regular {...common} />;
    case 'page':
      return <DocumentOnePage24Regular {...common} />;
    case 'margins':
      return <DocumentMargins24Regular {...common} />;
    case 'orientation':
      return <Orientation24Regular {...common} />;
    case 'scale':
      return <ZoomFit24Regular {...common} />;
    case 'print':
      return <Print24Regular {...common} />;
    case 'function':
      return <Calculator24Regular {...common} />;
    case 'names':
      return <Book24Regular {...common} />;
    case 'trace':
      return <ArrowTrendingLines24Regular {...common} />;
    case 'dependents':
      return <TextBulletListSquare24Regular {...common} />;
    case 'clearArrows':
      return <Delete24Regular {...common} />;
    case 'options':
      return <Settings24Regular {...common} />;
    case 'watch':
      return <Window24Regular {...common} />;
    case 'comment':
      return <Comment24Regular {...common} />;
    case 'protect':
      return <LockClosed24Regular {...common} />;
    case 'zoom':
      return <ZoomFit24Regular {...common} />;
    case 'script':
      return <Code24Regular {...common} />;
    case 'addIn':
      return <AppsAddIn24Regular {...common} />;
    case 'pdf':
      return <DocumentPdf24Regular {...common} />;
    case 'goTo':
      return <Search24Regular {...common} />;
    case 'find':
      return <Search24Regular {...common} />;
    case 'findSelect':
      return <SearchSquare24Regular {...common} />;
    case 'translate':
      return <Translate24Regular {...common} />;
    case 'spelling':
      return <TextProofingTools24Regular {...common} />;
    case 'chart':
      return <ChartMultiple24Regular {...common} />;
    case 'commentAdd':
      return <CommentAdd24Regular {...common} />;
    case 'commentMultiple':
      return <CommentMultiple24Regular {...common} />;
    case 'clear':
      return <Shield24Regular {...common} />;
  }
};

export const Toolbar = ({ instance, activeTab, onTabChange, locale }: Props): ReactElement => {
  const [active, setActive] = useState<ActiveState>(EMPTY);
  const [borderStyle, setBorderStyle] = useState<CellBorderStyle>('thin');
  const lang = locale === 'ja' ? 'ja' : 'en';
  const tr =
    lang === 'ja'
      ? {
          workbook: 'ブック',
          inspect: '検査',
          clipboard: 'クリップボード',
          paste: 'ペースト',
          cut: '切り取り',
          copy: 'コピー',
          formatPainter: '書式のコピー',
          number: '数値',
          font: 'フォント',
          alignment: '配置',
          cells: 'セル',
          editing: '編集',
          styles: 'スタイル',
          tables: 'テーブル',
          definedNames: '定義された名前',
          dataTools: 'データ ツール',
          window: 'ウィンドウ',
          names: '名前',
          functionLibrary: '関数ライブラリ',
          formulaAuditing: 'ワークシート分析',
          calculation: '計算方法',
          sortFilter: '並べ替えとフィルター',
          outline: 'アウトライン',
          workbookViews: 'ブックの表示',
          zoom: 'ズーム',
          protection: '保護',
          pageSetup: 'ページ設定',
          margins: '余白',
          orientation: '印刷の向き',
          scale: '拡大縮小',
          comments: 'コメント',
          accessibility: 'アクセシビリティ',
          script: 'スクリプト',
          addIn: 'アドイン',
          pdf: 'PDF',
          print: '印刷',
          links: 'リンク',
          formatCells: 'セルの書式設定',
          goTo: 'ジャンプ',
          general: '標準',
          clear: 'クリア',
          top: '上揃え',
          middle: '中央揃え',
          cellStyles: 'セル スタイル',
          conditional: '条件付き書式',
          rules: 'ルール',
          tracePrecedents: '参照元',
          traceDependents: '参照先',
          removeArrows: '矢印の削除',
          options: 'オプション',
          watch: 'ウォッチ',
          removeDuplicates: '重複の削除',
          showRows: '行の再表示',
          hideRows: '行を表示しない',
          showCols: '列の再表示',
          hideCols: '列を表示しない',
          protect: '保護',
          unprotect: '保護解除',
          findSelect: '検索と選択',
          find: '検索',
          replace: '置換',
          gotoSpecial: 'セル選択',
          newComment: 'メモを挿入',
          editComment: 'メモを編集',
          translate: '翻訳',
          spelling: 'スペル チェック',
          hyperlink: 'リンク',
          pivotTable: 'ピボットテーブル',
          formatTable: 'テーブルとして書式設定',
          chart: 'グラフ',
          pasteSpecial: 'クリップボード',
          portrait: '縦',
          landscape: '横',
          paperA4: 'A4',
          paperLetter: 'レター',
          marginsNormal: '標準',
          marginsWide: '広い',
          marginsNarrow: '狭い',
          marginsCustom: 'ユーザー設定',
          recalc: '再計算',
        }
      : {
          workbook: 'Workbook',
          inspect: 'Inspect',
          clipboard: 'Clipboard',
          paste: 'Paste',
          cut: 'Cut',
          copy: 'Copy',
          formatPainter: 'Format Painter',
          number: 'Number',
          font: 'Font',
          alignment: 'Alignment',
          cells: 'Cells',
          editing: 'Editing',
          styles: 'Styles',
          tables: 'Tables',
          definedNames: 'Defined Names',
          dataTools: 'Data Tools',
          window: 'Window',
          names: 'Names',
          functionLibrary: 'Function Library',
          formulaAuditing: 'Formula Auditing',
          calculation: 'Calculation',
          sortFilter: 'Sort & Filter',
          outline: 'Outline',
          workbookViews: 'Workbook Views',
          zoom: 'Zoom',
          protection: 'Protection',
          pageSetup: 'Page setup',
          margins: 'Margins',
          orientation: 'Orientation',
          scale: 'Scale',
          comments: 'Comments',
          accessibility: 'Accessibility',
          script: 'Script',
          addIn: 'Add-ins',
          pdf: 'PDF',
          print: 'Print',
          links: 'Links',
          formatCells: 'Format cells',
          goTo: 'Go To',
          general: 'General',
          clear: 'Clear',
          top: 'Top',
          middle: 'Middle',
          cellStyles: 'Cell styles',
          conditional: 'Conditional',
          rules: 'Rules',
          tracePrecedents: 'Trace precedents',
          traceDependents: 'Trace dependents',
          removeArrows: 'Remove arrows',
          options: 'Options',
          watch: 'Watch',
          removeDuplicates: 'Remove duplicates',
          showRows: 'Show Rows',
          hideRows: 'Hide Rows',
          showCols: 'Show Cols',
          hideCols: 'Hide Cols',
          protect: 'Protect',
          unprotect: 'Unprotect',
          findSelect: 'Find & Select',
          find: 'Find',
          replace: 'Replace',
          gotoSpecial: 'Go To Special',
          newComment: 'New Note',
          editComment: 'Edit Note',
          translate: 'Translate',
          spelling: 'Spelling',
          hyperlink: 'Link',
          pivotTable: 'PivotTable',
          formatTable: 'Format as Table',
          chart: 'Chart',
          pasteSpecial: 'Paste Special',
          portrait: 'Portrait',
          landscape: 'Landscape',
          paperA4: 'A4',
          paperLetter: 'Letter',
          marginsNormal: 'Normal',
          marginsWide: 'Wide',
          marginsNarrow: 'Narrow',
          marginsCustom: 'Custom',
          recalc: 'Calculate Now',
        };
  const borderPresets =
    lang === 'ja'
      ? [
          { value: 'none' as const, label: '罫線なし' },
          { value: 'outline' as const, label: '外枠' },
          { value: 'all' as const, label: '格子' },
          { value: 'top' as const, label: '上罫線' },
          { value: 'bottom' as const, label: '下罫線' },
          { value: 'left' as const, label: '左罫線' },
          { value: 'right' as const, label: '右罫線' },
          { value: 'doubleBottom' as const, label: '下二重罫線' },
        ]
      : BORDER_PRESETS;
  const borderStyles =
    lang === 'ja'
      ? [
          { value: 'thin' as const, label: '細線' },
          { value: 'medium' as const, label: '中線' },
          { value: 'thick' as const, label: '太線' },
          { value: 'dashed' as const, label: '破線' },
          { value: 'dotted' as const, label: '点線' },
          { value: 'double' as const, label: '二重線' },
        ]
      : BORDER_STYLES;
  const ribbonTabs = (Object.keys(RIBBON_TAB_LABELS) as RibbonTab[])
    .filter((id) => id !== 'file')
    .map((id) => ({
      id,
      label: RIBBON_TAB_LABELS[id][lang],
    }));

  useEffect(() => {
    if (!instance) return;
    setActive(project(instance));
    return instance.store.subscribe(() => setActive(project(instance)));
  }, [instance]);

  const wrapFormat = useCallback(
    (
      fn: (
        state: ReturnType<SpreadsheetInstance['store']['getState']>,
        store: SpreadsheetInstance['store'],
      ) => void,
    ) => {
      if (!instance) return;
      recordFormatChange(instance.history, instance.store, () =>
        fn(instance.store.getState(), instance.store),
      );
    },
    [instance],
  );

  const onUndo = useCallback(() => instance?.undo(), [instance]);
  const onRedo = useCallback(() => instance?.redo(), [instance]);
  // Re-focus the host (canvas region) before delegating to the system
  // clipboard so the host-bound copy/cut/paste listeners run with a real
  // selection. document.execCommand still works on Safari/Chrome for copy
  // and cut; paste falls back to the same listener as Ctrl/⌘+V.
  const dispatchClipboard = useCallback(
    (kind: 'copy' | 'cut' | 'paste') => {
      if (!instance) return;
      instance.host.focus();
      try {
        document.execCommand(kind);
      } catch {
        // execCommand can throw on some browsers — swallow so the button
        // still feels like a hint rather than blowing up the chrome.
      }
    },
    [instance],
  );
  const onCopy = useCallback(() => dispatchClipboard('copy'), [dispatchClipboard]);
  const onCut = useCallback(() => dispatchClipboard('cut'), [dispatchClipboard]);
  const onPaste = useCallback(() => dispatchClipboard('paste'), [dispatchClipboard]);
  const onFormatPainter = useCallback(() => {
    instance?.formatPainter?.activate(false);
  }, [instance]);

  const onAutoSum = useCallback(() => {
    if (!instance) return;
    const result = autoSum(instance.store.getState(), instance.workbook);
    if (!result) return;
    mutators.replaceCells(instance.store, instance.workbook.cells(result.addr.sheet));
    mutators.setActive(instance.store, result.addr);
  }, [instance]);

  const onMerge = useCallback(() => {
    if (!instance) return;
    const s = instance.store.getState();
    const r = s.selection.range;
    const anchor = s.merges.byAnchor.get(`${r.sheet}:${r.r0}:${r.c0}`);
    const isExact =
      anchor &&
      r.r0 === anchor.r0 &&
      r.c0 === anchor.c0 &&
      r.r1 === anchor.r1 &&
      r.c1 === anchor.c1;
    if (isExact) applyUnmerge(instance.store, instance.workbook, instance.history, r);
    else applyMerge(instance.store, instance.workbook, instance.history, r);
  }, [instance]);

  const onBorderPreset = useCallback(
    (preset: BorderPreset) => {
      wrapFormat((s, st) => {
        setBorderPreset(s, st, preset, borderStyle);
      });
    },
    [borderStyle, wrapFormat],
  );

  const onFreezeToggle = useCallback(() => {
    if (!instance) return;
    const s = instance.store.getState();
    if (s.layout.freezeRows > 0 || s.layout.freezeCols > 0) {
      setFreezePanes(instance.store, instance.history, 0, 0, instance.workbook);
    } else {
      // Freeze rows/cols up to active cell, or first row if at A1.
      const a = s.selection.active;
      const rows = a.row === 0 && a.col === 0 ? 1 : a.row;
      const cols = a.row === 0 && a.col === 0 ? 0 : a.col;
      setFreezePanes(instance.store, instance.history, rows, cols, instance.workbook);
    }
  }, [instance]);

  const onInsertRows = useCallback(() => {
    if (!instance) return;
    const r = instance.store.getState().selection.range;
    insertRows(instance.store, instance.workbook, instance.history, r.r0, r.r1 - r.r0 + 1);
  }, [instance]);

  const onDeleteRows = useCallback(() => {
    if (!instance) return;
    const r = instance.store.getState().selection.range;
    deleteRows(instance.store, instance.workbook, instance.history, r.r0, r.r1 - r.r0 + 1);
  }, [instance]);

  const onInsertCols = useCallback(() => {
    if (!instance) return;
    const r = instance.store.getState().selection.range;
    insertCols(instance.store, instance.workbook, instance.history, r.c0, r.c1 - r.c0 + 1);
  }, [instance]);

  const onDeleteCols = useCallback(() => {
    if (!instance) return;
    const r = instance.store.getState().selection.range;
    deleteCols(instance.store, instance.workbook, instance.history, r.c0, r.c1 - r.c0 + 1);
  }, [instance]);

  const onToggleRowsHidden = useCallback(() => {
    if (!instance) return;
    const s = instance.store.getState();
    const r = s.selection.range;
    if (hiddenInSelection(s.layout, 'row', r.r0, r.r1).length > 0) {
      showRows(instance.store, instance.history, r.r0, r.r1, instance.workbook);
    } else {
      hideRows(instance.store, instance.history, r.r0, r.r1, instance.workbook);
    }
  }, [instance]);

  const onToggleColsHidden = useCallback(() => {
    if (!instance) return;
    const s = instance.store.getState();
    const r = s.selection.range;
    if (hiddenInSelection(s.layout, 'col', r.c0, r.c1).length > 0) {
      showCols(instance.store, instance.history, r.c0, r.c1, instance.workbook);
    } else {
      hideCols(instance.store, instance.history, r.c0, r.c1, instance.workbook);
    }
  }, [instance]);

  const onFilterToggle = useCallback(() => {
    if (!instance) return;
    const s = instance.store.getState();
    if (s.ui.filterRange) clearFilter(s, instance.store, s.ui.filterRange);
    else setAutoFilter(instance.store, s.selection.range);
  }, [instance]);

  const onRemoveDuplicates = useCallback(() => {
    if (!instance) return;
    const s = instance.store.getState();
    const removed = removeDuplicates(s, instance.store, instance.workbook, s.selection.range);
    if (removed > 0) {
      mutators.replaceCells(instance.store, instance.workbook.cells(s.data.sheetIndex));
    }
  }, [instance]);

  const onZoom = useCallback(
    (zoom: number) => {
      if (!instance) return;
      setSheetZoom(instance.store, zoom, instance.workbook);
    },
    [instance],
  );

  const onSort = useCallback(
    (direction: 'asc' | 'desc') => {
      if (!instance) return;
      const s = instance.store.getState();
      const ok = sortRange(s, instance.store, instance.workbook, s.selection.range, {
        byCol: s.selection.active.col,
        direction,
        hasHeader: s.selection.range.r0 < s.selection.range.r1,
      });
      if (ok) mutators.replaceCells(instance.store, instance.workbook.cells(s.data.sheetIndex));
    },
    [instance],
  );

  const onPageOrientation = useCallback(
    (next: PageOrientation) => {
      if (!instance) return;
      const sheet = instance.store.getState().data.sheetIndex;
      recordPageSetupChange(instance.history, instance.store, () => {
        setPageOrientation(instance.store, sheet, next);
      });
    },
    [instance],
  );

  const onPaperSize = useCallback(
    (next: PaperSize) => {
      if (!instance) return;
      const sheet = instance.store.getState().data.sheetIndex;
      recordPageSetupChange(instance.history, instance.store, () => {
        setPaperSize(instance.store, sheet, next);
      });
    },
    [instance],
  );

  const onMarginPreset = useCallback(
    (next: MarginPreset) => {
      if (!instance) return;
      const sheet = instance.store.getState().data.sheetIndex;
      recordPageSetupChange(instance.history, instance.store, () => {
        setMarginPreset(instance.store, sheet, next);
      });
    },
    [instance],
  );

  // Insert tab > Format as Table — applies the default session table overlay
  // to the active range. Excel opens a style picker first; ours ships a
  // single default style today, so calling the command directly is honest
  // and skips a one-option dropdown.
  const onFormatAsTable = useCallback(() => {
    if (!instance) return;
    const r = instance.store.getState().selection.range;
    formatAsTable(instance.store, r);
  }, [instance]);

  const tool = (
    id: string,
    title: string,
    label: string | ReactElement,
    onClick: () => void,
    isActive = false,
    extra = '',
    disabled = false,
  ): ReactElement => (
    <button
      key={id}
      type="button"
      className={`demo__rb${extra}${isActive ? ' demo__rb--active' : ''}`}
      title={title}
      aria-label={title}
      onClick={onClick}
      disabled={disabled || !instance}
    >
      {label}
    </button>
  );

  const iconLabel = (icon: IconName, text: string): ReactElement => (
    <>
      <Icon name={icon} />
      <span>{text}</span>
    </>
  );

  const group = (title: string, children: ReactElement[], variant = ''): ReactElement => (
    <section
      key={`${title}-${variant || 'group'}`}
      className={`demo__ribbon-group${variant ? ` demo__ribbon-group--${variant}` : ''}`}
      aria-label={title}
    >
      <div className="demo__ribbon-tools">{children}</div>
      <div className="demo__ribbon-label">{title}</div>
    </section>
  );

  const rowBreak = (id: string): ReactElement => (
    <span key={id} className="demo__rb-break" aria-hidden="true" />
  );

  const select = (
    id: string,
    title: string,
    value: string | number,
    values: readonly (string | number)[],
    onChange: (value: string) => void,
    extra = '',
  ): ReactElement => (
    <Dropdown
      key={id}
      title={title}
      value={value}
      options={values.map((v) => ({ value: v, label: String(v) }))}
      onChange={(v) => onChange(String(v))}
      disabled={!instance}
      className={extra.trim()}
      display={String(value)}
    />
  );

  const optionSelect = <T extends string>(
    id: string,
    title: string,
    value: T,
    options: readonly { value: T; label: string }[],
    onChange: (value: T) => void,
    extra = '',
  ): ReactElement => (
    <Dropdown<T>
      key={id}
      title={title}
      value={value}
      options={options}
      onChange={onChange}
      disabled={!instance}
      className={extra.trim()}
    />
  );

  const color = (
    id: string,
    title: string,
    value: string,
    onChange: (value: string) => void,
    label: ReactElement,
  ): ReactElement => (
    <label key={id} className="demo__rb-color" title={title} aria-label={title}>
      <span>{label}</span>
      <input
        type="color"
        value={value}
        disabled={!instance}
        onChange={(e) => onChange(e.currentTarget.value)}
      />
    </label>
  );

  const ribbonGroups: Record<RibbonTab, ReactElement[]> = {
    file: [
      group(tr.workbook, [
        tool(
          'pageSetup',
          'Page setup',
          iconLabel('page', tr.pageSetup),
          () => instance?.openPageSetup(),
          false,
          ' demo__rb--wide',
        ),
        tool(
          'print',
          tr.print,
          iconLabel('print', tr.print),
          () => instance?.print(),
          false,
          ' demo__rb--wide',
        ),
        tool(
          'links',
          'Edit links',
          iconLabel('link', tr.links),
          () => instance?.openExternalLinksDialog(),
          false,
          ' demo__rb--wide',
        ),
      ]),
      group(tr.inspect, [
        tool(
          'formatCells',
          'Format cells',
          iconLabel('formatCells', tr.formatCells),
          () => instance?.openFormatDialog(),
          false,
          ' demo__rb--wide',
        ),
        tool(
          'gotoSpecial',
          'Go To Special',
          iconLabel('goTo', tr.goTo),
          () => instance?.openGoToSpecial(),
          false,
          ' demo__rb--wide',
        ),
      ]),
    ],
    home: [
      group(
        tr.clipboard,
        [
          tool(
            'paste',
            tr.paste,
            <>
              <Icon name="paste" />
              <span>{tr.paste}</span>
            </>,
            onPaste,
            false,
            ' demo__rb--large',
          ),
          tool('cut', tr.cut, <Icon name="cut" />, onCut),
          tool('copy', tr.copy, <Icon name="copy" />, onCopy),
          tool(
            'formatPainter',
            tr.formatPainter,
            <Icon name="paint" />,
            onFormatPainter,
            active.formatPainterArmed,
          ),
          tool(
            'clearFormat',
            'Clear formats',
            <Icon name="clear" />,
            () => wrapFormat(clearFormat),
            false,
            ' demo__rb--wide',
          ),
        ],
        'clipboard',
      ),
      group(
        tr.font,
        [
          select(
            'fontFamily',
            'Font',
            active.fontFamily,
            FONT_FAMILIES,
            (value) => wrapFormat((s, st) => setFont(s, st, { fontFamily: value })),
            ' demo__rb-select--font',
          ),
          select('fontSize', 'Font size', active.fontSize, FONT_SIZES, (value) =>
            wrapFormat((s, st) => setFont(s, st, { fontSize: Number(value) })),
          ),
          tool('fontGrow', 'Increase font size', <Icon name="fontGrow" />, () =>
            wrapFormat((s, st) => setFont(s, st, { fontSize: active.fontSize + 1 })),
          ),
          tool('fontShrink', 'Decrease font size', <Icon name="fontShrink" />, () =>
            wrapFormat((s, st) => setFont(s, st, { fontSize: Math.max(1, active.fontSize - 1) })),
          ),
          rowBreak('font-row-2'),
          tool(
            'bold',
            'Bold (⌘B)',
            <Icon name="bold" />,
            () => wrapFormat(toggleBold),
            active.bold,
            ' demo__rb--bold',
          ),
          tool(
            'italic',
            'Italic (⌘I)',
            <Icon name="italic" />,
            () => wrapFormat(toggleItalic),
            active.italic,
            ' demo__rb--italic',
          ),
          tool(
            'underline',
            'Underline (⌘U)',
            <Icon name="underline" />,
            () => wrapFormat(toggleUnderline),
            active.underline,
            ' demo__rb--underline',
          ),
          tool(
            'strike',
            'Strikethrough',
            <Icon name="strike" />,
            () => wrapFormat(toggleStrike),
            active.strike,
            ' demo__rb--strike',
          ),
          tool('borders', 'Borders', <Icon name="borders" />, () => wrapFormat(cycleBorders)),
          optionSelect(
            'borderPreset',
            'Border pattern',
            'outline',
            borderPresets,
            onBorderPreset,
            ' demo__rb-select--border',
          ),
          optionSelect(
            'borderStyle',
            'Border line style',
            borderStyle,
            borderStyles,
            setBorderStyle,
            ' demo__rb-select--border-style',
          ),
          color(
            'fontColor',
            'Font color',
            active.fontColor,
            (value) => wrapFormat((s, st) => setFontColor(s, st, value)),
            <Icon name="fontColor" />,
          ),
          color(
            'fillColor',
            'Fill color',
            active.fillColor,
            (value) => wrapFormat((s, st) => setFillColor(s, st, value)),
            <Icon name="fillColor" />,
          ),
        ],
        'font',
      ),
      group(
        tr.alignment,
        [
          tool(
            'top',
            'Top align',
            <Icon name="top" />,
            () => wrapFormat((s, st) => setVAlign(s, st, 'top')),
            false,
          ),
          tool(
            'middle',
            'Middle align',
            <Icon name="middle" />,
            () => wrapFormat((s, st) => setVAlign(s, st, 'middle')),
            false,
          ),
          rowBreak('alignment-row-2'),
          tool(
            'alignL',
            'Align left',
            <Icon name="alignLeft" />,
            () => wrapFormat((s, st) => setAlign(s, st, 'left')),
            active.alignLeft,
          ),
          tool(
            'alignC',
            'Align center',
            <Icon name="alignCenter" />,
            () => wrapFormat((s, st) => setAlign(s, st, 'center')),
            active.alignCenter,
          ),
          tool(
            'alignR',
            'Align right',
            <Icon name="alignRight" />,
            () => wrapFormat((s, st) => setAlign(s, st, 'right')),
            active.alignRight,
          ),
          tool('wrap', 'Wrap text', <Icon name="wrap" />, () => wrapFormat(toggleWrap)),
          tool('merge', 'Merge cells', <Icon name="merge" />, onMerge),
        ],
        'alignment',
      ),
      group(
        tr.number,
        [
          tool(
            'general',
            'General number format',
            iconLabel('formatCells', tr.general),
            () => wrapFormat((s, st) => setNumFmt(s, st, { kind: 'general' })),
            false,
            ' demo__rb--wide',
          ),
          rowBreak('number-row-2'),
          tool(
            'currency',
            'Currency',
            <Icon name="currency" />,
            () => wrapFormat(cycleCurrency),
            active.currency,
            ' demo__rb--mono',
          ),
          tool(
            'percent',
            'Percent',
            <Icon name="percent" />,
            () => wrapFormat(cyclePercent),
            active.percent,
            ' demo__rb--mono',
          ),
          tool('comma', 'Comma style', <Icon name="comma" />, () =>
            wrapFormat((s, st) => setNumFmt(s, st, { kind: 'fixed', decimals: 2 })),
          ),
          tool('decDown', 'Decrease decimals', <Icon name="decDown" />, () =>
            wrapFormat((s, st) => bumpDecimals(s, st, -1)),
          ),
          tool('decUp', 'Increase decimals', <Icon name="decUp" />, () =>
            wrapFormat((s, st) => bumpDecimals(s, st, 1)),
          ),
        ],
        'number',
      ),
      group(
        tr.styles,
        [
          tool(
            'conditional',
            'Conditional formatting',
            iconLabel('conditional', tr.conditional),
            () => instance?.openConditionalDialog(),
            false,
            ' demo__rb--wide',
          ),
          tool(
            'cellStyles',
            'Cell styles',
            iconLabel('tableStyle', tr.cellStyles),
            () => instance?.openCellStylesGallery(),
            false,
            ' demo__rb--wide',
          ),
          tool(
            'rules',
            'Manage conditional formatting rules',
            iconLabel('options', tr.rules),
            () => instance?.openCfRulesDialog(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'styles',
      ),
      group(
        tr.cells,
        [
          tool('insertRows', 'Insert selected rows', <Icon name="insertRows" />, onInsertRows),
          tool('deleteRows', 'Delete selected rows', <Icon name="deleteRows" />, onDeleteRows),
          tool('insertCols', 'Insert selected columns', <Icon name="insertCols" />, onInsertCols),
          tool('deleteCols', 'Delete selected columns', <Icon name="deleteCols" />, onDeleteCols),
          tool(
            'formatCellsHome',
            'Format cells',
            iconLabel('formatCells', tr.formatCells),
            () => instance?.openFormatDialog(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'cells',
      ),
      group(
        tr.editing,
        [
          tool('autosum', 'AutoSum (Σ)', <Icon name="autosum" />, onAutoSum),
          tool('undoHome', 'Undo (⌘Z)', <Icon name="undo" />, onUndo),
          tool('redoHome', 'Redo (⌘⇧Z)', <Icon name="redo" />, onRedo),
          tool('sortAscHome', 'Sort ascending', <Icon name="sortAsc" />, () => onSort('asc')),
          tool('filterHome', 'Filter', <Icon name="filter" />, onFilterToggle, active.filterOn),
          tool(
            'findHome',
            `${tr.find} (⌘F)`,
            iconLabel('find', tr.find),
            () => instance?.openFindReplace(),
            false,
            ' demo__rb--wide',
          ),
          tool(
            'gotoSpecialHome',
            'Go To Special',
            iconLabel('goTo', tr.gotoSpecial),
            () => instance?.openGoToSpecial(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'editing',
      ),
    ],
    insert: [
      group(
        tr.tables,
        [
          tool(
            'pivotTableInsert',
            'PivotTable',
            iconLabel('table', tr.pivotTable),
            () => instance?.openPivotTableDialog(),
            false,
            ' demo__rb--wide',
          ),
          tool(
            'formatTableInsert',
            'Format as Table',
            iconLabel('tableStyle', tr.formatTable),
            onFormatAsTable,
            false,
            ' demo__rb--wide',
          ),
          tool(
            'namedRangesInsert',
            'Name manager',
            iconLabel('names', tr.names),
            () => instance?.openNamedRangeDialog(),
            false,
            ' demo__rb--wide',
          ),
          tool(
            'removeDupesInsert',
            'Remove duplicates',
            iconLabel('removeDuplicates', tr.removeDuplicates),
            onRemoveDuplicates,
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        lang === 'ja' ? 'グラフ' : 'Charts',
        [
          tool(
            'chartInsert',
            'Recommended chart',
            iconLabel('chart', tr.chart),
            () => instance?.openQuickAnalysis(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        tr.links,
        [
          tool(
            'hyperlinkInsert',
            'Insert hyperlink (⌘K)',
            iconLabel('link', tr.hyperlink),
            () => instance?.openHyperlinkDialog(),
            false,
            ' demo__rb--wide',
          ),
          tool(
            'linksInsert',
            'Edit links',
            iconLabel('link', tr.links),
            () => instance?.openExternalLinksDialog(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        tr.comments,
        [
          tool(
            'commentInsert',
            active.hasComment ? 'Edit Note' : 'New Note',
            iconLabel(
              active.hasComment ? 'commentMultiple' : 'commentAdd',
              active.hasComment ? tr.editComment : tr.newComment,
            ),
            () => instance?.openCommentDialog(),
            active.hasComment,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        lang === 'ja' ? '記号と特殊文字' : 'Symbols',
        [
          tool(
            'fxInsert',
            'Insert function (Σ)',
            iconLabel('function', 'fx'),
            () => instance?.openFunctionArguments(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
    ],
    draw: [
      group(
        RIBBON_TAB_LABELS.draw[lang],
        [
          tool(
            'drawPen',
            RIBBON_TAB_LABELS.draw[lang],
            iconLabel('pen', lang === 'ja' ? 'ペン' : 'Pen'),
            () => undefined,
            false,
            ' demo__rb--wide',
            true,
          ),
          tool(
            'drawErase',
            'Eraser',
            iconLabel('eraser', lang === 'ja' ? '消しゴム' : 'Eraser'),
            () => undefined,
            false,
            ' demo__rb--wide',
            true,
          ),
        ],
        'tiles',
      ),
    ],
    pageLayout: [
      group(
        tr.pageSetup,
        [
          optionSelect<MarginPreset | 'custom'>(
            'marginsPreset',
            tr.margins,
            active.marginPreset ?? 'custom',
            [
              { value: 'normal', label: tr.marginsNormal },
              { value: 'wide', label: tr.marginsWide },
              { value: 'narrow', label: tr.marginsNarrow },
              // "Custom" is read-only — selecting it would have to round-trip
              // through Page Setup. We include it so the closed display can
              // honestly say "Custom" when the user has bespoke margins.
              { value: 'custom', label: tr.marginsCustom },
            ],
            (next) => {
              if (next === 'custom') {
                instance?.openPageSetup();
                return;
              }
              onMarginPreset(next);
            },
            ' demo__rb-select--border',
          ),
          optionSelect(
            'orientationPreset',
            tr.orientation,
            active.pageOrientation,
            [
              { value: 'portrait' as PageOrientation, label: tr.portrait },
              { value: 'landscape' as PageOrientation, label: tr.landscape },
            ],
            onPageOrientation,
            ' demo__rb-select--border',
          ),
          optionSelect(
            'paperSizePreset',
            'Paper size',
            active.paperSize,
            [
              { value: 'A4' as PaperSize, label: 'A4' },
              { value: 'A3' as PaperSize, label: 'A3' },
              { value: 'A5' as PaperSize, label: 'A5' },
              { value: 'letter' as PaperSize, label: tr.paperLetter },
              { value: 'legal' as PaperSize, label: 'Legal' },
              { value: 'tabloid' as PaperSize, label: 'Tabloid' },
            ],
            onPaperSize,
            ' demo__rb-select--border',
          ),
          tool(
            'pageSetupAdvanced',
            'Advanced page setup',
            iconLabel('options', tr.pageSetup),
            () => instance?.openPageSetup(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        tr.print,
        [
          tool(
            'printPageLayout',
            tr.print,
            iconLabel('print', tr.print),
            () => instance?.print(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
    ],
    formulas: [
      group(
        tr.functionLibrary,
        [
          tool(
            'fx',
            'Insert function',
            <Icon name="function" />,
            () => instance?.openFunctionArguments(),
            false,
            ' demo__rb--mono',
          ),
          tool(
            'autosumFormula',
            'AutoSum (Σ)',
            <>
              <Icon name="autosum" />
              <span>{lang === 'ja' ? 'オートSUM' : 'AutoSum'}</span>
            </>,
            onAutoSum,
          ),
          tool(
            'sum',
            'SUM arguments',
            iconLabel('function', 'SUM'),
            () => instance?.openFunctionArguments('SUM'),
            false,
            ' demo__rb--mono',
          ),
          tool(
            'avg',
            'AVERAGE arguments',
            iconLabel('function', 'AVG'),
            () => instance?.openFunctionArguments('AVERAGE'),
            false,
            ' demo__rb--mono',
          ),
        ],
        'tiles',
      ),
      group(
        tr.definedNames,
        [
          tool(
            'namedRanges',
            'Name manager',
            iconLabel('names', tr.names),
            () => instance?.openNamedRangeDialog(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        tr.formulaAuditing,
        [
          tool(
            'precedents',
            'Trace precedents',
            iconLabel('trace', tr.tracePrecedents),
            () => instance?.tracePrecedents(),
            false,
            ' demo__rb--wide',
          ),
          tool(
            'dependents',
            'Trace dependents',
            iconLabel('dependents', tr.traceDependents),
            () => instance?.traceDependents(),
            false,
            ' demo__rb--wide',
          ),
          tool(
            'clearArrows',
            'Remove arrows',
            iconLabel('clearArrows', tr.removeArrows),
            () => instance?.clearTraces(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        tr.calculation,
        [
          tool(
            'recalcNow',
            'Calculate Now (F9)',
            iconLabel('autosum', tr.recalc),
            () => instance?.recalc(),
            false,
            ' demo__rb--wide',
          ),
          tool(
            'calcOptions',
            'Calculation options',
            iconLabel('options', tr.options),
            () => instance?.openIterativeDialog(),
            false,
            ' demo__rb--wide',
          ),
          tool(
            'watch',
            'Watch Window',
            iconLabel('watch', tr.watch),
            () => instance?.toggleWatchWindow(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
    ],
    data: [
      group(
        tr.sortFilter,
        [
          tool(
            'filter',
            'Filter',
            <>
              <Icon name="filter" />
              <span>{lang === 'ja' ? 'フィルター' : 'Filter'}</span>
            </>,
            onFilterToggle,
            active.filterOn,
          ),
          tool(
            'sortAsc',
            'Sort ascending',
            <>
              <Icon name="sortAsc" />
              <span>A-Z</span>
            </>,
            () => onSort('asc'),
          ),
          tool(
            'sortDesc',
            'Sort descending',
            <>
              <Icon name="sortDesc" />
              <span>Z-A</span>
            </>,
            () => onSort('desc'),
          ),
        ],
        'tiles',
      ),
      group(
        tr.dataTools,
        [
          tool(
            'removeDupes',
            'Remove duplicates',
            iconLabel('removeDuplicates', tr.removeDuplicates),
            onRemoveDuplicates,
            false,
            ' demo__rb--wide',
          ),
          tool(
            'linksData',
            'Edit links',
            iconLabel('link', tr.links),
            () => instance?.openExternalLinksDialog(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        tr.outline,
        [
          tool(
            'hideRows',
            active.rowsHidden ? 'Show selected rows' : 'Hide selected rows',
            iconLabel('table', active.rowsHidden ? tr.showRows : tr.hideRows),
            onToggleRowsHidden,
            active.rowsHidden,
            ' demo__rb--wide',
          ),
          tool(
            'hideCols',
            active.colsHidden ? 'Show selected columns' : 'Hide selected columns',
            iconLabel('table', active.colsHidden ? tr.showCols : tr.hideCols),
            onToggleColsHidden,
            active.colsHidden,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
    ],
    review: [
      group(
        lang === 'ja' ? '文章校正' : 'Proofing',
        [
          tool(
            'spellingReview',
            tr.spelling,
            iconLabel('spelling', tr.spelling),
            () => undefined,
            false,
            ' demo__rb--wide',
            true,
          ),
        ],
        'tiles',
      ),
      group(
        lang === 'ja' ? '言語' : 'Language',
        [
          tool(
            'translateReview',
            tr.translate,
            iconLabel('translate', tr.translate),
            () => undefined,
            false,
            ' demo__rb--wide',
            true,
          ),
        ],
        'tiles',
      ),
      group(
        tr.comments,
        [
          tool(
            'newCommentReview',
            active.hasComment ? 'Edit Note' : 'New Note',
            iconLabel(
              active.hasComment ? 'commentMultiple' : 'commentAdd',
              active.hasComment ? tr.editComment : tr.newComment,
            ),
            () => instance?.openCommentDialog(),
            active.hasComment,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        lang === 'ja' ? '検索' : 'Find',
        [
          tool(
            'findReview',
            `${tr.find} (⌘F)`,
            iconLabel('find', tr.find),
            () => instance?.openFindReplace(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        tr.protection,
        [
          tool(
            'protectReview',
            active.protected ? 'Unprotect sheet' : 'Protect sheet',
            iconLabel('protect', active.protected ? tr.unprotect : tr.protect),
            () => instance?.toggleSheetProtection(),
            active.protected,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        tr.accessibility,
        [
          tool(
            'accessibility',
            tr.accessibility,
            iconLabel('accessibility', tr.accessibility),
            () => undefined,
            false,
            ' demo__rb--wide',
            true,
          ),
        ],
        'tiles',
      ),
    ],
    view: [
      group(
        tr.workbookViews,
        [
          tool(
            'watchView',
            'Watch Window',
            iconLabel('watch', tr.watch),
            () => instance?.toggleWatchWindow(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        tr.window,
        [
          tool(
            'freeze',
            'Freeze panes',
            <>
              <Icon name="freeze" />
              <span>{lang === 'ja' ? 'ウィンドウ枠' : 'Freeze'}</span>
            </>,
            onFreezeToggle,
            active.frozen,
          ),
        ],
        'tiles',
      ),
      group(
        tr.zoom,
        [
          tool(
            'zoom75',
            'Zoom to 75%',
            iconLabel('zoom', '75%'),
            () => onZoom(0.75),
            active.zoom === 0.75,
            ' demo__rb--mono',
          ),
          tool(
            'zoom100',
            'Zoom to 100%',
            iconLabel('zoom', '100%'),
            () => onZoom(1),
            active.zoom === 1,
            ' demo__rb--mono',
          ),
          tool(
            'zoom125',
            'Zoom to 125%',
            iconLabel('zoom', '125%'),
            () => onZoom(1.25),
            active.zoom === 1.25,
            ' demo__rb--mono',
          ),
        ],
        'tiles',
      ),
      group(
        tr.protection,
        [
          tool(
            'protect',
            active.protected ? 'Unprotect sheet' : 'Protect sheet',
            iconLabel('protect', active.protected ? tr.unprotect : tr.protect),
            () => instance?.toggleSheetProtection(),
            active.protected,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
    ],
    automate: [
      group(
        RIBBON_TAB_LABELS.automate[lang],
        [
          tool(
            'script',
            tr.script,
            iconLabel('script', tr.script),
            () => undefined,
            false,
            ' demo__rb--wide',
            true,
          ),
        ],
        'tiles',
      ),
    ],
    acrobat: [
      group(
        tr.addIn,
        [
          tool(
            'addIn',
            tr.addIn,
            iconLabel('addIn', tr.addIn),
            () => undefined,
            false,
            ' demo__rb--wide',
            true,
          ),
        ],
        'tiles',
      ),
      group(
        tr.pdf,
        [
          tool(
            'pdf',
            tr.pdf,
            iconLabel('pdf', tr.pdf),
            () => instance?.print(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
    ],
  };

  return (
    <div className="demo__ribbon-shell">
      <div className="demo__ribbon-tabs" role="tablist" aria-label="Ribbon tabs">
        {ribbonTabs.map((tab) => (
          <button
            key={tab.id}
            type="button"
            className={`demo__ribbon-tab${activeTab === tab.id ? ' demo__ribbon-tab--active' : ''}`}
            role="tab"
            aria-selected={activeTab === tab.id}
            onClick={() => onTabChange(tab.id)}
          >
            {tab.label}
          </button>
        ))}
      </div>
      <div className="demo__ribbon" role="toolbar" aria-label={`${activeTab} ribbon`}>
        {ribbonGroups[activeTab]}
      </div>
    </div>
  );
};
