export type ToolbarLang = 'ja' | 'en';

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

export interface RibbonCommand {
  id: string;
  title: string;
  label: string;
  icon?: string;
  kind?: 'button' | 'large' | 'wide' | 'mono' | 'select' | 'color' | 'break';
  disabled?: boolean;
  options?: readonly RibbonOption[];
  className?: string;
}

export interface RibbonOption {
  value: string;
  label: string;
}

export interface RibbonGroupModel {
  title: string;
  variant?: string;
  commands: RibbonCommand[];
}

export interface RibbonTabModel {
  id: RibbonTab;
  label: string;
  groups: RibbonGroupModel[];
}

export const RIBBON_TAB_LABELS: Record<RibbonTab, { en: string; ja: string }> = {
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

export const RIBBON_KEYSHORTCUTS: Readonly<Record<string, string>> = {
  copy: 'Control+C Meta+C',
  cut: 'Control+X Meta+X',
  findHome: 'Control+F Meta+F',
  findReview: 'Control+F Meta+F',
  formatCells: 'Control+1 Meta+1',
  formatCellsHome: 'Control+1 Meta+1',
  fx: 'Shift+F3',
  fxInsert: 'Shift+F3',
  gotoSpecial: 'F5 Control+G Meta+G',
  gotoSpecialHome: 'F5 Control+G Meta+G',
  hyperlinkInsert: 'Control+K Meta+K',
  namedRanges: 'Control+F3',
  namedRangesInsert: 'Control+F3',
  paste: 'Control+V Meta+V',
  recalcNow: 'F9',
  redoHome: 'Control+Y Meta+Y Meta+Shift+Z',
  undoHome: 'Control+Z Meta+Z',
};

export const FONT_FAMILIES = [
  'Aptos',
  'Calibri',
  'Arial',
  'Segoe UI',
  'Times New Roman',
  'Consolas',
] as const;

export const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36] as const;

export const toolbarText = (lang: ToolbarLang) =>
  lang === 'ja'
    ? {
        workbook: 'ブック',
        ribbonTabs: 'リボン タブ',
        ribbon: 'リボン',
        inspect: '検査',
        clipboard: 'クリップボード',
        paste: 'ペースト',
        cut: '切り取り',
        copy: 'コピー',
        formatPainter: '書式のコピー',
        clearFormats: '書式のクリア',
        number: '数値',
        font: 'フォント',
        fontSize: 'フォント サイズ',
        increaseFontSize: 'フォント サイズの拡大',
        decreaseFontSize: 'フォント サイズの縮小',
        bold: '太字',
        italic: '斜体',
        underline: '下線',
        strikethrough: '取り消し線',
        fontColor: '文字色',
        fillColor: '塗りつぶしの色',
        alignment: '配置',
        topAlign: '上揃え',
        middleAlign: '上下中央揃え',
        alignLeft: '左揃え',
        alignCenter: '中央揃え',
        alignRight: '右揃え',
        mergeCells: 'セルの結合',
        wrapText: '折り返して全体を表示',
        cells: 'セル',
        insertRows: '選択した行を挿入',
        deleteRows: '選択した行を削除',
        insertCols: '選択した列を挿入',
        deleteCols: '選択した列を削除',
        editing: '編集',
        undo: '元に戻す',
        redo: 'やり直し',
        autoSum: 'オートSUM',
        sortAscending: '昇順で並べ替え',
        sortDescending: '降順で並べ替え',
        filter: 'フィルター',
        styles: 'スタイル',
        conditionalFormatting: '条件付き書式',
        manageRules: '条件付き書式ルールの管理',
        tables: 'テーブル',
        charts: 'グラフ',
        symbols: '記号と特殊文字',
        insertFunction: '関数の挿入',
        definedNames: '定義された名前',
        dataTools: 'データ ツール',
        window: 'ウィンドウ',
        freeze: 'ウィンドウ枠',
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
        currency: '通貨',
        percent: 'パーセント',
        commaStyle: '桁区切りスタイル',
        decreaseDecimals: '小数点以下の桁数を減らす',
        increaseDecimals: '小数点以下の桁数を増やす',
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
        proofing: '文章校正',
        language: '言語',
        translate: '翻訳',
        spelling: 'スペル チェック',
        hyperlink: 'リンク',
        pivotTable: 'ピボットテーブル',
        formatTable: 'テーブルとして書式設定',
        chart: 'グラフ',
        pasteSpecial: 'クリップボード',
        borders: '罫線',
        borderPattern: '罫線パターン',
        borderLineStyle: '罫線のスタイル',
        noBorder: '罫線なし',
        outsideBorders: '外枠',
        allBorders: '格子',
        topBorder: '上罫線',
        bottomBorder: '下罫線',
        leftBorder: '左罫線',
        rightBorder: '右罫線',
        doubleBottomBorder: '下二重罫線',
        thin: '細線',
        medium: '中線',
        thick: '太線',
        dashed: '破線',
        dotted: '点線',
        double: '二重線',
        portrait: '縦',
        landscape: '横',
        paperSize: '用紙サイズ',
        paperA4: 'A4',
        paperA3: 'A3',
        paperA5: 'A5',
        paperLetter: 'レター',
        paperLegal: 'リーガル',
        paperTabloid: 'タブロイド',
        marginsNormal: '標準',
        marginsWide: '広い',
        marginsNarrow: '狭い',
        marginsCustom: 'ユーザー設定',
        recalc: '再計算',
        pen: 'ペン',
        eraser: '消しゴム',
        disabled: '未実装',
      }
    : {
        workbook: 'Workbook',
        ribbonTabs: 'Ribbon tabs',
        ribbon: 'ribbon',
        inspect: 'Inspect',
        clipboard: 'Clipboard',
        paste: 'Paste',
        cut: 'Cut',
        copy: 'Copy',
        formatPainter: 'Format Painter',
        clearFormats: 'Clear formats',
        number: 'Number',
        font: 'Font',
        fontSize: 'Font size',
        increaseFontSize: 'Increase font size',
        decreaseFontSize: 'Decrease font size',
        bold: 'Bold',
        italic: 'Italic',
        underline: 'Underline',
        strikethrough: 'Strikethrough',
        fontColor: 'Font color',
        fillColor: 'Fill color',
        alignment: 'Alignment',
        topAlign: 'Top align',
        middleAlign: 'Middle align',
        alignLeft: 'Align left',
        alignCenter: 'Align center',
        alignRight: 'Align right',
        mergeCells: 'Merge cells',
        wrapText: 'Wrap text',
        cells: 'Cells',
        insertRows: 'Insert selected rows',
        deleteRows: 'Delete selected rows',
        insertCols: 'Insert selected columns',
        deleteCols: 'Delete selected columns',
        editing: 'Editing',
        undo: 'Undo',
        redo: 'Redo',
        autoSum: 'AutoSum',
        sortAscending: 'Sort ascending',
        sortDescending: 'Sort descending',
        filter: 'Filter',
        styles: 'Styles',
        conditionalFormatting: 'Conditional formatting',
        manageRules: 'Manage conditional formatting rules',
        tables: 'Tables',
        charts: 'Charts',
        symbols: 'Symbols',
        insertFunction: 'Insert function',
        definedNames: 'Defined Names',
        dataTools: 'Data Tools',
        window: 'Window',
        freeze: 'Freeze',
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
        currency: 'Currency',
        percent: 'Percent',
        commaStyle: 'Comma style',
        decreaseDecimals: 'Decrease decimals',
        increaseDecimals: 'Increase decimals',
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
        proofing: 'Proofing',
        language: 'Language',
        translate: 'Translate',
        spelling: 'Spelling',
        hyperlink: 'Link',
        pivotTable: 'PivotTable',
        formatTable: 'Format as Table',
        chart: 'Chart',
        pasteSpecial: 'Paste Special',
        borders: 'Borders',
        borderPattern: 'Border pattern',
        borderLineStyle: 'Border line style',
        noBorder: 'No Border',
        outsideBorders: 'Outside Borders',
        allBorders: 'All Borders',
        topBorder: 'Top Border',
        bottomBorder: 'Bottom Border',
        leftBorder: 'Left Border',
        rightBorder: 'Right Border',
        doubleBottomBorder: 'Double Bottom',
        thin: 'Thin',
        medium: 'Medium',
        thick: 'Thick',
        dashed: 'Dashed',
        dotted: 'Dotted',
        double: 'Double',
        portrait: 'Portrait',
        landscape: 'Landscape',
        paperSize: 'Paper size',
        paperA4: 'A4',
        paperA3: 'A3',
        paperA5: 'A5',
        paperLetter: 'Letter',
        paperLegal: 'Legal',
        paperTabloid: 'Tabloid',
        marginsNormal: 'Normal',
        marginsWide: 'Wide',
        marginsNarrow: 'Narrow',
        marginsCustom: 'Custom',
        recalc: 'Calculate Now',
        pen: 'Pen',
        eraser: 'Eraser',
        disabled: 'Coming soon',
      };

export type ToolbarText = ReturnType<typeof toolbarText>;

const cmd = (
  id: string,
  label: string,
  title = label,
  icon?: string,
  kind: RibbonCommand['kind'] = 'button',
  disabled = false,
): RibbonCommand => ({ id, title, label, icon, kind, disabled });

const selectCmd = (
  id: string,
  label: string,
  title: string,
  options: readonly RibbonOption[],
  className?: string,
): RibbonCommand => ({ id, title, label, kind: 'select', options, className });

const colorCmd = (id: string, label: string, title: string, icon: string): RibbonCommand => ({
  id,
  title,
  label,
  icon,
  kind: 'color',
});

const breakCmd = (id: string): RibbonCommand => ({ id, title: id, label: '', kind: 'break' });

export function buildRibbonModel(lang: ToolbarLang = 'en'): RibbonTabModel[] {
  const tr = toolbarText(lang);
  const tab = (id: RibbonTab, groups: RibbonGroupModel[]): RibbonTabModel => ({
    id,
    label: RIBBON_TAB_LABELS[id][lang],
    groups,
  });
  const group = (
    title: string,
    commands: RibbonCommand[],
    variant: RibbonGroupModel['variant'] = 'tiles',
  ): RibbonGroupModel => ({ title, commands, variant });

  return [
    tab('file', [
      group(tr.workbook, [
        cmd('pageSetup', tr.pageSetup, 'Page setup', 'page', 'wide'),
        cmd('print', tr.print, tr.print, 'print', 'wide'),
        cmd('links', tr.links, 'Edit links', 'link', 'wide'),
      ]),
      group(tr.inspect, [
        cmd('formatCells', tr.formatCells, 'Format cells', 'formatCells', 'wide'),
        cmd('gotoSpecial', tr.gotoSpecial, 'Go To Special', 'goTo', 'wide'),
      ]),
    ]),
    tab('home', [
      group(
        tr.clipboard,
        [
          cmd('paste', tr.paste, tr.paste, 'paste', 'large'),
          cmd('cut', tr.cut, tr.cut, 'cut'),
          cmd('copy', tr.copy, tr.copy, 'copy'),
          cmd('formatPainter', tr.formatPainter, tr.formatPainter, 'paint'),
          cmd('clearFormat', tr.clear, 'Clear formats', 'clear', 'wide'),
        ],
        'clipboard',
      ),
      group(
        tr.font,
        [
          selectCmd(
            'fontFamily',
            tr.font,
            tr.font,
            FONT_FAMILIES.map((font) => ({ value: font, label: font })),
            'demo__rb-select--font',
          ),
          selectCmd(
            'fontSize',
            tr.fontSize,
            tr.fontSize,
            FONT_SIZES.map((size) => ({ value: String(size), label: String(size) })),
          ),
          cmd('fontGrow', '', tr.increaseFontSize, 'fontGrow'),
          cmd('fontShrink', '', tr.decreaseFontSize, 'fontShrink'),
          breakCmd('font-row-2'),
          cmd('bold', 'B', 'Bold (⌘B)', 'bold'),
          cmd('italic', 'I', 'Italic (⌘I)', 'italic'),
          cmd('underline', 'U', 'Underline (⌘U)', 'underline'),
          cmd('strike', 'S', 'Strikethrough', 'strike'),
          cmd('borders', tr.formatCells, 'Borders', 'borders', 'wide'),
          selectCmd(
            'borderPreset',
            tr.borderPattern,
            tr.borderPattern,
            [
              { value: 'none', label: tr.noBorder },
              { value: 'outline', label: tr.outsideBorders },
              { value: 'all', label: tr.allBorders },
              { value: 'top', label: tr.topBorder },
              { value: 'bottom', label: tr.bottomBorder },
              { value: 'left', label: tr.leftBorder },
              { value: 'right', label: tr.rightBorder },
              { value: 'doubleBottom', label: tr.doubleBottomBorder },
            ],
            'demo__rb-select--border',
          ),
          selectCmd(
            'borderStyle',
            tr.borderLineStyle,
            tr.borderLineStyle,
            [
              { value: 'thin', label: tr.thin },
              { value: 'medium', label: tr.medium },
              { value: 'thick', label: tr.thick },
              { value: 'dashed', label: tr.dashed },
              { value: 'dotted', label: tr.dotted },
              { value: 'double', label: tr.double },
            ],
            'demo__rb-select--border-style',
          ),
          colorCmd('fontColor', tr.fontColor, tr.fontColor, 'fontColor'),
          colorCmd('fillColor', tr.fillColor, tr.fillColor, 'fillColor'),
        ],
        'font',
      ),
      group(
        tr.alignment,
        [
          cmd('top', tr.top, 'Top align', 'top'),
          cmd('middle', tr.middle, 'Middle align', 'middle'),
          breakCmd('alignment-row-2'),
          cmd('alignL', tr.alignLeft, 'Align left', 'alignLeft'),
          cmd('alignC', tr.alignCenter, 'Align center', 'alignCenter'),
          cmd('alignR', tr.alignRight, 'Align right', 'alignRight'),
          cmd('wrap', 'Wrap', 'Wrap text', 'wrap', 'wide'),
          cmd('merge', 'Merge', 'Merge cells', 'merge', 'wide'),
        ],
        'alignment',
      ),
      group(
        tr.number,
        [
          cmd('general', tr.general, 'General number format', 'formatCells', 'wide'),
          breakCmd('number-row-2'),
          cmd('currency', '$', 'Currency', 'currency', 'mono'),
          cmd('percent', '%', 'Percent', 'percent', 'mono'),
          cmd('comma', ',', 'Comma style', 'comma'),
          cmd('decDown', '.0', 'Decrease decimals', 'decDown', 'mono'),
          cmd('decUp', '.00', 'Increase decimals', 'decUp', 'mono'),
        ],
        'number',
      ),
      group(
        tr.styles,
        [
          cmd('conditional', tr.conditional, 'Conditional formatting', 'conditional', 'wide'),
          cmd('cellStyles', tr.cellStyles, 'Cell styles', 'tableStyle', 'wide'),
          cmd('rules', tr.rules, 'Manage conditional formatting rules', 'options', 'wide'),
        ],
        'styles',
      ),
      group(
        tr.cells,
        [
          cmd('insertRows', tr.showRows, 'Insert selected rows', 'insertRows'),
          cmd('deleteRows', tr.hideRows, 'Delete selected rows', 'deleteRows'),
          cmd('insertCols', tr.showCols, 'Insert selected columns', 'insertCols'),
          cmd('deleteCols', tr.hideCols, 'Delete selected columns', 'deleteCols'),
          cmd('formatCellsHome', tr.formatCells, 'Format cells', 'formatCells', 'wide'),
        ],
        'cells',
      ),
      group(
        tr.editing,
        [
          cmd('autosum', 'Σ', 'AutoSum (Σ)', 'autosum'),
          cmd('undoHome', 'Undo', 'Undo (⌘Z)', 'undo'),
          cmd('redoHome', 'Redo', 'Redo (⌘⇧Z)', 'redo'),
          cmd('sortAscHome', 'A-Z', 'Sort ascending', 'sortAsc'),
          cmd('filterHome', 'Filter', 'Filter', 'filter', 'wide'),
          cmd('findHome', tr.find, `${tr.find} (⌘F)`, 'find', 'wide'),
          cmd('gotoSpecialHome', tr.gotoSpecial, 'Go To Special', 'goTo', 'wide'),
        ],
        'editing',
      ),
    ]),
    tab('insert', [
      group(tr.tables, [
        cmd('pivotTableInsert', tr.pivotTable, 'PivotTable', 'table', 'wide'),
        cmd('formatTableInsert', tr.formatTable, 'Format as Table', 'tableStyle', 'wide'),
        cmd('namedRangesInsert', tr.names, 'Name manager', 'names', 'wide'),
        cmd(
          'removeDupesInsert',
          tr.removeDuplicates,
          'Remove duplicates',
          'removeDuplicates',
          'wide',
        ),
      ]),
      group(lang === 'ja' ? 'グラフ' : 'Charts', [
        cmd('chartInsert', tr.chart, 'Recommended chart', 'chart', 'wide'),
      ]),
      group(tr.links, [
        cmd('hyperlinkInsert', tr.hyperlink, 'Insert hyperlink (⌘K)', 'link', 'wide'),
        cmd('linksInsert', tr.links, 'Edit links', 'link', 'wide'),
      ]),
      group(tr.comments, [cmd('commentInsert', tr.newComment, 'New Note', 'commentAdd', 'wide')]),
      group(lang === 'ja' ? '記号と特殊文字' : 'Symbols', [
        cmd('fxInsert', 'fx', 'Insert function (Σ)', 'function', 'wide'),
      ]),
    ]),
    tab('draw', [
      group(RIBBON_TAB_LABELS.draw[lang], [
        cmd('drawPen', lang === 'ja' ? 'ペン' : 'Pen', RIBBON_TAB_LABELS.draw[lang], 'pen', 'wide'),
        cmd('drawErase', lang === 'ja' ? '消しゴム' : 'Eraser', 'Eraser', 'eraser', 'wide'),
      ]),
    ]),
    tab('pageLayout', [
      group(tr.pageSetup, [
        selectCmd(
          'marginsPreset',
          tr.margins,
          tr.margins,
          [
            { value: 'normal', label: tr.marginsNormal },
            { value: 'wide', label: tr.marginsWide },
            { value: 'narrow', label: tr.marginsNarrow },
            { value: 'custom', label: tr.marginsCustom },
          ],
          'demo__rb-select--border',
        ),
        selectCmd(
          'orientationPreset',
          tr.orientation,
          tr.orientation,
          [
            { value: 'portrait', label: tr.portrait },
            { value: 'landscape', label: tr.landscape },
          ],
          'demo__rb-select--border',
        ),
        selectCmd(
          'paperSizePreset',
          tr.paperSize,
          tr.paperSize,
          [
            { value: 'A4', label: tr.paperA4 },
            { value: 'A3', label: tr.paperA3 },
            { value: 'A5', label: tr.paperA5 },
            { value: 'letter', label: tr.paperLetter },
            { value: 'legal', label: tr.paperLegal },
            { value: 'tabloid', label: tr.paperTabloid },
          ],
          'demo__rb-select--border',
        ),
        cmd('pageSetupAdvanced', tr.pageSetup, 'Advanced page setup', 'options', 'wide'),
      ]),
      group(tr.print, [cmd('printPageLayout', tr.print, tr.print, 'print', 'wide')]),
    ]),
    tab('formulas', [
      group(tr.functionLibrary, [
        cmd('fx', 'fx', 'Insert function', 'function', 'mono'),
        cmd(
          'autosumFormula',
          lang === 'ja' ? 'オートSUM' : 'AutoSum',
          'AutoSum (Σ)',
          'autosum',
          'wide',
        ),
        cmd('sum', 'SUM', 'SUM arguments', 'function', 'mono'),
        cmd('avg', 'AVG', 'AVERAGE arguments', 'function', 'mono'),
      ]),
      group(tr.definedNames, [cmd('namedRanges', tr.names, 'Name manager', 'names', 'wide')]),
      group(tr.formulaAuditing, [
        cmd('precedents', tr.tracePrecedents, 'Trace precedents', 'trace', 'wide'),
        cmd('dependents', tr.traceDependents, 'Trace dependents', 'dependents', 'wide'),
        cmd('clearArrows', tr.removeArrows, 'Remove arrows', 'clearArrows', 'wide'),
      ]),
      group(tr.calculation, [
        cmd('recalcNow', tr.recalc, 'Calculate Now (F9)', 'autosum', 'wide'),
        cmd('calcOptions', tr.options, 'Calculation options', 'options', 'wide'),
        cmd('watch', tr.watch, 'Watch Window', 'watch', 'wide'),
      ]),
    ]),
    tab('data', [
      group(tr.sortFilter, [
        cmd('filter', 'Filter', 'Filter', 'filter', 'wide'),
        cmd('sortAsc', 'A-Z', 'Sort ascending', 'sortAsc'),
        cmd('sortDesc', 'Z-A', 'Sort descending', 'sortDesc'),
      ]),
      group(tr.dataTools, [
        cmd('removeDupes', tr.removeDuplicates, 'Remove duplicates', 'removeDuplicates', 'wide'),
        cmd('linksData', tr.links, 'Edit links', 'link', 'wide'),
      ]),
      group(tr.outline, [
        cmd('hideRows', tr.hideRows, 'Hide selected rows', 'table', 'wide'),
        cmd('hideCols', tr.hideCols, 'Hide selected columns', 'table', 'wide'),
      ]),
    ]),
    tab('review', [
      group(lang === 'ja' ? '文章校正' : 'Proofing', [
        cmd('spellingReview', tr.spelling, tr.spelling, 'spelling', 'wide'),
      ]),
      group(lang === 'ja' ? '言語' : 'Language', [
        cmd('translateReview', tr.translate, tr.translate, 'translate', 'wide'),
      ]),
      group(tr.comments, [
        cmd('newCommentReview', tr.newComment, 'New Note', 'commentAdd', 'wide'),
      ]),
      group(lang === 'ja' ? '検索' : 'Find', [
        cmd('findReview', tr.find, `${tr.find} (⌘F)`, 'find', 'wide'),
      ]),
      group(tr.protection, [cmd('protectReview', tr.protect, 'Protect sheet', 'protect', 'wide')]),
      group(tr.accessibility, [
        cmd('accessibility', tr.accessibility, tr.accessibility, 'accessibility', 'wide'),
      ]),
    ]),
    tab('view', [
      group(tr.workbookViews, [cmd('watchView', tr.watch, 'Watch Window', 'watch', 'wide')]),
      group(tr.window, [
        cmd('freeze', lang === 'ja' ? 'ウィンドウ枠' : 'Freeze', 'Freeze panes', 'freeze', 'wide'),
      ]),
      group(tr.zoom, [
        cmd('zoom75', '75%', 'Zoom to 75%', 'zoom', 'mono'),
        cmd('zoom100', '100%', 'Zoom to 100%', 'zoom', 'mono'),
        cmd('zoom125', '125%', 'Zoom to 125%', 'zoom', 'mono'),
      ]),
      group(tr.protection, [cmd('protect', tr.protect, 'Protect sheet', 'protect', 'wide')]),
    ]),
    tab('automate', [
      group(RIBBON_TAB_LABELS.automate[lang], [
        cmd('script', tr.script, tr.script, 'script', 'wide'),
      ]),
    ]),
    tab('acrobat', [
      group(tr.addIn, [cmd('addIn', tr.addIn, tr.addIn, 'addIn', 'wide')]),
      group(tr.pdf, [cmd('pdf', tr.pdf, tr.pdf, 'pdf', 'wide')]),
    ]),
  ];
}
