/**
 * Central registry of every user-visible string in the core package.
 *
 * Adding a new dialog or menu? Append a section here, populate it in `ja`
 * and `en`, then read from the dictionary in the interact module — never
 * hard-code another label.
 *
 * Consumers can override individual strings (or whole sections) via
 * `Spreadsheet.mount({ strings: ... })`. We deep-merge the overlay onto
 * the locale base, so partial overrides are safe.
 */
export interface Strings {
  contextMenu: {
    copy: string;
    cut: string;
    paste: string;
    pasteSpecial: string;
    clear: string;
    bold: string;
    italic: string;
    underline: string;
    alignLeft: string;
    alignCenter: string;
    alignRight: string;
    borders: string;
    clearFormat: string;
    formatCells: string;
    selectAll: string;
    rowInsertAbove: string;
    rowInsertBelow: string;
    rowDelete: string;
    rowHide: string;
    rowUnhide: string;
    colInsertLeft: string;
    colInsertRight: string;
    colDelete: string;
    colHide: string;
    colUnhide: string;
    rowGroup: string;
    rowUngroup: string;
    colGroup: string;
    colUngroup: string;
    insertComment: string;
    deleteComment: string;
    insertHyperlink: string;
    addWatch: string;
    removeWatch: string;
  };
  formatDialog: {
    title: string;
    tabNumber: string;
    tabAlign: string;
    tabFont: string;
    tabBorder: string;
    tabFill: string;
    tabMore: string;
    catGeneral: string;
    catFixed: string;
    catCurrency: string;
    catAccounting: string;
    catPercent: string;
    catScientific: string;
    catDate: string;
    catTime: string;
    catDateTime: string;
    catText: string;
    catCustom: string;
    decimals: string;
    symbol: string;
    pattern: string;
    patternPlaceholder: string;
    alignDefault: string;
    alignLeft: string;
    alignCenter: string;
    alignRight: string;
    horizontalAlign: string;
    verticalAlign: string;
    vAlignTop: string;
    vAlignMiddle: string;
    vAlignBottom: string;
    wrap: string;
    indent: string;
    rotation: string;
    fontFamily: string;
    fontSize: string;
    fontBold: string;
    fontItalic: string;
    fontUnderline: string;
    fontStrike: string;
    fontStyle: string;
    color: string;
    resetToDefault: string;
    borderTop: string;
    borderRight: string;
    borderBottom: string;
    borderLeft: string;
    borderDiagonalDown: string;
    borderDiagonalUp: string;
    borderStyle: string;
    borderColor: string;
    borderStyleNone: string;
    borderStyleThin: string;
    borderStyleMedium: string;
    borderStyleThick: string;
    borderStyleDashed: string;
    borderStyleDotted: string;
    borderStyleDouble: string;
    borderPresetNone: string;
    borderPresetOutline: string;
    borderPresetAll: string;
    fill: string;
    fillNone: string;
    hyperlink: string;
    hyperlinkPlaceholder: string;
    comment: string;
    commentPlaceholder: string;
    validationListSource: string;
    validationListPlaceholder: string;
    validationListSourceKind: string;
    validationListSourceLiteral: string;
    validationListSourceRange: string;
    validationListRangePlaceholder: string;
    validationLegend: string;
    validationKind: string;
    validationKindNone: string;
    validationKindList: string;
    validationKindWhole: string;
    validationKindDecimal: string;
    validationKindDate: string;
    validationKindTime: string;
    validationKindTextLength: string;
    validationKindCustom: string;
    validationOp: string;
    validationOpBetween: string;
    validationOpNotBetween: string;
    validationOpEq: string;
    validationOpNeq: string;
    validationOpLt: string;
    validationOpLte: string;
    validationOpGt: string;
    validationOpGte: string;
    validationValueA: string;
    validationValueB: string;
    validationFormula: string;
    validationFormulaPlaceholder: string;
    validationAllowBlank: string;
    validationErrorStyle: string;
    validationErrorStop: string;
    validationErrorWarning: string;
    validationErrorInfo: string;
    clearField: string;
    preview: string;
    cancel: string;
    ok: string;
  };
  hyperlinkDialog: {
    title: string;
    url: string;
    urlPlaceholder: string;
    remove: string;
    cancel: string;
    ok: string;
    errorEmptyUrl: string;
  };
  pasteSpecialDialog: {
    title: string;
    sectionPaste: string;
    sectionOperation: string;
    pasteAll: string;
    pasteFormulas: string;
    pasteValues: string;
    pasteFormats: string;
    pasteFormulasAndNumFmt: string;
    pasteValuesAndNumFmt: string;
    opNone: string;
    opAdd: string;
    opSubtract: string;
    opMultiply: string;
    opDivide: string;
    skipBlanks: string;
    transpose: string;
    cancel: string;
    ok: string;
  };
  findReplace: {
    title: string;
    findLabel: string;
    replaceLabel: string;
    matchCase: string;
    prev: string;
    next: string;
    replaceOne: string;
    replaceAll: string;
    close: string;
  };
  toolbar: {
    formatPainter: string;
    formatPainterStickyHint: string;
    freezePanesMenu: string;
    freezeFirstRow: string;
    freezeFirstCol: string;
    freezeAtSelection: string;
    unfreeze: string;
  };
  conditionalDialog: {
    title: string;
    rangeLabel: string;
    rangeAuto: string;
    addRule: string;
    removeRule: string;
    clearAll: string;
    kindLabel: string;
    kindCellValue: string;
    kindColorScale: string;
    kindDataBar: string;
    kindIconSet: string;
    kindTopBottom: string;
    kindFormula: string;
    kindDuplicates: string;
    kindUnique: string;
    kindBlanks: string;
    kindNonBlanks: string;
    kindErrors: string;
    kindNoErrors: string;
    opLabel: string;
    opGt: string;
    opLt: string;
    opGte: string;
    opLte: string;
    opEq: string;
    opNeq: string;
    opBetween: string;
    opNotBetween: string;
    valueA: string;
    valueB: string;
    fillColor: string;
    fontColor: string;
    bold: string;
    italic: string;
    underline: string;
    strike: string;
    stopMin: string;
    stopMid: string;
    stopMax: string;
    useThreeStops: string;
    barColor: string;
    showValue: string;
    topBottomMode: string;
    topN: string;
    usePercent: string;
    iconSetArrows3: string;
    iconSetArrows5: string;
    iconSetTraffic3: string;
    iconSetStars3: string;
    formulaPlaceholder: string;
    reverseOrder: string;
    empty: string;
    close: string;
  };
  namedRangeDialog: {
    title: string;
    nameHeader: string;
    formulaHeader: string;
    empty: string;
    /** Read-only fallback note shown when the engine doesn't support write. */
    note: string;
    namePlaceholder: string;
    formulaPlaceholder: string;
    addButton: string;
    /** Per-row "Delete" action label. */
    deleteButton: string;
    /** Inline error: empty or invalid name. */
    errorEmptyName: string;
    /** Inline error: engine refused the write. */
    errorEngineFailed: string;
    close: string;
  };
  statusBar: {
    /** Right-click menu heading. */
    aggregatesHeading: string;
    sum: string;
    average: string;
    /** Count of non-blank cells (Excel "Count"). */
    count: string;
    /** Count of numeric cells only (Excel "Numerical Count"). */
    countNumbers: string;
    min: string;
    max: string;
    /** Calc-mode badge label. */
    calcLabel: string;
    calcAuto: string;
    calcManual: string;
    calcAutoNoTable: string;
    /** Tooltip on the badge — clarifies the F9 / Ctrl+Alt+F9 affordance. */
    calcRecalcHint: string;
  };
  iterativeDialog: {
    title: string;
    note: string;
    enable: string;
    maxIterations: string;
    maxChange: string;
    unsupported: string;
    cancel: string;
    ok: string;
  };
  externalLinksDialog: {
    /** Modal title — also used as aria-label. */
    title: string;
    /** Empty-state when the workbook has no `<externalReferences>` block. */
    empty: string;
    /** Column headers for the link table. */
    headerIndex: string;
    headerKind: string;
    headerTarget: string;
    headerPart: string;
    /** Hint shown above the table — Excel parity for the Edit Links dialog. */
    note: string;
    close: string;
  };
  cfRulesDialog: {
    /** Modal title — Excel "Manage Rules" parity. */
    title: string;
    /** Empty-state when the engine reports no CF rules on the active sheet. */
    empty: string;
    headerPriority: string;
    headerType: string;
    headerRange: string;
    headerActions: string;
    /** Note above the table — flags the read-only fallback for visual rules. */
    note: string;
    /** Per-row action labels. */
    remove: string;
    /** Footer button — drops every rule on the sheet. */
    clearAll: string;
    /** Confirmation text shown before clearAll fires. Inline confirmation, not a separate prompt. */
    clearAllConfirm: string;
    close: string;
  };
  fxDialog: {
    /** Modal title — appears in the header and the aria-label. */
    title: string;
    /** Search input placeholder on the function-picker step. */
    searchPlaceholder: string;
    /** Label above the live formula preview on the args step. */
    preview: string;
    /** Empty-state shown when the search yields zero matches. */
    empty: string;
    /** Hint shown when the function signature includes a `...` repeat marker. */
    variadicHint: string;
    /** "Back" button — returns to the picker from the args step. */
    back: string;
    cancel: string;
    insert: string;
    /** aria-label for the formula-bar fx button that opens this dialog. */
    fxButtonLabel: string;
  };
  watchPanel: {
    title: string;
    sheetHeader: string;
    cellHeader: string;
    nameHeader: string;
    valueHeader: string;
    formulaHeader: string;
    addWatch: string;
    removeWatch: string;
    clearAll: string;
    empty: string;
    close: string;
  };
  slicer: {
    /** Default panel header used when no column is bound yet. */
    title: string;
    /** "Select all" affordance label — clears the chip selection. */
    selectAll: string;
    /** "Clear" — empties the chip selection (synonym for select-all in v1). */
    clear: string;
    /** aria-label on the per-panel close (×) button. */
    close: string;
    /** Imperative API surface label, surfaced in chrome menus. */
    addSlicer: string;
    /** Prompt shown when picking which column drives a new slicer. */
    chooseColumn: string;
    /** Placeholder text shown when a slicer references a missing table. */
    tablePlaceholder: string;
  };
  errorMenu: {
    /** Heading shown above the action list when the menu is for a formula
     *  error (e.g. "#DIV/0! — 0"). */
    errorHeading: string;
    /** Heading shown above the action list for data-validation violations. */
    validationHeading: string;
    showInfo: string;
    editCell: string;
    traceError: string;
    ignore: string;
  };
  goToDialog: {
    /** Modal title — shown in the header bar and the aria-label. */
    title: string;
    /** Section legend above the scope radios. */
    scopeLabel: string;
    /** "Active sheet" radio — sweeps every cell on the current sheet. */
    scopeSheet: string;
    /** "Current selection" radio — sweeps only inside the active selection
     *  rectangle. Auto-disabled when the selection is a single cell. */
    scopeSelection: string;
    /** Section legend above the category radios. */
    kindLabel: string;
    kindBlanks: string;
    kindNonBlanks: string;
    kindFormulas: string;
    kindConstants: string;
    kindNumbers: string;
    kindText: string;
    kindErrors: string;
    kindDataValidation: string;
    kindConditionalFormat: string;
    /** Inline status when the predicate yields zero matches. The dialog
     *  stays open so the user can adjust the kind. */
    noResults: string;
    cancel: string;
    ok: string;
  };
  pageSetup: {
    /** Modal title — shown in the header bar and the aria-label. */
    title: string;
    orientation: string;
    orientPortrait: string;
    orientLandscape: string;
    paperSize: string;
    margins: string;
    marginTop: string;
    marginRight: string;
    marginBottom: string;
    marginLeft: string;
    headerLabel: string;
    footerLabel: string;
    /** Placeholder for the left header/footer slot. */
    slotLeftPlaceholder: string;
    /** Placeholder for the center header/footer slot. */
    slotCenterPlaceholder: string;
    /** Placeholder for the right header/footer slot. */
    slotRightPlaceholder: string;
    printTitleRows: string;
    printTitleRowsPlaceholder: string;
    printTitleCols: string;
    printTitleColsPlaceholder: string;
    /** Print scale (0.10 .. 4.00). */
    scale: string;
    /** Fit-to-N-pages-wide. */
    fitWidth: string;
    /** Fit-to-N-pages-tall. */
    fitHeight: string;
    showGridlines: string;
    showHeadings: string;
    cancel: string;
    ok: string;
  };
  protection: {
    /** Format-dialog tab label for the cell-lock section. */
    tabProtection: string;
    /** Checkbox label for the per-cell `locked` flag. */
    locked: string;
    /** Helper text under the locked checkbox explaining that the lock only
     *  takes effect when the sheet is itself protected. */
    lockedHint: string;
    /** Toolbar / menu label that turns sheet protection on. */
    protectSheet: string;
    /** Same control's label when protection is already on. */
    unprotectSheet: string;
    /** Field label for the (currently un-enforced) password input. */
    password: string;
    /** Placeholder text for the password input. */
    passwordPlaceholder: string;
  };
  a11y: {
    nameBox: string;
    formulaBar: string;
    spreadsheet: string;
  };
}

export const ja: Strings = {
  contextMenu: {
    copy: 'コピー',
    cut: '切り取り',
    paste: '貼り付け',
    pasteSpecial: '形式を選択して貼り付け…',
    clear: 'クリア',
    bold: '太字',
    italic: '斜体',
    underline: '下線',
    alignLeft: '左揃え',
    alignCenter: '中央揃え',
    alignRight: '右揃え',
    borders: '罫線',
    clearFormat: '書式のクリア',
    formatCells: 'セルの書式設定…',
    selectAll: 'すべて選択',
    rowInsertAbove: '上に行を挿入',
    rowInsertBelow: '下に行を挿入',
    rowDelete: '行の削除',
    rowHide: '行を非表示',
    rowUnhide: '行の再表示',
    rowGroup: '行をグループ化',
    rowUngroup: '行のグループ解除',
    colInsertLeft: '左に列を挿入',
    colInsertRight: '右に列を挿入',
    colDelete: '列の削除',
    colHide: '列を非表示',
    colUnhide: '列の再表示',
    colGroup: '列をグループ化',
    colUngroup: '列のグループ解除',
    insertComment: 'コメントを編集…',
    deleteComment: 'コメントを削除',
    insertHyperlink: 'ハイパーリンクを挿入…',
    addWatch: 'ウォッチを追加',
    removeWatch: 'ウォッチを削除',
  },
  formatDialog: {
    title: 'セルの書式設定',
    tabNumber: '表示形式',
    tabAlign: '配置',
    tabFont: 'フォント',
    tabBorder: '罫線',
    tabFill: '塗りつぶし',
    tabMore: 'その他',
    catGeneral: '標準',
    catFixed: '数値',
    catCurrency: '通貨',
    catAccounting: '会計',
    catPercent: 'パーセンテージ',
    catScientific: '指数',
    catDate: '日付',
    catTime: '時刻',
    catDateTime: '日付と時刻',
    catText: '文字列',
    catCustom: 'ユーザー定義',
    decimals: '小数点以下の桁数',
    symbol: '記号',
    pattern: '書式',
    patternPlaceholder: '例: 0.00 / yyyy-mm-dd',
    alignDefault: '標準',
    alignLeft: '左詰め',
    alignCenter: '中央',
    alignRight: '右詰め',
    horizontalAlign: '横位置',
    verticalAlign: '縦位置',
    vAlignTop: '上',
    vAlignMiddle: '中央',
    vAlignBottom: '下',
    wrap: '折り返して全体を表示',
    indent: 'インデント',
    rotation: '回転 (度)',
    fontFamily: 'フォント',
    fontSize: 'サイズ',
    fontBold: '太字',
    fontItalic: '斜体',
    fontUnderline: '下線',
    fontStrike: '取り消し線',
    fontStyle: 'スタイル',
    color: '色',
    resetToDefault: '標準に戻す',
    borderTop: '上',
    borderRight: '右',
    borderBottom: '下',
    borderLeft: '左',
    borderDiagonalDown: '対角線 ↘',
    borderDiagonalUp: '対角線 ↗',
    borderStyle: 'スタイル',
    borderColor: '線の色',
    borderStyleNone: 'なし',
    borderStyleThin: '細線',
    borderStyleMedium: '中線',
    borderStyleThick: '太線',
    borderStyleDashed: '破線',
    borderStyleDotted: '点線',
    borderStyleDouble: '二重線',
    borderPresetNone: 'なし',
    borderPresetOutline: '外枠',
    borderPresetAll: '格子',
    fill: '背景色',
    fillNone: '塗りつぶしなし',
    hyperlink: 'ハイパーリンク',
    hyperlinkPlaceholder: 'https://...',
    comment: 'コメント',
    commentPlaceholder: 'メモを入力',
    validationListSource: '入力規則 (リスト)',
    validationListPlaceholder: '値を改行で区切って入力',
    validationListSourceKind: 'ソース',
    validationListSourceLiteral: '値を直接入力',
    validationListSourceRange: 'セル範囲を参照',
    validationListRangePlaceholder: '例: Sheet1!$A$1:$A$10',
    validationLegend: '入力規則',
    validationKind: '種類',
    validationKindNone: 'なし',
    validationKindList: 'リスト',
    validationKindWhole: '整数',
    validationKindDecimal: '小数',
    validationKindDate: '日付',
    validationKindTime: '時刻',
    validationKindTextLength: '文字数',
    validationKindCustom: 'カスタム',
    validationOp: '条件',
    validationOpBetween: '範囲内',
    validationOpNotBetween: '範囲外',
    validationOpEq: '等しい',
    validationOpNeq: '等しくない',
    validationOpLt: 'より小さい',
    validationOpLte: '以下',
    validationOpGt: 'より大きい',
    validationOpGte: '以上',
    validationValueA: '値',
    validationValueB: '上限値',
    validationFormula: '数式',
    validationFormulaPlaceholder: '=A1>0',
    validationAllowBlank: '空白を許可',
    validationErrorStyle: 'エラーレベル',
    validationErrorStop: '中止',
    validationErrorWarning: '警告',
    validationErrorInfo: '情報',
    clearField: 'クリア',
    preview: 'プレビュー',
    cancel: 'キャンセル',
    ok: 'OK',
  },
  hyperlinkDialog: {
    title: 'ハイパーリンクの挿入',
    url: 'URL',
    urlPlaceholder: 'https://...',
    remove: 'リンクを削除',
    cancel: 'キャンセル',
    ok: 'OK',
    errorEmptyUrl: 'URL を入力してください',
  },
  pasteSpecialDialog: {
    title: '形式を選択して貼り付け',
    sectionPaste: '貼り付け',
    sectionOperation: '演算',
    pasteAll: 'すべて',
    pasteFormulas: '数式',
    pasteValues: '値',
    pasteFormats: '書式',
    pasteFormulasAndNumFmt: '数式と数値の書式',
    pasteValuesAndNumFmt: '値と数値の書式',
    opNone: 'なし',
    opAdd: '加算',
    opSubtract: '減算',
    opMultiply: '乗算',
    opDivide: '除算',
    skipBlanks: '空白セルを無視する',
    transpose: '行/列の入れ替え',
    cancel: 'キャンセル',
    ok: 'OK',
  },
  findReplace: {
    title: '検索と置換',
    findLabel: '検索',
    replaceLabel: '置換',
    matchCase: '大文字/小文字を区別',
    prev: '前へ',
    next: '次へ',
    replaceOne: '置換',
    replaceAll: 'すべて置換',
    close: '閉じる',
  },
  toolbar: {
    formatPainter: '書式のコピー/貼り付け',
    formatPainterStickyHint: 'ダブルクリックで連続適用',
    freezePanesMenu: 'ウィンドウ枠の固定',
    freezeFirstRow: '先頭行の固定',
    freezeFirstCol: '先頭列の固定',
    freezeAtSelection: '選択範囲で固定',
    unfreeze: '固定解除',
  },
  conditionalDialog: {
    title: '条件付き書式',
    rangeLabel: '対象範囲',
    rangeAuto: '選択範囲',
    addRule: 'ルールを追加',
    removeRule: '削除',
    clearAll: 'すべて削除',
    kindLabel: '種類',
    kindCellValue: 'セル値',
    kindColorScale: 'カラースケール',
    kindDataBar: 'データバー',
    kindIconSet: 'アイコンセット',
    kindTopBottom: '上位/下位',
    kindFormula: '数式',
    kindDuplicates: '重複する値',
    kindUnique: '一意の値',
    kindBlanks: '空白セル',
    kindNonBlanks: '空白以外のセル',
    kindErrors: 'エラー値',
    kindNoErrors: 'エラーなし',
    opLabel: '条件',
    opGt: 'より大きい',
    opLt: 'より小さい',
    opGte: '以上',
    opLte: '以下',
    opEq: '等しい',
    opNeq: '等しくない',
    opBetween: '範囲内',
    opNotBetween: '範囲外',
    valueA: '値',
    valueB: '値 (上限)',
    fillColor: '背景色',
    fontColor: '文字色',
    bold: '太字',
    italic: '斜体',
    underline: '下線',
    strike: '取り消し線',
    stopMin: '最小',
    stopMid: '中央',
    stopMax: '最大',
    useThreeStops: '3 段階',
    barColor: 'バーの色',
    showValue: '値も表示',
    topBottomMode: '対象',
    topN: '個数 (N)',
    usePercent: 'パーセント指定',
    iconSetArrows3: '3 矢印',
    iconSetArrows5: '5 矢印',
    iconSetTraffic3: '3 信号',
    iconSetStars3: '3 つ星',
    formulaPlaceholder: '例: >100 / <>"x" / =A1>0',
    reverseOrder: '並び順を反転',
    empty: 'ルールはまだありません',
    close: '閉じる',
  },
  namedRangeDialog: {
    title: '名前の管理',
    nameHeader: '名前',
    formulaHeader: '参照',
    empty: '名前付き範囲は登録されていません',
    note: '※ このエンジンでは編集に対応していません。一覧表示のみです。',
    namePlaceholder: '名前',
    formulaPlaceholder: '=Sheet1!$A$1:$B$5',
    addButton: '追加',
    deleteButton: '削除',
    errorEmptyName: '名前を入力してください',
    errorEngineFailed: '保存に失敗しました',
    close: '閉じる',
  },
  statusBar: {
    aggregatesHeading: '集計表示',
    sum: '合計',
    average: '平均',
    count: 'データの個数',
    countNumbers: '数値の個数',
    min: '最小値',
    max: '最大値',
    calcLabel: '計算',
    calcAuto: '自動',
    calcManual: '手動',
    calcAutoNoTable: '自動 (テーブル除く)',
    calcRecalcHint: 'クリックで再計算 (F9 / Ctrl+Alt+F9)',
  },
  iterativeDialog: {
    title: '反復計算',
    note: '循環参照を反復計算で解決します。各セルが収束基準を満たすか上限回数に達するまで再計算します。',
    enable: '反復計算を有効にする',
    maxIterations: '最大反復回数',
    maxChange: '変化の許容値',
    unsupported: 'このエンジンは反復計算に対応していません。',
    cancel: 'キャンセル',
    ok: 'OK',
  },
  externalLinksDialog: {
    title: '外部参照',
    empty: 'このブックには外部参照がありません。',
    headerIndex: '#',
    headerKind: '種類',
    headerTarget: 'リンク先',
    headerPart: 'パート',
    note: '読み取り専用 — リンクの編集は対応していません。書式は保存時に保持されます。',
    close: '閉じる',
  },
  cfRulesDialog: {
    title: '条件付き書式ルールの管理',
    empty: 'このシートには条件付き書式ルールがありません。',
    headerPriority: '優先度',
    headerType: '種類',
    headerRange: '範囲',
    headerActions: '操作',
    note: 'ビジュアルルール (カラースケール / データバー / アイコン) は読み取り専用です。削除のみ可能です。',
    remove: '削除',
    clearAll: 'すべて削除',
    clearAllConfirm: '本当にこのシートのすべてのルールを削除しますか？',
    close: '閉じる',
  },
  fxDialog: {
    title: '関数の引数',
    searchPlaceholder: '関数を検索…',
    preview: '数式の結果',
    empty: '一致する関数がありません',
    variadicHint: 'この関数は追加の引数を受け取ります (任意)。',
    back: '戻る',
    cancel: 'キャンセル',
    insert: '挿入',
    fxButtonLabel: '関数の引数を挿入',
  },
  watchPanel: {
    title: 'ウォッチ ウィンドウ',
    sheetHeader: 'シート',
    cellHeader: 'セル',
    nameHeader: '名前',
    valueHeader: '値',
    formulaHeader: '数式',
    addWatch: 'ウォッチを追加',
    removeWatch: 'ウォッチを削除',
    clearAll: 'すべて削除',
    empty: 'ウォッチはまだありません',
    close: '閉じる',
  },
  slicer: {
    title: 'スライサー',
    selectAll: 'すべて選択',
    clear: 'クリア',
    close: '閉じる',
    addSlicer: 'スライサーの追加',
    chooseColumn: '列を選択',
    tablePlaceholder: '対象テーブルが見つかりません',
  },
  errorMenu: {
    errorHeading: 'エラー',
    validationHeading: '入力規則違反',
    showInfo: 'エラーの詳細',
    editCell: 'セルを編集',
    traceError: 'エラーの参照元',
    ignore: '無視',
  },
  goToDialog: {
    title: '選択オプション',
    scopeLabel: '範囲',
    scopeSheet: 'アクティブなシート',
    scopeSelection: '現在の選択範囲',
    kindLabel: '種類',
    kindBlanks: '空白セル',
    kindNonBlanks: '空白以外のセル',
    kindFormulas: '数式',
    kindConstants: '定数',
    kindNumbers: '数値',
    kindText: '文字列',
    kindErrors: 'エラー',
    kindDataValidation: '入力規則',
    kindConditionalFormat: '条件付き書式',
    noResults: '該当するセルが見つかりません',
    cancel: 'キャンセル',
    ok: 'OK',
  },
  pageSetup: {
    title: 'ページ設定',
    orientation: '印刷の向き',
    orientPortrait: '縦',
    orientLandscape: '横',
    paperSize: '用紙サイズ',
    margins: '余白 (インチ)',
    marginTop: '上',
    marginRight: '右',
    marginBottom: '下',
    marginLeft: '左',
    headerLabel: 'ヘッダー',
    footerLabel: 'フッター',
    slotLeftPlaceholder: '左',
    slotCenterPlaceholder: '中央',
    slotRightPlaceholder: '右',
    printTitleRows: '印刷タイトル (行)',
    printTitleRowsPlaceholder: '例: 1:3',
    printTitleCols: '印刷タイトル (列)',
    printTitleColsPlaceholder: '例: A:B',
    scale: '倍率',
    fitWidth: '横方向のページ数',
    fitHeight: '縦方向のページ数',
    showGridlines: '枠線を印刷',
    showHeadings: '行列番号を印刷',
    cancel: 'キャンセル',
    ok: 'OK',
  },
  protection: {
    tabProtection: '保護',
    locked: 'ロック',
    lockedHint: 'シートが保護されている場合のみ、ロックされたセルへの書き込みがブロックされます。',
    protectSheet: 'シートを保護',
    unprotectSheet: 'シート保護を解除',
    password: 'パスワード',
    passwordPlaceholder: '任意 (現在は未検証)',
  },
  a11y: {
    nameBox: '名前ボックス',
    formulaBar: '数式バー',
    spreadsheet: 'スプレッドシート',
  },
};

export const en: Strings = {
  contextMenu: {
    copy: 'Copy',
    cut: 'Cut',
    paste: 'Paste',
    pasteSpecial: 'Paste Special…',
    clear: 'Clear',
    bold: 'Bold',
    italic: 'Italic',
    underline: 'Underline',
    alignLeft: 'Align Left',
    alignCenter: 'Align Center',
    alignRight: 'Align Right',
    borders: 'Borders',
    clearFormat: 'Clear Formatting',
    formatCells: 'Format Cells…',
    selectAll: 'Select All',
    rowInsertAbove: 'Insert row above',
    rowInsertBelow: 'Insert row below',
    rowDelete: 'Delete row',
    rowHide: 'Hide row',
    rowUnhide: 'Unhide rows',
    rowGroup: 'Group rows',
    rowUngroup: 'Ungroup rows',
    colInsertLeft: 'Insert column left',
    colInsertRight: 'Insert column right',
    colDelete: 'Delete column',
    colHide: 'Hide column',
    colUnhide: 'Unhide columns',
    colGroup: 'Group columns',
    colUngroup: 'Ungroup columns',
    insertComment: 'Edit comment…',
    deleteComment: 'Delete comment',
    insertHyperlink: 'Insert hyperlink…',
    addWatch: 'Add Watch',
    removeWatch: 'Remove Watch',
  },
  formatDialog: {
    title: 'Format Cells',
    tabNumber: 'Number',
    tabAlign: 'Alignment',
    tabFont: 'Font',
    tabBorder: 'Border',
    tabFill: 'Fill',
    tabMore: 'More',
    catGeneral: 'General',
    catFixed: 'Number',
    catCurrency: 'Currency',
    catAccounting: 'Accounting',
    catPercent: 'Percentage',
    catScientific: 'Scientific',
    catDate: 'Date',
    catTime: 'Time',
    catDateTime: 'Date & Time',
    catText: 'Text',
    catCustom: 'Custom',
    decimals: 'Decimal places',
    symbol: 'Symbol',
    pattern: 'Format code',
    patternPlaceholder: 'e.g. 0.00 / yyyy-mm-dd',
    alignDefault: 'General',
    alignLeft: 'Left',
    alignCenter: 'Center',
    alignRight: 'Right',
    horizontalAlign: 'Horizontal',
    verticalAlign: 'Vertical',
    vAlignTop: 'Top',
    vAlignMiddle: 'Middle',
    vAlignBottom: 'Bottom',
    wrap: 'Wrap text',
    indent: 'Indent',
    rotation: 'Rotation (deg)',
    fontFamily: 'Font',
    fontSize: 'Size',
    fontBold: 'Bold',
    fontItalic: 'Italic',
    fontUnderline: 'Underline',
    fontStrike: 'Strikethrough',
    fontStyle: 'Style',
    color: 'Color',
    resetToDefault: 'Reset to default',
    borderTop: 'Top',
    borderRight: 'Right',
    borderBottom: 'Bottom',
    borderLeft: 'Left',
    borderDiagonalDown: 'Diagonal ↘',
    borderDiagonalUp: 'Diagonal ↗',
    borderStyle: 'Line style',
    borderColor: 'Line color',
    borderStyleNone: 'None',
    borderStyleThin: 'Thin',
    borderStyleMedium: 'Medium',
    borderStyleThick: 'Thick',
    borderStyleDashed: 'Dashed',
    borderStyleDotted: 'Dotted',
    borderStyleDouble: 'Double',
    borderPresetNone: 'None',
    borderPresetOutline: 'Outline',
    borderPresetAll: 'All',
    fill: 'Background',
    fillNone: 'No fill',
    hyperlink: 'Hyperlink',
    hyperlinkPlaceholder: 'https://...',
    comment: 'Comment',
    commentPlaceholder: 'Enter a note',
    validationListSource: 'Data validation (list)',
    validationListPlaceholder: 'One value per line',
    validationListSourceKind: 'Source',
    validationListSourceLiteral: 'Literal values',
    validationListSourceRange: 'Cell range',
    validationListRangePlaceholder: 'e.g. Sheet1!$A$1:$A$10',
    validationLegend: 'Data validation',
    validationKind: 'Kind',
    validationKindNone: 'None',
    validationKindList: 'List',
    validationKindWhole: 'Whole number',
    validationKindDecimal: 'Decimal',
    validationKindDate: 'Date',
    validationKindTime: 'Time',
    validationKindTextLength: 'Text length',
    validationKindCustom: 'Custom',
    validationOp: 'Condition',
    validationOpBetween: 'between',
    validationOpNotBetween: 'not between',
    validationOpEq: 'equal to',
    validationOpNeq: 'not equal to',
    validationOpLt: 'less than',
    validationOpLte: 'less than or equal',
    validationOpGt: 'greater than',
    validationOpGte: 'greater than or equal',
    validationValueA: 'Value',
    validationValueB: 'Upper value',
    validationFormula: 'Formula',
    validationFormulaPlaceholder: '=A1>0',
    validationAllowBlank: 'Allow blank',
    validationErrorStyle: 'Error level',
    validationErrorStop: 'Stop',
    validationErrorWarning: 'Warning',
    validationErrorInfo: 'Information',
    clearField: 'Clear',
    preview: 'Preview',
    cancel: 'Cancel',
    ok: 'OK',
  },
  hyperlinkDialog: {
    title: 'Insert hyperlink',
    url: 'URL',
    urlPlaceholder: 'https://...',
    remove: 'Remove link',
    cancel: 'Cancel',
    ok: 'OK',
    errorEmptyUrl: 'Enter a URL',
  },
  pasteSpecialDialog: {
    title: 'Paste Special',
    sectionPaste: 'Paste',
    sectionOperation: 'Operation',
    pasteAll: 'All',
    pasteFormulas: 'Formulas',
    pasteValues: 'Values',
    pasteFormats: 'Formats',
    pasteFormulasAndNumFmt: 'Formulas and number formats',
    pasteValuesAndNumFmt: 'Values and number formats',
    opNone: 'None',
    opAdd: 'Add',
    opSubtract: 'Subtract',
    opMultiply: 'Multiply',
    opDivide: 'Divide',
    skipBlanks: 'Skip blanks',
    transpose: 'Transpose',
    cancel: 'Cancel',
    ok: 'OK',
  },
  findReplace: {
    title: 'Find and replace',
    findLabel: 'Find',
    replaceLabel: 'Replace',
    matchCase: 'Match case',
    prev: 'Previous',
    next: 'Next',
    replaceOne: 'Replace',
    replaceAll: 'Replace all',
    close: 'Close',
  },
  toolbar: {
    formatPainter: 'Format Painter',
    formatPainterStickyHint: 'Double-click for sticky mode',
    freezePanesMenu: 'Freeze Panes',
    freezeFirstRow: 'Freeze top row',
    freezeFirstCol: 'Freeze first column',
    freezeAtSelection: 'Freeze at selection',
    unfreeze: 'Unfreeze',
  },
  conditionalDialog: {
    title: 'Conditional Formatting',
    rangeLabel: 'Range',
    rangeAuto: 'Selection',
    addRule: 'Add rule',
    removeRule: 'Remove',
    clearAll: 'Remove all',
    kindLabel: 'Kind',
    kindCellValue: 'Cell value',
    kindColorScale: 'Color scale',
    kindDataBar: 'Data bar',
    kindIconSet: 'Icon set',
    kindTopBottom: 'Top / Bottom',
    kindFormula: 'Formula',
    kindDuplicates: 'Duplicate values',
    kindUnique: 'Unique values',
    kindBlanks: 'Blank cells',
    kindNonBlanks: 'Non-blank cells',
    kindErrors: 'Errors',
    kindNoErrors: 'No errors',
    opLabel: 'Condition',
    opGt: 'greater than',
    opLt: 'less than',
    opGte: 'greater than or equal',
    opLte: 'less than or equal',
    opEq: 'equal to',
    opNeq: 'not equal to',
    opBetween: 'between',
    opNotBetween: 'not between',
    valueA: 'Value',
    valueB: 'Upper value',
    fillColor: 'Fill color',
    fontColor: 'Font color',
    bold: 'Bold',
    italic: 'Italic',
    underline: 'Underline',
    strike: 'Strikethrough',
    stopMin: 'Min',
    stopMid: 'Mid',
    stopMax: 'Max',
    useThreeStops: '3-stop',
    barColor: 'Bar color',
    showValue: 'Show value',
    topBottomMode: 'Mode',
    topN: 'N',
    usePercent: 'Use percent',
    iconSetArrows3: '3 arrows',
    iconSetArrows5: '5 arrows',
    iconSetTraffic3: '3 traffic lights',
    iconSetStars3: '3 stars',
    formulaPlaceholder: 'e.g. >100 / <>"x" / =A1>0',
    reverseOrder: 'Reverse order',
    empty: 'No rules defined yet',
    close: 'Close',
  },
  namedRangeDialog: {
    title: 'Name Manager',
    nameHeader: 'Name',
    formulaHeader: 'Refers to',
    empty: 'No defined names registered',
    note: 'Editing is not supported by this engine. Listing only.',
    namePlaceholder: 'Name',
    formulaPlaceholder: '=Sheet1!$A$1:$B$5',
    addButton: 'Add',
    deleteButton: 'Delete',
    errorEmptyName: 'Enter a name',
    errorEngineFailed: 'Failed to save',
    close: 'Close',
  },
  statusBar: {
    aggregatesHeading: 'Customize Status Bar',
    sum: 'Sum',
    average: 'Average',
    count: 'Count',
    countNumbers: 'Numerical Count',
    min: 'Minimum',
    max: 'Maximum',
    calcLabel: 'Calc',
    calcAuto: 'Auto',
    calcManual: 'Manual',
    calcAutoNoTable: 'Auto (skip tables)',
    calcRecalcHint: 'Click to recalc (F9 / Ctrl+Alt+F9)',
  },
  iterativeDialog: {
    title: 'Iterative calculation',
    note: 'Resolve circular references by iterating until the change between cycles falls below the threshold or the iteration cap is hit.',
    enable: 'Enable iterative calculation',
    maxIterations: 'Maximum iterations',
    maxChange: 'Maximum change',
    unsupported: 'This engine does not support iterative calculation.',
    cancel: 'Cancel',
    ok: 'OK',
  },
  externalLinksDialog: {
    title: 'External Links',
    empty: 'This workbook has no external references.',
    headerIndex: '#',
    headerKind: 'Kind',
    headerTarget: 'Target',
    headerPart: 'Part',
    note: 'Read-only — editing external links is not supported. Records are preserved on save.',
    close: 'Close',
  },
  cfRulesDialog: {
    title: 'Conditional Formatting — Manage Rules',
    empty: 'This sheet has no conditional formatting rules.',
    headerPriority: 'Priority',
    headerType: 'Type',
    headerRange: 'Range',
    headerActions: 'Actions',
    note: 'Visual rules (color scale / data bar / icon set) are read-only. Removing them is supported.',
    remove: 'Remove',
    clearAll: 'Clear all',
    clearAllConfirm: 'Are you sure you want to remove every rule on this sheet?',
    close: 'Close',
  },
  fxDialog: {
    title: 'Function Arguments',
    searchPlaceholder: 'Search functions…',
    preview: 'Formula result',
    empty: 'No matching functions',
    variadicHint: 'This function accepts additional optional arguments.',
    back: 'Back',
    cancel: 'Cancel',
    insert: 'Insert',
    fxButtonLabel: 'Insert function arguments',
  },
  watchPanel: {
    title: 'Watch Window',
    sheetHeader: 'Sheet',
    cellHeader: 'Cell',
    nameHeader: 'Name',
    valueHeader: 'Value',
    formulaHeader: 'Formula',
    addWatch: 'Add Watch',
    removeWatch: 'Remove Watch',
    clearAll: 'Delete All',
    empty: 'No watches',
    close: 'Close',
  },
  slicer: {
    title: 'Slicer',
    selectAll: 'Select all',
    clear: 'Clear',
    close: 'Close',
    addSlicer: 'Insert Slicer',
    chooseColumn: 'Choose column',
    tablePlaceholder: 'Table not found',
  },
  errorMenu: {
    errorHeading: 'Error',
    validationHeading: 'Validation issue',
    showInfo: 'Show error info',
    editCell: 'Edit cell',
    traceError: 'Trace error',
    ignore: 'Ignore',
  },
  goToDialog: {
    title: 'Go To Special',
    scopeLabel: 'Scope',
    scopeSheet: 'Active sheet',
    scopeSelection: 'Current selection',
    kindLabel: 'Select',
    kindBlanks: 'Blanks',
    kindNonBlanks: 'Non-blanks',
    kindFormulas: 'Formulas',
    kindConstants: 'Constants',
    kindNumbers: 'Numbers',
    kindText: 'Text',
    kindErrors: 'Errors',
    kindDataValidation: 'Data validation',
    kindConditionalFormat: 'Conditional formats',
    noResults: 'No cells found',
    cancel: 'Cancel',
    ok: 'OK',
  },
  pageSetup: {
    title: 'Page Setup',
    orientation: 'Orientation',
    orientPortrait: 'Portrait',
    orientLandscape: 'Landscape',
    paperSize: 'Paper size',
    margins: 'Margins (inches)',
    marginTop: 'Top',
    marginRight: 'Right',
    marginBottom: 'Bottom',
    marginLeft: 'Left',
    headerLabel: 'Header',
    footerLabel: 'Footer',
    slotLeftPlaceholder: 'Left',
    slotCenterPlaceholder: 'Center',
    slotRightPlaceholder: 'Right',
    printTitleRows: 'Print title rows',
    printTitleRowsPlaceholder: 'e.g. 1:3',
    printTitleCols: 'Print title columns',
    printTitleColsPlaceholder: 'e.g. A:B',
    scale: 'Scale',
    fitWidth: 'Fit to width (pages)',
    fitHeight: 'Fit to height (pages)',
    showGridlines: 'Print gridlines',
    showHeadings: 'Print headings',
    cancel: 'Cancel',
    ok: 'OK',
  },
  protection: {
    tabProtection: 'Protection',
    locked: 'Locked',
    lockedHint: 'Locking cells only takes effect once the sheet is protected.',
    protectSheet: 'Protect Sheet',
    unprotectSheet: 'Unprotect Sheet',
    password: 'Password',
    passwordPlaceholder: 'Optional (not enforced yet)',
  },
  a11y: {
    nameBox: 'Name box',
    formulaBar: 'Formula bar',
    spreadsheet: 'Spreadsheet',
  },
};

export type Locale = 'ja' | 'en';

export const dictionaries: Record<Locale, Strings> = { ja, en };

/** Default — Japanese, matching the current playground demo. */
export const defaultStrings: Strings = ja;

/** Recursively merge a partial override onto a base dictionary. Lets
 *  consumers tweak a single label without re-supplying the entire tree. */
export function mergeStrings(base: Strings, overlay?: DeepPartial<Strings>): Strings {
  if (!overlay) return base;
  const out = structuredClone(base) as Strings;
  for (const sectionKey of Object.keys(overlay) as (keyof Strings)[]) {
    const section = overlay[sectionKey];
    if (!section) continue;
    Object.assign(out[sectionKey], section);
  }
  return out;
}

export type DeepPartial<T> = T extends object ? { [K in keyof T]?: DeepPartial<T[K]> } : T;
