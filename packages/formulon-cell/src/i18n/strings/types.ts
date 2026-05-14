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
    descGeneral: string;
    descFixed: string;
    descCurrency: string;
    descAccounting: string;
    descPercent: string;
    descScientific: string;
    descDate: string;
    descTime: string;
    descDateTime: string;
    descText: string;
    descCustom: string;
    decimals: string;
    symbol: string;
    patternType: string;
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
    previewText: string;
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
  commentDialog: {
    title: string;
    titleEdit: string;
    placeholder: string;
    remove: string;
    cancel: string;
    ok: string;
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
    /** Inline error: empty formula/ref. */
    errorEmptyFormula: string;
    /** Inline error: engine refused the write. */
    errorEngineFailed: string;
    close: string;
  };
  pivotTableDialog: {
    title: string;
    source: string;
    name: string;
    namePlaceholder: string;
    destination: string;
    destinationPlaceholder: string;
    rowField: string;
    columnField: string;
    valueField: string;
    aggregation: string;
    sum: string;
    count: string;
    rowSort: string;
    columnSort: string;
    sortNone: string;
    sortAsc: string;
    sortDesc: string;
    rowSubtotalTop: string;
    columnSubtotalTop: string;
    numberFormat: string;
    numberFormatPlaceholder: string;
    rowGrandTotals: string;
    columnGrandTotals: string;
    none: string;
    unsupported: string;
    invalidRange: string;
    invalidDestination: string;
    engineFailed: string;
    cancel: string;
    ok: string;
  };
  statusBar: {
    ready: string;
    cell: string;
    cells: string;
    /** Right-click menu heading. */
    aggregatesHeading: string;
    sum: string;
    average: string;
    /** Count of non-blank cells (spreadsheet "Count"). */
    count: string;
    /** Count of numeric cells only (spreadsheet "Numerical Count"). */
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
    zoom: string;
    zoomIn: string;
    zoomOut: string;
  };
  viewToolbar: {
    title: string;
    gridlines: string;
    headings: string;
    formulas: string;
    r1c1: string;
    freezeNone: string;
    freezeTopRow: string;
    freezeFirstColumn: string;
    freezePanes: string;
    zoom: string;
    zoom100: string;
    views: string;
    currentView: string;
    saveView: string;
    deleteView: string;
    objects: string;
  };
  workbookObjects: {
    title: string;
    preservedParts: string;
    tables: string;
    pivotTables: string;
    categories: string;
    tableNames: string;
    tableDetails: string;
    pivotDetails: string;
    sheet: string;
    columnSingular: string;
    columnPlural: string;
    cells: string;
    pivot: string;
    kindLabels: {
      charts: string;
      drawings: string;
      media: string;
      embeddings: string;
      comments: string;
      threadedComments: string;
      pivotTables: string;
      pivotCaches: string;
      queryTables: string;
      slicers: string;
      timelines: string;
      connections: string;
      externalLinks: string;
      controls: string;
      printerSettings: string;
      customXml: string;
      vbaProject: string;
      other: string;
    };
    compatibilityLabels: {
      cellFormatting: string;
      conditionalFormatting: string;
      dataValidation: string;
      hyperlinks: string;
      comments: string;
      definedNames: string;
      sheetProtection: string;
      sheetViews: string;
      loadedTables: string;
      formatAsTable: string;
      pivotLayouts: string;
      pivotAuthoring: string;
      sessionCharts: string;
      chartsDrawings: string;
      chartAuthoring: string;
      externalLinks: string;
    };
    compatibility: string;
    writable: string;
    readOnly: string;
    sessionOnly: string;
    unsupported: string;
    paths: string;
    noteLabel: string;
    readOnlyNote: string;
    empty: string;
    close: string;
  };
  sessionCharts: {
    close: string;
    resize: string;
    columnChart: string;
    lineChart: string;
  };
  autocomplete: {
    customFunction: string;
    structuredTableColumn: string;
    pickFromList: string;
  };
  argHelper: {
    implicitIntersection: string;
  };
  quickAnalysis: {
    title: string;
    groups: {
      formatting: string;
      charts: string;
      totals: string;
      tables: string;
      sparklines: string;
    };
    actions: {
      dataBar: string;
      colorScale: string;
      iconSet: string;
      greaterThan: string;
      top10: string;
      clearFormat: string;
      sumRow: string;
      sumCol: string;
      avgRow: string;
      countRow: string;
      formatAsTable: string;
      pivotStub: string;
      sparkLine: string;
      sparkColumn: string;
      sparkWinLoss: string;
      chartColumn: string;
      chartLine: string;
    };
  };
  sheetTabs: {
    workbookSheets: string;
    previousSheet: string;
    nextSheet: string;
    addSheet: string;
    rename: string;
    renameSheet: string;
    insertSheet: string;
    moveLeft: string;
    moveRight: string;
    deleteSheet: string;
    hideSheet: string;
    unhideSheet: string;
    unhideNamedSheet: string;
  };
  filterDropdown: {
    title: string;
    searchPlaceholder: string;
    selectAll: string;
    blanks: string;
    apply: string;
    clear: string;
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
    /** Hint shown above the table — spreadsheet parity for the Edit Links dialog. */
    note: string;
    close: string;
  };
  cfRulesDialog: {
    /** Modal title — spreadsheet "Manage Rules" parity. */
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
    cancelFormulaEdit: string;
    enterFormula: string;
    expandFormulaBar: string;
    collapseFormulaBar: string;
    spreadsheet: string;
  };
}

export type Locale = 'ja' | 'en';

export type DeepPartial<T> = T extends object ? { [K in keyof T]?: DeepPartial<T[K]> } : T;
