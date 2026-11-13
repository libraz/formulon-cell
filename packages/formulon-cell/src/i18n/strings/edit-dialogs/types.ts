export interface EditDialogsStrings {
  fillSeriesDialog: {
    title: string;
    seriesIn: string;
    columns: string;
    rows: string;
    up: string;
    left: string;
    type: string;
    autoFill: string;
    copy: string;
    linear: string;
    growth: string;
    day: string;
    weekday: string;
    month: string;
    year: string;
    ok: string;
    cancel: string;
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
    /** Header used by the native threaded-comment surface. */
    threadTitle?: string;
    /** Placeholder used while composing a native threaded comment. */
    threadPlaceholder?: string;
    /** Placeholder of the card that starts a new conversation on a cell. */
    threadStart: string;
    /** Send button of the conversation card. */
    send: string;
    /** Close button of the conversation card. */
    close: string;
    reply?: string;
    resolve?: string;
    reopen?: string;
    edit?: string;
    noReplies?: string;
    authorFallback?: string;
    noteConflict?: string;
    threadConflict?: string;
    writeFailed?: string;
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
    pasteComments: string;
    pasteValidation: string;
    pasteAllUsingSourceTheme: string;
    pasteAllExceptBorders: string;
    pasteColumnWidths: string;
    pasteAllMergingCF: string;
    /** The Paste Link button. */
    pasteLink: string;
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
  /** Row Height / Column Width: window titles and field labels. */
  dimensionDialog: {
    rowHeightTitle: string;
    rowHeightLabel: string;
    colWidthTitle: string;
    colWidthLabel: string;
  };
  /** Insert / Delete cells: the shift-or-whole-line chooser. */
  cellShiftDialog: {
    insertTitle: string;
    deleteTitle: string;
    shiftRight: string;
    shiftDown: string;
    shiftLeft: string;
    shiftUp: string;
    entireRow: string;
    entireColumn: string;
  };
  insertCopiedCellsDialog: {
    title: string;
    shiftRight: string;
    shiftDown: string;
    cancel: string;
    ok: string;
  };
  findReplace: {
    title: string;
    findTab: string;
    replaceTab: string;
    findWhat: string;
    replaceWith: string;
    optionsMore: string;
    optionsLess: string;
    within: string;
    sheet: string;
    workbook: string;
    search: string;
    byRows: string;
    byColumns: string;
    lookIn: string;
    formulas: string;
    values: string;
    comments: string;
    notes: string;
    matchCase: string;
    matchEntire: string;
    format: string;
    formatUnavailable: string;
    findAll: string;
    prev: string;
    next: string;
    replaceOne: string;
    replaceAll: string;
    close: string;
    bookHeader: string;
    cellHeader: string;
    valueHeader: string;
    cellsFound: string;
    replacedCount: string;
  };
  /** Review > Spelling, labelled as the desktop dialog. */
  spellingDialog: {
    /** Title; `{language}` is the language being checked. */
    title: string;
    notInDictionary: string;
    repeatedWord: string;
    changeTo: string;
    suggestions: string;
    ignoreOnce: string;
    ignoreAll: string;
    add: string;
    change: string;
    changeAll: string;
    /** Change and Change All while the flagged word is a repeated one. */
    delete: string;
    deleteAll: string;
    autoCorrect: string;
    suggest: string;
    undoLast: string;
    cancel: string;
    close: string;
    alwaysSuggest: string;
    ignoreUppercase: string;
    /** The Mac dialog's dictionary list: its label and the custom dictionaries entry. */
    addWordsTo: string;
    customDictionaries: string;
  };
  /** The AutoCorrect options, labelled as the desktop preference pane. */
  autoCorrectDialog: {
    /** Title; `{language}` is the language the lists are for. */
    title: string;
    /** The list's tooltip. */
    description: string;
    replaceAsYouType: string;
    correctTwoInitialCapitals: string;
    capitalizeSentences: string;
    capitalizeDays: string;
    replaceHyperlinks: string;
    replace: string;
    with: string;
    add: string;
    delete: string;
    close: string;
  };
  /** The Mac Preferences window: its panes and their settings, worded as the desktop window. */
  preferencesDialog: {
    /** The window title on Show All, where no pane is shown. */
    windowTitle: string;
    /** The toolbar's history segment: its two buttons and the segment's own name. */
    back: string;
    forward: string;
    history: string;
    /** The toolbar button that lists every pane. */
    showAll: string;
    /** The toolbar search field, which narrows Show All to the panes it names. */
    search: string;
    /** Show All's heading over the panes that set how a book is shown and edited. */
    authoringGroup: string;
    /** The View pane's name, which is also the window title while it is shown. */
    viewPane: string;
    zeroValues: string;
    zeroValuesTip: string;
  };
  goToDialog: {
    /** Go To Special modal title (header and aria-label). */
    title: string;
    /** Go To modal title. */
    goToTitle: string;
    /** Legend of the Go to list of names and recent references. */
    goToLabel: string;
    reference: string;
    referencePlaceholder: string;
    invalidReference: string;
    /** Button that opens Go To Special from Go To. */
    special: string;
    kindComments: string;
    kindConstants: string;
    kindFormulas: string;
    kindNumbers: string;
    kindText: string;
    kindLogical: string;
    kindErrors: string;
    kindBlanks: string;
    kindCurrentRegion: string;
    kindCurrentArray: string;
    kindRowDifferences: string;
    kindColumnDifferences: string;
    kindPrecedents: string;
    kindDependents: string;
    kindDirectOnly: string;
    kindAllLevels: string;
    kindLastCell: string;
    kindVisibleCells: string;
    kindConditionalFormat: string;
    kindDataValidation: string;
    /** Sub-radio for Conditional formats / Data validation: every cell with a rule. */
    ruleAll: string;
    /** Sub-radio: only cells whose rules match the active cell. */
    ruleSame: string;
    /** Alert text when the search finds nothing. */
    noResults: string;
    /** Alert text when the result exceeds the selection cap. */
    cancel: string;
    ok: string;
  };
  /** Drag-moving cells onto a range that holds data. */
  moveCells: {
    /** Confirm title, for assistive tech; the desktop panel shows none. */
    title: string;
    overwriteMessage: string;
    ok: string;
    cancel: string;
  };
  protection: {
    /** Alert raised by an attempt to change a locked cell of a protected sheet. */
    protectedCellEdit: string;
    /** Alert when a slicer click would filter a sheet whose protection refuses it. */
    filterRefused: string;
    /** Format-dialog tab label for the cell-lock section. */
    tabProtection: string;
    /** Checkbox label for the per-cell `locked` flag. */
    locked: string;
    /** Checkbox label for hiding formulas once the sheet is protected. */
    hiddenFormula: string;
    /** Helper text under the locked checkbox explaining that the lock only
     *  takes effect when the sheet is itself protected. */
    lockedHint: string;
    /** Toolbar / menu label that turns sheet protection on. */
    protectSheet: string;
    /** Same control's label when protection is already on. */
    unprotectSheet: string;
    /** Field label for the optional sheet protection password input. */
    password: string;
    /** Placeholder text for the password input. */
    passwordPlaceholder: string;
    /** Field label for confirming a new sheet protection password. */
    confirmPassword: string;
    /** Error message shown when confirmation does not match the password. */
    passwordMismatch: string;
    /** Heading above Protect Sheet permission checkboxes. */
    allowUsersTo: string;
    allowSelectLockedCells: string;
    allowSelectUnlockedCells: string;
    allowFormatCells: string;
    allowFormatColumns: string;
    allowFormatRows: string;
    allowInsertColumns: string;
    allowInsertRows: string;
    allowInsertHyperlinks: string;
    allowDeleteColumns: string;
    allowDeleteRows: string;
    allowSort: string;
    allowAutoFilter: string;
    allowPivotTables: string;
    allowEditObjects: string;
    allowEditScenarios: string;
  };
}
