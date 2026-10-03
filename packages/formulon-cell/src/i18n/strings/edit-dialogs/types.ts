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
    findLabel: string;
    replaceLabel: string;
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
    lookInRequiresFindTab: string;
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
  goToDialog: {
    /** Modal title — shown in the header bar and the aria-label. */
    title: string;
    goToTitle: string;
    reference: string;
    referencePlaceholder: string;
    invalidReference: string;
    /** Section legend above the scope radios. */
    scopeLabel: string;
    /** "Active sheet" radio — sweeps every cell on the current sheet. */
    scopeSheet: string;
    /** "Current selection" radio — sweeps only inside the active selection
     *  rectangle. Auto-disabled when the selection is a single cell. */
    scopeSelection: string;
    scopeSelectionRequiresMultiCell: string;
    /** Section legend above the category radios. */
    kindLabel: string;
    kindBlanks: string;
    kindNonBlanks: string;
    kindFormulas: string;
    kindConstants: string;
    kindNumbers: string;
    kindText: string;
    kindLogical: string;
    kindErrors: string;
    kindDataValidation: string;
    kindConditionalFormat: string;
    valueFilterLabel: string;
    /** Inline status when the predicate yields zero matches. The dialog
     *  stays open so the user can adjust the kind. */
    noResults: string;
    cancel: string;
    ok: string;
  };
  protection: {
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
