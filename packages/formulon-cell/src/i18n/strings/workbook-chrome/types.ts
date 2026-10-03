export interface WorkbookChromeStrings {
  statusBar: {
    ready: string;
    endMode: string;
    enter: string;
    edit: string;
    point: string;
    cell: string;
    cells: string;
    /** Right-click menu heading. */
    aggregatesHeading: string;
    /** Secondary heading for status-bar visibility toggles. */
    statusOptionsHeading: string;
    capsLock: string;
    numLock: string;
    scrollLock: string;
    uploadStatus: string;
    uploadStatusSaved: string;
    uploadStatusSaving: string;
    uploadStatusError: string;
    macroRecording: string;
    macroRecordingActive: string;
    viewShortcuts: string;
    zoomSlider: string;
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
    normalView: string;
    pageLayoutView: string;
    pageBreakPreview: string;
    zoom: string;
    zoomIn: string;
    zoomOut: string;
    zoomInUnavailable: string;
    zoomOutUnavailable: string;
  };
  viewToolbar: {
    title: string;
    normalView: string;
    pageLayoutView: string;
    pageBreakPreview: string;
    gridlines: string;
    headings: string;
    zeros: string;
    formulas: string;
    r1c1: string;
    rightToLeft: string;
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
    deleteViewRequiresActive: string;
    objects: string;
  };
  sheetTabs: {
    workbookSheets: string;
    previousSheet: string;
    nextSheet: string;
    addSheet: string;
    previousSheetUnavailable: string;
    nextSheetUnavailable: string;
    rename: string;
    renameSheet: string;
    insertSheet: string;
    moveLeft: string;
    moveRight: string;
    tabColor: string;
    noColor: string;
    tabColorRed: string;
    tabColorOrange: string;
    tabColorYellow: string;
    tabColorGreen: string;
    tabColorBlue: string;
    tabColorPurple: string;
    tabColorGray: string;
    deleteSheet: string;
    hideSheet: string;
    veryHideSheet: string;
    unhideSheet: string;
    unhideNamedSheet: string;
  };
  a11y: {
    nameBox: string;
    formulaBar: string;
    cancelFormulaEdit: string;
    enterFormula: string;
    cancelFormulaEditUnavailable: string;
    enterFormulaUnavailable: string;
    enterFormulaNoChanges: string;
    expandFormulaBar: string;
    collapseFormulaBar: string;
    spreadsheet: string;
    grid: string;
  };
  mountError: {
    title: string;
    engineHelp: string;
  };
}
