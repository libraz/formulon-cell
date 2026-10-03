export interface GridPopupsStrings {
  contextMenu: {
    title: string;
    copy: string;
    cut: string;
    paste: string;
    pasteSpecial: string;
    insertCopiedCells: string;
    /** Row/column-header variant. Inserts whole rows/columns straight away, so
     *  unlike `insertCopiedCells` it carries no dialog ellipsis. */
    insertCopiedBand: string;
    /** Row/column-header variant for a pending whole-row/column cut. */
    insertCutCells: string;
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
    insert: string;
    delete: string;
    rowHeight: string;
    colWidth: string;
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
    defineName: string;
    pasteFormulas: string;
    pasteFormulasNumFmt: string;
    pasteValues: string;
    pasteValuesNumFmt: string;
    pasteFormatsOnly: string;
    pasteTranspose: string;
    pasteSpecialDialog: string;
    filter: string;
    filterClear: string;
    filterReapply: string;
    filterByValue: string;
    sort: string;
    sortAsc: string;
    sortDesc: string;
    noHiddenRows: string;
    noHiddenColumns: string;
    clipboardUnavailable: string;
    pasteSpecialRequiresCopiedCells: string;
    editPhonetic: string;
    phoneticDialogTitle: string;
    phoneticDialogLabel: string;
    ok: string;
    cancel: string;
  };
  autoFillOptions: {
    title: string;
    copyCells: string;
    fillSeries: string;
    fillFormattingOnly: string;
    fillWithoutFormatting: string;
    fillDays: string;
    fillWeekdays: string;
    fillMonths: string;
    fillYears: string;
  };
  pasteOptions: {
    title: string;
    keepSourceFormatting: string;
    values: string;
    formattingOnly: string;
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
      pivotTable: string;
      sparkLine: string;
      sparkColumn: string;
      sparkWinLoss: string;
      chartColumn: string;
      chartLine: string;
    };
    disabledReasons: {
      requiresNumbers: string;
      requiresTwoNumbers: string;
      requiresThreeNumbers: string;
      requiresMultiCell: string;
      requiresHorizontalRun: string;
      pivotUnavailable: string;
      chartUnavailable: string;
    };
  };
  filterDropdown: {
    title: string;
    searchPlaceholder: string;
    selectAll: string;
    blanks: string;
    condition: string;
    conditionNone: string;
    conditionEquals: string;
    conditionNotEquals: string;
    conditionContains: string;
    conditionNotContains: string;
    conditionGreaterThan: string;
    conditionGreaterThanOrEqual: string;
    conditionLessThan: string;
    conditionLessThanOrEqual: string;
    conditionValue: string;
    apply: string;
    clear: string;
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
}
