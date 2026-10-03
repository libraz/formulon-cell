export interface FormulaDialogsStrings {
  namedRangeDialog: {
    title: string;
    nameHeader: string;
    /** Name Manager column header. */
    valueHeader: string;
    /** Name Manager reference column header. */
    formulaHeader: string;
    /** Name Manager quick reference edit label. */
    quickRefersLabel: string;
    /** Name Manager column header. */
    scopeHeader: string;
    /** Name Manager column header. */
    commentHeader: string;
    empty: string;
    /** Read-only fallback note shown when the engine doesn't support write. */
    note: string;
    namePlaceholder: string;
    formulaPlaceholder: string;
    newNameTitle: string;
    editNameTitle: string;
    newButton: string;
    editButton: string;
    filterButton: string;
    selectNameActionReason: string;
    filterRequiresNames: string;
    quickRefersRequiresSelection: string;
    scopeWorkbookOnly: string;
    filterAll: string;
    filterNamesWithErrors: string;
    filterNamesWithoutErrors: string;
    filterNamesScopedToWorkbook: string;
    addButton: string;
    /** Name Manager "Delete" action label. */
    deleteButton: string;
    workbookScope: string;
    valueUnavailable: string;
    /** Inline error: empty or invalid name. */
    errorEmptyName: string;
    /** Inline error: name violates spreadsheet naming rules. */
    errorInvalidName: string;
    /** Inline error: empty formula/ref. */
    errorEmptyFormula: string;
    /** Inline error: engine refused the write. */
    errorEngineFailed: string;
    ok: string;
    cancel: string;
    commitButton: string;
    cancelButton: string;
    confirmDeleteTitle: string;
    confirmDeleteMessage: string;
    close: string;
  };
  iterativeDialog: {
    title: string;
    note: string;
    enable: string;
    maxIterations: string;
    maxChange: string;
    inputsRequireEnabled: string;
    unsupported: string;
    cancel: string;
    ok: string;
  };
  evaluateFormulaDialog: {
    title: string;
    formula: string;
    evaluation: string;
    result: string;
    noFormula: string;
    evaluate: string;
    evaluateRequiresFormula: string;
    evaluateRequiresReference: string;
    evaluateComplete: string;
    close: string;
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
    kindExternalBook: string;
    kindOle: string;
    kindDde: string;
    kindUnknown: string;
    /** Hint shown above the table — spreadsheet parity for the Edit Links dialog. */
    note: string;
    updateValues: string;
    changeSource: string;
    openSource: string;
    breakLink: string;
    checkStatus: string;
    startupPrompt: string;
    readOnlyActionReason: string;
    noSelectionActionReason: string;
    close: string;
  };
  fxDialog: {
    /** Modal title — appears in the header and the aria-label. */
    title: string;
    /** Search input placeholder on the function-picker step. */
    searchPlaceholder: string;
    /** Label for the category picker on the function-picker step. */
    categoryLabel: string;
    categoryAll: string;
    categoryRecent: string;
    categoryLogical: string;
    categoryLookup: string;
    categoryText: string;
    categoryDateTime: string;
    categoryMath: string;
    categoryFinancial: string;
    categoryStatistical: string;
    categoryEngineering: string;
    categoryInformation: string;
    categoryDatabase: string;
    categoryCompatibility: string;
    categoryCube: string;
    categoryWeb: string;
    categoryDynamicArray: string;
    /** Label above the live formula preview on the args step. */
    preview: string;
    /** Empty-state shown when the search yields zero matches. */
    empty: string;
    /** Hint shown when the function signature includes a `...` repeat marker. */
    variadicHint: string;
    /** Generic label for an argument without a host-provided friendly name. */
    argumentLabel: string;
    /** Prefix for rows beyond the function's required arity. */
    optionalArgumentLabel: string;
    /** Button label for appending an argument row. */
    addArgument: string;
    /** Button label for removing the trailing optional argument row. */
    removeArgument: string;
    /** "Back" button — returns to the picker from the args step. */
    back: string;
    cancel: string;
    insert: string;
    insertRequiresFunction: string;
    /** Reason shown when the engine exposes a recognized but unavailable function. */
    functionUnavailable: string;
    /** Optional labels for the nonmodal Mac formula palette. */
    macPalette: {
      title: string;
      recent: string;
      all: string;
      showAll: string;
      insertFunction: string;
      result: string;
      done: string;
      close: string;
      rangePicker: string;
      help: string;
      searchPlaceholder: string;
      empty: string;
      unavailable: string;
      description: string;
      syntax: string;
      argument: string;
      pending: string;
      draftConflict: string;
    };
    /** aria-label for the formula-bar fx button that opens this dialog. */
    fxButtonLabel: string;
    fxButtonUnavailable: string;
  };
  watchPanel: {
    title: string;
    bookHeader: string;
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
}
