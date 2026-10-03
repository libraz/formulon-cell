export interface PageLayoutStrings {
  /** Page Layout and Page Break Preview chrome — the text painted onto the
   *  page itself rather than into any toolbar. */
  pageView: {
    /** Placeholder in an empty header band. */
    addHeader: string;
    /** Placeholder in an empty footer band. */
    addFooter: string;
    /** Page-number watermark. `{n}` is the 1-based page number. */
    pageNumber: string;
    /** Announced when a break line is picked up for dragging. */
    movePageBreak: string;
    /** Unit the Page Layout rulers are marked in. */
    rulerUnit: 'in' | 'cm';
  };
  pageSetup: {
    /** Modal title — shown in the header bar and the aria-label. */
    title: string;
    tabPage: string;
    tabMargins: string;
    tabHeaderFooter: string;
    tabSheet: string;
    orientation: string;
    orientPortrait: string;
    orientLandscape: string;
    printerProfile: string;
    printerProfileAutomatic: string;
    printerProfileRefresh: string;
    printerProfileRefreshInProgress: string;
    printerProfileRefreshFailed: string;
    paperSize: string;
    margins: string;
    marginTop: string;
    marginRight: string;
    marginBottom: string;
    marginLeft: string;
    marginHeader: string;
    marginFooter: string;
    printerMargins: string;
    printableTop: string;
    printableRight: string;
    printableBottom: string;
    printableLeft: string;
    printableMarginWarning: string;
    centerOnPage: string;
    centerHorizontally: string;
    centerVertically: string;
    headerLabel: string;
    footerLabel: string;
    headerBuiltin: string;
    footerBuiltin: string;
    customHeader: string;
    customFooter: string;
    headerNone: string;
    headerPageNumber: string;
    headerSheetName: string;
    footerPageNumber: string;
    footerWorkbookPath: string;
    differentOddEvenPages: string;
    differentFirstPage: string;
    scaleWithDocument: string;
    alignWithPageMargins: string;
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
    printArea: string;
    printAreaPlaceholder: string;
    invalidPrintArea: string;
    invalidPrintTitleRows: string;
    invalidPrintTitleCols: string;
    printOptions: string;
    blackAndWhite: string;
    draftQuality: string;
    comments: string;
    commentsNone: string;
    commentsAsDisplayed: string;
    commentsEndOfSheet: string;
    cellErrorsAs: string;
    cellErrorsDisplayed: string;
    cellErrorsBlank: string;
    cellErrorsDash: string;
    cellErrorsNA: string;
    pageOrder: string;
    pageOrderDownThenOver: string;
    pageOrderOverThenDown: string;
    scaling: string;
    adjustTo: string;
    percentNormalSize: string;
    fitTo: string;
    pagesWideBy: string;
    tall: string;
    printQuality: string;
    printQualityAutomatic: string;
    firstPageNumber: string;
    firstPageNumberPlaceholder: string;
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
  /** Page-layout scaling controls (Width / Height / Scale dropdowns). */
  pageScale: {
    width: string;
    height: string;
    scale: string;
    fitWidth: string;
    fitHeight: string;
    automatic: string;
    page: string;
    pages: string;
    custom: string;
    customScalePrompt: string;
    customPagesPrompt: string;
    invalidScale: string;
    invalidPages: string;
    ok: string;
    cancel: string;
  };
}
