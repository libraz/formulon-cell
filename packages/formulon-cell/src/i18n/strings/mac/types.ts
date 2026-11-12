import type { ExternalLinksPaneStrings } from '../../../interact/external-links-pane.js';
import type { TextToColumnsWizardStrings } from '../../../toolbar/dialogs/text-to-columns.js';
import type { CellFormatStrings } from '../cell-format/types.js';
import type { MacAccessibilityPaneStrings } from './accessibility-pane.js';
import type { MacAdvancedFilterStrings } from './advanced-filter.js';
import type { MacAlertStrings } from './alerts.js';
import type { MacCfRulesDialogStrings } from './cf-rules-dialog.js';
import type { MacChromeStrings } from './chrome.js';
import type { MacCreateTableStrings } from './create-table.js';
import type { MacDataValidationStrings } from './data-validation.js';
import type { MacErrorMenuStrings } from './error-menu.js';
import type { MacFilterPanelStrings } from './filter-panel.js';
import type { MacFindReplaceStrings } from './find-replace.js';
import type { MacGoToDialogStrings } from './goto-dialog.js';
import type { MacHyperlinkDialogStrings } from './hyperlink-dialog.js';
import type { MacIterativeDialogStrings } from './iterative-dialog.js';
import type { MacNamedRangeDialogStrings } from './named-range-dialog.js';
import type { MacPageSetupStrings } from './page-setup.js';
import type { MacPasteSpecialStrings } from './paste-special.js';
import type { MacPivotFieldSettingsStrings } from './pivot-field-settings.js';
import type { MacPivotOptionsStrings } from './pivot-options.js';
import type { MacPivotTableDialogStrings } from './pivot-table-dialog.js';
import type { MacProtectSheetStrings } from './protect-sheet.js';
import type { MacRemoveDuplicatesStrings } from './remove-duplicates.js';
import type { MacSheetTabMenuStrings } from './sheet-tab-menu.js';
import type { MacSortDialogStrings } from './sort-dialog.js';
import type { MacStyleDialogStrings } from './style-dialog.js';
import type { MacWatchWindowStrings } from './watch-window.js';
import type { MacZoomDialogStrings } from './zoom-dialog.js';

export interface MacStrings {
  /** Overrides of `formatDialog` that apply on the Mac host only. */
  formatDialog: Partial<CellFormatStrings['formatDialog']>;
  /** Overrides of `findReplace` that apply on the Mac host only. */
  errorMenu: MacErrorMenuStrings;
  filterPanel: MacFilterPanelStrings;
  sheetTabMenu: MacSheetTabMenuStrings;
  findReplace: MacFindReplaceStrings;
  /** Overrides of `iterativeDialog`, plus its Mac-only labels, that apply on the Mac host only. */
  iterativeDialog: MacIterativeDialogStrings;
  /** Overrides of `namedRangeDialog` that apply on the Mac host only. */
  namedRangeDialog: MacNamedRangeDialogStrings;
  /** Overrides of `pasteSpecialDialog` that apply on the Mac host only. */
  pasteSpecial: MacPasteSpecialStrings;
  /** Overrides of `pivotTableDialog`, plus its Mac-only labels, that apply on the Mac host only. */
  pivotTableDialog: MacPivotTableDialogStrings;
  /** The PivotTable Field dialog for value and row/column fields. */
  pivotFieldSettings: MacPivotFieldSettingsStrings;
  /** The PivotTable Options dialog. */
  pivotOptions: MacPivotOptionsStrings;
  /** Overrides of the Remove Duplicates wording that apply on the Mac host only. */
  removeDuplicates: MacRemoveDuplicatesStrings;
  /** Overrides of the Sort dialog wording, plus its Mac-only labels, that apply on the Mac host only. */
  sortDialog: MacSortDialogStrings;
  /** Overrides of `protection`, plus its Mac-only lines, for the Mac Protect Sheet dialog. */
  protectSheet: MacProtectSheetStrings;
  /** Overrides of the Advanced Filter wording, plus its Mac-only radios, on the Mac host. */
  advancedFilter: MacAdvancedFilterStrings;
  /** Overrides of `goToDialog`, plus its Mac-only lines, that apply on the Mac host only. */
  goToDialog: MacGoToDialogStrings;
  /** Accessible names of the Mac formula bar and status bar controls. */
  chrome: MacChromeStrings;
  /** The Mac three-step Convert Text to Columns wizard. */
  textToColumnsWizard: TextToColumnsWizardStrings;
  /** The Mac Style dialog opened from New Cell Style. */
  styleDialog: MacStyleDialogStrings;
  /** The Mac Workbook Links pane. */
  externalLinksPane: ExternalLinksPaneStrings;
  /** The Mac Watch Window task pane. */
  watchWindow: MacWatchWindowStrings;
  /** The Mac accessibility checker task pane. */
  accessibilityPane: MacAccessibilityPaneStrings;
  /** Overrides of `cfRulesDialog`, plus its Mac-only labels, that apply on the Mac host only. */
  cfRulesDialog: MacCfRulesDialogStrings;
  /** Overrides of `hyperlinkDialog`, plus its Mac-only labels, that apply on the Mac host only. */
  hyperlinkDialog: MacHyperlinkDialogStrings;
  /** Overrides of `pageSetup` that apply on the Mac host only. */
  pageSetup: MacPageSetupStrings;
  /** Captions only the Mac Data Validation message tabs show. */
  dataValidation: MacDataValidationStrings;
  /** Labels only the Mac Create Table dialog shows. */
  createTable: MacCreateTableStrings;
  /** Labels of the Mac New / Modify Table Style and PivotTable Quick Style dialog. */
  tableStyleDialog: {
    newTitle: string;
    modifyTitle: string;
    newPivotTitle: string;
    modifyPivotTitle: string;
    setDefault: string;
    name: string;
    format: string;
    clear: string;
    preview: string;
    stripeSize: string;
  };
  /** Labels of the Mac Zoom dialog. */
  zoomDialog: MacZoomDialogStrings;
  /** Texts of the Mac compact alert panels. */
  alerts: MacAlertStrings;
  data: {
    goalSeek: {
      title: string;
      formulaCell: string;
      targetValue: string;
      changingCell: string;
      iterations: string;
      run: string;
      cancel: string;
      result: string;
      invalidCell: string;
      invalidTarget: string;
      calculating: string;
      converged: string;
    };
    consolidate: {
      title: string;
      sources: string;
      sourcesHint: string;
      destination: string;
      function: string;
      replace: string;
      labels: string;
      links: string;
      unsupportedOption: string;
      run: string;
      cancel: string;
      /** Caption of the list of added source references. */
      referenceList: string;
      select: string;
      add: string;
      remove: string;
      /** Heading of the label-source checkboxes. */
      criteria: string;
      topRow: string;
      leftColumn: string;
    };
    subtotal: {
      title: string;
      range: string;
      groupBy: string;
      columns: string;
      function: string;
      replace: string;
      summaryBelow: string;
      run: string;
      cancel: string;
      /** Label of a subtotal row; `{group}` is the group's value. */
      groupTotal: string;
      /** Label of the grand-total row. */
      grandTotal: string;
      pageBreak: string;
      removeAll: string;
    };
    functions: { sum: string; average: string; count: string; min: string; max: string };
    /** Keyed by the error codes returned from the Data command planners. */
    errors: {
      invalidRange: string;
      invalidColumn: string;
      targetNotFinite: string;
      sameCell: string;
      noFormula: string;
      changingHasFormula: string;
      changingNotNumeric: string;
      scratchUnsupported: string;
      invalidTolerance: string;
      invalidIterations: string;
      nonFiniteResult: string;
      notConverged: string;
      stale: string;
      notEditable: string;
      writeFailed: string;
      failed: string;
      noSources: string;
      invalidFunction: string;
      invalidSource: string;
      areaLimit: string;
      dimensionMismatch: string;
      invalidDestination: string;
      destinationOverlapsSource: string;
      activeSheetOnly: string;
      needsDataRow: string;
      engineUnsupported: string;
      unsupportedObjects: string;
      protectedSheet: string;
      groupOutsideRange: string;
      groupLimit: string;
      noGroups: string;
      noRoom: string;
    };
  };
  sparkline: {
    title: string;
    dataRange: string;
    location: string;
    type: string;
    ok: string;
    cancel: string;
    invalidSource: string;
    locationOneCell: string;
    outOfBounds: string;
    chooseType: string;
    unavailable: string;
    protectedLocation: string;
    weightTitle: string;
    weightLabel: string;
    dateTitle: string;
    dateRangeLabel: string;
    /** The Axis dialog: its tabs, the horizontal axis options and the two vertical bounds. */
    axisHorizontal: string;
    axisVertical: string;
    axisGeneral: string;
    axisShow: string;
    axisPlotRtl: string;
    axisMinimum: string;
    axisMaximum: string;
    axisAutoEach: string;
    axisSameAll: string;
    /** Why Plot Data Right-to-Left stays greyed. */
    axisRtlUnsupported: string;
    axisTitle: string;
    customValueLabel: string;
    /** Edit Data > Hidden & Empty Cells: a submenu trigger, so no ellipsis. */
    hiddenEmptyMenu: string;
    /** Edit Sparklines refusing a location that spans rows and columns. */
    locationNotLine: string;
    /** The Axis dialog refusing a date range that spans rows and columns. */
    dateAxisNotLine: string;
  };
  slicer: {
    ok: string;
    cancel: string;
    noTable: string;
    noColumns: string;
    insertFailed: string;
  };
  workbookStats: {
    title: string;
    close: string;
    sheets: string;
    populatedCells: string;
    formulas: string;
    numbers: string;
    text: string;
    booleans: string;
    errors: string;
    tables: string;
    comments: string;
    hyperlinks: string;
    usedRows: string;
    usedColumns: string;
  };
  ink: {
    strokeLabel: string;
    drawLabel: string;
  };
  /** Cell context-menu labels where the Mac desktop menu differs from the
   * default profile's wording. */
  contextMenu: {
    cut: string;
    copy: string;
    paste: string;
    pasteSpecial: string;
    pasteSpecialDialog: string;
    pasteFormatsOnly: string;
    insert: string;
    delete: string;
    clear: string;
    newComment: string;
    newNote: string;
    editNote: string;
    deleteNote: string;
    formatCells: string;
    pickFromList: string;
    insertCopiedCells: string;
    insertBand: string;
    deleteBand: string;
    insertCopiedBand: string;
    rowHeight: string;
    colWidth: string;
    hide: string;
    unhide: string;
    showPhonetic: string;
    pasteKeepSource: string;
    pasteNoBorders: string;
    pasteColumnWidths: string;
    pasteLink: string;
    defineName: string;
    hyperlink: string;
  };
  /** Home ribbon labels where the Mac desktop ribbon differs from the
   * default profile's wording. */
  homeRibbon: {
    conditional: string;
    paste: string;
    cut: string;
    copy: string;
    formatPainter: string;
    cellStyles: string;
    autoSum: string;
  };
}
