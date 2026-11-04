import type { ExternalLinksPaneStrings } from '../../../interact/external-links-pane.js';
import type { TextToColumnsWizardStrings } from '../../../toolbar/dialogs/text-to-columns.js';
import type { CellFormatStrings } from '../cell-format/types.js';
import type { MacAdvancedFilterStrings } from './advanced-filter.js';
import type { MacCfRulesDialogStrings } from './cf-rules-dialog.js';
import type { MacChromeStrings } from './chrome.js';
import type { MacCreateTableStrings } from './create-table.js';
import type { MacDataValidationStrings } from './data-validation.js';
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
import type { MacSortDialogStrings } from './sort-dialog.js';
import type { MacStyleDialogStrings } from './style-dialog.js';
import type { MacZoomDialogStrings } from './zoom-dialog.js';

export interface MacStrings {
  /** Overrides of `formatDialog` that apply on the Mac host only. */
  macFormatDialog: Partial<CellFormatStrings['formatDialog']>;
  /** Overrides of `findReplace` that apply on the Mac host only. */
  macFindReplace: MacFindReplaceStrings;
  /** Overrides of `iterativeDialog`, plus its Mac-only labels, that apply on the Mac host only. */
  macIterativeDialog: MacIterativeDialogStrings;
  /** Overrides of `namedRangeDialog` that apply on the Mac host only. */
  macNamedRangeDialog: MacNamedRangeDialogStrings;
  /** Overrides of `pasteSpecialDialog` that apply on the Mac host only. */
  macPasteSpecial: MacPasteSpecialStrings;
  /** Overrides of `pivotTableDialog`, plus its Mac-only labels, that apply on the Mac host only. */
  macPivotTableDialog: MacPivotTableDialogStrings;
  /** The PivotTable Field dialog for value and row/column fields. */
  macPivotFieldSettings: MacPivotFieldSettingsStrings;
  /** The PivotTable Options dialog. */
  macPivotOptions: MacPivotOptionsStrings;
  /** Overrides of the Remove Duplicates wording that apply on the Mac host only. */
  macRemoveDuplicates: MacRemoveDuplicatesStrings;
  /** Overrides of the Sort dialog wording, plus its Mac-only labels, that apply on the Mac host only. */
  macSortDialog: MacSortDialogStrings;
  /** Overrides of `protection`, plus its Mac-only lines, for the Mac Protect Sheet dialog. */
  macProtectSheet: MacProtectSheetStrings;
  /** Overrides of the Advanced Filter wording, plus its Mac-only radios, on the Mac host. */
  macAdvancedFilter: MacAdvancedFilterStrings;
  /** Overrides of `goToDialog`, plus its Mac-only lines, that apply on the Mac host only. */
  macGoToDialog: MacGoToDialogStrings;
  /** Accessible names of the Mac formula bar and status bar controls. */
  macChrome: MacChromeStrings;
  /** The Mac three-step Convert Text to Columns wizard. */
  macTextToColumnsWizard: TextToColumnsWizardStrings;
  /** The Mac Style dialog opened from New Cell Style. */
  macStyleDialog: MacStyleDialogStrings;
  /** The Mac Workbook Links pane. */
  macExternalLinksPane: ExternalLinksPaneStrings;
  /** Overrides of `cfRulesDialog`, plus its Mac-only labels, that apply on the Mac host only. */
  macCfRulesDialog: MacCfRulesDialogStrings;
  /** Overrides of `hyperlinkDialog`, plus its Mac-only labels, that apply on the Mac host only. */
  macHyperlinkDialog: MacHyperlinkDialogStrings;
  /** Overrides of `pageSetup` that apply on the Mac host only. */
  macPageSetup: MacPageSetupStrings;
  /** Captions only the Mac Data Validation message tabs show. */
  macDataValidation: MacDataValidationStrings;
  /** Labels only the Mac Create Table dialog shows. */
  macCreateTable: MacCreateTableStrings;
  /** Labels of the Mac Zoom dialog. */
  macZoomDialog: MacZoomDialogStrings;
  macData: {
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
      unsupportedOption: string;
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
  macSparkline: {
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
  };
  macSlicer: {
    table: string;
    ok: string;
    cancel: string;
    noTable: string;
    noColumns: string;
    chooseTableAndColumn: string;
    insertFailed: string;
  };
  macWorkbookStats: {
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
  macInk: {
    strokeLabel: string;
    drawLabel: string;
  };
  /** Cell context-menu labels where the Mac desktop menu differs from the
   * default profile's wording. */
  macContextMenu: {
    cut: string;
    copy: string;
    paste: string;
    pasteSpecial: string;
    pasteSpecialDialog: string;
    pasteFormatsOnly: string;
    insert: string;
    delete: string;
    clear: string;
    newNote: string;
    editNote: string;
    deleteNote: string;
    formatCells: string;
    editPhonetic: string;
    defineName: string;
    hyperlink: string;
  };
  /** Home ribbon labels where the Mac desktop ribbon differs from the
   * default profile's wording. */
  macHomeRibbon: {
    conditional: string;
    paste: string;
    cut: string;
    copy: string;
    formatPainter: string;
    cellStyles: string;
    autoSum: string;
  };
}
