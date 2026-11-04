import { cellFormatEn } from '../cell-format/en.js';
import { dataObjectsEn } from '../data-objects/en.js';
import { editDialogsEn } from '../edit-dialogs/en.js';
import { formulaDialogsEn } from '../formula-dialogs/en.js';
import { gridPopupsEn } from '../grid-popups/en.js';
import { pageLayoutEn } from '../page-layout/en.js';
import { ribbonEn } from '../ribbon/en.js';
import { ribbonMenusEn } from '../ribbon-menus/en.js';
import { workbookChromeEn } from '../workbook-chrome/en.js';
import { macAdvancedFilterFrom } from './advanced-filter.js';
import { macAlertsFrom } from './alerts.js';
import { macCfRulesDialogFrom } from './cf-rules-dialog.js';
import { macChromeFrom } from './chrome.js';
import { macExternalLinksPaneEn } from './external-links-pane.js';
import { macFindReplaceFrom } from './find-replace.js';
import { macFormatDialogFrom } from './format-dialog.js';
import { macGoToDialogFrom } from './goto-dialog.js';
import { macHyperlinkDialogFrom } from './hyperlink-dialog.js';
import { macIterativeDialogFrom } from './iterative-dialog.js';
import { macNamedRangeDialogFrom } from './named-range-dialog.js';
import { macPageSetupFrom } from './page-setup.js';
import { macPasteSpecialFrom } from './paste-special.js';
import { macPivotFieldSettingsEn } from './pivot-field-settings.js';
import { macPivotOptionsEn } from './pivot-options.js';
import { macPivotTableDialogFrom } from './pivot-table-dialog.js';
import { macProtectSheetFrom } from './protect-sheet.js';
import { macRemoveDuplicatesFrom } from './remove-duplicates.js';
import { macSortDialogFrom } from './sort-dialog.js';
import { macStyleDialogEn } from './style-dialog.js';
import { macTextToColumnsWizardEn } from './text-to-columns-wizard.js';
import type { MacStrings } from './types.js';
import { macWatchWindowEn } from './watch-window.js';
import { macZoomDialogFrom } from './zoom-dialog.js';

export const macEn: MacStrings = {
  macData: {
    goalSeek: {
      title: 'Goal Seek',
      formulaCell: 'Set cell',
      targetValue: 'To value',
      changingCell: 'By changing cell',
      iterations: 'Maximum iterations',
      run: 'OK',
      cancel: 'Cancel',
      result: 'Result',
      invalidCell: 'Select a single cell.',
      invalidTarget: 'Enter a finite target value.',
      calculating: 'Calculating…',
      converged: 'Goal Seek converged.',
    },
    consolidate: {
      title: 'Consolidate',
      sources: 'References',
      sourcesHint: 'One range per line',
      destination: 'Copy results to',
      function: 'Function',
      replace: 'Replace current contents',
      labels: 'Use labels in top row and left column',
      links: 'Create links to source data',
      unsupportedOption: 'This option is not supported yet.',
      run: 'OK',
      cancel: 'Cancel',
      referenceList: 'All references:',
      select: 'Browse...',
      add: 'Add',
      remove: 'Delete',
      criteria: 'Use labels in:',
      topRow: 'Top row',
      leftColumn: 'Left column',
    },
    subtotal: {
      title: 'Subtotal',
      range: 'Range',
      groupBy: 'At each change in',
      columns: 'Add subtotal to',
      function: 'Use function',
      replace: 'Replace current subtotals',
      summaryBelow: 'Summary below data',
      run: 'OK',
      cancel: 'Cancel',
      groupTotal: '{group} Total',
      grandTotal: 'Grand Total',
      pageBreak: 'Page break between groups',
      removeAll: 'Remove All',
    },
    functions: { sum: 'Sum', average: 'Average', count: 'Count', min: 'Min', max: 'Max' },
    errors: {
      invalidRange: 'Enter a valid range.',
      invalidColumn: 'Enter valid columns inside the range.',
      targetNotFinite: 'Target value must be finite.',
      sameCell: 'The formula cell and changing cell must be different.',
      noFormula: 'The formula cell must contain a formula.',
      changingHasFormula: 'The changing cell must not contain a formula.',
      changingNotNumeric: 'The changing cell must be blank or contain a number.',
      scratchUnsupported: 'The workbook could not preserve the formula in a scratch calculation.',
      invalidTolerance: 'Tolerance must be a finite non-negative number.',
      invalidIterations: 'Maximum iterations must be a finite positive number.',
      nonFiniteResult: 'The formula did not produce a finite number.',
      notConverged: 'Goal Seek did not converge within the iteration limit.',
      stale: 'The workbook changed before the operation could be applied.',
      notEditable: 'The target cells are not editable.',
      writeFailed: 'The operation could not write the target cells.',
      failed: 'The operation could not be completed.',
      noSources: 'Add at least one source range.',
      invalidFunction: 'Choose a supported function.',
      invalidSource: 'A source range is not valid.',
      areaLimit: 'The source area exceeds the supported limit.',
      dimensionMismatch: 'All source ranges must have the same dimensions.',
      invalidDestination: 'Destination must be one cell or an equal-size range.',
      destinationOverlapsSource: 'The destination must not overlap a source range.',
      activeSheetOnly: 'Subtotal can only insert rows on the active sheet.',
      needsDataRow: 'Subtotal needs a header and at least one data row.',
      engineUnsupported: 'The workbook engine does not support SUBTOTAL.',
      unsupportedObjects:
        'Subtotal does not support tables, PivotTables, or merged cells in the selected range.',
      protectedSheet: 'Subtotal is unavailable on a protected sheet.',
      groupOutsideRange: 'The group column must be inside the selected range.',
      groupLimit: 'Subtotal exceeds the supported output size.',
      noGroups: 'No contiguous groups were found.',
      noRoom: 'There is no room to insert a subtotal row.',
    },
  },
  macSparkline: {
    title: 'Insert Sparkline',
    dataRange: 'Data range',
    location: 'Location',
    type: 'Type',
    ok: 'OK',
    cancel: 'Cancel',
    invalidSource: 'Enter a valid source range.',
    locationOneCell: 'Location must be one cell.',
    outOfBounds: 'The selected range is outside the worksheet.',
    chooseType: 'Choose a sparkline type.',
    unavailable: 'This operation is unavailable for the current selection.',
    protectedLocation: 'The sparkline location is protected.',
  },
  macSlicer: {
    table: 'Table',
    ok: 'OK',
    cancel: 'Cancel',
    noTable: 'No table is available.',
    noColumns: 'The table has no columns.',
    chooseTableAndColumn: 'Choose a table and column.',
    insertFailed: 'The slicer could not be inserted for this table.',
  },
  macWorkbookStats: {
    title: 'Workbook Statistics',
    close: 'Close',
    sheets: 'Sheets',
    populatedCells: 'Populated cells',
    formulas: 'Formula cells',
    numbers: 'Number cells',
    text: 'Text cells',
    booleans: 'Boolean cells',
    errors: 'Error cells',
    tables: 'Tables',
    comments: 'Comments / notes',
    hyperlinks: 'Hyperlinks',
    usedRows: 'Used rows',
    usedColumns: 'Used columns',
  },
  macInk: {
    strokeLabel: 'Ink stroke',
    drawLabel: 'Draw Ink',
  },
  macContextMenu: {
    cut: 'Cut',
    copy: 'Copy',
    paste: 'Paste',
    pasteSpecial: 'Paste Special',
    pasteSpecialDialog: 'Paste Special…',
    pasteFormatsOnly: 'Formatting',
    insert: 'Insert…',
    delete: 'Delete…',
    clear: 'Clear Contents',
    newNote: 'New Note',
    editNote: 'Edit Note',
    deleteNote: 'Delete Note',
    formatCells: 'Format Cells…',
    editPhonetic: 'Edit Phonetic Guide…',
    defineName: 'Define Name…',
    hyperlink: 'Hyperlink…',
  },
  macHomeRibbon: {
    conditional: 'Conditional Formatting',
    paste: 'Paste',
    cut: 'Cut',
    copy: 'Copy',
    formatPainter: 'Format',
    cellStyles: 'Cell Styles',
    autoSum: 'AutoSum',
  },
  macFormatDialog: macFormatDialogFrom(cellFormatEn.formatDialog),
  macIterativeDialog: macIterativeDialogFrom(formulaDialogsEn.iterativeDialog, {
    calcAutoNoTable: ribbonMenusEn.ribbonMenu.calcAutoNoTable,
  }),
  macNamedRangeDialog: macNamedRangeDialogFrom(formulaDialogsEn.namedRangeDialog, {
    defineTitle: 'Define Name',
    defineNames: 'Names:',
    defineNamePrompt: 'Enter a name for the data range:',
    defineRangePrompt: 'Select the range of cells:',
  }),
  macPasteSpecial: macPasteSpecialFrom(editDialogsEn.pasteSpecialDialog),
  macZoomDialog: macZoomDialogFrom(ribbonEn.ribbon, cellFormatEn.formatDialog),
  macAlerts: macAlertsFrom(ribbonEn.ribbon, {
    message: 'This sheet will be permanently deleted. Do you want to continue?',
    confirm: 'Delete',
    cancel: 'Cancel',
  }),
  macProtectSheet: macProtectSheetFrom(editDialogsEn.protection, {
    heading: 'Protect worksheet and contents of locked cells',
    note: 'By default, all cells are locked, but you can unlock them.',
  }),
  macAdvancedFilter: macAdvancedFilterFrom(ribbonMenusEn.ribbonMenu, {
    filterInPlace: 'Filter the list, in-place',
    copyToLocation: 'Copy to another location',
  }),
  macChrome: macChromeFrom(workbookChromeEn, formulaDialogsEn),
  macTextToColumnsWizard: macTextToColumnsWizardEn,
  macStyleDialog: macStyleDialogEn,
  macExternalLinksPane: macExternalLinksPaneEn,
  macWatchWindow: macWatchWindowEn,
  macCfRulesDialog: macCfRulesDialogFrom(cellFormatEn.cfRulesDialog, {
    orderLabel: 'Change rule order:',
    appliesToLabel: 'Applies to:',
  }),
  macHyperlinkDialog: macHyperlinkDialogFrom(editDialogsEn.hyperlinkDialog, {
    displayText: 'Text to display:',
    selectionInDocument: '<<Selection in Document>>',
    screenTip: 'ScreenTip...',
    screenTipTitle: 'Set Hyperlink ScreenTip',
    screenTipLabel: 'ScreenTip text:',
    tabWeb: 'Web Page or File',
    tabDocument: 'This Document',
    tabEmail: 'Email Address',
    webIntro: 'Link to an existing file or web page.',
    browse: 'Select...',
    browseUnavailable: 'Files cannot be chosen in a browser. Type the address.',
    cellReference: 'Type in the cell reference:',
    documentPlace: 'Or select a place in this document:',
    emailAddress: 'Email address:',
    subject: 'Subject:',
    recentEmails: 'Recently used email addresses:',
  }),
  macPageSetup: macPageSetupFrom(pageLayoutEn.pageSetup),
  macDataValidation: {
    inputCaption: 'When cell is selected, show this input message:',
    errorCaption: 'When user enters invalid data, show this error alert:',
  },
  macCreateTable: { nameLabel: 'What is the name of your table?' },
  macGoToDialog: macGoToDialogFrom(editDialogsEn.goToDialog, {
    selectCaption: 'Select',
    optionsGroup: 'Options',
    noOptions: 'This selection has no options.',
  }),
  macPivotFieldSettings: macPivotFieldSettingsEn,
  macPivotOptions: macPivotOptionsEn,
  macPivotTableDialog: macPivotTableDialogFrom(dataObjectsEn.pivotTableDialog, {
    chooseConnection: 'Choose Connection...',
  }),
  macRemoveDuplicates: macRemoveDuplicatesFrom(ribbonMenusEn.ribbonMenu),
  macSortDialog: macSortDialogFrom(ribbonMenusEn.ribbonMenu, {
    levelsHeading: 'Add a level to sort by:',
    columnHeading: 'Column',
    sortOnHeading: 'Sort On',
    sortOnValues: editDialogsEn.findReplace.values,
    colorHeading: 'Color/Icon',
    copyLevel: gridPopupsEn.contextMenu.copy,
    options: 'Options...',
  }),
  macFindReplace: macFindReplaceFrom(editDialogsEn.findReplace, 'Match half-width and full-width'),
};
