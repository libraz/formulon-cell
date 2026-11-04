import { cellFormatDe } from '../cell-format/de.js';
import { dataObjectsDe } from '../data-objects/de.js';
import { editDialogsDe } from '../edit-dialogs/de.js';
import { formulaDialogsDe } from '../formula-dialogs/de.js';
import { gridPopupsDe } from '../grid-popups/de.js';
import { pageLayoutDe } from '../page-layout/de.js';
import { ribbonDe } from '../ribbon/de.js';
import { ribbonMenusDe } from '../ribbon-menus/de.js';
import { workbookChromeDe } from '../workbook-chrome/de.js';
import { macAccessibilityPaneDe } from './accessibility-pane.js';
import { macAdvancedFilterFrom } from './advanced-filter.js';
import { macAlertsFrom } from './alerts.js';
import { macCfRulesDialogFrom } from './cf-rules-dialog.js';
import { macChromeFrom } from './chrome.js';
import { macExternalLinksPaneDe } from './external-links-pane.js';
import { macFindReplaceFrom } from './find-replace.js';
import { macFormatDialogFrom } from './format-dialog.js';
import { macGoToDialogFrom } from './goto-dialog.js';
import { macHyperlinkDialogFrom } from './hyperlink-dialog.js';
import { macIterativeDialogFrom } from './iterative-dialog.js';
import { macNamedRangeDialogFrom } from './named-range-dialog.js';
import { macPageSetupFrom } from './page-setup.js';
import { macPasteSpecialFrom } from './paste-special.js';
import { macPivotFieldSettingsDe } from './pivot-field-settings.js';
import { macPivotOptionsDe } from './pivot-options.js';
import { macPivotTableDialogFrom } from './pivot-table-dialog.js';
import { macProtectSheetFrom } from './protect-sheet.js';
import { macRemoveDuplicatesFrom } from './remove-duplicates.js';
import { macSortDialogFrom } from './sort-dialog.js';
import { macStyleDialogDe } from './style-dialog.js';
import { macTextToColumnsWizardDe } from './text-to-columns-wizard.js';
import type { MacStrings } from './types.js';
import { macWatchWindowDe } from './watch-window.js';
import { macZoomDialogFrom } from './zoom-dialog.js';

export const macDe: MacStrings = {
  macData: {
    goalSeek: {
      title: 'Zielwertsuche',
      formulaCell: 'Zielzelle',
      targetValue: 'Zielwert',
      changingCell: 'Veränderbare Zelle',
      iterations: 'Maximale Iterationen',
      run: 'OK',
      cancel: 'Abbrechen',
      result: 'Ergebnis',
      invalidCell: 'Wählen Sie eine einzelne Zelle aus.',
      invalidTarget: 'Geben Sie einen endlichen Zielwert ein.',
      calculating: 'Wird berechnet …',
      converged: 'Die Zielwertsuche ist konvergiert.',
    },
    consolidate: {
      title: 'Konsolidieren',
      sources: 'Verweise',
      sourcesHint: 'Ein Bereich pro Zeile',
      destination: 'Ergebnisse kopieren nach',
      function: 'Funktion',
      replace: 'Aktuelle Inhalte ersetzen',
      labels: 'Beschriftungen in oberster Zeile und linker Spalte verwenden',
      links: 'Verknüpfungen mit Quelldaten erstellen',
      unsupportedOption: 'Diese Option wird noch nicht unterstützt.',
      run: 'OK',
      cancel: 'Abbrechen',
      referenceList: 'Vorhandene Verweise:',
      select: 'Durchsuchen...',
      add: 'Hinzufügen',
      remove: 'Löschen',
      criteria: 'Beschriftung aus:',
      topRow: 'Oberster Zeile',
      leftColumn: 'Linker Spalte',
    },
    subtotal: {
      title: 'Teilergebnis',
      range: 'Bereich',
      groupBy: 'Bei jeder Änderung in',
      columns: 'Teilergebnis hinzufügen zu',
      function: 'Funktion verwenden',
      replace: 'Aktuelle Teilergebnisse ersetzen',
      summaryBelow: 'Zusammenfassung unterhalb der Daten',
      run: 'OK',
      cancel: 'Abbrechen',
      groupTotal: '{group} Ergebnis',
      grandTotal: 'Gesamtergebnis',
      pageBreak: 'Seitenumbrüche zwischen Gruppen einfügen',
      removeAll: 'Alle entfernen',
    },
    functions: { sum: 'Summe', average: 'Mittelwert', count: 'Anzahl', min: 'Min', max: 'Max' },
    errors: {
      invalidRange: 'Geben Sie einen gültigen Bereich ein.',
      invalidColumn: 'Geben Sie gültige Spalten innerhalb des Bereichs ein.',
      targetNotFinite: 'Der Zielwert muss endlich sein.',
      sameCell: 'Zielzelle und veränderbare Zelle müssen verschieden sein.',
      noFormula: 'Die Zielzelle muss eine Formel enthalten.',
      changingHasFormula: 'Die veränderbare Zelle darf keine Formel enthalten.',
      changingNotNumeric: 'Die veränderbare Zelle muss leer sein oder eine Zahl enthalten.',
      scratchUnsupported:
        'Die Formel konnte in einer temporären Berechnung nicht beibehalten werden.',
      invalidTolerance: 'Die Toleranz muss eine endliche, nicht negative Zahl sein.',
      invalidIterations: 'Die maximale Iterationszahl muss eine endliche positive Zahl sein.',
      nonFiniteResult: 'Die Formel hat keine endliche Zahl ergeben.',
      notConverged: 'Die Zielwertsuche ist innerhalb der Iterationsgrenze nicht konvergiert.',
      stale: 'Die Arbeitsmappe wurde geändert, bevor der Vorgang angewendet werden konnte.',
      notEditable: 'Die Zielzellen können nicht bearbeitet werden.',
      writeFailed: 'Die Zielzellen konnten nicht beschrieben werden.',
      failed: 'Der Vorgang konnte nicht abgeschlossen werden.',
      noSources: 'Fügen Sie mindestens einen Quellbereich hinzu.',
      invalidFunction: 'Wählen Sie eine unterstützte Funktion aus.',
      invalidSource: 'Ein Quellbereich ist ungültig.',
      areaLimit: 'Der Quellbereich überschreitet die unterstützte Grenze.',
      dimensionMismatch: 'Alle Quellbereiche müssen dieselben Abmessungen haben.',
      invalidDestination: 'Das Ziel muss eine Zelle oder ein gleich großer Bereich sein.',
      destinationOverlapsSource: 'Das Ziel darf keinen Quellbereich überlappen.',
      activeSheetOnly: 'Teilergebnisse können nur auf dem aktiven Blatt Zeilen einfügen.',
      needsDataRow:
        'Für Teilergebnisse sind eine Kopfzeile und mindestens eine Datenzeile erforderlich.',
      engineUnsupported: 'Die Arbeitsmappen-Engine unterstützt SUBTOTAL nicht.',
      unsupportedObjects:
        'Teilergebnisse unterstützen Tabellen, PivotTables oder verbundene Zellen im ausgewählten Bereich nicht.',
      protectedSheet: 'Auf einem geschützten Blatt sind keine Teilergebnisse verfügbar.',
      groupOutsideRange: 'Die Gruppierungsspalte muss innerhalb des ausgewählten Bereichs liegen.',
      groupLimit: 'Das Teilergebnis überschreitet die unterstützte Ausgabegröße.',
      noGroups: 'Keine zusammenhängenden Gruppen gefunden.',
      noRoom: 'Es ist kein Platz zum Einfügen einer Teilergebniszeile vorhanden.',
    },
  },
  macSparkline: {
    title: 'Sparkline einfügen',
    dataRange: 'Datenbereich',
    location: 'Position',
    type: 'Typ',
    ok: 'OK',
    cancel: 'Abbrechen',
    invalidSource: 'Geben Sie einen gültigen Quellbereich ein.',
    locationOneCell: 'Die Position muss eine Zelle sein.',
    outOfBounds: 'Der ausgewählte Bereich liegt außerhalb des Arbeitsblatts.',
    chooseType: 'Wählen Sie einen Sparkline-Typ aus.',
    unavailable: 'Dieser Vorgang ist für die aktuelle Auswahl nicht verfügbar.',
    protectedLocation: 'Die Sparkline-Position ist geschützt.',
  },
  macSlicer: {
    ok: 'OK',
    cancel: 'Abbrechen',
    noTable: 'Keine Tabelle verfügbar.',
    noColumns: 'Die Tabelle enthält keine Spalten.',
    insertFailed: 'Der Datenschnitt konnte für diese Tabelle nicht eingefügt werden.',
  },
  macWorkbookStats: {
    title: 'Arbeitsmappenstatistik',
    close: 'Schließen',
    sheets: 'Blätter',
    populatedCells: 'Ausgefüllte Zellen',
    formulas: 'Formelzellen',
    numbers: 'Zellen mit Zahlen',
    text: 'Textzellen',
    booleans: 'Boolesche Zellen',
    errors: 'Fehlerzellen',
    tables: 'Tabellen',
    comments: 'Kommentare / Notizen',
    hyperlinks: 'Hyperlinks',
    usedRows: 'Verwendete Zeilen',
    usedColumns: 'Verwendete Spalten',
  },
  macInk: {
    strokeLabel: 'Tintenstrich',
    drawLabel: 'Mit Tinte zeichnen',
  },
  macContextMenu: {
    cut: 'Ausschneiden',
    copy: 'Kopieren',
    paste: 'Einfügen',
    pasteSpecial: 'Inhalte einfügen',
    pasteSpecialDialog: 'Inhalte einfügen…',
    pasteFormatsOnly: 'Formatierung',
    insert: 'Einfügen…',
    delete: 'Löschen…',
    clear: 'Löschen',
    newNote: 'Neue Notiz',
    editNote: 'Notiz bearbeiten',
    deleteNote: 'Notiz löschen',
    formatCells: 'Zellen formatieren…',
    editPhonetic: 'Phonetischen Leitfaden bearbeiten…',
    defineName: 'Namen definieren…',
    hyperlink: 'Hyperlink einfügen…',
  },
  macHomeRibbon: {
    conditional: 'Bedingte Formatierung',
    paste: 'Einfügen',
    cut: 'Ausschneiden',
    copy: 'Kopieren',
    formatPainter: 'Format übertragen',
    cellStyles: 'Zellenformatvorlagen',
    autoSum: 'AutoSumme',
  },
  macFormatDialog: macFormatDialogFrom(cellFormatDe.formatDialog),
  macIterativeDialog: macIterativeDialogFrom(formulaDialogsDe.iterativeDialog, {
    calcAutoNoTable: ribbonMenusDe.ribbonMenu.calcAutoNoTable,
  }),
  macNamedRangeDialog: macNamedRangeDialogFrom(formulaDialogsDe.namedRangeDialog, {
    defineTitle: 'Namen definieren',
    defineNames: 'Namen:',
    defineNamePrompt: 'Geben Sie einen Namen für den Datenbereich ein:',
    defineRangePrompt: 'Wählen Sie den Zellbereich aus:',
  }),
  macPasteSpecial: macPasteSpecialFrom(editDialogsDe.pasteSpecialDialog),
  macZoomDialog: macZoomDialogFrom(ribbonDe.ribbon, cellFormatDe.formatDialog),
  macAlerts: macAlertsFrom(ribbonDe.ribbon, {
    message: 'Dieses Blatt wird endgültig gelöscht. Möchten Sie fortfahren?',
    confirm: 'Löschen',
    cancel: 'Abbrechen',
  }),
  macProtectSheet: macProtectSheetFrom(editDialogsDe.protection, {
    heading: 'Arbeitsblatt und Inhalt gesperrter Zellen schützen',
    note: 'Standardmäßig sind alle Zellen gesperrt, die Sperre kann aber aufgehoben werden.',
  }),
  macAdvancedFilter: macAdvancedFilterFrom(ribbonMenusDe.ribbonMenu, {
    filterInPlace: 'Liste an gleicher Stelle filtern',
    copyToLocation: 'An eine andere Stelle kopieren',
  }),
  macChrome: macChromeFrom(workbookChromeDe, formulaDialogsDe),
  macTextToColumnsWizard: macTextToColumnsWizardDe,
  macStyleDialog: macStyleDialogDe,
  macExternalLinksPane: macExternalLinksPaneDe,
  macWatchWindow: macWatchWindowDe,
  macAccessibilityPane: macAccessibilityPaneDe,
  macCfRulesDialog: macCfRulesDialogFrom(cellFormatDe.cfRulesDialog, {
    orderLabel: 'Regelreihenfolge ändern:',
    appliesToLabel: 'Wird angewendet auf:',
  }),
  macHyperlinkDialog: macHyperlinkDialogFrom(editDialogsDe.hyperlinkDialog, {
    displayText: 'Anzuzeigender Text:',
    selectionInDocument: '<<Auswahl im Dokument>>',
    screenTip: 'QuickInfo...',
    screenTipTitle: 'QuickInfo für Link festlegen',
    screenTipLabel: 'QuickInfo-Text:',
    tabWeb: 'Webseite oder Datei',
    tabDocument: 'Aktuelles Dokument',
    tabEmail: 'E-Mail-Adresse',
    webIntro: 'Link zu einer vorhandenen Datei oder Webseite.',
    browse: 'Auswählen...',
    browseUnavailable:
      'Im Browser können keine Dateien ausgewählt werden. Geben Sie die Adresse ein.',
    cellReference: 'Geben Sie den Zellbezug ein:',
    documentPlace: 'Oder wählen Sie eine Stelle im Dokument:',
    emailAddress: 'E-Mail-Adresse:',
    subject: 'Betreff:',
    recentEmails: 'Zuletzt verwendete E-Mail-Adressen:',
  }),
  macPageSetup: macPageSetupFrom(pageLayoutDe.pageSetup),
  macDataValidation: {
    inputCaption: 'Diese Eingabemeldung anzeigen, wenn die Zelle ausgewählt wird:',
    errorCaption: 'Diese Fehlermeldung anzeigen, wenn ungültige Daten eingegeben wurden:',
  },
  macCreateTable: { nameLabel: 'Wie lautet der Name Ihrer Tabelle?' },
  macGoToDialog: macGoToDialogFrom(editDialogsDe.goToDialog, {
    selectCaption: 'Auswählen',
    optionsGroup: 'Optionen',
    noOptions: 'Für diese Auswahl gibt es keine Optionen.',
  }),
  macPivotFieldSettings: macPivotFieldSettingsDe,
  macPivotOptions: macPivotOptionsDe,
  macPivotTableDialog: macPivotTableDialogFrom(dataObjectsDe.pivotTableDialog, {
    chooseConnection: 'Verbindung auswählen...',
  }),
  macRemoveDuplicates: macRemoveDuplicatesFrom(ribbonMenusDe.ribbonMenu),
  macSortDialog: macSortDialogFrom(ribbonMenusDe.ribbonMenu, {
    levelsHeading: 'Sortierebene hinzufügen:',
    columnHeading: 'Spalte',
    sortOnHeading: 'Sortieren nach',
    sortOnValues: editDialogsDe.findReplace.values,
    colorHeading: 'Farbe/Symbol',
    copyLevel: gridPopupsDe.contextMenu.copy,
    options: 'Optionen...',
  }),
  macFindReplace: macFindReplaceFrom(
    editDialogsDe.findReplace,
    'Halbbreite und Vollbreite unterscheiden',
  ),
};
