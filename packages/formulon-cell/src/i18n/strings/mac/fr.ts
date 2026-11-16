import { cellFormatFr } from '../cell-format/fr.js';
import { dataObjectsFr } from '../data-objects/fr.js';
import { editDialogsFr } from '../edit-dialogs/fr.js';
import { formulaDialogsFr } from '../formula-dialogs/fr.js';
import { gridPopupsFr } from '../grid-popups/fr.js';
import { pageLayoutFr } from '../page-layout/fr.js';
import { ribbonFr } from '../ribbon/fr.js';
import { ribbonMenusFr } from '../ribbon-menus/fr.js';
import { workbookChromeFr } from '../workbook-chrome/fr.js';
import { macAccessibilityPaneFr } from './accessibility-pane.js';
import { macAdvancedFilterFr, macAdvancedFilterFrom } from './advanced-filter.js';
import { macAlertsFr, macAlertsFrom, mergeMacAlerts } from './alerts.js';
import { macCfRulesDialogFr, macCfRulesDialogFrom } from './cf-rules-dialog.js';
import { macChromeFr, macChromeFrom } from './chrome.js';
import { macCommentsPaneFr } from './comments-pane.js';
import { macErrorMenuFr } from './error-menu.js';
import { macExternalLinksPaneFr } from './external-links-pane.js';
import { macFilterPanelFr } from './filter-panel.js';
import { macFindPanelFr, macFindReplaceFr, macFindReplaceFrom } from './find-replace.js';
import { macFormatDialogFr, macFormatDialogFrom } from './format-dialog.js';
import { macGoToDialogFr, macGoToDialogFrom } from './goto-dialog.js';
import { macHyperlinkDialogFr, macHyperlinkDialogFrom } from './hyperlink-dialog.js';
import { macIterativeDialogFr, macIterativeDialogFrom } from './iterative-dialog.js';
import { macNamedRangeDialogFr, macNamedRangeDialogFrom } from './named-range-dialog.js';
import { macPageSetupFr, macPageSetupFrom } from './page-setup.js';
import { macPasteSpecialFr, macPasteSpecialFrom } from './paste-special.js';
import { macPivotFieldSettingsFr } from './pivot-field-settings.js';
import { macPivotOptionsFr } from './pivot-options.js';
import { macPivotTableDialogFr, macPivotTableDialogFrom } from './pivot-table-dialog.js';
import { macProtectSheetFr, macProtectSheetFrom } from './protect-sheet.js';
import { macRemoveDuplicatesFr, macRemoveDuplicatesFrom } from './remove-duplicates.js';
import { macRibbonLabelsFr } from './ribbon-labels.js';
import { macRibbonMenuItemsFr } from './ribbon-menu-items.js';
import { macRibbonReasonsFr } from './ribbon-reasons.js';
import { macSheetTabMenuFr } from './sheet-tab-menu.js';
import { macSheetViewsFr } from './sheet-views.js';
import { macSortDialogFr, macSortDialogFrom } from './sort-dialog.js';
import { macStyleDialogFr } from './style-dialog.js';
import { macTextToColumnsWizardFr } from './text-to-columns-wizard.js';
import type { MacStrings } from './types.js';
import { macWatchWindowFr } from './watch-window.js';
import { macZoomDialogFr, macZoomDialogFrom } from './zoom-dialog.js';

export const macFr: MacStrings = {
  data: {
    goalSeek: {
      title: 'Valeur cible',
      formulaCell: 'Cellule à définir',
      targetValue: 'À la valeur',
      changingCell: 'Cellule variable',
      iterations: 'Nombre maximal d’itérations',
      run: 'OK',
      cancel: 'Annuler',
      result: 'Résultat',
      invalidCell: 'Sélectionnez une seule cellule.',
      invalidTarget: 'Entrez une valeur cible finie.',
      calculating: 'Calcul en cours…',
      converged: 'La recherche d’une valeur cible a convergé.',
    },
    consolidate: {
      title: 'Consolider',
      sources: 'Références',
      destination: 'Copier les résultats dans',
      function: 'Fonction',
      replace: 'Remplacer le contenu actuel',
      labels: 'Utiliser les étiquettes de la ligne supérieure et de la colonne de gauche',
      links: 'Créer des liens vers les données sources',
      unsupportedOption: 'Cette option n’est pas encore prise en charge.',
      run: 'OK',
      cancel: 'Annuler',
      referenceList: 'Toutes les références :',
      select: 'Parcourir...',
      add: 'Ajouter',
      remove: 'Supprimer',
      criteria: 'Étiquettes dans :',
      topRow: 'Ligne du haut',
      leftColumn: 'Colonne de gauche',
    },
    subtotal: {
      title: 'Sous-total',
      range: 'Plage',
      groupBy: 'À chaque modification de',
      columns: 'Ajouter un sous-total à',
      function: 'Utiliser la fonction',
      replace: 'Remplacer les sous-totaux actuels',
      summaryBelow: 'Résumé sous les données',
      run: 'OK',
      cancel: 'Annuler',
      groupTotal: 'Total {group}',
      grandTotal: 'Total général',
      pageBreak: 'Saut de page entre les groupes',
      removeAll: 'Supprimer tout',
    },
    functions: { sum: 'Somme', average: 'Moyenne', count: 'Nombre', min: 'Min', max: 'Max' },
    errors: {
      invalidRange: 'Entrez une plage valide.',
      invalidColumn: 'Entrez des colonnes valides dans la plage.',
      targetNotFinite: 'La valeur cible doit être finie.',
      sameCell: 'La cellule à définir et la cellule variable doivent être différentes.',
      noFormula: 'La cellule à définir doit contenir une formule.',
      changingHasFormula: 'La cellule variable ne doit pas contenir de formule.',
      changingNotNumeric: 'La cellule variable doit être vide ou contenir un nombre.',
      scratchUnsupported: 'Le classeur n’a pas pu conserver la formule dans un calcul temporaire.',
      invalidTolerance: 'La tolérance doit être un nombre fini positif ou nul.',
      invalidIterations: 'Le nombre maximal d’itérations doit être un nombre fini positif.',
      nonFiniteResult: 'La formule n’a pas produit de nombre fini.',
      notConverged: 'La recherche d’une valeur cible n’a pas convergé dans la limite d’itérations.',
      stale: 'Le classeur a changé avant l’application de l’opération.',
      notEditable: 'Les cellules cibles ne sont pas modifiables.',
      writeFailed: 'L’opération n’a pas pu écrire dans les cellules cibles.',
      failed: 'L’opération n’a pas pu être terminée.',
      noSources: 'Ajoutez au moins une plage source.',
      invalidFunction: 'Choisissez une fonction prise en charge.',
      invalidSource: 'Une plage source n’est pas valide.',
      areaLimit: 'La zone source dépasse la limite prise en charge.',
      dimensionMismatch: 'Toutes les plages sources doivent avoir les mêmes dimensions.',
      invalidDestination: 'La destination doit être une cellule ou une plage de même taille.',
      destinationOverlapsSource: 'La destination ne doit pas chevaucher une plage source.',
      activeSheetOnly: 'Les sous-totaux ne peuvent insérer des lignes que dans la feuille active.',
      needsDataRow:
        'Les sous-totaux nécessitent une ligne d’en-tête et au moins une ligne de données.',
      engineUnsupported: 'Le moteur du classeur ne prend pas en charge SOUS.TOTAL.',
      unsupportedObjects:
        'Les sous-totaux ne prennent pas en charge les tableaux, les tableaux croisés dynamiques ou les cellules fusionnées dans la plage sélectionnée.',
      protectedSheet: 'Les sous-totaux ne sont pas disponibles dans une feuille protégée.',
      groupOutsideRange: 'La colonne de regroupement doit se trouver dans la plage sélectionnée.',
      groupLimit: 'La sortie des sous-totaux dépasse la taille prise en charge.',
      noGroups: 'Aucun groupe contigu trouvé.',
      noRoom: 'Il n’y a pas assez de place pour insérer une ligne de sous-total.',
    },
  },
  sparkline: {
    title: 'Insérer un graphique sparkline',
    dataRange: 'Plage de données',
    location: 'Emplacement',
    type: 'Type',
    ok: 'OK',
    cancel: 'Annuler',
    invalidSource: 'Entrez une plage source valide.',
    locationOneCell: 'L’emplacement doit être une seule cellule.',
    chooseType: 'Choisissez un type de graphique sparkline.',
    unavailable: 'Cette opération n’est pas disponible pour la sélection actuelle.',
    protectedLocation: 'L’emplacement du graphique sparkline est protégé.',
    weightTitle: 'Épaisseur de trait du graphique sparkline',
    weightLabel: 'Entrez le poids du graphique sparkline :',
    dateTitle: 'Date sur les axes',
    dateRangeLabel: 'Sélectionnez la plage de dates pour le graphique sparkline :',
    axisHorizontal: 'Horizontal',
    axisVertical: 'Vertical',
    axisGeneral: 'Axe général',
    axisShow: 'Afficher l’axe',
    axisPlotRtl: 'Tracer les données de droite à gauche',
    axisMinimum: 'Valeur minimale :',
    axisMaximum: 'Valeur maximale :',
    axisAutoEach: 'Automatique pour chaque graphique sparkline',
    axisSameAll: 'Identique pour tous les graphiques sparkline',
    axisTitle: 'Axes',
    customValueLabel: 'Valeur personnalisée :',
    hiddenEmptyMenu: 'Cellules masquées et cellules vides',
    locationNotLine:
      'La référence d’emplacement n’est pas valide car les cellules ne se trouvent pas toutes dans la même colonne ou ligne. Sélectionnez des cellules qui se trouvent toutes dans une même colonne ou ligne.',
    dateAxisNotLine:
      'La référence de l’axe des dates n’est pas valide car les cellules ne se trouvent pas toutes dans la même colonne ou ligne. Sélectionnez des cellules qui se trouvent toutes dans une même colonne ou ligne.',
  },
  slicer: {
    ok: 'OK',
    cancel: 'Annuler',
    noTable: 'Aucun tableau n’est disponible.',
    noColumns: 'Le tableau ne contient aucune colonne.',
    insertFailed: 'Le segment n’a pas pu être inséré pour ce tableau.',
  },
  workbookStats: {
    title: 'Statistiques du classeur',
    close: 'Fermer',
    sheets: 'Feuilles',
    populatedCells: 'Cellules renseignées',
    formulas: 'Cellules de formule',
    numbers: 'Cellules numériques',
    text: 'Cellules de texte',
    booleans: 'Cellules booléennes',
    errors: 'Cellules d’erreur',
    tables: 'Tableaux',
    comments: 'Commentaires / notes',
    hyperlinks: 'Liens hypertexte',
    usedRows: 'Lignes utilisées',
    usedColumns: 'Colonnes utilisées',
  },
  ink: {
    strokeLabel: 'Trait manuscrit',
    drawLabel: 'Dessiner à l’encre',
  },
  contextMenu: {
    cut: 'Couper',
    copy: 'Copier',
    paste: 'Coller',
    pasteSpecial: 'Collage spécial',
    pasteSpecialDialog: 'Collage spécial…',
    pasteFormatsOnly: 'Mise en forme',
    insert: 'Insérer…',
    delete: 'Supprimer…',
    clear: 'Effacer',
    newComment: 'Nouveau commentaire',
    newNote: 'Nouvelle note',
    editNote: 'Modifier la note',
    deleteNote: 'Supprimer la note',
    formatCells: 'Format de cellule…',
    pickFromList: 'Liste déroulante de choix…',
    insertCopiedCells: 'Insérer les cellules copiées…',
    insertBand: 'Insérer',
    deleteBand: 'Supprimer',
    insertCopiedBand: 'Insérer les cellules copiées',
    rowHeight: 'Hauteur de ligne…',
    colWidth: 'Largeur de colonne…',
    hide: 'Masquer',
    unhide: 'Afficher',
    showPhonetic: 'Afficher le guide phonétique',
    pasteKeepSource: 'Conserver la mise en forme source',
    pasteNoBorders: 'Aucune bordure',
    pasteColumnWidths: 'Conserver la largeur des colonnes source',
    pasteLink: 'Coller avec liaison',
    defineName: 'Définir un nom…',
    hyperlink: 'Insérer un lien hypertexte…',
  },
  homeRibbon: {
    conditional: 'Mise en forme conditionnelle',
    paste: 'Coller',
    cut: 'Couper',
    copy: 'Copier',
    formatPainter: 'Reproduire la mise en forme',
    cellStyles: 'Styles de cellule',
    autoSum: 'Somme automatique',
  },
  formatDialog: {
    ...{ ...macFormatDialogFrom(cellFormatFr.formatDialog), ...macFormatDialogFr },
    ...macFormatDialogFr,
  },
  iterativeDialog: {
    ...{
      ...macIterativeDialogFrom(formulaDialogsFr.iterativeDialog, {
        calcAutoNoTable: ribbonMenusFr.ribbonMenu.calcAutoNoTable,
      }),
      ...macIterativeDialogFr,
    },
    ...macIterativeDialogFr,
  },
  namedRangeDialog: {
    ...{
      ...macNamedRangeDialogFrom(formulaDialogsFr.namedRangeDialog, {
        defineTitle: 'Définir un nom',
        defineNames: 'Noms :',
        defineNamePrompt: 'Entrez un nom pour la plage de données :',
        defineRangePrompt: 'Sélectionnez la plage de cellules :',
      }),
      ...macNamedRangeDialogFr,
    },
    ...macNamedRangeDialogFr,
  },
  pasteSpecial: {
    ...{ ...macPasteSpecialFrom(editDialogsFr.pasteSpecialDialog), ...macPasteSpecialFr },
    ...macPasteSpecialFr,
  },
  sheetViews: macSheetViewsFr,
  zoomDialog: {
    ...{
      ...macZoomDialogFrom(ribbonFr.ribbon, cellFormatFr.formatDialog),
      ...macZoomDialogFr,
    },
    ...macZoomDialogFr,
  },
  alerts: mergeMacAlerts(
    mergeMacAlerts(
      macAlertsFrom(
        ribbonFr.ribbon,
        {
          message: 'Cette feuille sera définitivement supprimée. Voulez-vous continuer ?',
          confirm: 'Supprimer',
          cancel: 'Annuler',
        },
        editDialogsFr.moveCells,
        {
          headline: 'Impossible de trouver les données que vous recherchez.',
          message:
            'Si vous êtes sûr qu’elles sont présentes dans la feuille actuelle, contrôlez ce que vous avez tapé et réessayez.',
          confirm: 'OK',
        },
      ),
      macAlertsFr,
    ),
    macAlertsFr,
  ),
  protectSheet: {
    ...{
      ...macProtectSheetFrom(editDialogsFr.protection, {
        heading: 'Protéger la feuille et le contenu des cellules verrouillées',
        note: 'Par défaut, toutes les cellules sont verrouillées, mais vous pouvez les déverrouiller.',
      }),
      ...macProtectSheetFr,
    },
    ...macProtectSheetFr,
  },
  advancedFilter: {
    ...{
      ...macAdvancedFilterFrom(ribbonMenusFr.ribbonMenu, {
        filterInPlace: 'Filtrer la liste, sur place',
        copyToLocation: 'Copier vers un autre emplacement',
      }),
      ...macAdvancedFilterFr,
    },
    ...macAdvancedFilterFr,
  },
  chrome: {
    ...{ ...macChromeFrom(workbookChromeFr, formulaDialogsFr), ...macChromeFr },
    ...macChromeFr,
  },
  textToColumnsWizard: macTextToColumnsWizardFr,
  styleDialog: macStyleDialogFr,
  externalLinksPane: macExternalLinksPaneFr,
  watchWindow: macWatchWindowFr,
  accessibilityPane: macAccessibilityPaneFr,
  commentsPane: macCommentsPaneFr,
  ribbonReasons: macRibbonReasonsFr,
  ribbonMenuItems: macRibbonMenuItemsFr,
  ribbonLabels: macRibbonLabelsFr,
  cfRulesDialog: {
    ...{
      ...macCfRulesDialogFrom(cellFormatFr.cfRulesDialog, {
        orderLabel: 'Modifier l’ordre des règles :',
        appliesToLabel: 'S’applique à :',
      }),
      ...macCfRulesDialogFr,
    },
    ...macCfRulesDialogFr,
  },
  hyperlinkDialog: {
    ...{
      ...macHyperlinkDialogFrom(editDialogsFr.hyperlinkDialog, {
        displayText: 'Texte à afficher :',
        selectionInDocument: '<<Sélection dans le document>>',
        screenTip: 'Info-bulle...',
        screenTipTitle: 'Info-bulle du lien hypertexte',
        screenTipLabel: 'Texte de l’info-bulle :',
        tabWeb: 'Page web ou fichier',
        tabDocument: 'Ce document',
        tabEmail: 'Adresse e-mail',
        webIntro: 'Lier à un fichier ou une page web existant.',
        browse: 'Sélectionner...',
        browseUnavailable: 'Impossible de choisir un fichier dans un navigateur. Tapez l’adresse.',
        cellReference: 'Tapez la référence de la cellule :',
        documentPlace: 'Ou sélectionnez un emplacement dans ce document :',
        emailAddress: 'Adresse e-mail :',
        subject: 'Objet :',
        recentEmails: 'Adresses e-mail récemment utilisées :',
      }),
      ...macHyperlinkDialogFr,
    },
    ...macHyperlinkDialogFr,
  },
  pageSetup: {
    ...{ ...macPageSetupFrom(pageLayoutFr.pageSetup), ...macPageSetupFr },
    ...macPageSetupFr,
  },
  dataValidation: {
    inputCaption: 'Quand la cellule est sélectionnée, afficher ce message de saisie :',
    errorCaption: 'Quand des données non valides sont tapées, afficher ce message d’alerte :',
  },
  createTable: { nameLabel: 'Quel est le nom de votre tableau ?' },
  tableStyleDialog: {
    newTitle: 'Nouveau style de tableau',
    modifyTitle: 'Modifier le style de tableau',
    newPivotTitle: 'Nouveau style rapide de tableau croisé dynamique',
    modifyPivotTitle: 'Modifier le style rapide du tableau croisé dynamique',
    setDefault: 'Définir en tant que style de tableau par défaut pour ce classeur',
    name: 'Nom :',
    format: 'Format',
    clear: 'Effacer',
    preview: 'Aperçu',
    stripeSize: 'Taille de rayure',
  },
  goToDialog: {
    ...macGoToDialogFrom(editDialogsFr.goToDialog, {
      selectCaption: 'Sélectionner',
      optionsGroup: 'Options',
      noOptions: 'Cette sélection n’a pas d’options.',
    }),
    ...macGoToDialogFr,
  },
  pivotFieldSettings: macPivotFieldSettingsFr,
  pivotOptions: macPivotOptionsFr,
  pivotTableDialog: {
    ...{
      ...macPivotTableDialogFrom(dataObjectsFr.pivotTableDialog, {
        chooseConnection: 'Choisir la connexion...',
        moveTitle: 'Déplacer le tableau croisé dynamique',
        moveLocation: 'Emplacement :',
      }),
      ...macPivotTableDialogFr,
    },
    ...macPivotTableDialogFr,
  },
  removeDuplicates: {
    ...macRemoveDuplicatesFrom(ribbonMenusFr.ribbonMenu),
    ...macRemoveDuplicatesFr,
  },
  sortDialog: {
    ...{
      ...macSortDialogFrom(ribbonMenusFr.ribbonMenu, {
        levelsHeading: 'Ajouter un niveau de tri :',
        columnHeading: 'Colonne',
        sortOnHeading: 'Trier sur',
        sortOnValues: editDialogsFr.findReplace.values,
        colorHeading: 'Couleur/Icône',
        copyLevel: gridPopupsFr.contextMenu.copy,
        options: 'Options...',
        rowHeading: 'Ligne',
        orientation: 'Orientation',
        sortTopToBottom: 'Trier de haut en bas',
        sortLeftToRight: 'Trier de gauche à droite',
        method: 'Méthode',
        usePhonetic: 'Utiliser le phonétique',
        noPhonetic: 'Ne pas utiliser le phonétique',
        caseSensitive: 'Respecter la casse',
      }),
      ...macSortDialogFr,
    },
    ...macSortDialogFr,
  },
  errorMenu: macErrorMenuFr,
  filterPanel: macFilterPanelFr(gridPopupsFr.filterDropdown),
  sheetTabMenu: macSheetTabMenuFr,
  findPanel: macFindPanelFr,
  findReplace: {
    ...{
      ...macFindReplaceFrom(editDialogsFr.findReplace, 'Distinguer demi-chasse et pleine chasse'),
      ...macFindReplaceFr,
    },
    ...macFindReplaceFr,
  },
};
