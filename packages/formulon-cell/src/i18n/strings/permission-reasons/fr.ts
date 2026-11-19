import type { PermissionReasonsStrings } from './types.js';

export const permissionReasonsFr: PermissionReasonsStrings = {
  permissionReasons: {
    cellHoldsNote: 'La cellule contient déjà une note.',
    allowEditRangesUnavailable:
      "Autoriser la modification des plages n'est pas disponible sur Mac.",
    selectionOverlapsTable: 'La sélection chevauche un tableau.',
    dialogNoChoiceAllowed: "Aucun choix de la boîte de dialogue n'est autorisé.",
    valuesOnlyClipboardCannotCut:
      'Un presse-papiers contenant uniquement des valeurs ne peut pas être coupé.',
    notBuiltInCommand: "Cette commande n'est pas disponible.",
    selectionDisabled: 'La sélection de cellules est désactivée.',
    copyDisabled: 'La copie est désactivée.',
    controllerDisposed: 'La feuille de calcul a été fermée.',
    workbookUnavailable: "Le classeur n'est pas disponible.",
    hyperlinkNeedsOneCell:
      "Les liens hypertextes automatiques s'appliquent à une seule cellule sans formule.",
    hyperlinkNeedsPermission:
      'Les liens hypertextes automatiques nécessitent les autorisations Mettre en forme les cellules et Insérer des liens hypertextes sur la feuille protégée.',
    mergedRangeTooLarge: 'La plage de cellules fusionnées est trop grande pour être vérifiée.',
    cellOutsideBounds: 'La cellule est en dehors de la zone autorisée.',
    cellOutsideWorkbook: "L'adresse de la cellule est en dehors du classeur.",
    cellNotEditable: "La cellule n'est pas modifiable.",
    sheetHasNoCells: 'Cette feuille ne contient pas de cellules.',
    sheetNotZoomable: 'Cette feuille ne peut pas être zoomée.',
    rangeOutsideBounds: 'La plage est en dehors de la zone autorisée.',
    rangeOutsideWorkbook: 'La plage est en dehors du classeur.',
    rangeNotEditable: 'La plage contient des cellules non modifiables.',
    sheetOutsideBounds: 'La feuille est en dehors de la zone autorisée.',
    sheetNotInWorkbook: 'La feuille ne figure pas dans le classeur.',
    restrictionHookAsync: 'Le hook de restriction doit être synchrone.',
    restrictionCallbackFailed: 'La vérification de la restriction a échoué.',
    restrictedByHost: "Cette action est restreinte par l'hôte.",
    readOnlyPolicy: 'La feuille de calcul est en lecture seule.',
    operationDenied: "Cette opération n'est pas autorisée.",
    affectedRangeMalformed: "La plage concernée n'est pas valide.",
    affectedRangeTooLarge: 'La plage concernée est trop grande.',
    tooManyCellsChanged:
      'Cette action modifierait un trop grand nombre de cellules. Sélectionnez un ensemble plus petit de cellules et réessayez.',
    hostCopyDisabled: "La copie est désactivée par la stratégie de l'hôte.",
    boundedViewportStructural:
      "Cette opération structurelle n'est pas disponible avec une zone d'affichage limitée.",
    restrictedEmbedding: "Non disponible dans l'intégration restreinte.",
    clipboardEmpty: 'Le presse-papiers ne contient rien à coller.',
    noLinkToPaste: 'Aucune liaison à coller.',
    valueNotWritten: 'La valeur n’a pas pu être écrite.',
    cellsNotFilled: 'Les cellules sélectionnées n’ont pas pu être remplies.',
    codeReadOnly: 'La feuille de calcul est en lecture seule.',
    codeOperationDenied: "Cette opération n'est pas autorisée.",
    codeCellIneligible: "La cellule n'est pas modifiable.",
    codeProtected: 'La feuille ou la cellule est protégée.',
    codeOutOfBounds: 'La cible est en dehors de la zone autorisée.',
    codeEffectTooLarge: 'La plage concernée est trop grande.',
    codeInvalid: "La demande n'est pas valide.",
  },
};
