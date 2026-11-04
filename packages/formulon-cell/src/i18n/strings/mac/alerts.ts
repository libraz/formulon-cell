import type { EditDialogsStrings } from '../edit-dialogs/types.js';
import type { RibbonStrings } from '../ribbon/types.js';

/** A two-button confirm of the Mac desktop's compact alert panel. */
export interface MacConfirmAlertStrings {
  message: string;
  /** The default button, on the right. */
  confirm: string;
  cancel: string;
}

/** Texts of the Mac desktop's compact alert panels. */
export interface MacAlertStrings {
  /** Before a merge that discards every value but the top-left one. */
  mergeLoseData: MacConfirmAlertStrings;
  /** Before deleting a sheet that holds data. */
  deleteSheet: MacConfirmAlertStrings;
  /** Before a drag-move overwrites data. */
  overwriteOnMove: MacConfirmAlertStrings;
}

/** A locale whose Mac wording is not measured keeps its shared merge warning. */
export const macAlertsFrom = (
  ribbon: RibbonStrings['ribbon'],
  deleteSheet: MacConfirmAlertStrings,
  moveCells: EditDialogsStrings['moveCells'],
): MacAlertStrings => ({
  mergeLoseData: {
    message: ribbon.mergeLoseDataMessage,
    confirm: ribbon.mergeLoseDataConfirm,
    cancel: ribbon.mergeLoseDataCancel,
  },
  deleteSheet,
  overwriteOnMove: {
    message: moveCells.overwriteMessage,
    confirm: moveCells.ok,
    cancel: moveCells.cancel,
  },
});

/** Japanese Mac app wording, measured from Excel for Mac. */
export const macAlertsJa: MacAlertStrings = {
  mergeLoseData: {
    message: 'セルを結合すると、左上の値のみが保持され、他のセルの値は破棄されます。',
    confirm: 'OK',
    cancel: 'キャンセル',
  },
  deleteSheet: {
    message: 'このシートは完全に削除されます。続けますか?',
    confirm: '削除',
    cancel: 'キャンセル',
  },
  overwriteOnMove: {
    message: '既にデータがありますが、置き換えますか?',
    confirm: 'OK',
    cancel: 'キャンセル',
  },
};
