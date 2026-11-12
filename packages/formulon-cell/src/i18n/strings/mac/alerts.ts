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

/** Each locale starts from its shared alert texts and overrides the fields the Excel bundle answers. */
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

/** Per-alert overrides of the shared texts. */
export type MacAlertsBundle = { [K in keyof MacAlertStrings]?: Partial<MacAlertStrings[K]> };

/** Applies a locale's bundle overrides over its shared alert texts. */
export const mergeMacAlerts = (
  base: MacAlertStrings,
  bundle: MacAlertsBundle,
): MacAlertStrings => ({
  mergeLoseData: { ...base.mergeLoseData, ...bundle.mergeLoseData },
  deleteSheet: { ...base.deleteSheet, ...bundle.deleteSheet },
  overwriteOnMove: { ...base.overwriteOnMove, ...bundle.overwriteOnMove },
});

/** English Mac app wording from the Excel bundle. The bundle's delete-sheet sentence names the
 *  vendor's product, so deleteSheet keeps its neutral wording in the same shape. */
export const macAlertsEn: MacAlertsBundle = {
  mergeLoseData: {
    message: 'Merging cells only keeps the upper-left value and discards other values.',
  },
};

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

/** Mac app wording from the Excel bundle; not answered by it: mergeLoseData.message, mergeLoseData.cancel, deleteSheet.message, deleteSheet.confirm, deleteSheet.cancel, overwriteOnMove.cancel. */
export const macAlertsDe: MacAlertsBundle = {
  overwriteOnMove: {
    message: 'Hier gibt es schon Daten. Möchten Sie diese ersetzen?',
  },
};

/** Mac app wording from the Excel bundle; not answered by it: mergeLoseData.message, deleteSheet.message, deleteSheet.confirm. */
export const macAlertsFr: MacAlertsBundle = {
  overwriteOnMove: {
    message: 'Cette cellule contient déjà des données. Voulez-vous les remplacer ?',
  },
};

/** Mac app wording from the Excel bundle; not answered by it: mergeLoseData.message, deleteSheet.message, deleteSheet.confirm, overwriteOnMove.message. */
export const macAlertsKo: MacAlertsBundle = {
  mergeLoseData: {
    confirm: '확인',
    cancel: '취소',
  },
  deleteSheet: {
    cancel: '취소',
  },
  overwriteOnMove: {
    confirm: '확인',
    cancel: '취소',
  },
};

/** Mac app wording from the Excel bundle; not answered by it: mergeLoseData.message, mergeLoseData.confirm, deleteSheet.message, deleteSheet.confirm, overwriteOnMove.confirm. */
export const macAlertsTh: MacAlertsBundle = {
  overwriteOnMove: {
    message: 'มีข้อมูลอยู่ที่นี่แล้ว คุณต้องการแทนที่ข้อมูลนั้นหรือไม่',
  },
};

/** Mac app wording from the Excel bundle; not answered by it: mergeLoseData.message, deleteSheet.message, deleteSheet.confirm. */
export const macAlertsZh: MacAlertsBundle = {
  mergeLoseData: {
    confirm: '确定',
    cancel: '取消',
  },
  deleteSheet: {
    cancel: '取消',
  },
  overwriteOnMove: {
    message: '此处已有数据。是否替换它?',
    confirm: '确定',
    cancel: '取消',
  },
};
