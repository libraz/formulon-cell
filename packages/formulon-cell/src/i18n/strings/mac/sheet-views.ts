/** Wording of the Mac Sheet View switcher and its Options window. */
export interface MacSheetViewStrings {
  /** The sheet's own view, outside every named sheet view. */
  defaultView: string;
  /** A view opened by New that has not been kept under a name yet. */
  temporaryView: string;
  /** Title of the Sheet View Options window. */
  optionsTitle: string;
  switchTo: string;
  rename: string;
  duplicate: string;
  delete: string;
  cancel: string;
  /** Title of the rename sheet. */
  renameTitle: string;
  /** Headline of the delete confirmation. */
  deleteTitle: string;
  /** The delete confirmation; `{name}` is the view's name. */
  deleteMessage: string;
  /** Refusal to rename Default. */
  renameDefault: string;
  /** Refusal of an empty name or one with a semicolon. */
  nameInvalid: string;
  /** Refusal of a name another view of the sheet holds. */
  nameTaken: string;
}

/** Mac app wording from the Excel bundle. */
export const macSheetViewsJa: MacSheetViewStrings = {
  defaultView: '既定',
  temporaryView: '一時ビュー',
  optionsTitle: 'シート ビューのオプション',
  switchTo: '切り替え',
  rename: '名前の変更',
  duplicate: '複製',
  delete: '削除',
  cancel: 'キャンセル',
  renameTitle: 'シート ビューの名前の変更',
  deleteTitle: 'このシート ビューを削除しますか?',
  deleteMessage: '"{name}" を削除しますか?',
  renameDefault: '他のシート ビューの名前は変更できますが、既定のビューは変更できません。',
  nameInvalid: '空ではなく、かつ ";" が含まれていない名前を入力してください',
  nameTaken:
    '同じ名前のシート ビューが既に存在します。別の名前を入力するか、もう一方のシート ビューを削除してから、もう一度お試しください。',
};

/** Mac app wording from the Excel bundle. */
export const macSheetViewsEn: MacSheetViewStrings = {
  defaultView: 'Default',
  temporaryView: 'Temporary View',
  optionsTitle: 'Sheet View options',
  switchTo: 'Switch to',
  rename: 'Rename',
  duplicate: 'Duplicate',
  delete: 'Delete',
  cancel: 'Cancel',
  renameTitle: 'Rename Sheet View',
  deleteTitle: 'Delete this sheet view?',
  deleteMessage: 'Are you sure you want to delete "{name}"?',
  renameDefault: 'You can rename other sheet views, but not the Default view.',
  nameInvalid: "Type a name that is not empty and doesn't contain the character ';'",
  nameTaken:
    "There's already a sheet view with that name. Type a different name, or delete the other one and try again.",
};

/** Mac app wording from the Excel bundle. */
export const macSheetViewsDe: MacSheetViewStrings = {
  defaultView: 'Standard',
  temporaryView: 'Temporäre Ansicht',
  optionsTitle: 'Tabellenansichtsoptionen',
  switchTo: 'Wechseln zu',
  rename: 'Umbenennen',
  duplicate: 'Duplizieren',
  delete: 'Löschen',
  cancel: 'Abbrechen',
  renameTitle: 'Tabellenansicht umbenennen',
  deleteTitle: 'Diese Tabellenansicht löschen?',
  deleteMessage: 'Möchten Sie "{name}" löschen?',
  renameDefault: 'Sie können andere Tabellenansichten umbenennen, aber nicht die Standardansicht.',
  nameInvalid: 'Geben Sie einen Namen ein, der nicht leer ist und nicht das Zeichen ";" enthält.',
  nameTaken:
    'Es ist bereits eine Tabellenansicht mit diesem Namen vorhanden. Geben Sie einen anderen Namen ein, oder löschen Sie den anderen, und versuchen Sie es noch mal.',
};

/** Mac app wording from the Excel bundle. */
export const macSheetViewsFr: MacSheetViewStrings = {
  defaultView: 'Par défaut',
  temporaryView: 'Affichage temporaire',
  optionsTitle: 'Options d’affichage de la Feuille',
  switchTo: 'Basculer vers',
  rename: 'Renommer',
  duplicate: 'Dupliquer',
  delete: 'Supprimer',
  cancel: 'Annuler',
  renameTitle: "Renommer l'affichage de la feuille",
  deleteTitle: 'Supprimer l’affichage de la feuille ?',
  deleteMessage: 'Voulez-vous vraiment supprimer «{name}» ?',
  renameDefault:
    "Vous pouvez renommer d'autres affichages Tableau, mais pas l'affichage par défaut.",
  nameInvalid: "Tapez un nom qui n'est pas vide et qui ne contient pas le caractère « ; »",
  nameTaken:
    "Un affichage tableau portant ce nom existe déjà. Entrez un autre nom ou supprimez l'autre nom, puis réessayez.",
};

/** Mac app wording from the Excel bundle. */
export const macSheetViewsKo: MacSheetViewStrings = {
  defaultView: '기본값',
  temporaryView: '임시 보기',
  optionsTitle: '시트 보기 옵션',
  switchTo: '전환',
  rename: '이름 바꾸기',
  duplicate: '복제',
  delete: '삭제',
  cancel: '취소',
  renameTitle: '시트 보기 이름 바꾸기',
  deleteTitle: '이 시트 보기를 삭제하시겠습니까?',
  deleteMessage: '"{name}"을(를) 삭제하시겠습니까?',
  renameDefault: '기본 보기가 아닌 다른 시트 보기의 이름을 바꿀 수 있습니다.',
  nameInvalid: "비어 있지 않고 ';' 문자를 포함하지 않는 이름을 입력하세요.",
  nameTaken:
    '동일한 이름의 시트 보기가 이미 있습니다. 다른 이름을 입력하거나 동일한 이름의 다른 시트 보기를 삭제하고 다시 시도하세요.',
};

/** Mac app wording from the Excel bundle. */
export const macSheetViewsTh: MacSheetViewStrings = {
  defaultView: 'ค่าเริ่มต้น',
  temporaryView: 'มุมมองชั่วคราว',
  optionsTitle: 'ตัวเลือกมุมมองแผ่นงาน',
  switchTo: 'สลับไปยัง',
  rename: 'เปลี่ยนชื่อ',
  duplicate: 'สำเนา',
  delete: 'ลบ',
  cancel: 'ยกเลิก',
  renameTitle: 'เปลี่ยนชื่อมุมมองแผ่นงาน',
  deleteTitle: 'ลบมุมมองแผ่นงานนี้หรือไม่',
  deleteMessage: 'คุณแน่ใจหรือไม่ว่าคุณต้องการลบ "{name}"',
  renameDefault: 'คุณสามารถเปลี่ยนชื่อมุมมองแผ่นงานอื่นได้แต่ไม่ใช่กับมุมมองเริ่มต้น',
  nameInvalid: "พิมพ์ชื่อที่ไม่ว่างเปล่าและไม่มีอักขระ ';'",
  nameTaken: 'มีมุมมองแผ่นงานที่มีชื่อนั้นอยู่แล้ว พิมพ์ชื่ออื่น หรือลบมุมมองแผ่นงานอื่นแล้วลองอีกครั้ง',
};

/** Mac app wording from the Excel bundle. */
export const macSheetViewsZh: MacSheetViewStrings = {
  defaultView: '默认',
  temporaryView: '临时视图',
  optionsTitle: '工作表视图选项',
  switchTo: '切换到',
  rename: '重命名',
  duplicate: '复制',
  delete: '删除',
  cancel: '取消',
  renameTitle: '重命名工作表视图',
  deleteTitle: '删除此工作表视图?',
  deleteMessage: '是否确定要删除 "{name}"?',
  renameDefault: '你可以重命名其他工作表视图，但不能重命名默认视图。',
  nameInvalid: '键入一个不为空且不包含字符 ";" 的名称',
  nameTaken: '已存在具有该名称的工作表视图。请键入其他名称或删除另一个名称，然后重试。',
};
