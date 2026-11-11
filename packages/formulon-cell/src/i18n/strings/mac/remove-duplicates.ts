import type { RibbonMenusStrings } from '../ribbon-menus/types.js';

type RibbonMenuStrings = RibbonMenusStrings['ribbonMenu'];

/** Overrides of the Remove Duplicates wording in `ribbonMenu` that apply on
 *  the Mac host only. */
export type MacRemoveDuplicatesStrings = Partial<RibbonMenuStrings>;

/** The `ribbonMenu` keys the Mac Remove Duplicates dialog can override. Each locale starts from its own shared text and overrides the fields the Excel bundle answers. */
const MAC_REMOVE_DUPLICATES_KEYS = [
  'sortDialogHeader',
] as const satisfies readonly (keyof RibbonMenuStrings)[];

export const macRemoveDuplicatesFrom = (menu: RibbonMenuStrings): MacRemoveDuplicatesStrings =>
  Object.fromEntries(MAC_REMOVE_DUPLICATES_KEYS.map((key) => [key, menu[key]]));

/** Japanese Mac app wording, measured from Excel for Mac. */
export const macRemoveDuplicatesJa: MacRemoveDuplicatesStrings = {
  sortDialogHeader: '先頭行を列見出しとする',
};

/** Mac app wording from the Excel bundle. */
export const macRemoveDuplicatesEn: Partial<MacRemoveDuplicatesStrings> = {
  sortDialogHeader: 'My list has headers',
};

/** Mac app wording from the Excel bundle. */
export const macRemoveDuplicatesDe: Partial<MacRemoveDuplicatesStrings> = {
  sortDialogHeader: 'Meine Liste hat Kopfzeilen',
};

/** Mac app wording from the Excel bundle. */
export const macRemoveDuplicatesFr: Partial<MacRemoveDuplicatesStrings> = {
  sortDialogHeader: 'Ma liste contient des en-têtes',
};

/** Mac app wording from the Excel bundle. */
export const macRemoveDuplicatesKo: Partial<MacRemoveDuplicatesStrings> = {
  sortDialogHeader: '목록에 머리글 포함',
};

/** Mac app wording from the Excel bundle. */
export const macRemoveDuplicatesTh: Partial<MacRemoveDuplicatesStrings> = {
  sortDialogHeader: 'รายการของฉันมีส่วนหัว',
};

/** Mac app wording from the Excel bundle. */
export const macRemoveDuplicatesZh: Partial<MacRemoveDuplicatesStrings> = {
  sortDialogHeader: '列表包含标题',
};
