import type { RibbonMenusStrings } from '../ribbon-menus/types.js';

type RibbonMenuStrings = RibbonMenusStrings['ribbonMenu'];

/** Overrides of the Remove Duplicates wording in `ribbonMenu` that apply on
 *  the Mac host only. */
export type MacRemoveDuplicatesStrings = Partial<RibbonMenuStrings>;

/** The `ribbonMenu` keys the Mac Remove Duplicates dialog can override. A
 *  locale whose Mac wording is not measured repeats its own shared text. */
const MAC_REMOVE_DUPLICATES_KEYS = [
  'sortDialogHeader',
] as const satisfies readonly (keyof RibbonMenuStrings)[];

export const macRemoveDuplicatesFrom = (menu: RibbonMenuStrings): MacRemoveDuplicatesStrings =>
  Object.fromEntries(MAC_REMOVE_DUPLICATES_KEYS.map((key) => [key, menu[key]]));

/** Japanese Mac app wording, measured from Excel for Mac. */
export const macRemoveDuplicatesJa: MacRemoveDuplicatesStrings = {
  sortDialogHeader: '先頭行を列見出しとする',
};
