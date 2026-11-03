import type { EditDialogsStrings } from '../edit-dialogs/types.js';

type FindReplaceStrings = EditDialogsStrings['findReplace'];

/** Overrides of `findReplace` that apply on the Mac host only, plus the
 *  half-width / full-width option the Mac app adds. */
export type MacFindReplaceStrings = Partial<FindReplaceStrings> & { matchWidth: string };

/** The `findReplace` keys the Mac host can override. A locale whose Mac wording
 *  is not measured repeats its own shared text for them. */
const MAC_FIND_REPLACE_KEYS = [
  'optionsMore',
  'optionsLess',
  'search',
  'lookIn',
  'matchCase',
  'matchEntire',
  'cellsFound',
] as const satisfies readonly (keyof FindReplaceStrings)[];

export const macFindReplaceFrom = (
  dialog: FindReplaceStrings,
  matchWidth: string,
): MacFindReplaceStrings => ({
  ...Object.fromEntries(MAC_FIND_REPLACE_KEYS.map((key) => [key, dialog[key]])),
  matchWidth,
});
