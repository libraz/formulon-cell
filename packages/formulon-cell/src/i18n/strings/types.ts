import type { CellFormatStrings } from './cell-format/types.js';
import type { DataObjectsStrings } from './data-objects/types.js';
import type { EditDialogsStrings } from './edit-dialogs/types.js';
import type { FormulaDialogsStrings } from './formula-dialogs/types.js';
import type { GridPopupsStrings } from './grid-popups/types.js';
import type { MacStrings } from './mac/types.js';
import type { PageLayoutStrings } from './page-layout/types.js';
import type { RibbonStrings } from './ribbon/types.js';
import type { RibbonMenusStrings } from './ribbon-menus/types.js';
import type { WorkbookChromeStrings } from './workbook-chrome/types.js';

/**
 * Central registry of every user-visible string in the core package.
 *
 * Adding a new dialog or menu? Append a section to the closest group under
 * `./<group>/types.ts`, populate it in that group's `ja.ts` and `en.ts`, then
 * read from the dictionary in the interact module — never hard-code another
 * label. A new group is registered in the `extends` list below and spread
 * into `en.ts` and `ja.ts`.
 *
 * Consumers can override individual strings (or whole sections) via
 * `Spreadsheet.mount({ strings: ... })`. We deep-merge the overlay onto
 * the locale base, so partial overrides are safe.
 */
export interface Strings
  extends RibbonStrings,
    RibbonMenusStrings,
    CellFormatStrings,
    GridPopupsStrings,
    EditDialogsStrings,
    FormulaDialogsStrings,
    DataObjectsStrings,
    MacStrings,
    PageLayoutStrings,
    WorkbookChromeStrings {}

export type Locale = 'ja' | 'en';

export type DeepPartial<T> = T extends object ? { [K in keyof T]?: DeepPartial<T[K]> } : T;
