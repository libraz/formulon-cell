// Shared toolbar type and constant declarations (sheet-format action ids,
// sheet range type, tab-colour actions). Type-and-constant only; handlers
// live in `toolbar-actions.ts`.

import type { SpreadsheetInstance } from '../mount/types.js';
import { SHEET_TAB_COLOR_CHOICES } from '../sheet-tab-colors.js';

export type CellFormatAction =
  | 'dialog'
  | 'rowHeight'
  | 'autoFitRowHeight'
  | 'colWidth'
  | 'autoFitColWidth'
  | 'hideRows'
  | 'showRows'
  | 'hideCols'
  | 'showCols'
  | 'renameSheet'
  | 'hideSheet'
  | 'unhideSheet'
  | 'moveSheetLeft'
  | 'moveSheetRight'
  | 'tabColorNone'
  | 'tabColorRed'
  | 'tabColorOrange'
  | 'tabColorYellow'
  | 'tabColorGreen'
  | 'tabColorBlue'
  | 'tabColorPurple'
  | 'tabColorGray'
  | 'protectSheet';

export type SheetRangeFor<I extends SpreadsheetInstance> = ReturnType<
  I['store']['getState']
>['selection']['range'];

export type SheetRange = SheetRangeFor<SpreadsheetInstance>;

export const SHEET_TAB_COLOR_ACTIONS = SHEET_TAB_COLOR_CHOICES.filter(
  (choice) => choice.color !== null,
).map((choice) => ({ action: choice.wrapperAction, color: choice.color })) as readonly {
  action: CellFormatAction;
  color: string;
}[];
