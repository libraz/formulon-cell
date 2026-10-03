import { cellFormatJa } from './cell-format/ja.js';
import { dataObjectsJa } from './data-objects/ja.js';
import { editDialogsJa } from './edit-dialogs/ja.js';
import { formulaDialogsJa } from './formula-dialogs/ja.js';
import { gridPopupsJa } from './grid-popups/ja.js';
import { macJa } from './mac/ja.js';
import { pageLayoutJa } from './page-layout/ja.js';
import { ribbonJa } from './ribbon/ja.js';
import { ribbonMenusJa } from './ribbon-menus/ja.js';
import type { Strings } from './types.js';
import { workbookChromeJa } from './workbook-chrome/ja.js';

export const ja: Strings = {
  ...ribbonJa,
  ...ribbonMenusJa,
  ...cellFormatJa,
  ...gridPopupsJa,
  ...editDialogsJa,
  ...formulaDialogsJa,
  ...dataObjectsJa,
  ...macJa,
  ...pageLayoutJa,
  ...workbookChromeJa,
};
