import { cellFormatEn } from './cell-format/en.js';
import { dataObjectsEn } from './data-objects/en.js';
import { editDialogsEn } from './edit-dialogs/en.js';
import { formulaDialogsEn } from './formula-dialogs/en.js';
import { gridPopupsEn } from './grid-popups/en.js';
import { macEn } from './mac/en.js';
import { pageLayoutEn } from './page-layout/en.js';
import { ribbonEn } from './ribbon/en.js';
import { ribbonMenusEn } from './ribbon-menus/en.js';
import type { Strings } from './types.js';
import { workbookChromeEn } from './workbook-chrome/en.js';

export const en: Strings = {
  ...ribbonEn,
  ...ribbonMenusEn,
  ...cellFormatEn,
  ...gridPopupsEn,
  ...editDialogsEn,
  ...formulaDialogsEn,
  ...dataObjectsEn,
  ...macEn,
  ...pageLayoutEn,
  ...workbookChromeEn,
};
