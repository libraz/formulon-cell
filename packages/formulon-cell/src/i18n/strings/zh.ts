import { cellFormatZh } from './cell-format/zh.js';
import { dataObjectsZh } from './data-objects/zh.js';
import { editDialogsZh } from './edit-dialogs/zh.js';
import { formulaDialogsZh } from './formula-dialogs/zh.js';
import { gridPopupsZh } from './grid-popups/zh.js';
import { localeSwitcherZh } from './locale-switcher/zh.js';
import { macZh } from './mac/zh.js';
import { pageLayoutZh } from './page-layout/zh.js';
import { ribbonZh } from './ribbon/zh.js';
import { ribbonMenusZh } from './ribbon-menus/zh.js';
import type { Strings } from './types.js';
import { workbookChromeZh } from './workbook-chrome/zh.js';

export const zh: Strings = {
  ...ribbonZh,
  ...ribbonMenusZh,
  ...cellFormatZh,
  ...gridPopupsZh,
  ...editDialogsZh,
  ...formulaDialogsZh,
  ...dataObjectsZh,
  ...macZh,
  ...pageLayoutZh,
  ...workbookChromeZh,
  ...localeSwitcherZh,
};
