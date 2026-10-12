import { cellFormatKo } from './cell-format/ko.js';
import { dataObjectsKo } from './data-objects/ko.js';
import { editDialogsKo } from './edit-dialogs/ko.js';
import { formulaDialogsKo } from './formula-dialogs/ko.js';
import { gridPopupsKo } from './grid-popups/ko.js';
import { localeSwitcherKo } from './locale-switcher/ko.js';
import { macKo } from './mac/ko.js';
import { pageLayoutKo } from './page-layout/ko.js';
import { ribbonKo } from './ribbon/ko.js';
import { ribbonMenusKo } from './ribbon-menus/ko.js';
import type { Strings } from './types.js';
import { workbookChromeKo } from './workbook-chrome/ko.js';

export const ko: Strings = {
  ...ribbonKo,
  ...ribbonMenusKo,
  ...cellFormatKo,
  ...gridPopupsKo,
  ...editDialogsKo,
  ...formulaDialogsKo,
  ...dataObjectsKo,
  ...macKo,
  ...pageLayoutKo,
  ...workbookChromeKo,
  ...localeSwitcherKo,
};
