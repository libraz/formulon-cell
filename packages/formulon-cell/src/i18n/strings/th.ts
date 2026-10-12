import { cellFormatTh } from './cell-format/th.js';
import { dataObjectsTh } from './data-objects/th.js';
import { editDialogsTh } from './edit-dialogs/th.js';
import { formulaDialogsTh } from './formula-dialogs/th.js';
import { gridPopupsTh } from './grid-popups/th.js';
import { localeSwitcherTh } from './locale-switcher/th.js';
import { macTh } from './mac/th.js';
import { pageLayoutTh } from './page-layout/th.js';
import { ribbonTh } from './ribbon/th.js';
import { ribbonMenusTh } from './ribbon-menus/th.js';
import type { Strings } from './types.js';
import { workbookChromeTh } from './workbook-chrome/th.js';

export const th: Strings = {
  ...ribbonTh,
  ...ribbonMenusTh,
  ...cellFormatTh,
  ...gridPopupsTh,
  ...editDialogsTh,
  ...formulaDialogsTh,
  ...dataObjectsTh,
  ...macTh,
  ...pageLayoutTh,
  ...workbookChromeTh,
  ...localeSwitcherTh,
};
