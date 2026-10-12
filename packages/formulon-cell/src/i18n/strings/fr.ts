import { cellFormatFr } from './cell-format/fr.js';
import { dataObjectsFr } from './data-objects/fr.js';
import { editDialogsFr } from './edit-dialogs/fr.js';
import { formulaDialogsFr } from './formula-dialogs/fr.js';
import { gridPopupsFr } from './grid-popups/fr.js';
import { localeSwitcherFr } from './locale-switcher/fr.js';
import { macFr } from './mac/fr.js';
import { pageLayoutFr } from './page-layout/fr.js';
import { ribbonFr } from './ribbon/fr.js';
import { ribbonMenusFr } from './ribbon-menus/fr.js';
import type { Strings } from './types.js';
import { workbookChromeFr } from './workbook-chrome/fr.js';

export const fr: Strings = {
  ...ribbonFr,
  ...ribbonMenusFr,
  ...cellFormatFr,
  ...gridPopupsFr,
  ...editDialogsFr,
  ...formulaDialogsFr,
  ...dataObjectsFr,
  ...macFr,
  ...pageLayoutFr,
  ...workbookChromeFr,
  ...localeSwitcherFr,
};
