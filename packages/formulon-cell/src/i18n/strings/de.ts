import { cellFormatDe } from './cell-format/de.js';
import { dataObjectsDe } from './data-objects/de.js';
import { editDialogsDe } from './edit-dialogs/de.js';
import { formulaDialogsDe } from './formula-dialogs/de.js';
import { gridPopupsDe } from './grid-popups/de.js';
import { localeSwitcherDe } from './locale-switcher/de.js';
import { macDe } from './mac/de.js';
import { pageLayoutDe } from './page-layout/de.js';
import { ribbonDe } from './ribbon/de.js';
import { ribbonMenusDe } from './ribbon-menus/de.js';
import type { Strings } from './types.js';
import { workbookChromeDe } from './workbook-chrome/de.js';

export const de: Strings = {
  ...ribbonDe,
  ...ribbonMenusDe,
  ...cellFormatDe,
  ...gridPopupsDe,
  ...editDialogsDe,
  ...formulaDialogsDe,
  ...dataObjectsDe,
  ...macDe,
  ...pageLayoutDe,
  ...workbookChromeDe,
  ...localeSwitcherDe,
};
