// Contract between the Page Setup dialog (`page-setup-dialog.ts`) and its
// per-tab builders. A tab builds its controls into its panel, registers
// listeners through `on` (so the dialog shell removes them on detach), and
// maps `PageSetup` to and from those controls.

import type { Strings } from '../../i18n/strings.js';
import type { SpreadsheetStore } from '../../store/store.js';
import type { DialogShell } from '../dialog-shell.js';

export interface PageSetupTabContext {
  readonly t: Strings['pageSetup'];
  readonly strings: Strings;
  readonly store: SpreadsheetStore;
  /** Shell-tracked listener registration. */
  readonly on: DialogShell['on'];
}
