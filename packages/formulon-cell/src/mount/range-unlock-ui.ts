import type { AllowedRangeUnlockUi } from '../commands/protection.js';
import type { Strings } from '../i18n/strings.js';
import { showMessage } from '../toolbar/dialogs/prompt.js';
import { showUnprotectSheetDialog } from '../toolbar/dialogs/protection.js';

/** Password prompts for password-guarded allowed edit ranges. */
export const createAllowedRangeUnlockUi = (getStrings: () => Strings): AllowedRangeUnlockUi => {
  const titleFor = (strings: Strings, range: string): string =>
    strings.ribbonMenu.allowedRangeUnlockTitle.replace('{title}', range);
  return {
    askPassword: (range) => {
      const strings = getStrings();
      return showUnprotectSheetDialog({
        title: titleFor(strings, range),
        password: strings.protection.password,
        ok: strings.pageSetup.ok,
        cancel: strings.pageSetup.cancel,
      });
    },
    reportIncorrect: async (range) => {
      const strings = getStrings();
      await showMessage({
        title: titleFor(strings, range),
        message: strings.ribbonMenu.workbookIncorrectPassword,
        okLabel: strings.pageSetup.ok,
      });
    },
  };
};
