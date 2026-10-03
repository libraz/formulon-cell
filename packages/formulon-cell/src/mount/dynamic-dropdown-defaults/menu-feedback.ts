import type { SpreadsheetInstance } from '../../index.js';
import { reportDialogLabels, showReport } from '../../toolbar/dialogs/report.js';
import { projectDisabledState } from '../../toolbar/menu-a11y.js';

export const setMenuControlDisabled = (
  button: HTMLButtonElement,
  disabled: boolean,
  reason?: string,
): void => {
  const baseTitle = button.dataset.menuBaseTitle ?? button.title;
  button.dataset.menuBaseTitle = baseTitle;
  projectDisabledState(button, disabled, reason ?? null, {
    datasetKey: 'menuDisabledReason',
    titlePrefix: baseTitle,
  });
};

export const showInstanceReport = async (
  instance: SpreadsheetInstance,
  title: string,
  items: { severity: 'info' | 'warning'; label: string; detail: string }[],
): Promise<void> => {
  const strings = instance.i18n.strings;
  await showReport({
    title,
    items,
    ...reportDialogLabels(strings),
  });
};
