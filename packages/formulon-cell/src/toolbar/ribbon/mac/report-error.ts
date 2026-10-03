import type { SpreadsheetInstance } from '../../../mount/types.js';
import { reportDialogLabels, showReport } from '../../dialogs/report.js';

/** Surface a failed Mac ribbon operation in the shared report dialog. */
export const reportMacRibbonError = (instance: SpreadsheetInstance, error: unknown): void => {
  const strings = instance.i18n.strings;
  void showReport({
    title: strings.ribbon.ribbon,
    items: [
      {
        severity: 'warning',
        label: error instanceof Error ? error.message : String(error),
        detail: '',
      },
    ],
    ...reportDialogLabels(strings),
  });
};
