import type { Strings } from '../../i18n/strings.js';
import { showNumberPrompt } from './prompt.js';

export interface DimensionDialogOptions {
  /** Row Height or Column Width. */
  axis: 'row' | 'col';
  strings: Strings;
  /** Undefined leaves the field empty, as for a selection of mixed sizes. */
  initial?: number;
  max: number;
  locale?: string | (() => string);
}

/** The Row Height / Column Width dialog: one number field, titled and laid out
 *  per axis (`fc-dimdlg--row` / `fc-dimdlg--col`). */
export const showDimensionDialog = (opts: DimensionDialogOptions): Promise<number | null> => {
  const t = opts.strings.dimensionDialog;
  const label = opts.axis === 'row' ? t.rowHeightLabel : t.colWidthLabel;
  return showNumberPrompt({
    title: opts.axis === 'row' ? t.rowHeightTitle : t.colWidthTitle,
    label,
    initial: opts.initial,
    min: 0,
    max: opts.max,
    step: 1,
    okLabel: opts.strings.hyperlinkDialog.ok,
    cancelLabel: opts.strings.hyperlinkDialog.cancel,
    invalidMessage: label,
    locale: opts.locale,
    panelClasses: ['fc-dimdlg', `fc-dimdlg--${opts.axis}`],
  });
};
