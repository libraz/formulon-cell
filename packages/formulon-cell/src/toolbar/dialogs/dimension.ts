import { showNumberPrompt } from './prompt.js';

export interface DimensionDialogOptions {
  title: string;
  label: string;
  /** Undefined leaves the field empty, as for a selection of mixed sizes. */
  initial?: number;
  max: number;
  okLabel: string;
  cancelLabel: string;
  locale?: string | (() => string);
}

export const showDimensionDialog = (opts: DimensionDialogOptions): Promise<number | null> =>
  showNumberPrompt({
    title: opts.title,
    label: opts.label,
    initial: opts.initial,
    min: 0,
    max: opts.max,
    step: 1,
    okLabel: opts.okLabel,
    cancelLabel: opts.cancelLabel,
    invalidMessage: opts.label,
    locale: opts.locale,
  });
