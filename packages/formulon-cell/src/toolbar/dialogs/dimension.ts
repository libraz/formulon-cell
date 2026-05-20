import { showNumberPrompt } from './prompt.js';

export interface DimensionDialogOptions {
  title: string;
  label: string;
  initial: number;
  max: number;
  okLabel: string;
  cancelLabel: string;
}

export const showDimensionDialog = (opts: DimensionDialogOptions): Promise<number | null> =>
  showNumberPrompt({
    title: opts.title,
    label: opts.label,
    initial: opts.initial,
    min: 1,
    max: opts.max,
    step: 1,
    okLabel: opts.okLabel,
    cancelLabel: opts.cancelLabel,
    invalidMessage: opts.label,
  });
