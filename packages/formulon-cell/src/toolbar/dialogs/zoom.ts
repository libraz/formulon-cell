import { showNumberPrompt } from './prompt.js';

export interface ZoomDialogOptions {
  title: string;
  label: string;
  initial: number;
  okLabel: string;
  cancelLabel: string;
  invalidMessage: string;
}

export const showZoomDialog = (opts: ZoomDialogOptions): Promise<number | null> =>
  showNumberPrompt({
    title: opts.title,
    label: opts.label,
    initial: opts.initial,
    min: 50,
    max: 400,
    step: 1,
    okLabel: opts.okLabel,
    cancelLabel: opts.cancelLabel,
    invalidMessage: opts.invalidMessage,
  });
