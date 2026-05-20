import { showNumberPrompt } from './prompt.js';

export interface PageScaleDialogOptions {
  title: string;
  label: string;
  initial: number;
  kind: 'scale' | 'pages';
  okLabel: string;
  cancelLabel: string;
  invalidMessage: string;
}

export const showPageScaleDialog = (opts: PageScaleDialogOptions): Promise<number | null> =>
  showNumberPrompt({
    title: opts.title,
    label: opts.label,
    initial: opts.initial,
    min: opts.kind === 'scale' ? 10 : 1,
    max: opts.kind === 'scale' ? 400 : 99,
    step: 1,
    okLabel: opts.okLabel,
    cancelLabel: opts.cancelLabel,
    invalidMessage: opts.invalidMessage,
  });
