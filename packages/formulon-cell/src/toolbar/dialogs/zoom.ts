import { ZOOM_MAX, ZOOM_MIN } from '../../store/render-scale.js';
import { showNumberPrompt } from './prompt.js';

export interface ZoomDialogOptions {
  title: string;
  label: string;
  initial: number;
  okLabel: string;
  cancelLabel: string;
  invalidMessage: string;
  locale?: string | (() => string);
}

export const showZoomDialog = (opts: ZoomDialogOptions): Promise<number | null> =>
  showNumberPrompt({
    title: opts.title,
    label: opts.label,
    initial: opts.initial,
    min: ZOOM_MIN * 100,
    max: ZOOM_MAX * 100,
    step: 1,
    integer: true,
    okLabel: opts.okLabel,
    cancelLabel: opts.cancelLabel,
    invalidMessage: opts.invalidMessage,
    locale: opts.locale,
  });
