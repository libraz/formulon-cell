import { parseLocalizedNumericInput } from '../../commands/coerce-input.js';
import type { Range } from '../../engine/types.js';
import { clampZoom, ZOOM_MAX, ZOOM_MIN } from '../../store/render-scale.js';
import { showNumberPrompt } from './prompt.js';
import {
  appendDialogActions,
  appendErrorRow,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
  showInputError,
} from './shell.js';

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

/** Zoom that fits the selection, as Zoom to Selection applies it. */
export const zoomForSelection = (range: Range): number => {
  const rows = Math.max(1, range.r1 - range.r0 + 1);
  const cols = Math.max(1, range.c1 - range.c0 + 1);
  return clampZoom(Math.min(20 / rows, 12 / cols));
};

/** Preset magnifications of the Mac Zoom dialog, top to bottom. */
const MAC_ZOOM_PRESETS = [200, 150, 125, 100, 75, 50, 25] as const;

export interface MacZoomDialogOptions {
  title: string;
  /** Caption above the choices. */
  caption: string;
  fitSelection: string;
  custom: string;
  /** Current zoom in percent; a non-preset value starts on Custom. */
  initial: number;
  okLabel: string;
  cancelLabel: string;
  invalidMessage: string;
  locale?: string | (() => string);
}

export type MacZoomChoice =
  | { readonly kind: 'percent'; readonly percent: number }
  | { readonly kind: 'selection' };

/** The Mac Zoom dialog: preset radios, Zoom to Selection and a custom
 *  percentage field. Resolves `null` on cancel. */
export const showMacZoomDialog = (opts: MacZoomDialogOptions): Promise<MacZoomChoice | null> =>
  new Promise<MacZoomChoice | null>((resolve) => {
    const shell = createDialogShell({ title: opts.title });
    shell.panel.classList.add('fc-tb__dlg--mac-zoom');
    const doc = shell.body.ownerDocument;
    const getLocale = (): string =>
      typeof opts.locale === 'function' ? opts.locale() : (opts.locale ?? 'en-US');
    const caption = doc.createElement('div');
    caption.className = 'fc-tb__dlg__label';
    caption.textContent = opts.caption;
    shell.body.appendChild(caption);
    const group = doc.createElement('div');
    group.setAttribute('role', 'radiogroup');
    group.setAttribute('aria-label', opts.caption);
    shell.body.appendChild(group);
    const radio = (value: string, text: string): HTMLInputElement => {
      const row = doc.createElement('div');
      row.className = 'fc-fmtdlg__row';
      const label = doc.createElement('label');
      const input = doc.createElement('input');
      input.type = 'radio';
      input.name = 'fc-mac-zoom';
      input.value = value;
      label.append(input, doc.createTextNode(text));
      row.appendChild(label);
      group.appendChild(row);
      return input;
    };
    const presets = MAC_ZOOM_PRESETS.map((percent) => radio(String(percent), `${percent}%`));
    const selection = radio('selection', opts.fitSelection);
    const custom = radio('custom', opts.custom);
    const customInput = doc.createElement('input');
    customInput.type = 'number';
    customInput.className = 'fc-tb__dlg__input';
    customInput.min = String(ZOOM_MIN * 100);
    customInput.max = String(ZOOM_MAX * 100);
    customInput.step = '1';
    customInput.value = String(opts.initial);
    customInput.setAttribute('aria-label', opts.custom);
    const percentSign = doc.createElement('span');
    percentSign.textContent = '%';
    custom.closest('.fc-fmtdlg__row')?.append(customInput, percentSign);
    const preset = MAC_ZOOM_PRESETS.indexOf(opts.initial as (typeof MAC_ZOOM_PRESETS)[number]);
    ((preset >= 0 ? presets[preset] : undefined) ?? custom).checked = true;
    customInput.addEventListener('focus', () => {
      custom.checked = true;
    });
    const errorRow = appendErrorRow(shell.body);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: opts.cancelLabel,
      okLabel: opts.okLabel,
    });

    const lifecycle = installDialogLifecycle<MacZoomChoice | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
    });
    const onOk = (): void => {
      if (selection.checked) {
        lifecycle.finish({ kind: 'selection' });
        return;
      }
      const checked = presets.findIndex((input) => input.checked);
      if (checked >= 0) {
        lifecycle.finish({ kind: 'percent', percent: MAC_ZOOM_PRESETS[checked] as number });
        return;
      }
      const n = parseLocalizedNumericInput(customInput.value, getLocale());
      if (n === null || !Number.isInteger(n) || n < ZOOM_MIN * 100 || n > ZOOM_MAX * 100) {
        showInputError(errorRow, customInput, opts.invalidMessage);
        return;
      }
      lifecycle.finish({ kind: 'percent', percent: n });
    };
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, () =>
      (group.querySelector<HTMLInputElement>('input:checked') ?? custom).focus(),
    );
  });
