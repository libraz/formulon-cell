import { createInteractionButton } from './chip-button.js';

export interface RangePickerControlOptions {
  label: string;
  getValue: () => string;
  onPicked?: (value: string) => void;
  subscribeToRangeChanges?: (listener: () => void) => () => void;
  kind?: string;
}

export const attachRangePickerButton = (
  input: HTMLInputElement,
  opts: RangePickerControlOptions,
): HTMLButtonElement => {
  const wrap = document.createElement('span');
  wrap.className = 'fc-range-picker';
  input.replaceWith(wrap);
  wrap.appendChild(input);

  const button = createInteractionButton({
    className: 'fc-range-picker__btn',
    ariaLabel: opts.label,
    dataset: { rangePicker: opts.kind ?? 'range' },
    pressed: false,
  });
  button.title = opts.label;
  button.setAttribute('aria-pressed', 'false');

  let unsubscribe: (() => void) | null = null;
  let observer: MutationObserver | null = null;
  let activeDialog: HTMLElement | null = null;
  const dialog = (): HTMLElement | null => input.closest('.fc-fmtdlg');
  const shouldStopForDialogState = (): boolean => {
    const dlg = dialog();
    return !wrap.isConnected || dlg?.hidden === true || dlg?.getAttribute('aria-hidden') === 'true';
  };
  const onDocumentKeydown = (event: KeyboardEvent): void => {
    if (event.key !== 'Escape') return;
    event.preventDefault();
    event.stopPropagation();
    stopPicking();
    input.focus({ preventScroll: true });
    input.select();
  };
  const applyPickedValue = (): void => {
    const value = opts.getValue();
    input.value = value;
    input.dispatchEvent(new Event('input', { bubbles: true }));
    opts.onPicked?.(value);
  };
  const stopPicking = (): void => {
    if (unsubscribe) {
      unsubscribe();
      unsubscribe = null;
    }
    document.removeEventListener('keydown', onDocumentKeydown, true);
    activeDialog?.removeEventListener('fc-range-picker-stop-all', stopPicking);
    activeDialog = null;
    observer?.disconnect();
    observer = null;
    wrap.classList.remove('fc-range-picker--picking');
    button.dataset.rangePickerActive = 'false';
    button.setAttribute('aria-pressed', 'false');
    dialog()?.classList.toggle(
      'fc-fmtdlg--range-picking',
      !!dialog()?.querySelector('.fc-range-picker__btn[data-range-picker-active="true"]'),
    );
  };
  const startPicking = (): void => {
    if (!opts.subscribeToRangeChanges || unsubscribe) return;
    dialog()
      ?.querySelectorAll<HTMLElement>('.fc-range-picker--picking')
      .forEach((picker) => {
        if (picker !== wrap) picker.dispatchEvent(new CustomEvent('fc-range-picker-stop'));
      });
    wrap.classList.add('fc-range-picker--picking');
    dialog()?.classList.add('fc-fmtdlg--range-picking');
    button.dataset.rangePickerActive = 'true';
    button.setAttribute('aria-pressed', 'true');
    document.addEventListener('keydown', onDocumentKeydown, true);
    activeDialog = dialog();
    activeDialog?.addEventListener('fc-range-picker-stop-all', stopPicking);
    unsubscribe = opts.subscribeToRangeChanges(() => {
      if (shouldStopForDialogState()) {
        stopPicking();
        return;
      }
      applyPickedValue();
    });
    observer = new MutationObserver(() => {
      if (shouldStopForDialogState()) stopPicking();
    });
    if (document.body) observer.observe(document.body, { childList: true, subtree: true });
    const dlg = activeDialog ?? dialog();
    if (dlg)
      observer.observe(dlg, { attributes: true, attributeFilter: ['hidden', 'aria-hidden'] });
  };
  button.addEventListener('click', () => {
    applyPickedValue();
    if (opts.subscribeToRangeChanges) {
      if (unsubscribe) stopPicking();
      else startPicking();
    }
    input.focus({ preventScroll: true });
    input.select();
  });
  input.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      stopPicking();
    }
  });
  wrap.addEventListener('fc-range-picker-stop', stopPicking);
  wrap.appendChild(button);
  return button;
};

export const updateRangePickerLabel = (button: HTMLButtonElement, label: string): void => {
  button.title = label;
  button.setAttribute('aria-label', label);
};
