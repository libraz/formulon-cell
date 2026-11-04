import { createInteractionButton } from './chip-button.js';

export interface RangePickerControlOptions {
  label: string;
  getValue: () => string;
  onPicked?: (value: string) => void;
  subscribeToRangeChanges?: (listener: () => void) => () => void;
  kind?: string;
}

const INTERACTIVE_SELECTOR =
  'button, input, select, textarea, a[href], [role="button"], [role="menuitem"], [role="tab"]';

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
  // Armed: the focused field receives grid picks, but the dialog stays expanded.
  let collapsed = false;
  let lastSeen = '';
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
  // A grid pick ends on pointerup; hand focus back so Enter/Escape/typing reach the field.
  const onDocumentPointerUp = (event: PointerEvent): void => {
    if (!unsubscribe) return;
    const target = event.target;
    if (target instanceof Node && dialog()?.contains(target)) return;
    // Real controls (ribbon, tabs, inputs) keep their own focus.
    if (target instanceof Element && target.closest(INTERACTIVE_SELECTOR)) return;
    // Let the grid's own pointerup handling (host.focus) run first.
    queueMicrotask(() => {
      if (!unsubscribe || shouldStopForDialogState()) return;
      if (!collapsed && document.activeElement === input) return;
      if (!collapsed) applyPickedValueIfChanged();
      input.focus({ preventScroll: true });
    });
  };
  const applyPickedValueIfChanged = (): void => {
    const picked = opts.getValue();
    if (picked === lastSeen && input.value === picked) return;
    lastSeen = picked;
    applyPickedValue();
  };
  const stopPicking = (): void => {
    document.removeEventListener('pointerup', onDocumentPointerUp, true);
    if (unsubscribe) {
      unsubscribe();
      unsubscribe = null;
    }
    document.removeEventListener('keydown', onDocumentKeydown, true);
    activeDialog?.removeEventListener('fc-range-picker-stop-all', stopPicking);
    activeDialog = null;
    observer?.disconnect();
    observer = null;
    collapsed = false;
    wrap.classList.remove('fc-range-picker--picking', 'fc-range-picker--armed');
    button.dataset.rangePickerActive = 'false';
    button.setAttribute('aria-pressed', 'false');
    syncDialogClasses();
  };
  const syncDialogClasses = (): void => {
    const dlg = dialog();
    if (!dlg) return;
    dlg.classList.toggle(
      'fc-fmtdlg--range-picking',
      !!dlg.querySelector('.fc-range-picker__btn[data-range-picker-active="true"]'),
    );
    dlg.classList.toggle('fc-fmtdlg--range-armed', !!dlg.querySelector('.fc-range-picker--armed'));
  };
  const startPicking = (collapse: boolean): void => {
    if (!opts.subscribeToRangeChanges || unsubscribe) return;
    dialog()
      ?.querySelectorAll<HTMLElement>('.fc-range-picker--picking, .fc-range-picker--armed')
      .forEach((picker) => {
        if (picker !== wrap) picker.dispatchEvent(new CustomEvent('fc-range-picker-stop'));
      });
    collapsed = collapse;
    lastSeen = opts.getValue();
    if (collapse) {
      wrap.classList.add('fc-range-picker--picking');
      button.dataset.rangePickerActive = 'true';
      button.setAttribute('aria-pressed', 'true');
      document.addEventListener('keydown', onDocumentKeydown, true);
    } else {
      wrap.classList.add('fc-range-picker--armed');
    }
    syncDialogClasses();
    document.addEventListener('pointerup', onDocumentPointerUp, true);
    activeDialog = dialog();
    activeDialog?.addEventListener('fc-range-picker-stop-all', stopPicking);
    unsubscribe = opts.subscribeToRangeChanges(() => {
      if (shouldStopForDialogState()) {
        stopPicking();
        return;
      }
      // Only a changed selection feeds an armed field; unrelated store writes must not.
      if (!collapsed) {
        const picked = opts.getValue();
        if (picked === lastSeen) return;
        lastSeen = picked;
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
      const wasCollapsed = collapsed;
      if (unsubscribe) stopPicking();
      if (!wasCollapsed) startPicking(true);
    }
    input.focus({ preventScroll: true });
    input.select();
  });
  input.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') {
      // An inactive picker belongs to the dialog's normal keyboard route.
      // Let its Escape handler dismiss the dialog; only an active range
      // selection owns Escape and needs to cancel itself first.
      if (!unsubscribe || !collapsed) return;
      event.preventDefault();
      event.stopPropagation();
      stopPicking();
    }
  });
  input.addEventListener('focus', () => {
    if (input.disabled || input.readOnly) return;
    startPicking(false);
  });
  input.addEventListener('blur', (event) => {
    // Focus moving to the grid keeps the field armed; another dialog control ends it.
    const next = event.relatedTarget;
    if (collapsed || !(next instanceof Node) || !dialog()?.contains(next)) return;
    stopPicking();
  });
  wrap.addEventListener('fc-range-picker-stop', stopPicking);
  wrap.appendChild(button);
  return button;
};

export const updateRangePickerLabel = (button: HTMLButtonElement, label: string): void => {
  button.title = label;
  button.setAttribute('aria-label', label);
};
