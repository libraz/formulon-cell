import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { createInteractionButton } from './chip-button.js';

const ENHANCED = 'fcCustomSelectEnhanced';
const HANDLES = new WeakMap<HTMLSelectElement, CustomSelectHandle>();

interface CustomSelectHandle {
  sync(): void;
  dispose(): void;
}

function optionLabel(select: HTMLSelectElement): string {
  const selected = Array.from(select.options).find((option) => option.value === select.value);
  return selected?.textContent?.trim() || select.value || '';
}

function selectValue(select: HTMLSelectElement, value: string): void {
  if (select.value === value) return;
  select.value = value;
  select.dispatchEvent(new Event('change', { bubbles: true }));
}

function optionDisabled(opt: HTMLOptionElement): boolean {
  return opt.disabled || opt.closest('optgroup')?.disabled === true;
}

function optionDisabledReason(opt: HTMLOptionElement): string | null {
  return opt.dataset.disabledReason ?? opt.closest('optgroup')?.dataset.disabledReason ?? null;
}

function createCustomSelectButton(select: HTMLSelectElement): HTMLButtonElement {
  const button = createInteractionButton({
    className: 'fc-select__button',
    role: 'combobox',
    ariaLabel: select.getAttribute('aria-label') ?? '',
  });
  // WebKit skips implicit button tab stops when full keyboard access is off.
  button.tabIndex = 0;
  button.setAttribute('aria-haspopup', 'listbox');
  button.setAttribute('aria-expanded', 'false');

  const value = document.createElement('span');
  value.className = 'fc-select__value';
  const arrow = document.createElement('span');
  arrow.className = 'fc-select__arrow';
  arrow.setAttribute('aria-hidden', 'true');
  button.append(value, arrow);
  return button;
}

function createCustomSelectOptionRow(opt: HTMLOptionElement): HTMLButtonElement {
  const row = createInteractionButton({
    className: 'fc-select__option',
    role: 'option',
    dataset: { value: opt.value },
    text: opt.textContent ?? '',
  });
  projectDisabledState(row, optionDisabled(opt), optionDisabledReason(opt), {
    datasetKey: 'disabledReason',
  });
  return row;
}

export function enhanceCustomSelect(select: HTMLSelectElement): CustomSelectHandle | null {
  if ((select.dataset as DOMStringMap)[ENHANCED] === '1') return null;
  (select.dataset as DOMStringMap)[ENHANCED] = '1';

  const wrap = document.createElement('span');
  wrap.className = 'fc-select';

  const button = createCustomSelectButton(select);
  const value = button.querySelector<HTMLElement>('.fc-select__value');
  if (!value) return null;

  const list = document.createElement('div');
  list.className = 'fc-select__list';
  list.setAttribute('role', 'listbox');
  list.hidden = true;

  select.classList.add('fc-select__native');
  select.tabIndex = -1;
  select.setAttribute('aria-hidden', 'true');
  select.after(wrap);
  wrap.append(select, button, list);

  let activeIndex = Math.max(0, select.selectedIndex);
  let disposed = false;
  let pointerDownInside = false;

  const options = (): HTMLOptionElement[] => Array.from(select.options);
  const rows = (): HTMLElement[] =>
    Array.from(list.querySelectorAll<HTMLElement>('[role="option"]'));

  const syncSelected = (): void => {
    value.textContent = optionLabel(select);
    projectDisabledState(button, select.disabled, select.dataset.disabledReason ?? null, {
      datasetKey: 'disabledReason',
    });
    const opts = options();
    const selectedOption = opts[select.selectedIndex];
    activeIndex =
      selectedOption && !optionDisabled(selectedOption)
        ? select.selectedIndex
        : opts.findIndex((opt) => !optionDisabled(opt));
    for (const [i, row] of rows().entries()) {
      const selected = opts[i]?.value === select.value;
      row.setAttribute('aria-selected', selected ? 'true' : 'false');
      row.classList.toggle('fc-select__option--active', i === activeIndex);
    }
  };

  const close = (): void => {
    pointerDownInside = false;
    list.hidden = true;
    button.setAttribute('aria-expanded', 'false');
  };

  const move = (delta: number): void => {
    const opts = options();
    const count = opts.length;
    if (count === 0) {
      activeIndex = -1;
      renderActive();
      return;
    }

    const start = activeIndex < 0 ? (delta > 0 ? -1 : 0) : activeIndex;
    activeIndex = -1;
    for (let distance = 1; distance <= count; distance += 1) {
      const candidate = (start + delta * distance + count) % count;
      const candidateOption = opts[candidate];
      if (candidateOption && !optionDisabled(candidateOption)) {
        activeIndex = candidate;
        break;
      }
    }
    renderActive(true);
  };

  const setActive = (index: number): void => {
    const opt = options()[index];
    activeIndex = opt && !optionDisabled(opt) ? index : -1;
    renderActive(true);
  };

  const renderActive = (scroll = false): void => {
    for (const [i, row] of rows().entries()) {
      const active = i === activeIndex;
      row.classList.toggle('fc-select__option--active', active);
      if (scroll && active && typeof row.scrollIntoView === 'function') {
        row.scrollIntoView({ block: 'nearest' });
      }
    }
  };

  const open = (): void => {
    if (select.disabled) return;
    renderOptions();
    list.hidden = false;
    button.setAttribute('aria-expanded', 'true');
    syncSelected();
  };

  function renderOptions(): void {
    list.replaceChildren();
    for (const [i, opt] of options().entries()) {
      const row = createCustomSelectOptionRow(opt);
      row.addEventListener('click', () => {
        if (select.disabled || optionDisabled(opt)) return;
        selectValue(select, opt.value);
        activeIndex = i;
        syncSelected();
        close();
        button.focus({ preventScroll: true });
      });
      list.appendChild(row);
    }
  }

  const onButtonClick = (): void => {
    if (select.disabled) return;
    if (list.hidden) open();
    else close();
  };

  const onButtonKey = (e: KeyboardEvent): void => {
    if (select.disabled) return;
    if (e.key === 'Tab') {
      if (!list.hidden) close();
      return;
    }
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      if (list.hidden) open();
      move(e.key === 'ArrowDown' ? 1 : -1);
      return;
    }
    if (e.key === 'Home' || e.key === 'End') {
      e.preventDefault();
      if (list.hidden) open();
      const opts = options();
      const candidates = opts.map((opt, index) => ({ opt, index }));
      const target =
        e.key === 'Home'
          ? candidates.find(({ opt }) => !optionDisabled(opt))
          : [...candidates].reverse().find(({ opt }) => !optionDisabled(opt));
      setActive(target?.index ?? -1);
      return;
    }
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      if (list.hidden) {
        open();
        return;
      }
      const opt = options()[activeIndex];
      if (opt && !optionDisabled(opt)) selectValue(select, opt.value);
      syncSelected();
      close();
      return;
    }
    if (e.key === 'Escape') {
      if (list.hidden) return;
      e.preventDefault();
      syncSelected();
      close();
    }
  };

  const onFocusOut = (e: FocusEvent): void => {
    if (pointerDownInside) return;
    if (!(e.relatedTarget instanceof Node) || !wrap.contains(e.relatedTarget)) close();
  };

  const onWrapPointerDown = (): void => {
    pointerDownInside = true;
  };

  const onDocumentPointerEnd = (e: Event): void => {
    pointerDownInside = false;
    const pointerTargetInside = e.target instanceof Node && wrap.contains(e.target);
    const focusOutside = !wrap.contains(document.activeElement);
    if (!pointerTargetInside || (e.type === 'pointercancel' && focusOutside)) close();
  };

  const onWindowBlur = (): void => {
    pointerDownInside = false;
    if (!wrap.contains(document.activeElement)) close();
  };

  const onDocumentPointer = (e: Event): void => {
    if (!wrap.contains(e.target as Node)) {
      pointerDownInside = false;
      close();
    }
  };

  const observer = new MutationObserver(() => {
    renderOptions();
    syncSelected();
    if (select.disabled) close();
  });
  observer.observe(select, { childList: true, subtree: true, attributes: true });

  button.addEventListener('click', onButtonClick);
  button.addEventListener('keydown', onButtonKey);
  wrap.addEventListener('focusout', onFocusOut);
  wrap.addEventListener('pointerdown', onWrapPointerDown);
  select.addEventListener('change', syncSelected);
  document.addEventListener('pointerdown', onDocumentPointer);
  document.addEventListener('pointerup', onDocumentPointerEnd);
  document.addEventListener('pointercancel', onDocumentPointerEnd);
  select.ownerDocument.defaultView?.addEventListener('blur', onWindowBlur);

  renderOptions();
  syncSelected();

  const handle: CustomSelectHandle = {
    sync: syncSelected,
    dispose() {
      if (disposed) return;
      disposed = true;
      pointerDownInside = false;
      observer.disconnect();
      button.removeEventListener('click', onButtonClick);
      button.removeEventListener('keydown', onButtonKey);
      wrap.removeEventListener('focusout', onFocusOut);
      wrap.removeEventListener('pointerdown', onWrapPointerDown);
      select.removeEventListener('change', syncSelected);
      document.removeEventListener('pointerdown', onDocumentPointer);
      document.removeEventListener('pointerup', onDocumentPointerEnd);
      document.removeEventListener('pointercancel', onDocumentPointerEnd);
      select.ownerDocument.defaultView?.removeEventListener('blur', onWindowBlur);
      select.classList.remove('fc-select__native');
      select.removeAttribute('aria-hidden');
      select.tabIndex = 0;
      wrap.replaceWith(select);
      delete (select.dataset as DOMStringMap)[ENHANCED];
      HANDLES.delete(select);
    },
  };
  HANDLES.set(select, handle);
  return handle;
}

export function syncCustomSelects(root: HTMLElement): void {
  root
    .querySelectorAll<HTMLSelectElement>('select[data-fc-custom-select-enhanced="1"]')
    .forEach((select) => {
      HANDLES.get(select)?.sync();
    });
}
