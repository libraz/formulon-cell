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

export function enhanceCustomSelect(select: HTMLSelectElement): CustomSelectHandle | null {
  if ((select.dataset as DOMStringMap)[ENHANCED] === '1') return null;
  (select.dataset as DOMStringMap)[ENHANCED] = '1';

  const wrap = document.createElement('span');
  wrap.className = 'fc-select';

  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'fc-select__button';
  button.setAttribute('aria-haspopup', 'listbox');
  button.setAttribute('aria-expanded', 'false');
  button.setAttribute('role', 'combobox');
  button.setAttribute('aria-label', select.getAttribute('aria-label') ?? '');

  const value = document.createElement('span');
  value.className = 'fc-select__value';
  const arrow = document.createElement('span');
  arrow.className = 'fc-select__arrow';
  arrow.setAttribute('aria-hidden', 'true');
  button.append(value, arrow);

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

  const options = (): HTMLOptionElement[] => Array.from(select.options);
  const rows = (): HTMLElement[] =>
    Array.from(list.querySelectorAll<HTMLElement>('[role="option"]'));

  const syncSelected = (): void => {
    value.textContent = optionLabel(select);
    const opts = options();
    activeIndex = Math.max(0, select.selectedIndex);
    for (const [i, row] of rows().entries()) {
      const selected = opts[i]?.value === select.value;
      row.setAttribute('aria-selected', selected ? 'true' : 'false');
      row.classList.toggle('fc-select__option--active', i === activeIndex);
    }
  };

  const close = (): void => {
    list.hidden = true;
    button.setAttribute('aria-expanded', 'false');
  };

  const move = (delta: number): void => {
    const count = options().length;
    if (count === 0) return;
    activeIndex = (activeIndex + delta + count) % count;
    for (const [i, row] of rows().entries()) {
      row.classList.toggle('fc-select__option--active', i === activeIndex);
      if (i === activeIndex) row.scrollIntoView({ block: 'nearest' });
    }
  };

  const open = (): void => {
    renderOptions();
    list.hidden = false;
    button.setAttribute('aria-expanded', 'true');
    syncSelected();
  };

  function renderOptions(): void {
    list.replaceChildren();
    for (const [i, opt] of options().entries()) {
      const row = document.createElement('button');
      row.type = 'button';
      row.className = 'fc-select__option';
      row.setAttribute('role', 'option');
      row.dataset.value = opt.value;
      row.textContent = opt.textContent;
      row.disabled = opt.disabled;
      row.addEventListener('click', () => {
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
    if (list.hidden) open();
    else close();
  };

  const onButtonKey = (e: KeyboardEvent): void => {
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      if (list.hidden) open();
      move(e.key === 'ArrowDown' ? 1 : -1);
      return;
    }
    if (e.key === 'Home' || e.key === 'End') {
      e.preventDefault();
      if (list.hidden) open();
      activeIndex = e.key === 'Home' ? 0 : Math.max(0, options().length - 1);
      syncSelected();
      return;
    }
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      if (list.hidden) {
        open();
        return;
      }
      const opt = options()[activeIndex];
      if (opt && !opt.disabled) selectValue(select, opt.value);
      syncSelected();
      close();
      return;
    }
    if (e.key === 'Escape') {
      e.preventDefault();
      close();
    }
  };

  const onDocumentPointer = (e: Event): void => {
    if (!wrap.contains(e.target as Node)) close();
  };

  const observer = new MutationObserver(() => {
    renderOptions();
    syncSelected();
  });
  observer.observe(select, { childList: true, subtree: true, attributes: true });

  button.addEventListener('click', onButtonClick);
  button.addEventListener('keydown', onButtonKey);
  select.addEventListener('change', syncSelected);
  document.addEventListener('pointerdown', onDocumentPointer);

  renderOptions();
  syncSelected();

  const handle: CustomSelectHandle = {
    sync: syncSelected,
    dispose() {
      if (disposed) return;
      disposed = true;
      observer.disconnect();
      button.removeEventListener('click', onButtonClick);
      button.removeEventListener('keydown', onButtonKey);
      select.removeEventListener('change', syncSelected);
      document.removeEventListener('pointerdown', onDocumentPointer);
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
