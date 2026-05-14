import { applyFilter, clearFilter, distinctValues } from '../commands/filter.js';
import { addrKey } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';
import { inheritHostTokens } from './inherit-host-tokens.js';

export interface FilterDropdownDeps {
  store: SpreadsheetStore;
  strings?: Strings;
  /** Where to anchor the popover. When omitted the popover is centred. */
  anchorRect?: { x: number; y: number; w: number; h: number };
}

export interface FilterDropdownHandle {
  /** Open against the active selection's column. The active selection acts
   *  as the data range; first row is treated as header. */
  open(range: Range, col: number, anchor: { x: number; y: number; h: number }): void;
  close(): void;
  isOpen(): boolean;
  detach(): void;
}

/**
 * Lightweight column-filter popover. Lists distinct values in the column with
 * a checkbox each; "Apply" calls `applyFilter`, "Clear" calls `clearFilter`.
 *
 * The popover lives in `document.body` so it can escape any clipping ancestors,
 * and is positioned via fixed coordinates from the anchor rect.
 */
export function attachFilterDropdown(deps: FilterDropdownDeps): FilterDropdownHandle {
  const strings = deps.strings ?? defaultStrings;
  const t = strings.filterDropdown;
  let root: HTMLDivElement | null = null;
  let activeRange: Range | null = null;
  let activeCol = 0;
  let activeHidden = new Set<string>();

  const close = (): void => {
    if (!root) return;
    root.remove();
    root = null;
    activeRange = null;
    activeHidden = new Set();
    document.removeEventListener('mousedown', onDocMouseDown, true);
    document.removeEventListener('keydown', onDocKey, true);
  };

  const onDocMouseDown = (e: MouseEvent): void => {
    if (!root) return;
    if (root.contains(e.target as Node)) return;
    close();
  };
  const onDocKey = (e: KeyboardEvent): void => {
    if (e.key === 'Escape') {
      e.preventDefault();
      close();
    } else if (e.key === 'Enter' && root?.contains(e.target as Node)) {
      e.preventDefault();
      applyActiveFilter();
    }
  };

  const applyActiveFilter = (): void => {
    if (!activeRange || !root) return;
    clearFilter(deps.store.getState(), deps.store, activeRange);
    const state2 = deps.store.getState();
    applyFilter(state2, deps.store, activeRange, activeCol, (cell) => {
      const key = cellToKey(cell?.value);
      return !activeHidden.has(key);
    });
    close();
  };

  const open = (range: Range, col: number, anchor: { x: number; y: number; h: number }): void => {
    close();
    activeRange = range;
    activeCol = col;

    const state = deps.store.getState();
    const distinct = distinctValues(state, range, col);
    const hidden = new Set<string>();
    activeHidden = hidden;
    // Pre-mark currently-hidden values as unchecked so the dropdown reflects
    //  the live filter state.
    for (let r = range.r0 + 1; r <= range.r1; r += 1) {
      if (!state.layout.hiddenRows.has(r)) continue;
      const cell = state.data.cells.get(addrKey({ sheet: range.sheet, row: r, col }));
      const key = cellToKey(cell?.value);
      hidden.add(key);
    }

    const r = document.createElement('div');
    r.className = 'fc-filter-dropdown';
    r.style.position = 'fixed';
    r.style.left = `${anchor.x}px`;
    r.style.top = `${anchor.y + anchor.h}px`;
    r.setAttribute('role', 'dialog');
    r.setAttribute('aria-label', t.title);

    const search = document.createElement('input');
    search.className = 'fc-filter-dropdown__search';
    search.type = 'search';
    search.placeholder = t.searchPlaceholder;
    search.spellcheck = false;

    const list = document.createElement('div');
    list.className = 'fc-filter-dropdown__list';
    list.tabIndex = -1;

    const renderRows = (filter: string): void => {
      list.innerHTML = '';
      const f = filter.toLowerCase();
      // (Select All) header
      const allRow = document.createElement('label');
      allRow.className = 'fc-filter-dropdown__row fc-filter-dropdown__row--all';
      const allCb = document.createElement('input');
      allCb.type = 'checkbox';
      allCb.checked = distinct.every((v) => !hidden.has(v));
      allCb.indeterminate = !allCb.checked && distinct.some((v) => !hidden.has(v));
      allCb.addEventListener('change', () => {
        if (allCb.checked) {
          hidden.clear();
        } else {
          for (const v of distinct) hidden.add(v);
        }
        renderRows(search.value);
      });
      const allLabel = document.createElement('span');
      allLabel.textContent = t.selectAll;
      allRow.append(allCb, allLabel);
      list.appendChild(allRow);

      for (const v of distinct) {
        const display = v === '' ? t.blanks : v;
        if (f && !display.toLowerCase().includes(f)) continue;
        const row = document.createElement('label');
        row.className = 'fc-filter-dropdown__row';
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.value = v;
        cb.checked = !hidden.has(v);
        cb.addEventListener('change', () => {
          if (cb.checked) hidden.delete(v);
          else hidden.add(v);
          // Refresh select-all indeterminate state without rebuilding rows.
          allCb.checked = distinct.every((vv) => !hidden.has(vv));
          allCb.indeterminate = !allCb.checked && distinct.some((vv) => !hidden.has(vv));
        });
        const text = document.createElement('span');
        text.textContent = display;
        row.append(cb, text);
        list.appendChild(row);
      }
    };
    renderRows('');
    search.addEventListener('input', () => renderRows(search.value));

    const actions = document.createElement('div');
    actions.className = 'fc-filter-dropdown__actions';
    const apply = document.createElement('button');
    apply.type = 'button';
    apply.className = 'fc-filter-dropdown__apply';
    apply.textContent = t.apply;
    apply.addEventListener('click', () => applyActiveFilter());
    const clear = document.createElement('button');
    clear.type = 'button';
    clear.className = 'fc-filter-dropdown__clear';
    clear.textContent = t.clear;
    clear.addEventListener('click', () => {
      if (!activeRange) return;
      clearFilter(deps.store.getState(), deps.store, activeRange);
      close();
    });
    actions.append(clear, apply);

    r.append(search, list, actions);
    // Borrow theme tokens from the first .fc-host on the page (filter is
    // body-attached, so `[data-fc-theme]` doesn't cascade automatically).
    const host = document.querySelector('.fc-host');
    if (host) inheritHostTokens(host, r);
    document.body.appendChild(r);
    root = r;
    const rect = r.getBoundingClientRect();
    const left = Math.max(4, Math.min(anchor.x, window.innerWidth - rect.width - 4));
    const top = Math.max(4, Math.min(anchor.y + anchor.h, window.innerHeight - rect.height - 4));
    r.style.left = `${left}px`;
    r.style.top = `${top}px`;

    requestAnimationFrame(() => search.focus());

    document.addEventListener('mousedown', onDocMouseDown, true);
    document.addEventListener('keydown', onDocKey, true);
  };

  return {
    open,
    close,
    isOpen: () => root != null,
    detach() {
      close();
    },
  };
}

const cellToKey = (v: unknown): string => {
  if (!v || typeof v !== 'object') return '';
  const cv = v as { kind: string; value?: unknown };
  if (cv.kind === 'number') return String(cv.value);
  if (cv.kind === 'text') return String(cv.value ?? '');
  if (cv.kind === 'bool') return cv.value ? 'TRUE' : 'FALSE';
  return '';
};
