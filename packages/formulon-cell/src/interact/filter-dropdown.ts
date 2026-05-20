import {
  applyConditionFilter,
  applyValueFilter,
  clearFilter,
  type ConditionFilterOp,
  distinctValues,
  filterValueKey,
  recordFilterChange,
} from '../commands/filter.js';
import type { History } from '../commands/history.js';
import { addrKey } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';
import { createDialogSelect } from '../toolbar/dialogs/form-controls.js';
import { createInteractionButton } from './chip-button.js';
import { inheritHostTokens } from './inherit-host-tokens.js';
import { clampPanelToViewport } from './overlay-position.js';

export interface FilterDropdownDeps {
  store: SpreadsheetStore;
  history?: History | null;
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

const createFilterDropdownActionButton = (
  className: string,
  label: string,
): HTMLButtonElement => {
  return createInteractionButton({
    className,
    text: label,
  });
};

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
  let restoreFocus: HTMLElement | null = null;

  const close = (): void => {
    if (!root) return;
    root.remove();
    root = null;
    activeRange = null;
    activeHidden = new Set();
    document.removeEventListener('mousedown', onDocMouseDown, true);
    document.removeEventListener('keydown', onDocKey, true);
    restoreFocus?.focus({ preventScroll: true });
    restoreFocus = null;
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
    const condition = root.querySelector<HTMLSelectElement>('.fc-filter-dropdown__condition-op');
    const conditionValue = root.querySelector<HTMLInputElement>(
      '.fc-filter-dropdown__condition-value',
    );
    recordFilterChange(deps.history ?? null, deps.store, () => {
      if (condition?.value && conditionValue?.value.trim()) {
        applyConditionFilter(deps.store.getState(), deps.store, activeRange as Range, activeCol, {
          op: condition.value as ConditionFilterOp,
          value: conditionValue.value,
        });
      } else {
        applyValueFilter(
          deps.store.getState(),
          deps.store,
          activeRange as Range,
          activeCol,
          Array.from(activeHidden),
        );
      }
    });
    close();
  };

  const open = (range: Range, col: number, anchor: { x: number; y: number; h: number }): void => {
    close();
    restoreFocus =
      document.activeElement instanceof HTMLElement && document.activeElement !== document.body
        ? document.activeElement
        : null;
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
      const key = filterValueKey(cell?.value);
      hidden.add(key);
    }

    const r = document.createElement('div');
    r.className = 'fc-filter-dropdown';
    r.style.position = 'fixed';
    r.style.left = `${anchor.x}px`;
    r.style.top = `${anchor.y + anchor.h}px`;
    r.setAttribute('role', 'dialog');
    r.setAttribute('aria-modal', 'false');
    r.setAttribute('aria-label', t.title);

    const search = document.createElement('input');
    search.className = 'fc-filter-dropdown__search';
    search.type = 'search';
    search.placeholder = t.searchPlaceholder;
    search.setAttribute('aria-label', t.searchPlaceholder);
    search.spellcheck = false;

    const conditionPanel = document.createElement('div');
    conditionPanel.className = 'fc-filter-dropdown__condition';
    const conditionLabel = document.createElement('label');
    conditionLabel.className = 'fc-filter-dropdown__condition-label';
    conditionLabel.textContent = t.condition;
    const conditionOptions: Array<{ value: '' | ConditionFilterOp; label: string }> = [
      { value: '', label: t.conditionNone },
      { value: 'equals', label: t.conditionEquals },
      { value: 'notEquals', label: t.conditionNotEquals },
      { value: 'contains', label: t.conditionContains },
      { value: 'notContains', label: t.conditionNotContains },
      { value: 'greaterThan', label: t.conditionGreaterThan },
      { value: 'greaterThanOrEqual', label: t.conditionGreaterThanOrEqual },
      { value: 'lessThan', label: t.conditionLessThan },
      { value: 'lessThanOrEqual', label: t.conditionLessThanOrEqual },
    ];
    const conditionSelect = createDialogSelect(conditionOptions, '', {
      ariaLabel: t.condition,
      className: 'fc-filter-dropdown__condition-op',
    });
    const conditionInput = document.createElement('input');
    conditionInput.type = 'text';
    conditionInput.className = 'fc-filter-dropdown__condition-value';
    conditionInput.placeholder = t.conditionValue;
    conditionInput.setAttribute('aria-label', t.conditionValue);
    conditionLabel.appendChild(conditionSelect);
    conditionPanel.append(conditionLabel, conditionInput);

    const list = document.createElement('div');
    list.className = 'fc-filter-dropdown__list';
    list.setAttribute('role', 'group');
    list.setAttribute('aria-label', t.title);
    list.tabIndex = -1;

    const rowCheckboxes = (): HTMLInputElement[] =>
      Array.from(list.querySelectorAll<HTMLInputElement>('input[type="checkbox"]'));
    const focusCheckbox = (idx: number): void => {
      const boxes = rowCheckboxes();
      if (boxes.length === 0) return;
      const next = (idx + boxes.length) % boxes.length;
      boxes[next]?.focus({ preventScroll: true });
    };
    const handleRowKey = (event: KeyboardEvent): void => {
      const boxes = rowCheckboxes();
      const idx = boxes.indexOf(event.currentTarget as HTMLInputElement);
      if (idx < 0) return;
      if (event.key === 'ArrowDown') {
        event.preventDefault();
        focusCheckbox(idx + 1);
      } else if (event.key === 'ArrowUp') {
        event.preventDefault();
        focusCheckbox(idx - 1);
      } else if (event.key === 'Home') {
        event.preventDefault();
        focusCheckbox(0);
      } else if (event.key === 'End') {
        event.preventDefault();
        focusCheckbox(boxes.length - 1);
      }
    };

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
      allCb.addEventListener('keydown', handleRowKey);
      allCb.addEventListener('change', () => {
        if (allCb.checked) {
          hidden.clear();
        } else {
          for (const v of distinct) hidden.add(v);
        }
        renderRows(search.value);
        requestAnimationFrame(() => focusCheckbox(0));
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
        cb.addEventListener('keydown', handleRowKey);
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
    search.addEventListener('keydown', (event) => {
      if (event.key === 'ArrowDown') {
        event.preventDefault();
        focusCheckbox(0);
      } else if (event.key === 'End' && (event.ctrlKey || event.metaKey)) {
        event.preventDefault();
        focusCheckbox(rowCheckboxes().length - 1);
      }
    });

    const actions = document.createElement('div');
    actions.className = 'fc-filter-dropdown__actions';
    const apply = createFilterDropdownActionButton('fc-filter-dropdown__apply', t.apply);
    apply.addEventListener('click', () => applyActiveFilter());
    const clear = createFilterDropdownActionButton('fc-filter-dropdown__clear', t.clear);
    clear.addEventListener('click', () => {
      if (!activeRange) return;
      recordFilterChange(deps.history ?? null, deps.store, () => {
        clearFilter(deps.store.getState(), deps.store, activeRange as Range);
      });
      close();
    });
    actions.append(clear, apply);

    r.append(conditionPanel, search, list, actions);
    // Borrow theme tokens from the first .fc-host on the page (filter is
    // body-attached, so `[data-fc-theme]` doesn't cascade automatically).
    const host = document.querySelector('.fc-host');
    if (host) inheritHostTokens(host, r);
    document.body.appendChild(r);
    root = r;
    const position = clampPanelToViewport(r, anchor.x, anchor.y + anchor.h, {
      pad: 4,
      fallbackWidth: 260,
      fallbackHeight: 320,
    });
    r.style.left = `${position.x}px`;
    r.style.top = `${position.y}px`;

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
