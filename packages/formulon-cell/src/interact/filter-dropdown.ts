import { parseLocalizedNumericInput } from '../commands/coerce-input.js';
import {
  applyConditionFilter,
  applyValueFilter,
  type ConditionFilterOp,
  clearFilter,
  type DistinctFilterFailureReason,
  distinctFilterItemsResult,
  recordFilterChange,
} from '../commands/filter.js';
import type { History } from '../commands/history.js';
import { createProfileTextSemantics, localeForWorkbook } from '../engine/profile-text-semantics.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { sameRange } from '../store/selection-geometry.js';
import type { SpreadsheetStore, ValueFilterCriteria } from '../store/store.js';
import { createDialogSelect } from '../toolbar/dialogs/form-controls.js';
import { createInteractionButton } from './chip-button.js';
import { overlayPortalFor } from './overlay-portal.js';
import { clampPanelToViewport } from './overlay-position.js';

export interface FilterDropdownDeps {
  store: SpreadsheetStore;
  /** Host whose themed overlay portal owns the dropdown. */
  host?: HTMLElement;
  history?: History | null;
  /** Optional engine used to apply typed AutoFilter criteria and native row
   * visibility as the dropdown commits a change. */
  getWb?: () => WorkbookHandle;
  strings?: Strings;
  /** Locale used to format the checklist labels (dates, currency). Defaults to
   *  `en-US`. */
  locale?: string;
  /** Receives an explicit native checklist refusal/failure. The result still
   *  contains a safe locale-aware approximation for rendering. */
  onDistinctFailure?: (reason: DistinctFilterFailureReason) => void;
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

const createFilterDropdownActionButton = (className: string, label: string): HTMLButtonElement => {
  return createInteractionButton({
    className,
    text: label,
  });
};

const findColumnCriteria = (
  criteria: readonly ValueFilterCriteria[],
  range: Range,
  col: number,
): ValueFilterCriteria | undefined =>
  criteria.find((c) => sameRange(c.range, range) && c.byCol === col);

/**
 * Lightweight column-filter popover. Lists distinct values in the column with
 * a checkbox each; "Apply" calls `applyFilter`, "Clear" calls `clearFilter`.
 *
 * The popover lives in the overlay portal so it can escape any clipping
 * ancestors, and is positioned via fixed coordinates from the anchor rect.
 */
export function attachFilterDropdown(deps: FilterDropdownDeps): FilterDropdownHandle {
  const strings = deps.strings ?? defaultStrings;
  const t = strings.filterDropdown;
  const ownerDocument = deps.host?.ownerDocument ?? document;
  const frame = ownerDocument.defaultView
    ? ownerDocument.defaultView.requestAnimationFrame.bind(ownerDocument.defaultView)
    : requestAnimationFrame;
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
    ownerDocument.removeEventListener('mousedown', onDocMouseDown, true);
    ownerDocument.removeEventListener('keydown', onDocKey, true);
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
    const wb = deps.getWb?.();
    const condition = root.querySelector<HTMLSelectElement>('.fc-filter-dropdown__condition-op');
    const conditionValue = root.querySelector<HTMLInputElement>(
      '.fc-filter-dropdown__condition-value',
    );
    recordFilterChange(
      deps.history ?? null,
      deps.store,
      () => {
        if (condition?.value && conditionValue?.value.trim()) {
          applyConditionFilter(
            deps.store.getState(),
            deps.store,
            activeRange as Range,
            activeCol,
            {
              op: condition.value as ConditionFilterOp,
              value: conditionValue.value,
            },
            wb,
          );
        } else {
          applyValueFilter(
            deps.store.getState(),
            deps.store,
            activeRange as Range,
            activeCol,
            Array.from(activeHidden),
            wb,
          );
        }
      },
      wb,
    );
    close();
  };

  const open = (range: Range, col: number, anchor: { x: number; y: number; h: number }): void => {
    close();
    const activeElement = ownerDocument.activeElement;
    restoreFocus =
      activeElement?.nodeType === 1 && activeElement !== ownerDocument.body
        ? (activeElement as HTMLElement)
        : null;
    activeRange = range;
    activeCol = col;

    const state = deps.store.getState();
    const wb = deps.getWb?.();
    const activeLocale = deps.locale ?? localeForWorkbook(wb);
    const semantics = createProfileTextSemantics(activeLocale);
    const distinctResult = distinctFilterItemsResult(state, range, col, activeLocale, wb);
    if (distinctResult.failure) deps.onDistinctFailure?.(distinctResult.failure);
    const distinct = distinctResult.items;
    const activeCriteria = findColumnCriteria(state.ui.filterCriteria, range, col);
    const hidden = new Set<string>(
      activeCriteria && !activeCriteria.condition && !activeCriteria.color
        ? activeCriteria.hiddenValues
        : [],
    );
    activeHidden = hidden;

    const r = ownerDocument.createElement('div');
    r.className = 'fc-filter-dropdown';
    r.style.position = 'fixed';
    r.style.left = `${anchor.x}px`;
    r.style.top = `${anchor.y + anchor.h}px`;
    r.setAttribute('role', 'dialog');
    r.setAttribute('aria-modal', 'false');
    r.setAttribute('aria-label', t.title);

    const search = ownerDocument.createElement('input');
    search.className = 'fc-filter-dropdown__search';
    search.type = 'search';
    search.placeholder = t.searchPlaceholder;
    search.setAttribute('aria-label', t.searchPlaceholder);
    search.spellcheck = false;

    const conditionPanel = ownerDocument.createElement('div');
    conditionPanel.className = 'fc-filter-dropdown__condition';
    const conditionLabel = ownerDocument.createElement('label');
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
    if (activeCriteria?.condition) conditionSelect.value = activeCriteria.condition.op;
    const conditionInput = ownerDocument.createElement('input');
    conditionInput.type = 'text';
    conditionInput.className = 'fc-filter-dropdown__condition-value';
    conditionInput.placeholder = t.conditionValue;
    conditionInput.setAttribute('aria-label', t.conditionValue);
    if (activeCriteria?.condition) {
      const activeValue = activeCriteria.condition.value;
      if (
        activeCriteria.condition.op === 'equals' ||
        activeCriteria.condition.op === 'notEquals' ||
        activeCriteria.condition.op === 'greaterThan' ||
        activeCriteria.condition.op === 'greaterThanOrEqual' ||
        activeCriteria.condition.op === 'lessThan' ||
        activeCriteria.condition.op === 'lessThanOrEqual'
      ) {
        // Criteria are stored in the engine's canonical decimal grammar;
        // accept the active locale first, then decode that canonical form so
        // a reopened German/French dropdown displays `1,5` again.
        const numeric =
          parseLocalizedNumericInput(activeValue, activeLocale) ??
          parseLocalizedNumericInput(activeValue, 'en-US');
        conditionInput.value =
          numeric !== null ? new Intl.NumberFormat(activeLocale).format(numeric) : activeValue;
      } else {
        conditionInput.value = activeValue;
      }
    }
    conditionLabel.appendChild(conditionSelect);
    conditionPanel.append(conditionLabel, conditionInput);

    const list = ownerDocument.createElement('div');
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
      const f = filter.toLocaleLowerCase(activeLocale);
      const hiddenHas = (value: string): boolean => {
        const key = semantics.key(value);
        for (const hiddenValue of hidden) {
          if (semantics.key(hiddenValue) === key) return true;
        }
        return false;
      };
      const hiddenDelete = (value: string): void => {
        const key = semantics.key(value);
        for (const hiddenValue of hidden) {
          if (semantics.key(hiddenValue) === key) hidden.delete(hiddenValue);
        }
      };
      // (Select All) header
      const allRow = ownerDocument.createElement('label');
      allRow.className = 'fc-filter-dropdown__row fc-filter-dropdown__row--all';
      const allCb = ownerDocument.createElement('input');
      allCb.type = 'checkbox';
      allCb.checked = distinct.every((v) => !hiddenHas(v.key));
      allCb.indeterminate = !allCb.checked && distinct.some((v) => !hiddenHas(v.key));
      allCb.addEventListener('keydown', handleRowKey);
      allCb.addEventListener('change', () => {
        if (allCb.checked) {
          hidden.clear();
        } else {
          for (const v of distinct) hidden.add(v.key);
        }
        renderRows(search.value);
        frame(() => focusCheckbox(0));
      });
      const allLabel = ownerDocument.createElement('span');
      allLabel.textContent = t.selectAll;
      allRow.append(allCb, allLabel);
      list.appendChild(allRow);

      for (const item of distinct) {
        const display = item.key === '' ? t.blanks : item.label;
        if (f && !display.toLocaleLowerCase(activeLocale).includes(f)) continue;
        const row = ownerDocument.createElement('label');
        row.className = 'fc-filter-dropdown__row';
        const cb = ownerDocument.createElement('input');
        cb.type = 'checkbox';
        cb.value = item.key;
        cb.checked = !hiddenHas(item.key);
        cb.addEventListener('keydown', handleRowKey);
        cb.addEventListener('change', () => {
          if (cb.checked) hiddenDelete(item.key);
          else hidden.add(item.key);
          // Refresh select-all indeterminate state without rebuilding rows.
          allCb.checked = distinct.every((vv) => !hiddenHas(vv.key));
          allCb.indeterminate = !allCb.checked && distinct.some((vv) => !hiddenHas(vv.key));
        });
        const text = ownerDocument.createElement('span');
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

    const actions = ownerDocument.createElement('div');
    actions.className = 'fc-filter-dropdown__actions';
    const apply = createFilterDropdownActionButton('fc-filter-dropdown__apply', t.apply);
    apply.addEventListener('click', () => applyActiveFilter());
    const clear = createFilterDropdownActionButton('fc-filter-dropdown__clear', t.clear);
    clear.addEventListener('click', () => {
      if (!activeRange) return;
      const wb = deps.getWb?.();
      recordFilterChange(
        deps.history ?? null,
        deps.store,
        () => {
          clearFilter(deps.store.getState(), deps.store, activeRange as Range, wb);
        },
        wb,
      );
      close();
    });
    actions.append(clear, apply);

    r.append(conditionPanel, search, list, actions);
    overlayPortalFor(deps.host).appendChild(r);
    root = r;
    const position = clampPanelToViewport(r, anchor.x, anchor.y + anchor.h, {
      pad: 4,
      fallbackWidth: 260,
      fallbackHeight: 320,
    });
    r.style.left = `${position.x}px`;
    r.style.top = `${position.y}px`;

    frame(() => search.focus());

    ownerDocument.addEventListener('mousedown', onDocMouseDown, true);
    ownerDocument.addEventListener('keydown', onDocKey, true);
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
