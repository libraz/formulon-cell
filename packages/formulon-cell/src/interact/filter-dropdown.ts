import { canExecuteBuiltIn } from '../commands/built-in-command-policy.js';
import {
  inputLocaleForWorkbook,
  parseLocalizedNumericInput,
  workbookDate1904,
} from '../commands/coerce-input.js';
import {
  DATE_PERIOD_GROUP_STARTS,
  DATE_PERIODS,
  numFmtShowsDate,
} from '../commands/date-periods.js';
import {
  applyColorFilter,
  applyConditionFilter,
  applyValueFilter,
  type ConditionFilterOp,
  type ConditionFilterOptions,
  type DistinctFilterFailureReason,
  distinctFilterItemsResult,
  type FilterValueItem,
  isFilterComparisonOp,
  resetFilterCriteria,
} from '../commands/filter.js';
import { filterUseWrite } from '../commands/filter-actions.js';
import { columnColorChoices } from '../commands/filter-colors.js';
import { columnDateSerials } from '../commands/filter-date-groups.js';
import type { History } from '../commands/history.js';
import { parseDateTimeInput } from '../commands/input-grammar.js';
import { sortRangeWithHistory } from '../commands/sort.js';
import { addrKey } from '../engine/address.js';
import { createProfileTextSemantics, localeForWorkbook } from '../engine/profile-text-semantics.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { excelSerialToDateParts, excelSerialToIsoDate } from '../format/date-serial.js';
import type { DatePeriodKey } from '../i18n/strings/data-objects/types.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { drawConditionalIconArt } from '../render/painters/cf-icons.js';
import { sameRange } from '../store/selection-geometry.js';
import type { SpreadsheetStore, ValueFilterCriteria } from '../store/store.js';
import type { ConditionalIconSet } from '../store/types/conditional-rule.js';
import type { FilterColorSpec } from '../store/types.js';
import { createDialogSelect } from '../toolbar/dialogs/form-controls.js';
import { createExcelRibbonSvg } from '../toolbar/excel-ribbon-icons.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { createInteractionButton } from './chip-button.js';
import { deferInitialFocus } from './dialog-shell.js';
import { onOutsidePress } from './outside-press.js';
import { overlayPortalFor } from './overlay-portal.js';
import { placePanel } from './overlay-position.js';

/** Select value prefix of a relative date period. */
const DATE_PERIOD_PREFIX = 'datePeriod:';

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
 * a checkbox each; "Apply" calls `applyFilter`, "Clear" calls `resetFilterCriteria`.
 *
 * The popover lives in the overlay portal so it can escape any clipping
 * ancestors, and is positioned via fixed coordinates from the anchor rect.
 */
export function attachFilterDropdown(deps: FilterDropdownDeps): FilterDropdownHandle {
  const strings = deps.strings ?? defaultStrings;
  const t = strings.filterDropdown;
  const m = strings.mac.filterPanel;
  const fcHost = deps.host?.closest<HTMLElement>('.fc-host') ?? deps.host;
  const mac = fcHost?.dataset.fcPlatform === 'mac';
  const ownerDocument = deps.host?.ownerDocument ?? document;
  const frame = ownerDocument.defaultView
    ? ownerDocument.defaultView.requestAnimationFrame.bind(ownerDocument.defaultView)
    : requestAnimationFrame;
  let root: HTMLDivElement | null = null;
  let activeRange: Range | null = null;
  let activeCol = 0;
  let activeHidden = new Set<string>();
  let restoreFocus: HTMLElement | null = null;
  let detachOutside: (() => void) | null = null;

  const close = (): void => {
    if (!root) return;
    root.remove();
    root = null;
    activeRange = null;
    activeHidden = new Set();
    detachOutside?.();
    detachOutside = null;
    ownerDocument.removeEventListener('keydown', onDocKey, true);
    restoreFocus?.focus({ preventScroll: true });
    restoreFocus = null;
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

  /** A typed date as its serial, read in the workbook's locale; the text itself when it is no date. */
  const dateOperand = (text: string): string => {
    const wb = deps.getWb?.();
    const parsed = parseDateTimeInput(
      text.trim(),
      wb ? inputLocaleForWorkbook(wb) : (deps.locale ?? 'en-US'),
      wb ? workbookDate1904(wb) : false,
    );
    return parsed ? String(parsed.value) : text;
  };

  /** The condition the panel's fields describe, or null while an operand is missing. */
  const readCondition = (picked: string): ConditionFilterOptions | null => {
    if (!root) return null;
    if (picked.startsWith(DATE_PERIOD_PREFIX)) {
      return {
        op: 'datePeriod',
        value: '',
        period: picked.slice(DATE_PERIOD_PREFIX.length) as DatePeriodKey,
      };
    }
    const op = picked as ConditionFilterOp;
    const dates = root.dataset.fcDateColumn === 'true';
    const field = (name: string): string => {
      const raw =
        root?.querySelector<HTMLInputElement | HTMLSelectElement>(`.fc-filter-dropdown__${name}`)
          ?.value ?? '';
      return dates && !name.endsWith('unit') ? dateOperand(raw) : raw;
    };
    const value = field('condition-value').trim();
    if (op === 'aboveAverage' || op === 'belowAverage') return { op, value: '' };
    if (op === 'top10' || op === 'bottom10') {
      return value
        ? { op, value, unit: field('condition-unit') === 'percent' ? 'percent' : 'items' }
        : null;
    }
    if (op === 'between') {
      const upper = field('condition-value2').trim();
      return value && upper ? { op, value, value2: upper } : null;
    }
    if (!value) return null;
    const first = { op, value: field('condition-value') };
    const secondOp =
      root.querySelector<HTMLSelectElement>('.fc-filter-dropdown__condition-op2')?.value ?? '';
    const secondValue = field('condition-second').trim();
    if (!isFilterComparisonOp(op) || !isFilterComparisonOp(secondOp) || !secondValue) {
      return first;
    }
    const or = root.querySelector<HTMLInputElement>(
      '.fc-filter-dropdown__condition-join input[value="or"]',
    )?.checked;
    return { ...first, second: { op: secondOp, value: secondValue }, join: or ? 'or' : 'and' };
  };

  /** A filter change made from this dropdown under `commandId`. */
  const dropdownFilter = (commandId: string, wb: WorkbookHandle | undefined) => ({
    store: deps.store,
    history: deps.history ?? null,
    ...(wb ? { workbook: wb } : {}),
    origin: 'contextMenu' as const,
    commandId,
  });

  /** Apply the panel's state; the Mac panel keeps itself open while it auto-applies. */
  const applyActiveFilter = (closeAfter = true): void => {
    if (!activeRange || !root) return;
    const wb = deps.getWb?.();
    const condition = root.querySelector<HTMLSelectElement>('.fc-filter-dropdown__condition-op');
    filterUseWrite(dropdownFilter('filterByValue', wb), activeRange, undefined, () => {
      const picked = condition?.value ?? '';
      const spec = picked ? readCondition(picked) : null;
      if (spec) {
        applyConditionFilter(
          deps.store.getState(),
          deps.store,
          activeRange as Range,
          activeCol,
          spec,
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
    });
    if (closeAfter) close();
  };

  /** Whether the column's data is mostly numbers, which picks Excel's numeric condition list. */
  const numericColumn = (range: Range, col: number): boolean => {
    const cells = deps.store.getState().data.cells;
    let numbers = 0;
    let texts = 0;
    for (let row = range.r0 + 1; row <= range.r1; row += 1) {
      const kind = cells.get(addrKey({ sheet: range.sheet, row, col }))?.value.kind;
      if (kind === 'number') numbers += 1;
      else if (kind === 'text') texts += 1;
    }
    return numbers > 0 && numbers >= texts;
  };

  /** Whether the column's numbers are mostly dates, which picks the date condition list. */
  const dateColumn = (range: Range, col: number): boolean => {
    const state = deps.store.getState();
    let dates = 0;
    let others = 0;
    for (let row = range.r0 + 1; row <= range.r1; row += 1) {
      const key = addrKey({ sheet: range.sheet, row, col });
      const kind = state.data.cells.get(key)?.value.kind;
      if (kind !== 'number' && kind !== 'text') continue;
      if (kind === 'number' && numFmtShowsDate(state.format.formats.get(key)?.numFmt)) dates += 1;
      else others += 1;
    }
    return dates > 0 && dates >= others;
  };

  /** The column's header text, which titles the Mac panel. */
  const columnTitle = (range: Range, col: number): string => {
    const value = deps.store
      .getState()
      .data.cells.get(addrKey({ sheet: range.sheet, row: range.r0, col }))?.value;
    if (value?.kind === 'text') return value.value;
    return value?.kind === 'number' ? String(value.value) : '';
  };

  /** A conditional formatting icon drawn at 14 px, as the grid paints it. */
  const iconCanvas = (set: ConditionalIconSet, slot: number): HTMLCanvasElement => {
    const canvas = ownerDocument.createElement('canvas');
    const ratio = ownerDocument.defaultView?.devicePixelRatio ?? 1;
    canvas.width = Math.round(14 * ratio);
    canvas.height = Math.round(14 * ratio);
    canvas.className = 'fc-filter-dropdown__swatch-icon';
    const ctx = canvas.getContext('2d');
    if (ctx) {
      ctx.scale((14 * ratio) / 12, (14 * ratio) / 12);
      drawConditionalIconArt(ctx, set, slot);
    }
    return canvas;
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
    const searchLabel = mac ? m.search : t.searchPlaceholder;
    search.placeholder = searchLabel;
    search.setAttribute('aria-label', searchLabel);
    search.spellcheck = false;

    let autoApplyIfOn = (): void => {};
    let colorBy: HTMLElement | null = null;
    const conditionPanel = ownerDocument.createElement('div');
    conditionPanel.className = 'fc-filter-dropdown__condition';
    const conditionLabel = ownerDocument.createElement('label');
    conditionLabel.className = 'fc-filter-dropdown__condition-label';
    conditionLabel.textContent = t.condition;
    const dates = mac && dateColumn(range, col);
    const numeric = mac && !dates && numericColumn(range, col);
    r.dataset.fcDateColumn = dates ? 'true' : 'false';
    const conditionOptions: Array<{ value: string; label: string; disabled?: boolean }> = dates
      ? [
          // The Mac pop-up's order for a date column is provisional until measured on the Mac app.
          { value: '', label: m.chooseOne },
          { value: 'equals', label: m.equals },
          { value: 'notEquals', label: m.notEquals },
          { value: 'lessThan', label: m.before },
          { value: 'lessThanOrEqual', label: m.beforeOrEqual },
          { value: 'greaterThan', label: m.after },
          { value: 'greaterThanOrEqual', label: m.afterOrEqual },
          { value: 'between', label: m.between },
          ...DATE_PERIODS.flatMap(([key], index) => [
            ...(DATE_PERIOD_GROUP_STARTS.has(index)
              ? [{ value: `-${index}`, label: '──────────', disabled: true }]
              : []),
            { value: `${DATE_PERIOD_PREFIX}${key}`, label: m.periods[key] },
          ]),
        ]
      : mac
        ? [
            { value: '', label: m.chooseOne },
            { value: 'equals', label: m.equals },
            { value: 'notEquals', label: m.notEquals },
            ...(numeric
              ? ([
                  { value: 'greaterThan', label: m.greaterThan },
                  { value: 'greaterThanOrEqual', label: m.greaterThanOrEqual },
                  { value: 'lessThan', label: m.lessThan },
                  { value: 'lessThanOrEqual', label: m.lessThanOrEqual },
                  { value: 'between', label: m.between },
                  { value: 'top10', label: m.top10 },
                  { value: 'bottom10', label: m.bottom10 },
                  { value: 'aboveAverage', label: m.aboveAverage },
                  { value: 'belowAverage', label: m.belowAverage },
                ] as const)
              : ([
                  { value: 'contains', label: m.contains },
                  { value: 'notContains', label: m.notContains },
                ] as const)),
          ]
        : [
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
    if (activeCriteria?.condition) {
      const active = activeCriteria.condition;
      conditionSelect.value =
        active.op === 'datePeriod' ? `${DATE_PERIOD_PREFIX}${active.period ?? ''}` : active.op;
    }
    const conditionInput = ownerDocument.createElement('input');
    conditionInput.type = 'text';
    conditionInput.className = 'fc-filter-dropdown__condition-value';
    conditionInput.placeholder = t.conditionValue;
    conditionInput.setAttribute('aria-label', t.conditionValue);
    // A stored operand as typed back: a date column's serial as its date, any other number in
    // the active locale (the store keeps the engine's canonical decimal grammar).
    const operandText = (stored: string): string => {
      const numeric =
        parseLocalizedNumericInput(stored, activeLocale) ??
        parseLocalizedNumericInput(stored, 'en-US');
      if (numeric === null) return stored;
      if (dates) return excelSerialToIsoDate(numeric, state.ui.date1904 === true) ?? stored;
      return new Intl.NumberFormat(activeLocale).format(numeric);
    };
    if (activeCriteria?.condition) {
      const activeValue = activeCriteria.condition.value;
      if (
        activeCriteria.condition.op === 'equals' ||
        activeCriteria.condition.op === 'notEquals' ||
        activeCriteria.condition.op === 'greaterThan' ||
        activeCriteria.condition.op === 'greaterThanOrEqual' ||
        activeCriteria.condition.op === 'lessThan' ||
        activeCriteria.condition.op === 'lessThanOrEqual' ||
        activeCriteria.condition.op === 'between' ||
        activeCriteria.condition.op === 'top10' ||
        activeCriteria.condition.op === 'bottom10'
      ) {
        // Criteria are stored in the engine's canonical decimal grammar;
        // accept the active locale first, then decode that canonical form so
        // a reopened German/French dropdown displays `1,5` again.
        conditionInput.value = operandText(activeValue);
      } else {
        conditionInput.value = activeValue;
      }
    }
    if (mac) {
      conditionInput.placeholder = '';
      const upperInput = ownerDocument.createElement('input');
      upperInput.type = 'text';
      upperInput.className = 'fc-filter-dropdown__condition-value2';
      upperInput.setAttribute('aria-label', t.conditionValue);
      const unitSelect = createDialogSelect(
        [
          { value: 'items', label: m.items },
          { value: 'percent', label: m.percent },
        ],
        activeCriteria?.condition?.unit ?? 'items',
        { ariaLabel: m.items, className: 'fc-filter-dropdown__condition-unit' },
      );
      if (activeCriteria?.condition?.value2 !== undefined) {
        upperInput.value = operandText(activeCriteria.condition.value2);
      }
      // The operand fields each operator needs: none, one value, two values, or a count and unit.
      const syncFields = (): void => {
        const op = conditionSelect.value;
        const ranked = op === 'top10' || op === 'bottom10';
        const average =
          op === 'aboveAverage' || op === 'belowAverage' || op.startsWith(DATE_PERIOD_PREFIX);
        conditionInput.hidden = average;
        upperInput.hidden = op !== 'between';
        unitSelect.hidden = !ranked;
        conditionInput.type = ranked ? 'number' : 'text';
        if (ranked) {
          conditionInput.min = '1';
          if (!conditionInput.value) conditionInput.value = '10';
        } else {
          conditionInput.removeAttribute('min');
        }
      };
      syncFields();
      conditionSelect.addEventListener('change', () => {
        syncFields();
        const op = conditionSelect.value;
        if (op === 'aboveAverage' || op === 'belowAverage' || op.startsWith(DATE_PERIOD_PREFIX)) {
          autoApplyIfOn();
        } else if (op) conditionInput.focus();
        else autoApplyIfOn();
      });
      conditionInput.addEventListener('change', () => autoApplyIfOn());
      upperInput.addEventListener('change', () => autoApplyIfOn());
      unitSelect.addEventListener('change', () => autoApplyIfOn());
      // A second comparison joined by And or Or, offered once the first takes one operand. The
      // row's layout on the Mac app is provisional until measured.
      const active = activeCriteria?.condition;
      const joinRow = ownerDocument.createElement('div');
      joinRow.className = 'fc-filter-dropdown__condition-join';
      joinRow.setAttribute('role', 'radiogroup');
      const joinName = `fc-filter-join-${col}`;
      for (const [value, label] of [
        ['and', m.and],
        ['or', m.or],
      ] as const) {
        const radioLabel = ownerDocument.createElement('label');
        const radio = ownerDocument.createElement('input');
        radio.type = 'radio';
        radio.name = joinName;
        radio.value = value;
        radio.checked = (active?.join ?? 'and') === value;
        radio.addEventListener('change', () => autoApplyIfOn());
        radioLabel.append(radio, label);
        joinRow.appendChild(radioLabel);
      }
      const secondSelect = createDialogSelect(
        conditionOptions.filter(
          (option) => option.value === '' || isFilterComparisonOp(option.value),
        ),
        active?.second?.op ?? '',
        { ariaLabel: t.condition, className: 'fc-filter-dropdown__condition-op2' },
      );
      const secondInput = ownerDocument.createElement('input');
      secondInput.type = 'text';
      secondInput.className = 'fc-filter-dropdown__condition-second';
      secondInput.setAttribute('aria-label', t.conditionValue);
      if (active?.second) secondInput.value = operandText(active.second.value);
      const syncSecond = (): void => {
        const offered = isFilterComparisonOp(conditionSelect.value);
        joinRow.hidden = !offered;
        secondSelect.hidden = !offered;
        secondInput.hidden = !offered || secondSelect.value === '';
      };
      syncSecond();
      conditionSelect.addEventListener('change', syncSecond);
      secondSelect.addEventListener('change', () => {
        syncSecond();
        if (secondSelect.value) secondInput.focus();
        else autoApplyIfOn();
      });
      secondInput.addEventListener('change', () => autoApplyIfOn());
      conditionPanel.append(
        conditionSelect,
        conditionInput,
        upperInput,
        unitSelect,
        joinRow,
        secondSelect,
        secondInput,
      );
    } else {
      conditionLabel.appendChild(conditionSelect);
      conditionPanel.append(conditionLabel, conditionInput);
    }

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

    // A column's dated numbers group by year and month, as on the desktop.
    const dateSerials = columnDateSerials(state, range, col);
    const date1904 = state.ui.date1904 === true;
    const expanded = new Set<string>();
    const monthFormat = new Intl.DateTimeFormat(activeLocale, { month: 'long', timeZone: 'UTC' });
    const monthName = (month: number): string =>
      monthFormat.format(new Date(Date.UTC(2000, month - 1, 1)));

    const renderRows = (filter: string): void => {
      list.innerHTML = '';
      const f = filter.toLocaleLowerCase(activeLocale);
      // Semantic keys of `hidden`, rebuilt after each mutation.
      let hiddenKeys: Set<string> | null = null;
      const hiddenHas = (value: string): boolean => {
        if (!hiddenKeys) hiddenKeys = new Set([...hidden].map((v) => semantics.key(v)));
        return hiddenKeys.has(semantics.key(value));
      };
      const hiddenAdd = (value: string): void => {
        hidden.add(value);
        hiddenKeys = null;
      };
      const hiddenDelete = (value: string): void => {
        const key = semantics.key(value);
        for (const hiddenValue of hidden) {
          if (semantics.key(hiddenValue) === key) hidden.delete(hiddenValue);
        }
        hiddenKeys = null;
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
          hiddenKeys = null;
        } else {
          for (const v of distinct) hiddenAdd(v.key);
        }
        renderRows(search.value);
        autoApplyIfOn();
        frame(() => focusCheckbox(0));
      });
      const allLabel = ownerDocument.createElement('span');
      allLabel.textContent = mac ? m.selectAll : t.selectAll;
      allRow.append(allCb, allLabel);
      list.appendChild(allRow);

      // In the date tree a change re-renders the rows, so the parents' states follow; the focus
      // returns to the checkbox that changed.
      const rerender = (focusValue: string): void => {
        renderRows(search.value);
        frame(() =>
          rowCheckboxes()
            .find((box) => box.value === focusValue)
            ?.focus({ preventScroll: true }),
        );
      };
      const itemRow = (item: FilterValueItem, display: string, level: number, tree: boolean) => {
        const row = ownerDocument.createElement('label');
        row.className = 'fc-filter-dropdown__row';
        if (level > 0) row.dataset.fcLevel = String(level);
        const cb = ownerDocument.createElement('input');
        cb.type = 'checkbox';
        cb.value = item.key;
        cb.checked = !hiddenHas(item.key);
        cb.addEventListener('keydown', handleRowKey);
        cb.addEventListener('change', () => {
          if (cb.checked) hiddenDelete(item.key);
          else hiddenAdd(item.key);
          autoApplyIfOn();
          if (tree) {
            rerender(item.key);
            return;
          }
          // Refresh select-all indeterminate state without rebuilding rows.
          allCb.checked = distinct.every((vv) => !hiddenHas(vv.key));
          allCb.indeterminate = !allCb.checked && distinct.some((vv) => !hiddenHas(vv.key));
        });
        const text = ownerDocument.createElement('span');
        text.textContent = display;
        row.append(cb, text);
        list.appendChild(row);
      };
      // A year or month of the date tree: a disclosure toggle and a checkbox for all its dates.
      const groupRow = (path: string, label: string, items: FilterValueItem[], level: number) => {
        const row = ownerDocument.createElement('div');
        row.className = 'fc-filter-dropdown__row fc-filter-dropdown__row--group';
        row.dataset.fcLevel = String(level);
        const isOpen = expanded.has(path);
        const toggle = createInteractionButton({
          className: 'fc-filter-dropdown__disclosure',
          ariaLabel: label,
        });
        toggle.setAttribute('aria-expanded', isOpen ? 'true' : 'false');
        toggle.dataset.fcDatePath = path;
        toggle.addEventListener('click', () => {
          if (expanded.has(path)) expanded.delete(path);
          else expanded.add(path);
          renderRows(search.value);
        });
        const cb = ownerDocument.createElement('input');
        cb.type = 'checkbox';
        cb.value = `group:${path}`;
        const shown = items.filter((item) => !hiddenHas(item.key)).length;
        cb.checked = shown === items.length;
        cb.indeterminate = shown > 0 && shown < items.length;
        cb.addEventListener('keydown', handleRowKey);
        cb.addEventListener('change', () => {
          for (const item of items) {
            if (cb.checked) hiddenDelete(item.key);
            else hiddenAdd(item.key);
          }
          autoApplyIfOn();
          rerender(cb.value);
        });
        const text = ownerDocument.createElement('span');
        text.textContent = label;
        const box = ownerDocument.createElement('label');
        box.append(cb, text);
        row.append(toggle, box);
        list.appendChild(row);
        return isOpen;
      };

      const dated = f ? [] : distinct.filter((item) => dateSerials.has(item.key));
      if (dated.length > 0) {
        const years = new Map<number, Map<number, { item: FilterValueItem; day: number }[]>>();
        for (const item of dated) {
          const parts = excelSerialToDateParts(dateSerials.get(item.key) as number, date1904);
          if (!parts) continue;
          const months = years.get(parts.year) ?? new Map();
          years.set(parts.year, months);
          const days = months.get(parts.month) ?? [];
          months.set(parts.month, days);
          days.push({ item, day: parts.day });
        }
        for (const [year, months] of [...years].sort(([a], [b]) => a - b)) {
          const yearItems = [...months.values()].flat().map(({ item }) => item);
          if (!groupRow(String(year), String(year), yearItems, 0)) continue;
          for (const [month, days] of [...months].sort(([a], [b]) => a - b)) {
            const path = `${year}-${month}`;
            const items = days.map(({ item }) => item);
            if (!groupRow(path, monthName(month), items, 1)) continue;
            for (const { item, day } of [...days].sort((a, b) => a.day - b.day)) {
              itemRow(item, String(day), 2, true);
            }
          }
        }
      }
      for (const item of distinct) {
        if (dated.includes(item)) continue;
        const display = item.key === '' ? (mac ? m.blanks : t.blanks) : item.label;
        if (f && !display.toLocaleLowerCase(activeLocale).includes(f)) continue;
        itemRow(item, display, 0, dated.length > 0);
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
    const apply = createFilterDropdownActionButton(
      'fc-filter-dropdown__apply',
      mac ? m.apply : t.apply,
    );
    apply.addEventListener('click', () => applyActiveFilter());
    const clear = createFilterDropdownActionButton(
      'fc-filter-dropdown__clear',
      mac ? m.clear : t.clear,
    );
    clear.addEventListener('click', () => {
      if (!activeRange) return;
      const wb = deps.getWb?.();
      filterUseWrite(dropdownFilter('filterClear', wb), activeRange, undefined, () => {
        resetFilterCriteria(deps.store, activeRange as Range, wb);
      });
      // The desktop panel stays open on a cleared column, showing the idle state.
      if (mac) open(range, col, anchor);
      else close();
    });
    if (mac) {
      const autoApplyRow = ownerDocument.createElement('label');
      autoApplyRow.className = 'fc-filter-dropdown__auto';
      const autoApplyBox = ownerDocument.createElement('input');
      autoApplyBox.type = 'checkbox';
      autoApplyBox.checked = true;
      const autoApplyText = ownerDocument.createElement('span');
      autoApplyText.textContent = m.autoApply;
      autoApplyRow.append(autoApplyBox, autoApplyText);
      const syncApply = (): void => projectDisabledState(apply, autoApplyBox.checked, null);
      autoApplyBox.addEventListener('change', syncApply);
      syncApply();
      // The clear button is live only while the column carries a filter.
      const syncClear = (): void =>
        projectDisabledState(
          clear,
          !findColumnCriteria(deps.store.getState().ui.filterCriteria, range, col),
          null,
        );
      syncClear();
      autoApplyIfOn = () => {
        if (!autoApplyBox.checked) return;
        applyActiveFilter(false);
        syncClear();
      };
      actions.append(autoApplyRow, apply, clear);
      const sortRow = ownerDocument.createElement('div');
      sortRow.className = 'fc-filter-dropdown__sort';
      for (const direction of ['asc', 'desc'] as const) {
        const sortId = direction === 'asc' ? 'sortAsc' : 'sortDesc';
        const button = createFilterDropdownActionButton(
          `fc-filter-dropdown__sort-${direction}`,
          direction === 'asc' ? m.sortAsc : m.sortDesc,
        );
        button.addEventListener('click', () => {
          const wb = deps.getWb?.();
          if (!activeRange || !wb || !deps.history) return;
          sortRangeWithHistory({
            store: deps.store,
            workbook: wb,
            history: deps.history,
            range: activeRange,
            options: { byCol: activeCol, direction, hasHeader: true },
            origin: 'contextMenu',
            commandId: sortId,
          });
        });
        projectDisabledState(
          button,
          !deps.getWb ||
            !deps.history ||
            !canExecuteBuiltIn(deps.store, sortId, 'contextMenu').allowed,
          m.unavailable,
        );
        sortRow.append(button);
      }
      // The colour or icon the column was last sorted by, which the sort side shows as picked.
      const sortColor = (): FilterColorSpec | undefined => {
        const key = deps.store
          .getState()
          .ui.filterSorts?.find((spec) => sameRange(spec.range, range))
          ?.keys.find((k) => k.byCol === col);
        if (key?.sortOn === 'cellIcon' && key.icon) return { kind: 'cellIcon', icon: key.icon };
        if ((key?.sortOn === 'cellColor' || key?.sortOn === 'fontColor') && key.color) {
          return { kind: key.sortOn, color: key.color };
        }
        return undefined;
      };
      // Rows with the colour or icon go first, as the desktop's Sort by color does.
      const sortByColor = (wb: WorkbookHandle, history: History, color: FilterColorSpec): void => {
        sortRangeWithHistory({
          store: deps.store,
          workbook: wb,
          history,
          range,
          options: {
            byCol: col,
            direction: 'asc',
            hasHeader: true,
            keys: [
              color.kind === 'cellIcon'
                ? { byCol: col, direction: 'asc', sortOn: 'cellIcon', icon: color.icon }
                : { byCol: col, direction: 'asc', sortOn: color.kind, color: color.color },
            ],
          },
          origin: 'contextMenu',
          commandId: 'sortAsc',
        });
      };
      // By color, on the sort side and on the filter side: a kind, then the fills, font colours or
      // icons the column shows. The pop-up's layout on the Mac app is provisional until measured.
      const choices = columnColorChoices(state, range, col);
      const colorPopup = (side: 'sort' | 'filter'): HTMLElement => {
        const active = side === 'filter' ? activeCriteria?.color : sortColor();
        const wrap = ownerDocument.createElement('div');
        wrap.className = 'fc-filter-dropdown__color-filter';
        const row = ownerDocument.createElement('label');
        row.className = 'fc-filter-dropdown__color';
        const text = ownerDocument.createElement('span');
        text.textContent = m.byColor;
        const select = createDialogSelect(
          [
            { value: '', label: m.colorNone },
            { value: 'cellColor', label: m.cellColor, disabled: choices.fills.length === 0 },
            { value: 'fontColor', label: m.fontColor, disabled: choices.fonts.length === 0 },
            { value: 'cellIcon', label: m.cellIcon, disabled: choices.icons.length === 0 },
          ],
          active?.kind ?? '',
          { ariaLabel: m.byColor, className: 'fc-filter-dropdown__color-kind' },
        );
        row.append(text, select);
        const swatches = ownerDocument.createElement('div');
        swatches.className = 'fc-filter-dropdown__swatches';
        swatches.setAttribute('role', 'group');
        const applyColor = (color: FilterColorSpec | null): void => {
          const wb = deps.getWb?.();
          if (side === 'sort') {
            if (color && wb && deps.history) sortByColor(wb, deps.history, color);
          } else {
            filterUseWrite(dropdownFilter('filterByValue', wb), range, undefined, () => {
              if (color) {
                applyColorFilter(deps.store.getState(), deps.store, range, col, color, wb);
              } else {
                applyValueFilter(deps.store.getState(), deps.store, range, col, [], wb);
              }
            });
          }
          open(range, col, anchor);
        };
        const swatch = (label: string, pressed: boolean, color: FilterColorSpec) => {
          const button = createInteractionButton({
            className: 'fc-filter-dropdown__swatch',
            ariaLabel: label,
          });
          button.setAttribute('aria-pressed', pressed ? 'true' : 'false');
          button.addEventListener('click', () => applyColor(color));
          return button;
        };
        const renderSwatches = (): void => {
          swatches.replaceChildren();
          const kind = select.value;
          swatches.hidden = kind === '';
          if (kind === 'cellIcon') {
            for (const icon of choices.icons) {
              const pressed =
                active?.kind === 'cellIcon' &&
                active.icon.set === icon.set &&
                active.icon.slot === icon.slot;
              const button = swatch(`${icon.set} ${icon.slot + 1}`, pressed, {
                kind: 'cellIcon',
                icon,
              });
              button.dataset.fcIcon = `${icon.set}:${icon.slot}`;
              button.appendChild(iconCanvas(icon.set, icon.slot));
              swatches.appendChild(button);
            }
            return;
          }
          if (kind !== 'cellColor' && kind !== 'fontColor') return;
          for (const color of kind === 'cellColor' ? choices.fills : choices.fonts) {
            const pressed = active?.kind === kind && active.color === color;
            const button = swatch(color, pressed, { kind, color });
            button.dataset.fcColor = color;
            button.style.setProperty('--fc-swatch', color);
            swatches.appendChild(button);
          }
        };
        select.addEventListener('change', () => {
          if (select.value === '' && active && side === 'filter') applyColor(null);
          else renderSwatches();
        });
        if (side === 'sort') {
          projectDisabledState(
            select,
            !deps.getWb ||
              !deps.history ||
              !canExecuteBuiltIn(deps.store, 'sortAsc', 'contextMenu').allowed,
            m.unavailable,
          );
        }
        renderSwatches();
        wrap.append(row, swatches);
        return wrap;
      };
      const title = ownerDocument.createElement('div');
      title.className = 'fc-filter-dropdown__title';
      title.textContent = columnTitle(range, col);
      r.setAttribute('aria-label', title.textContent || t.title);
      const closeButton = createFilterDropdownActionButton('fc-filter-dropdown__close', '');
      closeButton.setAttribute('aria-label', m.close);
      closeButton.addEventListener('click', close);
      const section = (text: string): HTMLElement => {
        const label = ownerDocument.createElement('div');
        label.className = 'fc-filter-dropdown__section';
        label.textContent = text;
        return label;
      };
      const rule = ownerDocument.createElement('hr');
      rule.className = 'fc-filter-dropdown__rule';
      r.append(
        closeButton,
        title,
        section(m.sortSection),
        sortRow,
        colorPopup('sort'),
        rule,
        section(m.filterSection),
      );
      colorBy = colorPopup('filter');
    } else {
      actions.append(clear, apply);
    }

    if (colorBy) r.append(colorBy);
    let searchHost: HTMLElement = search;
    if (mac) {
      searchHost = ownerDocument.createElement('div');
      searchHost.className = 'fc-filter-dropdown__search-field';
      const icon = createExcelRibbonSvg('find', 'fc-filter-dropdown__search-icon');
      if (icon) searchHost.append(icon);
      searchHost.append(search);
    }
    r.append(conditionPanel, searchHost, list, actions);
    overlayPortalFor(deps.host).appendChild(r);
    root = r;
    placePanel(r, anchor.x, anchor.y + anchor.h, {
      fallbackWidth: 260,
      fallbackHeight: 320,
    });

    deferInitialFocus(r, () => search.focus());

    detachOutside = onOutsidePress(
      ownerDocument,
      () => [root],
      () => {
        if (!root) return;
        // The desktop panel is a floating window that only its close button dismisses.
        if (mac) return;
        close();
      },
      { event: 'mousedown' },
    );
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
