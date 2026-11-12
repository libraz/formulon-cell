import { DATE_PERIOD_GROUP_STARTS } from '../commands/date-periods.js';
import {
  clearPivotFieldFilter,
  type PivotFieldFilterModel,
  type PivotFilterItem,
  pivotFieldFilterModel,
  setPivotFieldCondition,
  setPivotFieldItemsChecked,
} from '../commands/pivot-filter.js';
import { pivotStyleLayer } from '../engine/pivot-style.js';
import { type Addr, PivotAxis, PivotFilterKind, PivotSortOrder } from '../engine/types.js';
import { isMacHost } from '../extensions/ui-options.js';
import type { Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';
import { createExcelRibbonSvg } from '../toolbar/excel-ribbon-icons.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { createInteractionButton } from './chip-button.js';
import { enhanceCustomSelect } from './custom-select.js';
import { deferInitialFocus } from './dialog-shell.js';
import { onOutsidePress } from './outside-press.js';
import { overlayPortalFor } from './overlay-portal.js';
import { placePanel } from './overlay-position.js';
import {
  pivotFilterConditionToSpec,
  pivotFilterInputFor,
  pivotFilterSpecToCondition,
} from './pivot-field-settings.js';
import { commitPivotChange, type PivotCommitContext } from './pivot-field-settings-dialog-apply.js';
import { showPivotFilterDialog } from './pivot-filter-dialog.js';
import { PIVOT_FILTER_PERIODS, pivotFilterOperand } from './pivot-filter-kinds.js';

/**
 * The filter panel a report's drop-down button opens: sort buttons, label and
 * value filter pickers, a searchable item checklist and Clear Filter. A
 * checkbox applies at once, as one undo step; there is no OK. A label or value
 * filter is entered in the PivotTable filter dialog and applied the same way.
 */

/** Everything the panel needs to open on a report. */
export interface PivotFilterPanelContext extends PivotCommitContext {
  host: HTMLElement;
  store: SpreadsheetStore;
  strings: Strings;
}

type PanelStrings = Strings['workbookObjects']['pivotFilterPanel'];
type ConditionKey = keyof PanelStrings['conditions'];

/** The desktop's label and value filter lists; null marks a separator. Every
 *  entry opens the dialog for its condition kind. */
const LABEL_CONDITIONS: readonly (readonly [ConditionKey, PivotFilterKind] | null)[] = [
  null,
  ['equals', PivotFilterKind.LabelEqual],
  ['notEquals', PivotFilterKind.LabelNotEqual],
  null,
  ['beginsWith', PivotFilterKind.LabelBeginsWith],
  ['notBeginsWith', PivotFilterKind.LabelNotBeginsWith],
  ['endsWith', PivotFilterKind.LabelEndsWith],
  ['notEndsWith', PivotFilterKind.LabelNotEndsWith],
  null,
  ['contains', PivotFilterKind.LabelContains],
  ['notContains', PivotFilterKind.LabelNotContains],
  null,
  ['greater', PivotFilterKind.LabelGreaterThan],
  ['greaterOrEqual', PivotFilterKind.LabelGreaterThanOrEqual],
  ['less', PivotFilterKind.LabelLessThan],
  ['lessOrEqual', PivotFilterKind.LabelLessThanOrEqual],
  null,
  ['between', PivotFilterKind.LabelBetween],
  ['notBetween', PivotFilterKind.LabelNotBetween],
];
const VALUE_CONDITIONS: readonly (readonly [ConditionKey, PivotFilterKind] | null)[] = [
  null,
  ['equals', PivotFilterKind.ValueEqual],
  ['notEquals', PivotFilterKind.ValueNotEqual],
  null,
  ['greater', PivotFilterKind.ValueGreaterThan],
  ['greaterOrEqual', PivotFilterKind.ValueGreaterThanOrEqual],
  ['less', PivotFilterKind.ValueLessThan],
  ['lessOrEqual', PivotFilterKind.ValueLessThanOrEqual],
  null,
  ['between', PivotFilterKind.ValueBetween],
  ['notBetween', PivotFilterKind.ValueNotBetween],
  null,
  ['top10', PivotFilterKind.ValueTopCount],
  ['bottom10', PivotFilterKind.ValueBottomCount],
];

/** A date field's label filters: the desktop's Date Filters comparisons. */
const DATE_CONDITIONS: readonly (readonly [ConditionKey, PivotFilterKind] | null)[] = [
  null,
  ['equals', PivotFilterKind.DateEqual],
  ['before', PivotFilterKind.DateOlderThan],
  ['after', PivotFilterKind.DateNewerThan],
  ['between', PivotFilterKind.DateBetween],
];
const open = new WeakMap<HTMLElement, () => void>();

/** The report field the cell at `addr` opens: the field behind a drop-down
 *  button, else the row, column or filter field of a label cell. */
function fieldAt(
  ctx: PivotFilterPanelContext,
  addr: Addr,
): { pivotIndex: number; fieldIndex: number } | null {
  const button = pivotStyleLayer(ctx.wb, ctx.store.getState()).buttonAt(addr.row, addr.col);
  if (button && button.fieldIndex >= 0) return button;
  const info = ctx.wb.pivotCellInfo(addr);
  if (!info || info.isValue || info.fieldIndex < 0) return null;
  const definition = ctx.wb.getPivotDefinition(addr.sheet, info.pivotIndex);
  const field = definition?.fields.find(
    (f) => f.cacheFieldIndex === info.fieldIndex && f.axis !== PivotAxis.Value,
  );
  return field ? { pivotIndex: info.pivotIndex, fieldIndex: field.index } : null;
}

/** Open the filter panel of the field at `addr` (a drop-down button or a
 *  label cell), hanging from `anchor` in viewport coordinates. False when the
 *  cell belongs to no report field. */
export function openPivotFieldFilter(
  ctx: PivotFilterPanelContext,
  addr: Addr,
  anchor?: { x: number; y: number },
): boolean {
  const target = fieldAt(ctx, addr);
  if (!target) return false;
  open.get(ctx.host)?.();
  const { wb, strings } = ctx;
  const t = strings.workbookObjects.pivotFilterPanel;
  const sheet = addr.sheet;
  const { pivotIndex, fieldIndex } = target;
  const doc = ctx.host.ownerDocument;
  const read = (): PivotFieldFilterModel | null =>
    pivotFieldFilterModel(wb, sheet, pivotIndex, fieldIndex);
  let model = read();
  if (!model) return false;
  const input = pivotFilterInputFor(wb);

  const panel = doc.createElement('div');
  panel.className = 'fc-pivot-filter-panel';
  panel.setAttribute('role', 'dialog');
  panel.setAttribute('aria-modal', 'false');
  panel.setAttribute('aria-label', model.label);
  panel.style.position = 'fixed';

  const commit = (apply: () => boolean): void => {
    commitPivotChange(ctx, sheet, pivotIndex, apply);
    model = read() ?? model;
    render();
  };

  const titleBar = doc.createElement('div');
  titleBar.className = 'fc-pivot-filter-panel__title';
  const closeBtn = createInteractionButton(
    { className: 'fc-pivot-filter-panel__close', ariaLabel: t.close },
    doc,
  );
  const title = doc.createElement('span');
  title.textContent = model.label;
  titleBar.append(closeBtn, title);

  const heading = (text: string): HTMLElement => {
    const h = doc.createElement('div');
    h.className = 'fc-pivot-filter-panel__heading';
    h.textContent = text;
    return h;
  };
  const row = (label: string, control: HTMLSelectElement): HTMLElement => {
    control.setAttribute('aria-label', label.replace(/[:：]\s*$/, ''));
    const r = doc.createElement('div');
    r.className = 'fc-pivot-filter-panel__row';
    const span = doc.createElement('span');
    span.textContent = label;
    r.append(span, control);
    return r;
  };
  const selectHandles = new Map<HTMLSelectElement, { sync(): void; dispose(): void }>();
  const select = (): HTMLSelectElement => {
    const s = doc.createElement('select');
    s.className = 'fc-pivot-filter-panel__select';
    return s;
  };

  const sortBy = select();
  sortBy.dataset.pivotFilterSortBy = 'true';
  const sortButton = (label: string, aria: string, order: PivotSortOrder): HTMLButtonElement => {
    const b = createInteractionButton({ className: 'fc-pivot-filter-panel__sort' });
    const glyph = createExcelRibbonSvg(
      order === PivotSortOrder.Ascending ? 'sortAscSmall' : 'sortDescSmall',
      'fc-pivot-filter-panel__sort-icon',
    );
    const text = doc.createElement('span');
    text.textContent = label;
    if (glyph) b.append(glyph);
    b.append(text);
    b.setAttribute('aria-label', aria);
    b.dataset.pivotFilterSort = order === PivotSortOrder.Ascending ? 'asc' : 'desc';
    b.addEventListener('click', () =>
      commit(() => wb.setPivotFieldSort(sheet, pivotIndex, fieldIndex, order, sortBy.value)),
    );
    return b;
  };
  const sortRow = doc.createElement('div');
  sortRow.className = 'fc-pivot-filter-panel__sort-row';
  sortRow.append(
    sortButton(t.sortAscending, t.sortAscendingAria, PivotSortOrder.Ascending),
    sortButton(t.sortDescending, t.sortDescendingAria, PivotSortOrder.Descending),
  );

  const conditionSelect = (
    list: readonly (readonly [string, PivotFilterKind] | null)[],
    which: 'label' | 'value',
  ): HTMLSelectElement => {
    const s = select();
    s.dataset.pivotFilterBy = which;
    for (const entry of [['', null] as const, ...list]) {
      const option = doc.createElement('option');
      if (entry === null) {
        projectDisabledState(option, true, null);
        option.textContent = '──────────';
      } else if (entry[1] === null) {
        option.value = '';
        option.textContent = t.chooseOne;
      } else {
        option.value = String(entry[1]);
        option.textContent = entry[0];
      }
      s.appendChild(option);
    }
    s.addEventListener('change', () => {
      const picked = s.value;
      s.value = '';
      selectHandles.get(s)?.sync();
      if (!picked || !model) return;
      const kind = Number(picked) as PivotFilterKind;
      // A relative period takes no operand and applies at once.
      if (pivotFilterOperand(kind) === 'none') {
        const spec = { axis: model.axis ?? PivotAxis.Row, fieldName: model.sourceName, kind };
        commit(() => setPivotFieldCondition(wb, sheet, pivotIndex, fieldIndex, spec));
        return;
      }
      const current = pivotFilterSpecToCondition(model.condition ?? undefined, input);
      void showPivotFilterDialog({
        host: ctx.host,
        strings: strings.pivotTableDialog,
        fieldName: model.label,
        condition: { kind, value: current?.kind === kind ? current.value : '' },
        okLabel: strings.pageSetup.ok,
        cancelLabel: strings.pageSetup.cancel,
        dataFields: model.dataFields,
        mac: isMacHost(ctx.host),
      }).then((condition) => {
        if (!condition || !model) return;
        const spec = pivotFilterConditionToSpec(
          model.sourceName,
          condition,
          model.axis ?? PivotAxis.Row,
          input,
        );
        if (spec) commit(() => setPivotFieldCondition(wb, sheet, pivotIndex, fieldIndex, spec));
      });
    });
    return s;
  };

  const labeled = (list: typeof LABEL_CONDITIONS) =>
    list.map((entry) => (entry ? ([t.conditions[entry[0]], entry[1]] as const) : null));
  const dateConditions = () => [
    ...labeled(DATE_CONDITIONS),
    ...PIVOT_FILTER_PERIODS.flatMap(([kind, key], index) => [
      ...(DATE_PERIOD_GROUP_STARTS.has(index) ? [null] : []),
      [strings.datePeriods[key], kind] as const,
    ]),
  ];

  const search = doc.createElement('input');
  search.type = 'search';
  search.className = 'fc-pivot-filter-panel__search';
  search.placeholder = t.search;
  search.setAttribute('aria-label', t.search);
  const searchBox = doc.createElement('div');
  searchBox.className = 'fc-pivot-filter-panel__search-box';
  const lens = createExcelRibbonSvg('find', 'fc-pivot-filter-panel__search-icon');
  if (lens) searchBox.append(lens);
  searchBox.append(search);
  const list = doc.createElement('div');
  list.className = 'fc-pivot-filter-panel__list';
  list.setAttribute('role', 'group');
  const clear = createInteractionButton({
    className: 'fc-pivot-filter-panel__clear',
    text: t.clearFilter,
  });
  clear.addEventListener('click', () =>
    commit(() => clearPivotFieldFilter(wb, sheet, pivotIndex, fieldIndex)),
  );

  const checkbox = (label: string, checked: boolean, onChange: (on: boolean) => void) => {
    const l = doc.createElement('label');
    l.className = 'fc-pivot-filter-panel__item';
    const box = doc.createElement('input');
    box.type = 'checkbox';
    box.checked = checked;
    box.addEventListener('change', () => onChange(box.checked));
    const span = doc.createElement('span');
    span.textContent = label;
    l.append(box, span);
    return { l, box };
  };

  const applyItems = (items: PivotFilterItem[]): void =>
    commit(() => setPivotFieldItemsChecked(wb, sheet, pivotIndex, fieldIndex, items));

  const render = (): void => {
    if (!model) return;
    sortBy.replaceChildren(
      ...[['', model.label] as const, ...model.dataFields.map((d) => [d, d] as const)].map(
        ([value, label]) => {
          const o = doc.createElement('option');
          o.value = value;
          o.textContent = label;
          return o;
        },
      ),
    );
    sortBy.value = model.sortByField;
    selectHandles.get(sortBy)?.sync();
    const needle = search.value.trim().toLocaleLowerCase();
    const shown = model.items.filter(
      (i) => !needle || i.value.toLocaleLowerCase().includes(needle),
    );
    const all = shown.length > 0 && shown.every((i) => i.checked);
    const none = shown.every((i) => !i.checked);
    const items = model.items;
    const selectAll = checkbox(t.selectAll, all, (on) => {
      const hit = new Set(shown.map((i) => i.value));
      applyItems(items.map((i) => (hit.has(i.value) ? { ...i, checked: on } : i)));
    });
    selectAll.box.indeterminate = !all && !none;
    selectAll.box.dataset.pivotFilterSelectAll = 'true';
    list.replaceChildren(
      selectAll.l,
      ...shown.map((item) => {
        const entry = checkbox(item.value === '' ? t.blank : item.value, item.checked, (on) =>
          applyItems(items.map((i) => (i === item ? { ...i, checked: on } : i))),
        );
        entry.box.dataset.pivotFilterItem = item.value;
        return entry.l;
      }),
    );
    projectDisabledState(clear, !model.filtered, null);
  };
  search.addEventListener('input', render);

  const separator = doc.createElement('hr');
  separator.className = 'fc-pivot-filter-panel__separator';
  const footer = doc.createElement('div');
  footer.className = 'fc-pivot-filter-panel__footer';
  footer.appendChild(clear);
  panel.append(
    titleBar,
    heading(t.sortHeading),
    sortRow,
    row(t.sortBy, sortBy),
    separator,
    heading(t.filterHeading),
    row(
      t.byLabel,
      conditionSelect(model.dateField ? dateConditions() : labeled(LABEL_CONDITIONS), 'label'),
    ),
    row(t.byValue, conditionSelect(labeled(VALUE_CONDITIONS), 'value')),
    searchBox,
    list,
    footer,
  );
  // The desktop's pop-up buttons: the shared custom select over each native one.
  for (const s of panel.querySelectorAll<HTMLSelectElement>('select')) {
    const handle = enhanceCustomSelect(s);
    if (handle) selectHandles.set(s, handle);
  }
  render();

  const close = (): void => {
    for (const handle of selectHandles.values()) handle.dispose();
    panel.remove();
    detachOutside();
    doc.removeEventListener('keydown', onDocKey, true);
    if (open.get(ctx.host) === close) open.delete(ctx.host);
  };
  const onDocKey = (e: KeyboardEvent): void => {
    if (e.key === 'Escape' && panel.contains(e.target as Node)) {
      e.preventDefault();
      close();
    }
  };
  closeBtn.addEventListener('click', close);

  overlayPortalFor(ctx.host).appendChild(panel);
  const rect = ctx.host.getBoundingClientRect();
  const at = anchor ?? { x: rect.left + 40, y: rect.top + 40 };
  // The desktop hangs the panel 6 pt below the button's bottom-left corner.
  placePanel(panel, at.x, at.y + 6, {
    fallbackWidth: 257,
    fallbackHeight: 434,
  });
  deferInitialFocus(panel, () => search.focus({ preventScroll: true }));
  const detachOutside = onOutsidePress(
    doc,
    () => [panel],
    (e) => {
      // The filter dialog the panel opened sits over it; leave the panel be.
      if ((e.target as Element | null)?.closest?.('.fc-pivotdlg--filter')) return;
      close();
    },
    { event: 'mousedown' },
  );
  doc.addEventListener('keydown', onDocKey, true);
  open.set(ctx.host, close);
  return true;
}
