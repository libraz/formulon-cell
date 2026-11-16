import type { PivotCacheOptions, PivotOptions } from '@libraz/formulon';
import { pivotReportEffects, pivotWrite } from '../commands/pivot-actions.js';
import type { Strings } from '../i18n/strings.js';
import { createDialogSelect } from '../toolbar/dialogs/form-controls.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import {
  appendDialogButton,
  appendDialogFrame,
  appendDialogTabPair,
  createDialogShell,
  deferInitialFocus,
  focusAndSelectInput,
} from './dialog-shell.js';
import type { PivotCommitContext } from './pivot-field-settings-dialog-apply.js';

/** The report options the dialog edits. */
export type PivotOptionsEdit = Pick<
  PivotOptions,
  | 'showDrill'
  | 'printDrill'
  | 'showDataTips'
  | 'showHeaders'
  | 'gridDropZones'
  | 'showError'
  | 'errorCaption'
  | 'showMissing'
  | 'missingCaption'
  | 'showEmptyRow'
  | 'showEmptyColumn'
  | 'itemPrintTitles'
  | 'fieldPrintTitles'
  | 'multipleFieldFilters'
  | 'autoFormatting'
  | 'preserveFormatting'
  | 'mergeItem'
  | 'indent'
  | 'pageOverThenDown'
  | 'pageWrap'
  | 'customListSort'
  | 'enableDrill'
  | 'altText'
  | 'altTextSummary'
>;

/** The options of the report's cache the dialog edits. */
export type PivotCacheOptionsEdit = Pick<
  PivotCacheOptions,
  'saveData' | 'refreshOnLoad' | 'missingItemsLimit'
>;

/** The report settings the dialog edits. */
export interface PivotOptionsValues {
  name: string;
  options: PivotOptionsEdit;
  cache: PivotCacheOptionsEdit;
}

export interface PivotOptionsDialogOptions extends PivotOptionsValues {
  host: HTMLElement;
  strings: Strings;
}

export type PivotOptionsResult = PivotOptionsValues;

type TabId = 'display' | 'layout' | 'data' | 'altText';
type OptionKey = keyof PivotOptionsEdit;
type CacheKey = keyof PivotCacheOptionsEdit;
type CaptionKey = 'errorCaption' | 'missingCaption';

/** The largest compact-form indent the engine stores; the dialog shows each value plus one. */
const MAX_INDENT = 126;
/** The largest number of report filter fields per column or row. */
const MAX_PAGE_WRAP = 255;
/** The largest number of retained deleted items per field the engine stores. */
const MAX_MISSING_ITEMS = 1048576;

/** One control row of a tab: a check box, a radio group, a popup or a text field.
 *  A row bound to an option or a cache option edits it; an unbound row renders the
 *  desktop default and stays disabled. */
type OptionRow =
  | {
      kind: 'check';
      label: string;
      checked: boolean;
      indent?: boolean;
      option?: OptionKey;
      cache?: CacheKey;
      caption?: { key: CaptionKey; value: string | null };
    }
  | {
      kind: 'radios';
      label?: string;
      options: readonly string[];
      selected: number;
      option?: 'pageOverThenDown';
      /** The Automatic / Max pair of the retained-items limit; it follows the
       *  preceding `missingItemsLimit` check box. */
      limit?: boolean;
    }
  | {
      kind: 'popup';
      label: string;
      value: number;
      option?: { key: 'indent' | 'pageWrap'; offset: number; max: number };
    };

interface OptionGroup {
  label: string;
  rows: readonly OptionRow[];
}

const disable = (el: HTMLElement & { disabled?: boolean }, reason: string): void => {
  projectDisabledState(el, true, reason, { datasetKey: 'disabledReason' });
};

const range = (from: number, to: number): number[] =>
  Array.from({ length: to - from + 1 }, (_, index) => from + index);

function renderGroups(
  panel: HTMLElement,
  groups: readonly (OptionGroup | 'rule')[],
  reason: string,
): void {
  let retainBox: HTMLInputElement | null = null;
  for (const group of groups) {
    if (group === 'rule') {
      const rule = document.createElement('hr');
      rule.className = 'fc-pvoptdlg__rule';
      panel.appendChild(rule);
      continue;
    }
    const section = document.createElement('div');
    section.className = 'fc-pvoptdlg__group';
    const label = document.createElement('div');
    label.className = 'fc-pvoptdlg__group-label';
    label.textContent = group.label;
    const controls = document.createElement('div');
    controls.className = 'fc-pvoptdlg__controls';
    for (const row of group.rows) {
      if (row.kind === 'check') {
        const line = document.createElement('label');
        line.className = `fc-pvoptdlg__check${row.indent ? ' fc-pvoptdlg__check--indent' : ''}`;
        const box = document.createElement('input');
        box.type = 'checkbox';
        box.checked = row.checked;
        if (row.option) box.dataset.pvoptOption = row.option;
        else if (row.cache) box.dataset.pvoptCache = row.cache;
        else disable(box, reason);
        if (row.cache === 'missingItemsLimit') retainBox = box;
        line.append(box, document.createTextNode(row.label));
        if (row.caption) {
          const text = document.createElement('input');
          text.type = 'text';
          text.className = 'fc-pvoptdlg__field';
          text.setAttribute('aria-label', row.label);
          text.value = row.caption.value ?? '';
          text.dataset.pvoptOption = row.caption.key;
          // The caption applies while its check box is on, as on the desktop.
          const follow = (): void => projectDisabledState(text, !box.checked, null);
          follow();
          box.addEventListener('change', follow);
          line.appendChild(text);
        }
        controls.appendChild(line);
      } else if (row.kind === 'radios') {
        const wrap = document.createElement('div');
        wrap.className = 'fc-pvoptdlg__radio-row';
        if (row.label) {
          const caption = document.createElement('span');
          caption.className = 'fc-pvoptdlg__caption';
          caption.textContent = row.label;
          wrap.appendChild(caption);
        }
        const list = document.createElement('div');
        list.className = 'fc-pvoptdlg__radios';
        list.setAttribute('role', 'radiogroup');
        if (row.label) list.setAttribute('aria-label', row.label);
        if (row.option) list.dataset.pvoptOption = row.option;
        if (row.limit) list.dataset.pvoptLimit = '';
        const groupName = `fc-pvoptdlg-${row.option ?? row.options.join('|')}`;
        const radios: HTMLInputElement[] = [];
        for (const [index, option] of row.options.entries()) {
          const line = document.createElement('label');
          line.className = 'fc-pvoptdlg__radio';
          const radio = document.createElement('input');
          radio.type = 'radio';
          radio.name = groupName;
          radio.value = String(index);
          radio.checked = index === row.selected;
          if (!row.option && !row.limit) disable(radio, reason);
          radios.push(radio);
          line.append(radio, document.createTextNode(option));
          list.appendChild(line);
        }
        const box = row.limit ? retainBox : null;
        if (box) {
          // Off, the limit is 0 and the pair is disabled with Max shown, as on the desktop.
          const follow = (): void => {
            if (!box.checked && radios[1]) radios[1].checked = true;
            for (const radio of radios) projectDisabledState(radio, !box.checked, null);
          };
          follow();
          box.addEventListener('change', follow);
        }
        wrap.appendChild(list);
        controls.appendChild(wrap);
      } else {
        const line = document.createElement('label');
        line.className = 'fc-pvoptdlg__popup-row';
        const caption = document.createElement('span');
        caption.className = 'fc-pvoptdlg__caption';
        caption.textContent = row.label;
        const bind = row.option;
        const values = bind ? range(bind.offset, bind.max + bind.offset) : [row.value];
        const shown = String(row.value);
        const select = createDialogSelect(
          values.map((value) => ({ value: String(value), label: String(value) })),
          shown,
          { className: 'fc-fmtdlg__select fc-pvoptdlg__popup' },
        );
        select.setAttribute('aria-label', row.label);
        if (bind) select.dataset.pvoptOption = bind.key;
        else disable(select, reason);
        line.append(caption, select);
        controls.appendChild(line);
      }
    }
    section.append(label, controls);
    panel.appendChild(section);
  }
}

/** Reads the bound controls of the dialog back into option values. */
function readControls(
  root: HTMLElement,
  initial: PivotOptionsValues,
): Pick<PivotOptionsValues, 'options' | 'cache'> {
  const options: Record<string, unknown> = { ...initial.options };
  const cache: Record<string, unknown> = { ...initial.cache };
  for (const el of root.querySelectorAll<HTMLElement>('[data-pvopt-option]')) {
    const key = el.dataset.pvoptOption as OptionKey;
    if (el instanceof HTMLInputElement && el.type === 'checkbox') options[key] = el.checked;
    else if (el instanceof HTMLSelectElement)
      options[key] = Number(el.value) - (key === 'indent' ? 1 : 0);
    else if (el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement)
      options[key] = el.value === '' ? null : el.value;
    else options[key] = el.querySelector<HTMLInputElement>('input:checked')?.value === '1';
  }
  for (const el of root.querySelectorAll<HTMLInputElement>('[data-pvopt-cache]')) {
    const key = el.dataset.pvoptCache as CacheKey;
    if (key !== 'missingItemsLimit') {
      cache[key] = el.checked;
      continue;
    }
    const max =
      root.querySelector<HTMLInputElement>('[data-pvopt-limit] input:checked')?.value === '1';
    cache[key] = !el.checked ? 0 : max ? MAX_MISSING_ITEMS : null;
  }
  return {
    options: options as unknown as PivotOptionsEdit,
    cache: cache as unknown as PivotCacheOptionsEdit,
  };
}

/** The PivotTable Options dialog; resolves null on cancel. */
export const showPivotOptionsDialog = (
  opts: PivotOptionsDialogOptions,
): Promise<PivotOptionsResult | null> =>
  new Promise<PivotOptionsResult | null>((resolve) => {
    const t = opts.strings.mac.pivotOptions;
    let done = false;
    const finish = (result: PivotOptionsResult | null): void => {
      if (done) return;
      done = true;
      shell.dispose();
      resolve(result);
    };
    const shell = createDialogShell({
      host: opts.host,
      className: 'fc-pvoptdlg',
      ariaLabel: t.title,
      onDismiss: () => finish(null),
    });
    shell.overlay.classList.add('fc-fmtdlg');
    const { body, footer } = appendDialogFrame(shell, {
      title: t.title,
      panelClasses: ['fc-fmtdlg__panel', 'fc-pvoptdlg__frame'],
      bodyClass: 'fc-fmtdlg__body fc-pvoptdlg__body',
      footerClass: 'fc-fmtdlg__footer fc-pvoptdlg__footer',
    });
    const nameRow = document.createElement('label');
    nameRow.className = 'fc-pvoptdlg__name-row';
    const nameLabel = document.createElement('span');
    nameLabel.textContent = t.name;
    const name = document.createElement('input');
    name.type = 'text';
    name.className = 'fc-pvoptdlg__name';
    name.value = opts.name;
    name.dataset.pvoptName = '';
    nameRow.append(nameLabel, name);
    body.appendChild(nameRow);

    const group = document.createElement('div');
    group.className = 'fc-pvoptdlg__tab-group';
    const tabs = document.createElement('div');
    tabs.className = 'fc-pvoptdlg__tabs';
    tabs.setAttribute('role', 'tablist');
    const panels = document.createElement('div');
    panels.className = 'fc-pvoptdlg__panels';
    group.append(tabs, panels);
    body.appendChild(group);
    const pairs = new Map<TabId, ReturnType<typeof appendDialogTabPair>>();
    for (const [id, label] of [
      ['display', t.tabDisplay],
      ['layout', t.tabLayout],
      ['data', t.tabData],
      ['altText', t.tabAltText],
    ] as const) {
      pairs.set(
        id,
        appendDialogTabPair(tabs, panels, {
          id,
          label,
          tabId: `fc-pvoptdlg-tab-${id}`,
          panelId: `fc-pvoptdlg-panel-${id}`,
          tabClass: 'fc-pvfsdlg__tab',
          panelClass: 'fc-pvoptdlg__panel-tab',
          tabDatasetKey: 'pvoptTab',
          panelDatasetKey: 'pvoptPanel',
        }),
      );
    }
    const activate = (id: TabId): void => {
      for (const [key, pair] of pairs) {
        const on = key === id;
        pair.button.setAttribute('aria-selected', on ? 'true' : 'false');
        pair.button.tabIndex = on ? 0 : -1;
        pair.panel.hidden = !on;
      }
    };
    for (const [id, pair] of pairs) pair.button.addEventListener('click', () => activate(id));

    const reason = t.unavailable;
    const o = opts.options;
    const d = t.display;
    const display = pairs.get('display')?.panel;
    if (display) {
      renderGroups(
        display,
        [
          {
            label: d.group,
            rows: [
              { kind: 'check', label: d.expandButtons, checked: o.showDrill, option: 'showDrill' },
              {
                kind: 'check',
                label: d.printExpandButtons,
                checked: o.printDrill,
                indent: true,
                option: 'printDrill',
              },
              {
                kind: 'check',
                label: d.contextualTooltips,
                checked: o.showDataTips,
                option: 'showDataTips',
              },
              {
                kind: 'check',
                label: d.fieldCaptions,
                checked: o.showHeaders,
                option: 'showHeaders',
              },
              {
                kind: 'check',
                label: d.classicLayout,
                checked: o.gridDropZones,
                option: 'gridDropZones',
              },
              {
                kind: 'check',
                label: d.errorValues,
                checked: o.showError,
                option: 'showError',
                caption: { key: 'errorCaption', value: o.errorCaption },
              },
              {
                kind: 'check',
                label: d.emptyCells,
                checked: o.showMissing,
                option: 'showMissing',
                caption: { key: 'missingCaption', value: o.missingCaption },
              },
            ],
          },
          'rule',
          {
            label: d.noDataGroup,
            rows: [
              {
                kind: 'check',
                label: d.noDataRows,
                checked: o.showEmptyRow,
              },
              {
                kind: 'check',
                label: d.noDataColumns,
                checked: o.showEmptyColumn,
              },
            ],
          },
          'rule',
          {
            label: d.printGroup,
            rows: [
              {
                kind: 'check',
                label: d.repeatRowLabels,
                checked: o.itemPrintTitles,
                option: 'itemPrintTitles',
              },
              {
                kind: 'check',
                label: d.printTitles,
                checked: o.fieldPrintTitles,
                option: 'fieldPrintTitles',
              },
              { kind: 'check', label: d.itemLabels, checked: true },
            ],
          },
        ],
        reason,
      );
    }
    const l = t.layout;
    const layout = pairs.get('layout')?.panel;
    if (layout) {
      renderGroups(
        layout,
        [
          {
            label: l.group,
            rows: [
              {
                kind: 'check',
                label: l.multipleFilters,
                checked: o.multipleFieldFilters,
                option: 'multipleFieldFilters',
              },
              {
                kind: 'check',
                label: l.autofitColumns,
                checked: o.autoFormatting,
                option: 'autoFormatting',
              },
              {
                kind: 'check',
                label: l.preserveFormatting,
                checked: o.preserveFormatting,
                option: 'preserveFormatting',
              },
              { kind: 'check', label: l.mergeLabels, checked: o.mergeItem, option: 'mergeItem' },
              {
                kind: 'popup',
                label: l.compactIndent,
                value: o.indent + 1,
                option: { key: 'indent', offset: 1, max: MAX_INDENT },
              },
            ],
          },
          'rule',
          {
            label: l.reportFilterGroup,
            rows: [
              {
                kind: 'radios',
                label: l.fieldOrder,
                options: [l.downThenOver, l.overThenDown],
                selected: o.pageOverThenDown ? 1 : 0,
                option: 'pageOverThenDown',
              },
              {
                kind: 'popup',
                label: l.fieldsPerColumn,
                value: o.pageWrap,
                option: { key: 'pageWrap', offset: 0, max: MAX_PAGE_WRAP },
              },
            ],
          },
          'rule',
          {
            label: l.sortGroup,
            rows: [
              {
                kind: 'check',
                label: l.customLists,
                checked: o.customListSort,
                option: 'customListSort',
              },
            ],
          },
        ],
        reason,
      );
    }
    const dt = t.data;
    const data = pairs.get('data')?.panel;
    const limit = opts.cache.missingItemsLimit;
    if (data) {
      renderGroups(
        data,
        [
          {
            label: dt.group,
            rows: [
              {
                kind: 'check',
                label: dt.saveSourceData,
                checked: opts.cache.saveData,
                cache: 'saveData',
              },
              {
                kind: 'check',
                label: dt.enableShowDetails,
                checked: o.enableDrill,
                option: 'enableDrill',
              },
              {
                kind: 'check',
                label: dt.retainDeletedItems,
                checked: limit !== 0,
                cache: 'missingItemsLimit',
              },
              {
                kind: 'radios',
                label: dt.itemsToRetain,
                options: [dt.retainAutomatic, dt.retainMax],
                selected: limit === null ? 0 : 1,
                limit: true,
              },
              {
                kind: 'check',
                label: dt.refreshOnOpen,
                checked: opts.cache.refreshOnLoad,
                cache: 'refreshOnLoad',
              },
              { kind: 'check', label: dt.refreshOnChange, checked: false },
            ],
          },
          'rule',
          {
            label: dt.connectionGroup,
            rows: [
              { kind: 'check', label: dt.savePassword, checked: false },
              { kind: 'check', label: dt.backgroundQuery, checked: false },
            ],
          },
        ],
        reason,
      );
    }
    const altText = pairs.get('altText')?.panel;
    if (altText) {
      const a = t.altText;
      const titleLabel = document.createElement('label');
      titleLabel.className = 'fc-pvoptdlg__alt-label';
      titleLabel.textContent = a.title;
      const title = document.createElement('input');
      title.type = 'text';
      title.className = 'fc-pvoptdlg__alt-title';
      title.setAttribute('aria-label', a.title);
      title.value = o.altText ?? '';
      title.dataset.pvoptOption = 'altText';
      const descLabel = document.createElement('label');
      descLabel.className = 'fc-pvoptdlg__alt-label';
      descLabel.textContent = a.description;
      const desc = document.createElement('textarea');
      desc.className = 'fc-pvoptdlg__alt-description';
      desc.setAttribute('aria-label', a.description);
      desc.value = o.altTextSummary ?? '';
      desc.dataset.pvoptOption = 'altTextSummary';
      const help = document.createElement('p');
      help.className = 'fc-pvoptdlg__alt-help';
      help.textContent = a.help;
      altText.append(titleLabel, title, descLabel, desc, help);
    }
    activate('display');

    const spacer = document.createElement('span');
    spacer.className = 'fc-pvfsdlg__spacer';
    footer.appendChild(spacer);
    const cancel = appendDialogButton(footer, {
      label: t.cancel,
      variant: 'secondary',
      baseClass: 'fc-fmtdlg__btn',
    });
    const ok = appendDialogButton(footer, {
      label: t.ok,
      variant: 'primary',
      baseClass: 'fc-fmtdlg__btn',
    });
    cancel.dataset.pvoptCancel = '';
    ok.dataset.pvoptOk = '';
    shell.on(cancel, 'click', () => finish(null));
    shell.on(ok, 'click', () => finish({ name: name.value, ...readControls(body, opts) }));
    shell.on(name, 'keydown', (event) => {
      if ((event as KeyboardEvent).key !== 'Enter') return;
      event.preventDefault();
      ok.click();
    });
    shell.open();
    focusAndSelectInput(name);
    deferInitialFocus(shell.overlay, () => focusAndSelectInput(name));
  });

/** The members of `next` that differ from `from`. */
const changed = <T extends object>(from: T, next: T): Partial<T> =>
  Object.fromEntries(
    Object.entries(next).filter(([key, value]) => from[key as keyof T] !== value),
  ) as Partial<T>;

const isEmpty = (patch: object): boolean => Object.keys(patch).length === 0;

/** Applies the name and the options as one undo step; an empty name keeps the
 *  current one, and nothing changed does nothing. */
export function applyPivotOptions(
  ctx: PivotCommitContext,
  sheetIndex: number,
  pivotIndex: number,
  previous: PivotOptionsValues,
  result: PivotOptionsResult,
): boolean {
  const next: PivotOptionsValues = {
    name: result.name.trim() || previous.name,
    options: result.options,
    cache: result.cache,
  };
  if (
    next.name === previous.name &&
    isEmpty(changed(previous.options, next.options)) &&
    isEmpty(changed(previous.cache, next.cache))
  )
    return true;
  return pivotWrite(
    ctx,
    ctx.history,
    pivotReportEffects(ctx.wb, sheetIndex, pivotIndex),
    false,
    () => writePivotOptions(ctx, sheetIndex, pivotIndex, previous, next),
  );
}

const writePivotOptions = (
  ctx: PivotCommitContext,
  sheetIndex: number,
  pivotIndex: number,
  previous: PivotOptionsValues,
  next: PivotOptionsValues,
): boolean => {
  const { wb } = ctx;
  // Only a changed setting is written: the engine refuses a rename onto the current name.
  const move = (from: PivotOptionsValues, to: PivotOptionsValues): boolean => {
    const options = changed(from.options, to.options);
    const cache = changed(from.cache, to.cache);
    return (
      (from.name === to.name || wb.renamePivotTable(sheetIndex, pivotIndex, to.name)) &&
      (isEmpty(options) || wb.updatePivotOptions(sheetIndex, pivotIndex, options)) &&
      (isEmpty(cache) || wb.updatePivotCacheOptions(sheetIndex, pivotIndex, cache))
    );
  };
  if (!move(previous, next)) return false;
  const history = ctx.history && !ctx.history.isReplaying() ? ctx.history : null;
  history?.push({
    undo: () => {
      move(next, previous);
      ctx.onChanged();
    },
    redo: () => {
      move(previous, next);
      ctx.onChanged();
    },
  });
  ctx.onChanged();
  return true;
};

/** The settings the dialog edits, read from the report; null when the report cannot be read. */
function readPivotOptions(
  ctx: PivotCommitContext,
  sheetIndex: number,
  pivotIndex: number,
): PivotOptionsValues | null {
  const options = ctx.wb.pivotOptions(sheetIndex, pivotIndex);
  const cache = ctx.wb.pivotCacheOptions(sheetIndex, pivotIndex);
  if (!options || !cache) return null;
  return {
    name: ctx.wb.pivotTableName(sheetIndex, pivotIndex) ?? '',
    options: {
      showDrill: options.showDrill,
      printDrill: options.printDrill,
      showDataTips: options.showDataTips,
      showHeaders: options.showHeaders,
      gridDropZones: options.gridDropZones,
      showError: options.showError,
      errorCaption: options.errorCaption,
      showMissing: options.showMissing,
      missingCaption: options.missingCaption,
      showEmptyRow: options.showEmptyRow,
      showEmptyColumn: options.showEmptyColumn,
      itemPrintTitles: options.itemPrintTitles,
      fieldPrintTitles: options.fieldPrintTitles,
      multipleFieldFilters: options.multipleFieldFilters,
      autoFormatting: options.autoFormatting,
      preserveFormatting: options.preserveFormatting,
      mergeItem: options.mergeItem,
      indent: options.indent,
      pageOverThenDown: options.pageOverThenDown,
      pageWrap: options.pageWrap,
      customListSort: options.customListSort,
      enableDrill: options.enableDrill,
      altText: options.altText,
      altTextSummary: options.altTextSummary,
    },
    cache: {
      saveData: cache.saveData,
      missingItemsLimit: cache.missingItemsLimit,
      refreshOnLoad: cache.refreshOnLoad,
    },
  };
}

/** The report options the PivotTable Analyze Show group turns on and off. */
export type PivotShowOption = 'showDrill' | 'showHeaders';

/** Turns one Show option of a report over as one undo step; false when the report cannot be read
 *  or the write is refused. */
export function togglePivotShowOption(
  ctx: PivotCommitContext,
  sheetIndex: number,
  pivotIndex: number,
  option: PivotShowOption,
): boolean {
  const previous = readPivotOptions(ctx, sheetIndex, pivotIndex);
  if (!previous) return false;
  return applyPivotOptions(ctx, sheetIndex, pivotIndex, previous, {
    ...previous,
    options: { ...previous.options, [option]: !previous.options[option] },
  });
}

/** Opens the dialog on a report and applies the result; false when the report cannot be read. */
export function openPivotOptions(
  ctx: PivotCommitContext & { host: HTMLElement; strings: Strings },
  sheetIndex: number,
  pivotIndex: number,
): Promise<boolean> {
  const previous = readPivotOptions(ctx, sheetIndex, pivotIndex);
  if (!previous) return Promise.resolve(false);
  return showPivotOptionsDialog({ host: ctx.host, strings: ctx.strings, ...previous }).then(
    (result) => (result ? applyPivotOptions(ctx, sheetIndex, pivotIndex, previous, result) : false),
  );
}
