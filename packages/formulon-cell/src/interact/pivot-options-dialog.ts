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

export interface PivotOptionsDialogOptions {
  host: HTMLElement;
  strings: Strings;
  /** The report's current name. */
  name: string;
  /** Whether the report's columns autofit when it updates. */
  autofitColumns: boolean;
}

export interface PivotOptionsResult {
  name: string;
  autofitColumns: boolean;
}

/** The report settings the dialog edits. */
export interface PivotOptionsValues {
  name: string;
  autofitColumns: boolean;
}

type TabId = 'display' | 'layout' | 'data' | 'altText';

/** One control row of a tab: a check box, a radio group, a popup or a text field.
 *  Only the name and the autofit check box have an engine setter; every other
 *  option renders with the desktop default and stays disabled. */
type OptionRow =
  | {
      kind: 'check';
      label: string;
      checked: boolean;
      indent?: boolean;
      field?: boolean;
      autofit?: boolean;
    }
  | { kind: 'radios'; label?: string; options: readonly string[]; selected: number }
  | { kind: 'popup'; label: string; value: string }
  | { kind: 'rule' };

interface OptionGroup {
  label: string;
  rows: readonly OptionRow[];
}

const disable = (el: HTMLElement & { disabled?: boolean }, reason: string): void => {
  projectDisabledState(el, true, reason, { datasetKey: 'disabledReason' });
};

function renderGroups(
  panel: HTMLElement,
  groups: readonly (OptionGroup | 'rule')[],
  reason: string,
): void {
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
      if (row.kind === 'rule') continue;
      if (row.kind === 'check') {
        const line = document.createElement('label');
        line.className = `fc-pvoptdlg__check${row.indent ? ' fc-pvoptdlg__check--indent' : ''}`;
        const box = document.createElement('input');
        box.type = 'checkbox';
        box.checked = row.checked;
        if (row.autofit) box.dataset.pvoptAutofit = '';
        else disable(box, reason);
        line.append(box, document.createTextNode(row.label));
        if (row.field) {
          const text = document.createElement('input');
          text.type = 'text';
          text.className = 'fc-pvoptdlg__field';
          text.setAttribute('aria-label', row.label);
          disable(text, reason);
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
        for (const [index, option] of row.options.entries()) {
          const line = document.createElement('label');
          line.className = 'fc-pvoptdlg__radio';
          const radio = document.createElement('input');
          radio.type = 'radio';
          radio.checked = index === row.selected;
          disable(radio, reason);
          line.append(radio, document.createTextNode(option));
          list.appendChild(line);
        }
        wrap.appendChild(list);
        controls.appendChild(wrap);
      } else {
        const line = document.createElement('label');
        line.className = 'fc-pvoptdlg__popup-row';
        const caption = document.createElement('span');
        caption.className = 'fc-pvoptdlg__caption';
        caption.textContent = row.label;
        const select = createDialogSelect([{ value: row.value, label: row.value }], row.value, {
          className: 'fc-fmtdlg__select fc-pvoptdlg__popup',
        });
        select.setAttribute('aria-label', row.label);
        disable(select, reason);
        line.append(caption, select);
        controls.appendChild(line);
      }
    }
    section.append(label, controls);
    panel.appendChild(section);
  }
}

/** The PivotTable Options dialog. Only the name and the autofit option apply; resolves null on cancel. */
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
    const d = t.display;
    const display = pairs.get('display')?.panel;
    if (display) {
      renderGroups(
        display,
        [
          {
            label: d.group,
            rows: [
              { kind: 'check', label: d.expandButtons, checked: true },
              { kind: 'check', label: d.printExpandButtons, checked: false, indent: true },
              { kind: 'check', label: d.contextualTooltips, checked: true },
              { kind: 'check', label: d.fieldCaptions, checked: true },
              { kind: 'check', label: d.classicLayout, checked: false },
              { kind: 'check', label: d.errorValues, checked: false, field: true },
              { kind: 'check', label: d.emptyCells, checked: true, field: true },
            ],
          },
          'rule',
          {
            label: d.noDataGroup,
            rows: [
              { kind: 'check', label: d.noDataRows, checked: false },
              { kind: 'check', label: d.noDataColumns, checked: false },
            ],
          },
          'rule',
          {
            label: d.printGroup,
            rows: [
              { kind: 'check', label: d.repeatRowLabels, checked: true },
              { kind: 'check', label: d.printTitles, checked: false },
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
              { kind: 'check', label: l.multipleFilters, checked: false },
              {
                kind: 'check',
                label: l.autofitColumns,
                checked: opts.autofitColumns,
                autofit: true,
              },
              { kind: 'check', label: l.preserveFormatting, checked: true },
              { kind: 'check', label: l.mergeLabels, checked: false },
              { kind: 'popup', label: l.compactIndent, value: '1' },
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
                selected: 0,
              },
              { kind: 'popup', label: l.fieldsPerColumn, value: '0' },
            ],
          },
          'rule',
          {
            label: l.sortGroup,
            rows: [
              { kind: 'check', label: l.customLists, checked: true },
              { kind: 'radios', options: [l.sortAscending, l.sortDataSource], selected: 1 },
            ],
          },
        ],
        reason,
      );
    }
    const dt = t.data;
    const data = pairs.get('data')?.panel;
    if (data) {
      renderGroups(
        data,
        [
          {
            label: dt.group,
            rows: [
              { kind: 'check', label: dt.saveSourceData, checked: true },
              { kind: 'check', label: dt.enableShowDetails, checked: true },
              { kind: 'check', label: dt.retainDeletedItems, checked: true },
              {
                kind: 'radios',
                label: dt.itemsToRetain,
                options: [dt.retainAutomatic, dt.retainMax],
                selected: 0,
              },
              { kind: 'check', label: dt.refreshOnOpen, checked: false },
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
      disable(title, reason);
      const descLabel = document.createElement('label');
      descLabel.className = 'fc-pvoptdlg__alt-label';
      descLabel.textContent = a.description;
      const desc = document.createElement('textarea');
      desc.className = 'fc-pvoptdlg__alt-description';
      desc.setAttribute('aria-label', a.description);
      disable(desc, reason);
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
    const autofit = layout?.querySelector<HTMLInputElement>('[data-pvopt-autofit]');
    shell.on(ok, 'click', () =>
      finish({ name: name.value, autofitColumns: autofit?.checked ?? opts.autofitColumns }),
    );
    shell.on(name, 'keydown', (event) => {
      if ((event as KeyboardEvent).key !== 'Enter') return;
      event.preventDefault();
      ok.click();
    });
    shell.open();
    focusAndSelectInput(name);
    deferInitialFocus(shell.overlay, () => focusAndSelectInput(name));
  });

/** Applies the name and autofit option as one undo step; an empty or unchanged name and
 *  an unchanged option do nothing. */
export function applyPivotOptions(
  ctx: PivotCommitContext,
  sheetIndex: number,
  pivotIndex: number,
  previous: PivotOptionsValues,
  result: PivotOptionsResult,
): boolean {
  const next: PivotOptionsValues = {
    name: result.name.trim() || previous.name,
    autofitColumns: result.autofitColumns,
  };
  if (next.name === previous.name && next.autofitColumns === previous.autofitColumns) return true;
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
  const move = (from: PivotOptionsValues, to: PivotOptionsValues): boolean =>
    (from.name === to.name || wb.renamePivotTable(sheetIndex, pivotIndex, to.name)) &&
    (from.autofitColumns === to.autofitColumns ||
      wb.setPivotAutofitColumns(sheetIndex, pivotIndex, to.autofitColumns));
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

/** Opens the dialog on a report and applies the result. */
export function openPivotOptions(
  ctx: PivotCommitContext & { host: HTMLElement; strings: Strings },
  sheetIndex: number,
  pivotIndex: number,
): Promise<boolean> {
  const previous: PivotOptionsValues = {
    name: ctx.wb.pivotTableName(sheetIndex, pivotIndex) ?? '',
    autofitColumns: ctx.wb.pivotAutofitColumns(sheetIndex, pivotIndex),
  };
  return showPivotOptionsDialog({ host: ctx.host, strings: ctx.strings, ...previous }).then(
    (result) => (result ? applyPivotOptions(ctx, sheetIndex, pivotIndex, previous, result) : false),
  );
}
