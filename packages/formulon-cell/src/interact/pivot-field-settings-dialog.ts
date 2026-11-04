import { formatCodeToNumFmt, numFmtToFormatCode } from '../engine/format-writeback.js';
import {
  PIVOT_SHOW_AS_BASE_NEXT,
  PIVOT_SHOW_AS_BASE_PREVIOUS,
  PivotAggregation,
  PivotShowValuesAs,
  type PivotShowValuesAs as PivotShowValuesAsValue,
} from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';
import { createDialogSelect } from '../toolbar/dialogs/form-controls.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import {
  appendDialogButton,
  appendDialogFrame,
  appendDialogOptionButton,
  appendDialogTabPair,
  createDialogShell,
  deferInitialFocus,
  focusAndSelectInput,
} from './dialog-shell.js';
import { attachFormatDialog, type FormatDialogHandle } from './format-dialog.js';
import { pivotShowAsNeedsBaseField, pivotShowAsNeedsBaseItem } from './pivot-field-settings.js';

/** Opens Format Cells on `current` and calls `apply` with the chosen code
 *  (empty for General) when the user confirms. */
export type PivotNumberFormatPicker = (current: string, apply: (code: string) => void) => void;

/** A picker backed by its own Format Cells instance in number-only mode, so
 *  the choice never touches the sheet. */
export function createPivotNumberFormatPicker(deps: {
  host: HTMLElement;
  store: SpreadsheetStore;
  getStrings: () => Strings;
  getWb: () => WorkbookHandle | null;
  getLocale?: () => string;
}): { pick: PivotNumberFormatPicker; detach(): void } {
  let dialog: FormatDialogHandle | null = null;
  return {
    pick: (current, apply) => {
      dialog ??= attachFormatDialog({
        host: deps.host,
        store: deps.store,
        strings: deps.getStrings(),
        getWb: deps.getWb,
        ...(deps.getLocale ? { getLocale: deps.getLocale } : {}),
      });
      const numFmt = formatCodeToNumFmt(current);
      dialog.open('number', {
        mode: 'numberFormat',
        initialFormat: numFmt ? { numFmt } : {},
        onApplyNumberFormat: (next) => apply(numFmtToFormatCode(next, deps.getWb()) ?? ''),
      });
    },
    detach: () => {
      dialog?.detach();
      dialog = null;
    },
  };
}

/** What the dialog shows for the field it was opened on. */
export type PivotFieldSettingsTarget =
  | {
      kind: 'value';
      sourceName: string;
      name: string;
      aggregation: PivotAggregation;
      numberFormat: string;
      showValuesAs: number;
      baseField?: number;
      baseItem?: number;
      /** The name the workbook profile gives each function, when known; the name
       *  field follows the function while it still shows one of these. */
      autoNames?: Partial<Record<PivotAggregation, string>>;
      /** Every pivot field, offered as the base of a calculation. */
      baseFields: readonly { index: number; name: string; items: readonly string[] }[];
    }
  | {
      kind: 'field';
      name: string;
      subtotal: 'auto' | 'custom' | 'none';
      subtotalFns: readonly PivotAggregation[];
    };

export type PivotFieldSettingsResult =
  | {
      kind: 'value';
      name: string;
      /** Write `name` as is. False for an untouched name after a function
       *  change the dialog could not rename for, so the engine names it. */
      keepName: boolean;
      aggregation: PivotAggregation;
      numberFormat: string;
      /** Running Total is reported as RunningTotalInRow; the caller picks the axis. */
      showValuesAs: number;
      baseField?: number;
      baseItem?: number;
    }
  | { kind: 'field'; subtotal: 'auto' | 'custom' | 'none'; subtotalFns: PivotAggregation[] }
  | { kind: 'hide' };

export interface PivotFieldSettingsDialogOptions {
  host: HTMLElement;
  strings: Strings;
  target: PivotFieldSettingsTarget;
  /** Opens Format Cells on the number tab; the Number Format button is
   *  disabled without it. */
  pickNumberFormat?: PivotNumberFormatPicker;
}

/** Functions in the order of the desktop list. */
type PivotDialogLabels = Strings['pivotTableDialog'];

const AGGREGATION_ORDER: readonly [PivotAggregation, keyof PivotDialogLabels][] = [
  [PivotAggregation.Sum, 'sum'],
  [PivotAggregation.Count, 'count'],
  [PivotAggregation.Average, 'average'],
  [PivotAggregation.Max, 'max'],
  [PivotAggregation.Min, 'min'],
  [PivotAggregation.Product, 'product'],
  [PivotAggregation.CountNumbers, 'countNumbers'],
  [PivotAggregation.StdDev, 'stdDev'],
  [PivotAggregation.StdDevP, 'stdDevP'],
  [PivotAggregation.Var, 'variance'],
  [PivotAggregation.VarP, 'varianceP'],
];

/** Calculations in the order of the desktop popup; Running Total stands for both
 *  engine running-total kinds. */
const SHOW_AS_ORDER: readonly [number, keyof PivotDialogLabels][] = [
  [PivotShowValuesAs.Normal, 'showValuesAsNormal'],
  [PivotShowValuesAs.DifferenceFrom, 'showValuesAsDifferenceFrom'],
  [PivotShowValuesAs.PercentOf, 'showValuesAsPercentOf'],
  [PivotShowValuesAs.PercentDifferenceFrom, 'showValuesAsPercentDifferenceFrom'],
  [PivotShowValuesAs.RunningTotalInRow, 'showValuesAsRunningTotal'],
  [PivotShowValuesAs.PercentOfRow, 'showValuesAsPercentOfRow'],
  [PivotShowValuesAs.PercentOfCol, 'showValuesAsPercentOfColumn'],
  [PivotShowValuesAs.PercentOfTotal, 'showValuesAsPercentOfTotal'],
  [PivotShowValuesAs.Index, 'showValuesAsIndex'],
  [PivotShowValuesAs.PercentOfParentRow, 'showValuesAsPercentOfParentRow'],
  [PivotShowValuesAs.PercentOfParentCol, 'showValuesAsPercentOfParentColumn'],
  [PivotShowValuesAs.PercentOfParent, 'showValuesAsPercentOfParentTotal'],
  [PivotShowValuesAs.PercentOfRunningTotal, 'showValuesAsPercentRunningTotal'],
  [PivotShowValuesAs.RankAscending, 'showValuesAsRankAscending'],
  [PivotShowValuesAs.RankDescending, 'showValuesAsRankDescending'],
];

/** Running Total also takes a base field: the axis it runs along. */
const usesBaseField = (kind: number): boolean =>
  kind === PivotShowValuesAs.RunningTotalInRow ||
  pivotShowAsNeedsBaseField(kind as PivotShowValuesAsValue);
const usesBaseItem = (kind: number): boolean =>
  pivotShowAsNeedsBaseItem(kind as PivotShowValuesAsValue);

/** A single-select list of option buttons: arrow keys move the selection. */
function createListbox(
  label: string,
  options: readonly { value: string; label: string }[],
  selected: string | null,
  onSelect: (value: string) => void,
): { root: HTMLDivElement; value(): string | null; setDisabled(disabled: boolean): void } {
  const root = document.createElement('div');
  root.className = 'fc-pvfsdlg__list';
  root.setAttribute('role', 'listbox');
  root.setAttribute('aria-label', label);
  root.tabIndex = 0;
  let current = selected;
  const buttons = options.map((option) =>
    appendDialogOptionButton(root, {
      label: option.label,
      baseClass: 'fc-pvfsdlg__option',
      datasetKey: 'pvfsValue',
      value: option.value,
      selected: option.value === selected,
    }),
  );
  const select = (value: string): void => {
    current = value;
    for (const button of buttons) {
      button.setAttribute('aria-selected', button.dataset.pvfsValue === value ? 'true' : 'false');
    }
    buttons
      .find((button) => button.dataset.pvfsValue === value)
      ?.scrollIntoView?.({
        block: 'nearest',
      });
    onSelect(value);
  };
  root.addEventListener('click', (event) => {
    if (root.getAttribute('aria-disabled') === 'true') return;
    const button = (event.target as Element | null)?.closest<HTMLElement>('[data-pvfs-value]');
    if (button?.dataset.pvfsValue !== undefined) select(button.dataset.pvfsValue);
  });
  root.addEventListener('keydown', (event) => {
    if (root.getAttribute('aria-disabled') === 'true') return;
    if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp') return;
    event.preventDefault();
    const at = buttons.findIndex((button) => button.dataset.pvfsValue === current);
    const next =
      buttons[Math.max(0, Math.min(buttons.length - 1, at + (event.key === 'ArrowDown' ? 1 : -1)))];
    if (next?.dataset.pvfsValue !== undefined) select(next.dataset.pvfsValue);
  });
  return {
    root,
    value: () => current,
    setDisabled: (disabled) => {
      projectDisabledState(root, disabled, null);
      root.tabIndex = disabled ? -1 : 0;
    },
  };
}

/** Pushes the buttons after it to the right edge of the footer. */
const footerSpacer = (): HTMLSpanElement => {
  const spacer = document.createElement('span');
  spacer.className = 'fc-pvfsdlg__spacer';
  return spacer;
};

const labelledRow = (label: string, control: HTMLElement): HTMLDivElement => {
  const row = document.createElement('div');
  row.className = 'fc-pvfsdlg__name-row';
  const text = document.createElement('label');
  text.className = 'fc-pvfsdlg__label';
  text.textContent = label;
  row.append(text, control);
  return row;
};

/** The PivotTable Field dialog. A value field gets Summarize by / Show data as
 *  tabs; a row, column or filter field gets its subtotals. Resolves null on
 *  cancel. */
export const showPivotFieldSettingsDialog = (
  opts: PivotFieldSettingsDialogOptions,
): Promise<PivotFieldSettingsResult | null> =>
  new Promise<PivotFieldSettingsResult | null>((resolve) => {
    const t = opts.strings.macPivotFieldSettings;
    const labels = opts.strings.pivotTableDialog;
    const { target } = opts;
    let done = false;
    const finish = (result: PivotFieldSettingsResult | null): void => {
      if (done) return;
      done = true;
      shell.dispose();
      resolve(result);
    };
    const shell = createDialogShell({
      host: opts.host,
      className: 'fc-pvfsdlg',
      ariaLabel: t.title,
      onDismiss: () => finish(null),
    });
    shell.overlay.classList.add('fc-fmtdlg');
    shell.overlay.dataset.pvfsKind = target.kind;
    const { body, footer } = appendDialogFrame(shell, {
      title: t.title,
      panelClasses: ['fc-fmtdlg__panel', 'fc-pvfsdlg__frame'],
      bodyClass: 'fc-fmtdlg__body fc-pvfsdlg__body',
      footerClass: 'fc-fmtdlg__footer fc-pvfsdlg__footer',
    });
    const nameInput = document.createElement('input');
    nameInput.type = 'text';
    nameInput.className = 'fc-pvfsdlg__name';
    nameInput.value = target.name;
    nameInput.setAttribute('aria-label', t.customName);
    nameInput.dataset.pvfsName = '';

    let collect: () => PivotFieldSettingsResult | null;
    let initialFocus: () => void;
    const primary: HTMLButtonElement[] = [];

    if (target.kind === 'value') {
      const source = document.createElement('span');
      source.className = 'fc-pvfsdlg__source';
      source.textContent = target.sourceName;
      source.dataset.pvfsSource = '';
      let nameEdited = false;
      let shownAggregation = target.aggregation;
      // Without the profile's names the field cannot follow the function, so an
      // untouched name is left for the engine to rename on OK.
      const namesKnown = Object.keys(target.autoNames ?? {}).length > 0;
      nameInput.addEventListener('input', () => {
        nameEdited = true;
      });
      body.append(labelledRow(t.sourceName, source), labelledRow(t.customName, nameInput));

      const tabs = document.createElement('div');
      tabs.className = 'fc-pvfsdlg__tabs';
      tabs.setAttribute('role', 'tablist');
      const panels = document.createElement('div');
      panels.className = 'fc-pvfsdlg__tab-panels';
      const summarize = appendDialogTabPair(tabs, panels, {
        id: 'summarize',
        label: t.summarizeTab,
        tabId: 'fc-pvfsdlg-tab-summarize',
        panelId: 'fc-pvfsdlg-panel-summarize',
        tabClass: 'fc-pvfsdlg__tab',
        panelClass: 'fc-pvfsdlg__panel-tab',
        tabDatasetKey: 'pvfsTab',
      });
      const showAs = appendDialogTabPair(tabs, panels, {
        id: 'showAs',
        label: t.showAsTab,
        tabId: 'fc-pvfsdlg-tab-show-as',
        panelId: 'fc-pvfsdlg-panel-show-as',
        tabClass: 'fc-pvfsdlg__tab',
        panelClass: 'fc-pvfsdlg__panel-tab',
        tabDatasetKey: 'pvfsTab',
      });
      const activate = (tab: 'summarize' | 'showAs'): void => {
        for (const [id, pair] of [
          ['summarize', summarize],
          ['showAs', showAs],
        ] as const) {
          const on = id === tab;
          pair.button.setAttribute('aria-selected', on ? 'true' : 'false');
          pair.button.tabIndex = on ? 0 : -1;
          pair.panel.hidden = !on;
        }
        shell.overlay.dataset.pvfsTab = tab;
      };
      summarize.button.addEventListener('click', () => activate('summarize'));
      showAs.button.addEventListener('click', () => activate('showAs'));
      const group = document.createElement('div');
      group.className = 'fc-pvfsdlg__group';
      group.append(tabs, panels);
      body.appendChild(group);

      const aggregation = createListbox(
        t.summarizeTab,
        AGGREGATION_ORDER.map(([value, key]) => ({
          value: String(value),
          label: String(labels[key]),
        })),
        String(target.aggregation),
        (value) => {
          const next = Number(value) as PivotAggregation;
          if (next === shownAggregation) return;
          shownAggregation = next;
          // An untouched name, stored caption or default alike, is regenerated
          // for the chosen function, as the desktop renames it on every change.
          const nextName = target.autoNames?.[next];
          if (!nameEdited && nextName) nameInput.value = nextName;
        },
      );
      aggregation.root.dataset.pvfsAggregation = '';
      summarize.panel.appendChild(aggregation.root);

      // The engine keeps a row or a column running total; both read as one.
      const initialShowAs =
        target.showValuesAs === PivotShowValuesAs.RunningTotalInCol
          ? PivotShowValuesAs.RunningTotalInRow
          : target.showValuesAs;
      const kind = createDialogSelect(
        SHOW_AS_ORDER.map(([value, key]) => ({ value: String(value), label: String(labels[key]) })),
        String(initialShowAs),
        { className: 'fc-fmtdlg__select fc-pvfsdlg__show-as' },
      );
      kind.dataset.pvfsShowAs = '';
      kind.setAttribute('aria-label', t.showAsTab);
      const baseFieldLabel = document.createElement('div');
      baseFieldLabel.className = 'fc-pvfsdlg__label';
      baseFieldLabel.textContent = t.baseField;
      const baseItemLabel = document.createElement('div');
      baseItemLabel.className = 'fc-pvfsdlg__label';
      baseItemLabel.textContent = t.baseItem;
      let baseItem: ReturnType<typeof createListbox> | null = null;
      const baseItemWrap = document.createElement('div');
      baseItemWrap.className = 'fc-pvfsdlg__base-item';
      const firstBase = target.baseFields[0]?.index;
      const baseField = createListbox(
        t.baseField,
        target.baseFields.map((field) => ({ value: String(field.index), label: field.name })),
        target.baseField !== undefined
          ? String(target.baseField)
          : firstBase !== undefined
            ? String(firstBase)
            : null,
        () => renderBaseItems(),
      );
      baseField.root.dataset.pvfsBaseField = '';
      const renderBaseItems = (): void => {
        const field = target.baseFields.find(
          (candidate) => String(candidate.index) === baseField.value(),
        );
        const usesItem = usesBaseItem(Number(kind.value));
        const items = usesItem
          ? [
              {
                value: String(PIVOT_SHOW_AS_BASE_PREVIOUS),
                label: labels.showValuesAsBaseItemPrevious,
              },
              { value: String(PIVOT_SHOW_AS_BASE_NEXT), label: labels.showValuesAsBaseItemNext },
              ...(field?.items ?? []).map((label, index) => ({ value: String(index), label })),
            ]
          : [];
        const keep =
          baseItem?.value() ??
          (target.baseItem !== undefined
            ? String(target.baseItem)
            : String(PIVOT_SHOW_AS_BASE_PREVIOUS));
        baseItem = createListbox(
          t.baseItem,
          items,
          items.some((item) => item.value === keep) ? keep : (items[0]?.value ?? null),
          () => {},
        );
        baseItem.root.dataset.pvfsBaseItem = '';
        baseItem.setDisabled(!usesItem);
        baseItemLabel.classList.toggle('fc-pvfsdlg__label--disabled', !usesItem);
        baseItemWrap.replaceChildren(baseItem.root);
      };
      const syncShowAs = (): void => {
        const usesField = usesBaseField(Number(kind.value));
        baseField.setDisabled(!usesField);
        baseFieldLabel.classList.toggle('fc-pvfsdlg__label--disabled', !usesField);
        renderBaseItems();
      };
      kind.addEventListener('change', syncShowAs);
      const bases = document.createElement('div');
      bases.className = 'fc-pvfsdlg__bases';
      const fieldColumn = document.createElement('div');
      fieldColumn.append(baseFieldLabel, baseField.root);
      const itemColumn = document.createElement('div');
      itemColumn.append(baseItemLabel, baseItemWrap);
      bases.append(fieldColumn, itemColumn);
      showAs.panel.append(kind, bases);
      syncShowAs();
      activate('summarize');

      let numberFormat = target.numberFormat;
      const formatBtn = appendDialogButton(footer, {
        label: t.numberFormat,
        baseClass: 'fc-fmtdlg__btn',
      });
      formatBtn.classList.add('fc-pvfsdlg__format');
      formatBtn.dataset.pvfsNumberFormat = '';
      footer.appendChild(footerSpacer());
      const pick = opts.pickNumberFormat;
      if (pick) {
        formatBtn.addEventListener('click', () => {
          pick(numberFormat, (code) => {
            numberFormat = code;
          });
        });
      } else {
        projectDisabledState(formatBtn, true, t.unavailable, { datasetKey: 'disabledReason' });
      }
      collect = () => {
        const showValuesAs = Number(kind.value);
        const field = usesBaseField(showValuesAs) ? baseField.value() : null;
        const item = usesBaseItem(showValuesAs) ? baseItem?.value() : null;
        return {
          kind: 'value',
          name: nameInput.value,
          keepName:
            namesKnown ||
            nameEdited ||
            Number(aggregation.value() ?? target.aggregation) === target.aggregation,
          aggregation: Number(aggregation.value() ?? target.aggregation) as PivotAggregation,
          numberFormat,
          showValuesAs,
          ...(field !== null && field !== undefined ? { baseField: Number(field) } : {}),
          ...(item !== null && item !== undefined ? { baseItem: Number(item) } : {}),
        };
      };
      initialFocus = () => focusAndSelectInput(nameInput);
    } else {
      // The engine has no setter for a field's caption: the name stays focusable
      // and selectable, as on the desktop, but read-only.
      nameInput.readOnly = true;
      nameInput.setAttribute('aria-readonly', 'true');
      nameInput.title = t.unavailable;
      body.appendChild(labelledRow(t.customName, nameInput));
      const heading = document.createElement('div');
      heading.className = 'fc-pvfsdlg__heading';
      heading.textContent = t.subtotals;
      const radios = document.createElement('div');
      radios.className = 'fc-pvfsdlg__radios';
      radios.setAttribute('role', 'radiogroup');
      radios.setAttribute('aria-label', t.subtotals);
      const radio = (value: 'auto' | 'custom' | 'none', text: string): HTMLInputElement => {
        const label = document.createElement('label');
        label.className = 'fc-pvfsdlg__radio';
        const input = document.createElement('input');
        input.type = 'radio';
        input.name = 'fc-pvfsdlg-subtotal';
        input.value = value;
        input.checked = target.subtotal === value;
        input.dataset.pvfsSubtotal = value;
        label.append(input, document.createTextNode(text));
        radios.appendChild(label);
        return input;
      };
      const auto = radio('auto', t.subtotalAutomatic);
      const custom = radio('custom', t.subtotalCustom);
      const none = radio('none', t.subtotalNone);
      // Turning the automatic subtotal off has no engine setter.
      projectDisabledState(none, target.subtotal !== 'none', t.unavailable, {
        datasetKey: 'disabledReason',
      });
      const chosen = new Set<number>(target.subtotalFns);
      const fnList = document.createElement('div');
      fnList.className = 'fc-pvfsdlg__list fc-pvfsdlg__fn-list';
      fnList.setAttribute('role', 'listbox');
      fnList.setAttribute('aria-multiselectable', 'true');
      fnList.setAttribute('aria-label', t.subtotals);
      fnList.dataset.pvfsSubtotalFns = '';
      const fnButtons = AGGREGATION_ORDER.map(([value, key]) =>
        appendDialogOptionButton(fnList, {
          label: String(labels[key]),
          baseClass: 'fc-pvfsdlg__option',
          datasetKey: 'pvfsValue',
          value: String(value),
          selected: chosen.has(value),
        }),
      );
      const ok = (): HTMLButtonElement | undefined => primary[0];
      const sync = (): void => {
        for (const button of fnButtons) {
          button.setAttribute(
            'aria-selected',
            chosen.has(Number(button.dataset.pvfsValue)) ? 'true' : 'false',
          );
        }
        const button = ok();
        if (button) projectDisabledState(button, custom.checked && chosen.size === 0, null);
      };
      fnList.addEventListener('click', (event) => {
        const button = (event.target as Element | null)?.closest<HTMLElement>('[data-pvfs-value]');
        if (!button) return;
        const value = Number(button.dataset.pvfsValue);
        // Picking a function is choosing Custom, as on the desktop.
        custom.checked = true;
        if (chosen.has(value)) chosen.delete(value);
        else chosen.add(value);
        sync();
      });
      for (const input of [auto, custom, none]) input.addEventListener('change', sync);
      const subtotals = document.createElement('div');
      subtotals.className = 'fc-pvfsdlg__subtotals';
      subtotals.append(radios, fnList);
      const noData = document.createElement('input');
      noData.type = 'checkbox';
      noData.dataset.pvfsShowNoData = '';
      // Showing items with no data has no engine setter.
      projectDisabledState(noData, true, t.unavailable, { datasetKey: 'disabledReason' });
      const noDataLabel = document.createElement('label');
      noDataLabel.className = 'fc-pvfsdlg__check';
      noDataLabel.append(noData, document.createTextNode(t.showItemsWithNoData));
      body.append(heading, subtotals, noDataLabel);
      footer.appendChild(footerSpacer());
      const hide = appendDialogButton(footer, { label: t.hide, baseClass: 'fc-fmtdlg__btn' });
      hide.dataset.pvfsHide = '';
      hide.addEventListener('click', () => finish({ kind: 'hide' }));
      collect = () => ({
        kind: 'field',
        subtotal: auto.checked ? 'auto' : custom.checked ? 'custom' : 'none',
        subtotalFns: [...chosen].sort((a, b) => a - b) as PivotAggregation[],
      });
      initialFocus = () => focusAndSelectInput(nameInput);
      queueMicrotask(sync);
    }

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
    primary.push(ok);
    ok.dataset.pvfsOk = '';
    cancel.dataset.pvfsCancel = '';
    shell.on(cancel, 'click', () => finish(null));
    shell.on(ok, 'click', () => {
      if (ok.disabled) return;
      finish(collect());
    });
    shell.on(shell.overlay, 'keydown', (event) => {
      const key = (event as KeyboardEvent).key;
      if (key !== 'Enter' || (event.target as Element | null)?.closest('button, select')) return;
      event.preventDefault();
      ok.click();
    });
    shell.open();
    // The shell focuses its first control; the desktop opens on the name, selected.
    initialFocus();
    deferInitialFocus(shell.overlay, () => initialFocus());
  });
