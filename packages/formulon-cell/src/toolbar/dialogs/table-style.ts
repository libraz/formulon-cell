import type { StripeElement } from '../../commands/format-as-table.js';
import type {
  StyleElementFormat,
  StyleElementFormats,
  TableStyleElementType,
} from '../../engine/table-style-elements.js';
import type { Strings } from '../../i18n/strings.js';
import type { CellFormat } from '../../store/types.js';
import { projectDisabledState } from '../menu-a11y.js';
import { tableStyleSwatchGrid } from '../ribbon/menus/styles.js';
import { appendCheckboxRow, appendSelectRow } from './form-controls.js';
import {
  appendDialogActions,
  appendDialogButton,
  appendDialogNameField,
  appendErrorRow,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
} from './shell.js';

export type TableStyleKind = 'table' | 'pivot';
type ElementLabel = keyof Strings['ribbonMenu']['tableStyleElements'];

/** The elements the dialog lists for a table style, in the desktop order, and the label each
 *  shows. The Mac order is measured down to First Column; the rest follows the desktop list. */
const TABLE_ELEMENTS: readonly (readonly [TableStyleElementType, ElementLabel])[] = [
  ['wholeTable', 'wholeTable'],
  ['firstColumnStripe', 'firstColumnStripe'],
  ['secondColumnStripe', 'secondColumnStripe'],
  ['firstRowStripe', 'firstRowStripe'],
  ['secondRowStripe', 'secondRowStripe'],
  ['lastColumn', 'lastColumn'],
  ['firstColumn', 'firstColumn'],
  ['headerRow', 'headerRow'],
  ['totalRow', 'totalRow'],
  ['firstHeaderCell', 'firstHeaderCell'],
  ['lastHeaderCell', 'lastHeaderCell'],
  ['firstTotalCell', 'firstTotalCell'],
  ['lastTotalCell', 'lastTotalCell'],
];

/** The PivotTable style elements, in the desktop order; unmeasured on the Mac app. */
const PIVOT_ELEMENTS: readonly (readonly [TableStyleElementType, ElementLabel])[] = [
  ['wholeTable', 'wholeTable'],
  ['pageFieldLabels', 'pageFieldLabels'],
  ['pageFieldValues', 'pageFieldValues'],
  ['firstColumnStripe', 'firstColumnStripe'],
  ['secondColumnStripe', 'secondColumnStripe'],
  ['firstRowStripe', 'firstRowStripe'],
  ['secondRowStripe', 'secondRowStripe'],
  ['firstColumn', 'firstColumn'],
  ['headerRow', 'headerRow'],
  ['firstHeaderCell', 'firstHeaderCell'],
  ['firstSubtotalColumn', 'firstSubtotalColumn'],
  ['secondSubtotalColumn', 'secondSubtotalColumn'],
  ['thirdSubtotalColumn', 'thirdSubtotalColumn'],
  ['firstSubtotalRow', 'firstSubtotalRow'],
  ['secondSubtotalRow', 'secondSubtotalRow'],
  ['thirdSubtotalRow', 'thirdSubtotalRow'],
  ['blankRow', 'blankRow'],
  ['firstColumnSubheading', 'firstColumnSubheading'],
  ['secondColumnSubheading', 'secondColumnSubheading'],
  ['thirdColumnSubheading', 'thirdColumnSubheading'],
  ['firstRowSubheading', 'firstRowSubheading'],
  ['secondRowSubheading', 'secondRowSubheading'],
  ['thirdRowSubheading', 'thirdRowSubheading'],
  ['lastColumn', 'grandTotalColumn'],
  ['totalRow', 'grandTotalRow'],
];

const STRIPES: ReadonlySet<TableStyleElementType> = new Set<StripeElement>([
  'firstRowStripe',
  'secondRowStripe',
  'firstColumnStripe',
  'secondColumnStripe',
]);
const STRIPE_SIZES = [1, 2, 3, 4, 5, 6, 7, 8, 9];

export interface TableStyleDialogValue {
  name: string;
  elements: StyleElementFormats;
  stripeSizes: Partial<Record<StripeElement, number>>;
  /** Set as the workbook's default style of its kind. */
  setAsDefault: boolean;
}

export interface TableStyleDialogOptions {
  kind: TableStyleKind;
  mode: 'new' | 'modify';
  /** The Mac host: its own titles and labels, and no PivotTable default checkbox. */
  mac: boolean;
  strings: Strings;
  initial: Omit<TableStyleDialogValue, 'setAsDefault'> & { setAsDefault?: boolean };
  /** Opens the differential Format Cells for one element; `apply` receives the new format. */
  editFormat: (initial: Partial<CellFormat>, apply: (format: Partial<CellFormat>) => void) => void;
  /** Describes an element's format for the Element Formatting line. */
  describe: (format: Partial<CellFormat>) => string;
}

const hasFormat = (format: StyleElementFormat | undefined): boolean =>
  !!format && Object.keys(format).length > 0;

/** The New / Modify Table Style and PivotTable Style dialog: a name, the style's elements with
 *  Format and Clear per element, the stripe size, a preview, and the default checkbox. */
export const showTableStyleDialog = (
  opts: TableStyleDialogOptions,
): Promise<TableStyleDialogValue | null> =>
  new Promise<TableStyleDialogValue | null>((resolve) => {
    const { strings, mac, kind } = opts;
    const t = strings.ribbonMenu;
    const m = strings.mac.tableStyleDialog;
    const title =
      kind === 'table'
        ? opts.mode === 'new'
          ? mac
            ? m.newTitle
            : t.tableStyleNewTitle
          : mac
            ? m.modifyTitle
            : t.tableStyleModifyTitle
        : opts.mode === 'new'
          ? mac
            ? m.newPivotTitle
            : t.tableStyleNewPivotTitle
          : mac
            ? m.modifyPivotTitle
            : t.tableStyleModifyPivotTitle;
    const elements: StyleElementFormats = structuredClone(opts.initial.elements);
    const stripeSizes: Partial<Record<StripeElement, number>> = { ...opts.initial.stripeSizes };
    const list = kind === 'table' ? TABLE_ELEMENTS : PIVOT_ELEMENTS;
    let selected: TableStyleElementType = list[0]?.[0] ?? 'wholeTable';

    const shell = createDialogShell({ title });
    shell.panel.classList.add('fc-tablestyle-dlg');
    const name = appendDialogNameField(
      shell.body,
      mac ? m.name : t.tableStyleName,
      opts.initial.name,
    );

    const columns = document.createElement('div');
    columns.className = 'fc-tablestyle-dlg__columns';
    const left = document.createElement('div');
    left.className = 'fc-tablestyle-dlg__elements';
    const listLabel = document.createElement('div');
    listLabel.className = 'fc-tablestyle-dlg__label';
    listLabel.id = 'fc-tablestyle-dlg-elements';
    listLabel.textContent = t.tableStyleElement;
    const listbox = document.createElement('div');
    listbox.className = 'fc-tablestyle-dlg__list';
    listbox.setAttribute('role', 'listbox');
    listbox.setAttribute('aria-labelledby', listLabel.id);
    listbox.tabIndex = 0;
    const buttons = document.createElement('div');
    buttons.className = 'fc-tablestyle-dlg__buttons';
    const format = appendDialogButton(buttons, {
      label: mac ? m.format : t.tableStyleFormat,
      variant: 'secondary',
    });
    const clear = appendDialogButton(buttons, {
      label: mac ? m.clear : t.tableStyleClearElement,
      variant: 'secondary',
    });
    left.append(listLabel, listbox, buttons);
    const stripe = appendSelectRow(
      left,
      mac ? m.stripeSize : t.tableStyleStripeSize,
      STRIPE_SIZES.map((n) => ({ value: String(n), label: String(n) })),
      '1',
      'stripeSize',
    );

    const right = document.createElement('div');
    right.className = 'fc-tablestyle-dlg__preview';
    const previewLabel = document.createElement('div');
    previewLabel.className = 'fc-tablestyle-dlg__label';
    previewLabel.textContent = mac ? m.preview : t.tableStylePreview;
    const previewBox = document.createElement('div');
    previewBox.className = 'fc-tablestyle-dlg__preview-box';
    right.append(previewLabel, previewBox);
    columns.append(left, right);
    shell.body.appendChild(columns);

    // The Mac dialog has no Element Formatting line.
    const describeRow = document.createElement('div');
    describeRow.className = 'fc-tablestyle-dlg__describe';
    describeRow.hidden = mac;
    shell.body.appendChild(describeRow);

    // The Mac PivotTable Quick Style dialog has no default checkbox.
    const setDefault =
      kind === 'pivot' && mac
        ? null
        : appendCheckboxRow(
            shell.body,
            kind === 'table'
              ? mac
                ? m.setDefault
                : t.tableStyleSetDefaultTable
              : t.tableStyleSetDefaultPivot,
            opts.initial.setAsDefault === true,
            'setAsDefault',
          );

    const errorRow = appendErrorRow(shell.body);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: strings.hyperlinkDialog.cancel,
      okLabel: strings.hyperlinkDialog.ok,
    });

    const render = (): void => {
      listbox.replaceChildren(
        ...list.map(([type, label]) => {
          const option = document.createElement('div');
          option.className = 'fc-tablestyle-dlg__option';
          option.setAttribute('role', 'option');
          option.dataset.element = type;
          option.textContent = t.tableStyleElements[label];
          option.setAttribute('aria-selected', String(type === selected));
          option.classList.toggle(
            'fc-tablestyle-dlg__option--formatted',
            hasFormat(elements[type]),
          );
          return option;
        }),
      );
      const current = elements[selected];
      projectDisabledState(clear, !hasFormat(current), null);
      const isStripe = STRIPES.has(selected);
      projectDisabledState(stripe, !isStripe, null);
      stripe.value = String(isStripe ? (stripeSizes[selected as StripeElement] ?? 1) : 1);
      describeRow.textContent = hasFormat(current)
        ? `${t.tableStyleElementFormatting} ${opts.describe(current ?? {})}`
        : '';
      previewBox.replaceChildren(tableStyleSwatchGrid(elements));
    };

    const select = (type: TableStyleElementType): void => {
      selected = type;
      render();
      listbox
        .querySelector<HTMLElement>(`[data-element="${type}"]`)
        ?.scrollIntoView({ block: 'nearest' });
    };
    listbox.addEventListener('click', (event) => {
      const option = (event.target as HTMLElement).closest<HTMLElement>('[role="option"]');
      const type = option?.dataset.element as TableStyleElementType | undefined;
      if (type) select(type);
    });
    listbox.addEventListener('keydown', (event) => {
      if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp') return;
      event.preventDefault();
      const at = list.findIndex(([type]) => type === selected);
      const next =
        list[Math.min(list.length - 1, Math.max(0, at + (event.key === 'ArrowDown' ? 1 : -1)))];
      if (next) select(next[0]);
    });
    format.addEventListener('click', () => {
      const element = selected;
      opts.editFormat({ ...(elements[element] ?? {}) }, (next) => {
        elements[element] = next as StyleElementFormat;
        render();
      });
    });
    clear.addEventListener('click', () => {
      delete elements[selected];
      render();
    });
    stripe.addEventListener('change', () => {
      if (!STRIPES.has(selected)) return;
      const size = Number(stripe.value);
      if (size > 1) stripeSizes[selected as StripeElement] = size;
      else delete stripeSizes[selected as StripeElement];
    });

    const lifecycle = installDialogLifecycle<TableStyleDialogValue | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
    });
    const onOk = (): TableStyleDialogValue | null => {
      const label = name.valueOrError(errorRow, strings.namedRangeDialog.errorEmptyName);
      if (label === null) return null;
      const value: TableStyleDialogValue = {
        name: label,
        elements,
        stripeSizes,
        setAsDefault: setDefault?.checked === true,
      };
      lifecycle.finish(value);
      return value;
    };
    okBtn.addEventListener('click', () => {
      onOk();
    });
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    render();
    mountDialog(shell, name.focus);
  });
