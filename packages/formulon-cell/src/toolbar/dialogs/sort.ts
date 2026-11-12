import type { ColumnColorChoices } from '../../commands/filter-colors.js';
import type { MacSortDialogLabels } from '../../i18n/strings/mac/sort-dialog.js';
import { conditionalIconCanvas } from '../../render/painters/cf-icons.js';
import type { ConditionalIconRef } from '../../store/types/conditional-rule.js';
import { projectDisabledState } from '../menu-a11y.js';
import { appendDialogSelectOptions, createDialogSelect } from './form-controls.js';
import {
  appendDialogActions,
  appendDialogButton,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
} from './shell.js';

export interface SortDialogColumn {
  value: string;
  label: string;
  /** The header cell's text, shown in place of `label` on the Mac host while
   *  the first row counts as headers. */
  heading?: string;
}

export interface SortDialogOptions {
  /** The spreadsheet host; anchors the dialog and picks its platform layout. */
  host?: HTMLElement;
  /** Labels of the Mac layout; the Mac host uses the shared layout without them. */
  mac?: MacSortDialogLabels;
  title: string;
  columnLabel: string;
  thenByLabel: string;
  noThenByLabel: string;
  orderLabel: string;
  headerLabel: string;
  addLevelLabel: string;
  deleteLevelLabel: string;
  copyLevelLabel: string;
  levelUnavailableLabel: string;
  ascendingLabel: string;
  descendingLabel: string;
  columns: readonly SortDialogColumn[];
  /** Key choices once the sort runs left to right (Mac options sheet); values
   *  are absolute row indexes. Without them the column list stays. */
  rows?: readonly SortDialogColumn[];
  /** Sort On choices and the order of a colour or icon level. */
  sortOn: SortDialogSortOnLabels;
  /** The fills, font colours and conditional formatting icons a column's data shows, keyed by
   *  the column's value; a colour or icon level picks from them. */
  colorChoices?: (column: string) => ColumnColorChoices;
  /** Order › Custom List…: its label and the Custom Lists dialog it opens, which resolves with
   *  the chosen list or `null`. Without it the order offers no custom list. */
  customList?: { label: string; open(initial: readonly string[] | null): Promise<string[] | null> };
  initialColumn: string;
  initialDirection: 'asc' | 'desc';
  initialHasHeader: boolean;
  okLabel: string;
  cancelLabel: string;
}

export interface SortDialogSortOnLabels {
  heading: string;
  values: string;
  cellColor: string;
  fontColor: string;
  cellIcon: string;
  onTop: string;
  onBottom: string;
}

/** One level of the dialog: a value level, or a colour or icon to put on top (`asc`) or at the
 *  bottom. */
export interface SortDialogLevel {
  column: string;
  direction: 'asc' | 'desc';
  sortOn?: 'cellColor' | 'fontColor' | 'cellIcon' | 'customList';
  color?: string;
  icon?: ConditionalIconRef;
  /** The list order of a custom-list level. */
  customList?: string[];
}

export interface SortDialogResult {
  column: string;
  direction: 'asc' | 'desc';
  levels: SortDialogLevel[];
  hasHeader: boolean;
  orientation: 'topToBottom' | 'leftToRight';
  matchCase: boolean;
}

const buildColumnSelect = (
  columns: readonly SortDialogColumn[],
  initial: string,
): HTMLSelectElement => {
  return createDialogSelect(columns, initial);
};

interface SortLevelControls {
  row: HTMLDivElement;
  label: HTMLSpanElement;
  column: HTMLSelectElement;
  direction: HTMLSelectElement;
  sortOn: HTMLSelectElement;
  /** The level as the dialog's result reports it. */
  value(): SortDialogLevel;
}

const appendSortLevelRow = (
  body: HTMLElement,
  opts: SortDialogOptions,
  mac: MacSortDialogLabels | undefined,
  keys: readonly SortDialogColumn[],
  rowLabel: string,
  initial: SortDialogLevel,
): SortLevelControls => {
  const doc = body.ownerDocument;
  const row = doc.createElement('div');
  row.className = 'fc-sortdlg__level';
  row.setAttribute('role', 'row');

  const label = doc.createElement('span');
  label.className = 'fc-sortdlg__level-label';
  label.textContent = rowLabel;
  label.setAttribute('role', 'rowheader');

  const column = buildColumnSelect(keys, initial.column);
  column.classList.add('fc-sortdlg__column');
  column.setAttribute('aria-label', rowLabel);
  column.setAttribute('role', 'cell');

  const t = opts.sortOn;
  const sortOn = createDialogSelect([], '', { ariaLabel: t.heading, role: 'cell' });
  sortOn.classList.add('fc-sortdlg__sort-on');
  const direction = createDialogSelect([], '', { ariaLabel: opts.orderLabel, role: 'cell' });
  direction.classList.add('fc-sortdlg__direction');
  // A colour or icon level's choices: the swatches of what the column shows. The Mac layout
  // gives them its Color column; the shared layout sets them beside the order. Provisional until
  // measured.
  const swatches = doc.createElement('div');
  swatches.className = 'fc-sortdlg__color';
  swatches.setAttribute('role', 'cell');
  swatches.setAttribute('aria-label', mac?.colorHeading ?? t.heading);
  let picked: { color?: string; icon?: ConditionalIconRef } = {
    ...(initial.color ? { color: initial.color } : {}),
    ...(initial.icon ? { icon: initial.icon } : {}),
  };

  // A value level ordered by a custom list keeps the list here; the order then reads `list`.
  let customList: string[] | null =
    initial.sortOn === 'customList' && initial.customList ? [...initial.customList] : null;
  const kind = (): 'cellColor' | 'fontColor' | 'cellIcon' | undefined =>
    sortOn.value === 'values'
      ? undefined
      : (sortOn.value as 'cellColor' | 'fontColor' | 'cellIcon');
  const choices = (): ColumnColorChoices =>
    opts.colorChoices?.(column.value) ?? { fills: [], fonts: [], icons: [] };

  const renderSortOn = (): void => {
    const current =
      kind() ?? (initial.sortOn === 'customList' ? undefined : initial.sortOn) ?? 'values';
    const c = choices();
    sortOn.replaceChildren();
    appendDialogSelectOptions(sortOn, [
      { value: 'values', label: mac?.sortOnValues ?? t.values },
      { value: 'cellColor', label: t.cellColor, disabled: c.fills.length === 0 },
      { value: 'fontColor', label: t.fontColor, disabled: c.fonts.length === 0 },
      { value: 'cellIcon', label: t.cellIcon, disabled: c.icons.length === 0 },
    ]);
    const option = [...sortOn.options].find((o) => o.value === current);
    sortOn.value = option && !option.disabled ? current : 'values';
  };
  const renderOrder = (value: 'asc' | 'desc' | 'list'): void => {
    direction.replaceChildren();
    appendDialogSelectOptions(
      direction,
      kind()
        ? [
            { value: 'asc', label: t.onTop },
            { value: 'desc', label: t.onBottom },
          ]
        : [
            { value: 'asc', label: opts.ascendingLabel },
            { value: 'desc', label: opts.descendingLabel },
            ...(customList ? [{ value: 'list', label: customList.join(', ') }] : []),
            ...(opts.customList ? [{ value: 'custom', label: opts.customList.label }] : []),
          ],
    );
    direction.value = value === 'list' && !customList ? 'asc' : value;
  };
  // Order › Custom List… opens the Custom Lists dialog; Cancel puts the order back.
  let lastOrder = direction.value;
  direction.addEventListener('change', () => {
    if (direction.value !== 'custom' || !opts.customList) {
      lastOrder = direction.value;
      return;
    }
    const previous = lastOrder;
    void opts.customList.open(customList).then((list) => {
      if (list && list.length > 0) customList = list;
      renderOrder(list && list.length > 0 ? 'list' : (previous as 'asc' | 'desc' | 'list'));
      lastOrder = direction.value;
    });
  });
  const renderSwatches = (): void => {
    swatches.replaceChildren();
    const k = kind();
    swatches.hidden = !k;
    if (!k) return;
    const c = choices();
    const buttons: HTMLButtonElement[] = [];
    const add = (labelText: string, choice: { color?: string; icon?: ConditionalIconRef }) => {
      const button = appendDialogButton(swatches, { label: '' });
      button.className = 'fc-sortdlg__swatch';
      button.setAttribute('aria-label', labelText);
      const on =
        choice.color !== undefined
          ? picked.color === choice.color
          : picked.icon?.set === choice.icon?.set && picked.icon?.slot === choice.icon?.slot;
      button.setAttribute('aria-pressed', on ? 'true' : 'false');
      button.addEventListener('click', () => {
        picked = choice;
        for (const other of buttons) other.setAttribute('aria-pressed', 'false');
        button.setAttribute('aria-pressed', 'true');
      });
      buttons.push(button);
      return button;
    };
    if (k === 'cellIcon') {
      for (const icon of c.icons) {
        const button = add(`${icon.set} ${icon.slot + 1}`, { icon });
        button.dataset.fcIcon = `${icon.set}:${icon.slot}`;
        button.appendChild(conditionalIconCanvas(doc, icon.set, icon.slot));
      }
    } else {
      for (const color of k === 'cellColor' ? c.fills : c.fonts) {
        const button = add(color, { color });
        button.dataset.fcColor = color;
        button.style.setProperty('--fc-swatch', color);
      }
    }
    // A level starts on the column's first colour or icon, as the desktop does.
    if (!buttons.some((b) => b.getAttribute('aria-pressed') === 'true')) buttons[0]?.click();
  };

  renderSortOn();
  renderOrder(customList ? 'list' : initial.direction);
  lastOrder = direction.value;
  renderSwatches();
  sortOn.addEventListener('change', () => {
    picked = {};
    renderOrder(direction.value === 'desc' ? 'desc' : 'asc');
    renderSwatches();
  });
  column.addEventListener('change', () => {
    picked = {};
    renderSortOn();
    renderOrder(direction.value === 'desc' ? 'desc' : 'asc');
    renderSwatches();
  });

  if (mac) {
    row.append(label, column, sortOn, direction, swatches);
  } else {
    const order = doc.createElement('div');
    order.className = 'fc-sortdlg__order';
    order.setAttribute('role', 'cell');
    order.append(swatches, direction);
    direction.removeAttribute('role');
    row.append(label, column, sortOn, order);
  }
  body.appendChild(row);
  const value = (): SortDialogLevel => {
    const k = kind();
    const dir = direction.value === 'desc' ? 'desc' : 'asc';
    if (!k && direction.value === 'list' && customList) {
      return { column: column.value, direction: 'asc', sortOn: 'customList', customList };
    }
    if (k === 'cellIcon' && picked.icon) {
      return { column: column.value, direction: dir, sortOn: k, icon: picked.icon };
    }
    if ((k === 'cellColor' || k === 'fontColor') && picked.color) {
      return { column: column.value, direction: dir, sortOn: k, color: picked.color };
    }
    return { column: column.value, direction: dir };
  };
  return { row, label, column, direction, sortOn, value };
};

interface SortOptionsChoice {
  orientation: SortDialogResult['orientation'];
  matchCase: boolean;
}

/** The Mac options sheet layered over the Sort dialog. Resolves the choices on
 *  OK and `null` on Cancel. The phonetic radios are shown as measured but do
 *  not affect order: the workbook carries no phonetic guide text. */
const showSortOptionsSheet = (
  host: HTMLElement | undefined,
  mac: MacSortDialogLabels,
  actions: { okLabel: string; cancelLabel: string },
  initial: SortOptionsChoice,
): Promise<SortOptionsChoice | null> =>
  new Promise<SortOptionsChoice | null>((resolve) => {
    const shell = createDialogShell({ host, title: mac.options });
    shell.panel.classList.add('fc-sortdlg-options');
    const doc = shell.body.ownerDocument;

    const group = (title: string): HTMLDivElement => {
      const box = doc.createElement('div');
      box.className = 'fc-sortdlg-options__group';
      box.setAttribute('role', 'group');
      box.setAttribute('aria-label', title);
      const heading = doc.createElement('span');
      heading.className = 'fc-sortdlg-options__heading';
      heading.textContent = title;
      box.appendChild(heading);
      return box;
    };
    const control = (
      box: HTMLElement,
      type: 'radio' | 'checkbox',
      name: string,
      text: string,
      checked: boolean,
    ): HTMLInputElement => {
      const label = doc.createElement('label');
      label.className = 'fc-sortdlg-options__choice';
      const input = doc.createElement('input');
      input.type = type;
      input.name = name;
      input.checked = checked;
      label.append(input, doc.createTextNode(text));
      box.appendChild(label);
      return input;
    };

    const orientationBox = group(mac.orientation);
    const topToBottom = control(
      orientationBox,
      'radio',
      'fc-sort-orientation',
      mac.sortTopToBottom,
      initial.orientation === 'topToBottom',
    );
    const leftToRight = control(
      orientationBox,
      'radio',
      'fc-sort-orientation',
      mac.sortLeftToRight,
      initial.orientation === 'leftToRight',
    );
    const caseBox = doc.createElement('div');
    caseBox.className = 'fc-sortdlg-options__case';
    const caseSensitive = control(caseBox, 'checkbox', '', mac.caseSensitive, initial.matchCase);
    orientationBox.appendChild(caseBox);

    const methodBox = group(mac.method);
    control(methodBox, 'radio', 'fc-sort-method', mac.usePhonetic, true);
    control(methodBox, 'radio', 'fc-sort-method', mac.noPhonetic, false);
    shell.body.append(orientationBox, methodBox);

    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: actions.cancelLabel,
      okLabel: actions.okLabel,
    });
    const lifecycle = installDialogLifecycle<SortOptionsChoice | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => lifecycle.finish(choice()),
    });
    const choice = (): SortOptionsChoice => ({
      orientation: leftToRight.checked && !topToBottom.checked ? 'leftToRight' : 'topToBottom',
      matchCase: caseSensitive.checked,
    });
    okBtn.addEventListener('click', () => lifecycle.finish(choice()));
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));
    mountDialog(shell, topToBottom, { immediate: true });
  });

export const showSortDialog = (opts: SortDialogOptions): Promise<SortDialogResult | null> =>
  new Promise<SortDialogResult | null>((resolve) => {
    const shell = createDialogShell({ host: opts.host, title: opts.title });
    shell.panel.classList.add('fc-sortdlg');
    const mac =
      shell.overlay.closest<HTMLElement>('[data-fc-platform]')?.dataset.fcPlatform === 'mac'
        ? opts.mac
        : undefined;

    const toolbar = document.createElement('div');
    toolbar.className = 'fc-sortdlg__toolbar';
    const addLevelBtn = appendDialogButton(toolbar, { label: opts.addLevelLabel });
    addLevelBtn.classList.add('fc-sortdlg__add-level');
    const deleteLevelBtn = appendDialogButton(toolbar, { label: opts.deleteLevelLabel });
    deleteLevelBtn.classList.add('fc-sortdlg__delete-level');
    const copyLevelBtn = appendDialogButton(toolbar, { label: opts.copyLevelLabel });
    copyLevelBtn.classList.add('fc-sortdlg__copy-level');
    if (mac) {
      // The Mac app captions the level buttons + / − / Copy under the table.
      addLevelBtn.textContent = '+';
      deleteLevelBtn.textContent = '−';
      copyLevelBtn.textContent = mac.copyLevel;
      addLevelBtn.setAttribute('aria-label', opts.addLevelLabel);
      deleteLevelBtn.setAttribute('aria-label', opts.deleteLevelLabel);
      copyLevelBtn.setAttribute('aria-label', opts.copyLevelLabel);
      const heading = document.createElement('span');
      heading.className = 'fc-sortdlg__heading';
      heading.textContent = mac.levelsHeading;
      shell.body.appendChild(heading);
    }
    shell.body.appendChild(toolbar);

    const headerRow = document.createElement('div');
    headerRow.className = 'fc-sortdlg__grid-head';
    headerRow.setAttribute('role', 'row');
    const headings = mac
      ? ['', mac.columnHeading, mac.sortOnHeading, opts.orderLabel, mac.colorHeading]
      : ['', opts.columnLabel, opts.sortOn.heading, opts.orderLabel];
    const headCells: HTMLSpanElement[] = [];
    for (const label of headings) {
      const cell = document.createElement('span');
      cell.setAttribute('role', 'columnheader');
      cell.textContent = label;
      headerRow.appendChild(cell);
      headCells.push(cell);
    }
    shell.body.appendChild(headerRow);

    const levelsWrap = document.createElement('div');
    levelsWrap.className = 'fc-sortdlg__levels';
    levelsWrap.setAttribute('role', 'grid');
    levelsWrap.setAttribute('aria-label', opts.title);
    shell.body.appendChild(levelsWrap);

    const levels: SortLevelControls[] = [];
    let selectedLevelIndex = 0;
    let orientation: SortDialogResult['orientation'] = 'topToBottom';
    let matchCase = false;
    // The key list names rows once the sort runs left to right.
    const keyItems = (): readonly SortDialogColumn[] =>
      orientation === 'leftToRight' && opts.rows ? opts.rows : opts.columns;

    const dataHeaderRow = document.createElement('label');
    dataHeaderRow.className = 'fc-fmtdlg__row fc-tb__dlg__label fc-sortdlg__header';
    const hasHeader = document.createElement('input');
    hasHeader.type = 'checkbox';
    hasHeader.checked = opts.initialHasHeader;
    dataHeaderRow.append(hasHeader, document.createTextNode(` ${opts.headerLabel}`));

    // On the Mac host header text names the columns while the first row counts as headers.
    const relabelColumns = (level: SortLevelControls): void => {
      if (!mac) return;
      keyItems().forEach((item, index) => {
        const option = level.column.options[index];
        if (option) option.text = hasHeader.checked && item.heading ? item.heading : item.label;
      });
    };
    hasHeader.addEventListener('change', () => {
      for (const level of levels) relabelColumns(level);
    });

    const selectLevel = (index: number): void => {
      selectedLevelIndex = Math.max(0, Math.min(index, levels.length - 1));
      for (const [idx, level] of levels.entries()) {
        const selected = idx === selectedLevelIndex;
        level.row.classList.toggle('fc-sortdlg__level--selected', selected);
        level.row.setAttribute('aria-selected', String(selected));
      }
      const hasSelected = levels.length > 0;
      projectDisabledState(
        deleteLevelBtn,
        !hasSelected || levels.length === 1,
        opts.levelUnavailableLabel,
        { datasetKey: 'disabledReason' },
      );
      projectDisabledState(copyLevelBtn, !hasSelected, opts.levelUnavailableLabel, {
        datasetKey: 'disabledReason',
      });
    };

    const refreshLabels = (): void => {
      levels.forEach((level, index) => {
        level.label.textContent = index === 0 ? opts.columnLabel : opts.thenByLabel;
        level.column.setAttribute('aria-label', level.label.textContent);
      });
      selectLevel(selectedLevelIndex);
    };

    const addLevel = (
      initial: SortDialogLevel = { column: opts.initialColumn, direction: opts.initialDirection },
      focus = true,
    ): void => {
      const level = appendSortLevelRow(
        levelsWrap,
        opts,
        mac,
        keyItems(),
        levels.length === 0 ? opts.columnLabel : opts.thenByLabel,
        initial,
      );
      level.row.addEventListener('click', () => selectLevel(levels.indexOf(level)));
      level.row.addEventListener('focusin', () => selectLevel(levels.indexOf(level)));
      levels.push(level);
      relabelColumns(level);
      selectedLevelIndex = levels.length - 1;
      refreshLabels();
      if (focus) level.column.focus({ preventScroll: true });
    };

    addLevel(undefined, false);

    addLevelBtn.addEventListener('click', () => {
      const source = levels[selectedLevelIndex]?.value();
      addLevel(
        source
          ? { column: source.column, direction: source.direction }
          : { column: opts.initialColumn, direction: opts.initialDirection },
      );
    });
    deleteLevelBtn.addEventListener('click', () => {
      if (levels.length <= 1) return;
      const [removed] = levels.splice(selectedLevelIndex, 1);
      removed?.row.remove();
      selectedLevelIndex = Math.max(0, selectedLevelIndex - 1);
      refreshLabels();
      levels[selectedLevelIndex]?.column.focus({ preventScroll: true });
    });
    copyLevelBtn.addEventListener('click', () => {
      const source = levels[selectedLevelIndex];
      if (!source) return;
      addLevel(source.value());
    });

    shell.body.appendChild(dataHeaderRow);

    if (mac) {
      const optionsBtn = appendDialogButton(shell.footer, { label: mac.options });
      optionsBtn.classList.add('fc-sortdlg__options');
      optionsBtn.addEventListener('click', () => {
        void showSortOptionsSheet(opts.host, mac, opts, { orientation, matchCase }).then(
          (choice) => {
            if (!choice) return;
            const flipped = choice.orientation !== orientation;
            ({ orientation, matchCase } = choice);
            if (!flipped) return;
            headCells[1]?.replaceChildren(
              orientation === 'leftToRight' && opts.rows ? mac.rowHeading : mac.columnHeading,
            );
            const items = keyItems();
            for (const level of levels) {
              level.column.replaceChildren();
              appendDialogSelectOptions(level.column, items);
              level.column.value = items[0]?.value ?? '';
              relabelColumns(level);
            }
          },
        );
      });
    }
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: opts.cancelLabel,
      okLabel: opts.okLabel,
    });

    const buildResult = (): SortDialogResult => ({
      column: levels[0]?.column.value ?? opts.initialColumn,
      direction: levels[0]?.direction.value === 'desc' ? 'desc' : 'asc',
      levels: levels.map((level) => level.value()),
      hasHeader: hasHeader.checked,
      orientation,
      matchCase,
    });

    const lifecycle = installDialogLifecycle<SortDialogResult | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => lifecycle.finish(buildResult()),
    });
    okBtn.addEventListener('click', () => lifecycle.finish(buildResult()));
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, () => levels[0]?.column.focus({ preventScroll: true }), { immediate: true });
  });
