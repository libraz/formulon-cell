import type { MacSortDialogLabels } from '../../i18n/strings/mac/sort-dialog.js';
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
  initialColumn: string;
  initialDirection: 'asc' | 'desc';
  initialHasHeader: boolean;
  okLabel: string;
  cancelLabel: string;
}

export interface SortDialogResult {
  column: string;
  direction: 'asc' | 'desc';
  levels: Array<{ column: string; direction: 'asc' | 'desc' }>;
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

const buildDirectionSelect = (
  asc: string,
  desc: string,
  initial: 'asc' | 'desc',
): HTMLSelectElement => {
  return createDialogSelect(
    [
      { value: 'asc', label: asc },
      { value: 'desc', label: desc },
    ],
    initial,
  );
};

interface SortLevelControls {
  row: HTMLDivElement;
  label: HTMLSpanElement;
  column: HTMLSelectElement;
  direction: HTMLSelectElement;
}

const appendSortLevelRow = (
  body: HTMLElement,
  opts: SortDialogOptions,
  mac: MacSortDialogLabels | undefined,
  keys: readonly SortDialogColumn[],
  rowLabel: string,
  columnValue: string,
  directionValue: 'asc' | 'desc',
): SortLevelControls => {
  const row = document.createElement('div');
  row.className = 'fc-sortdlg__level';
  row.setAttribute('role', 'row');

  const label = document.createElement('span');
  label.className = 'fc-sortdlg__level-label';
  label.textContent = rowLabel;
  label.setAttribute('role', 'rowheader');

  const column = buildColumnSelect(keys, columnValue);
  column.classList.add('fc-sortdlg__column');
  column.setAttribute('aria-label', rowLabel);
  column.setAttribute('role', 'cell');

  const direction = buildDirectionSelect(opts.ascendingLabel, opts.descendingLabel, directionValue);
  direction.classList.add('fc-sortdlg__direction');
  direction.setAttribute('aria-label', opts.orderLabel);
  direction.setAttribute('role', 'cell');

  if (mac) {
    // Values are the only sort-on key; colour and icon sorts are not offered.
    const sortOn = createDialogSelect([{ value: 'values', label: mac.sortOnValues }], 'values', {
      ariaLabel: mac.sortOnHeading,
      role: 'cell',
    });
    sortOn.classList.add('fc-sortdlg__sort-on');
    const color = createDialogSelect([], '', { ariaLabel: mac.colorHeading, role: 'cell' });
    color.classList.add('fc-sortdlg__color');
    projectDisabledState(color, true, null);
    row.append(label, column, sortOn, direction, color);
  } else {
    row.append(label, column, direction);
  }
  body.appendChild(row);
  return { row, label, column, direction };
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
      : ['', opts.columnLabel, opts.orderLabel];
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
      columnValue = opts.initialColumn,
      directionValue: 'asc' | 'desc' = opts.initialDirection,
      focus = true,
    ): void => {
      const level = appendSortLevelRow(
        levelsWrap,
        opts,
        mac,
        keyItems(),
        levels.length === 0 ? opts.columnLabel : opts.thenByLabel,
        columnValue,
        directionValue,
      );
      level.row.addEventListener('click', () => selectLevel(levels.indexOf(level)));
      level.row.addEventListener('focusin', () => selectLevel(levels.indexOf(level)));
      levels.push(level);
      relabelColumns(level);
      selectedLevelIndex = levels.length - 1;
      refreshLabels();
      if (focus) level.column.focus({ preventScroll: true });
    };

    addLevel(opts.initialColumn, opts.initialDirection, false);

    addLevelBtn.addEventListener('click', () => {
      const source = levels[selectedLevelIndex];
      addLevel(
        source?.column.value ?? opts.initialColumn,
        source?.direction.value === 'desc' ? 'desc' : 'asc',
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
      addLevel(source.column.value, source.direction.value === 'desc' ? 'desc' : 'asc');
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
      levels: levels.map((level) => ({
        column: level.column.value,
        direction: level.direction.value === 'desc' ? 'desc' : 'asc',
      })),
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
