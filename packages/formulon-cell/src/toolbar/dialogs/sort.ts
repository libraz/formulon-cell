import { projectDisabledState } from '../menu-a11y.js';
import { createDialogSelect } from './form-controls.js';
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
}

export interface SortDialogOptions {
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

  const column = buildColumnSelect(opts.columns, columnValue);
  column.setAttribute('aria-label', rowLabel);
  column.setAttribute('role', 'cell');

  const direction = buildDirectionSelect(opts.ascendingLabel, opts.descendingLabel, directionValue);
  direction.setAttribute('aria-label', opts.orderLabel);
  direction.setAttribute('role', 'cell');

  row.append(label, column, direction);
  body.appendChild(row);
  return { row, label, column, direction };
};

export const showSortDialog = (opts: SortDialogOptions): Promise<SortDialogResult | null> =>
  new Promise<SortDialogResult | null>((resolve) => {
    const shell = createDialogShell({ title: opts.title });

    const toolbar = document.createElement('div');
    toolbar.className = 'fc-sortdlg__toolbar';
    const addLevelBtn = appendDialogButton(toolbar, { label: opts.addLevelLabel });
    addLevelBtn.classList.add('fc-sortdlg__add-level');
    const deleteLevelBtn = appendDialogButton(toolbar, { label: opts.deleteLevelLabel });
    deleteLevelBtn.classList.add('fc-sortdlg__delete-level');
    const copyLevelBtn = appendDialogButton(toolbar, { label: opts.copyLevelLabel });
    copyLevelBtn.classList.add('fc-sortdlg__copy-level');
    shell.body.appendChild(toolbar);

    const headerRow = document.createElement('div');
    headerRow.className = 'fc-sortdlg__grid-head';
    headerRow.setAttribute('role', 'row');
    for (const label of ['', opts.columnLabel, opts.orderLabel]) {
      const cell = document.createElement('span');
      cell.setAttribute('role', 'columnheader');
      cell.textContent = label;
      headerRow.appendChild(cell);
    }
    shell.body.appendChild(headerRow);

    const levelsWrap = document.createElement('div');
    levelsWrap.className = 'fc-sortdlg__levels';
    levelsWrap.setAttribute('role', 'grid');
    levelsWrap.setAttribute('aria-label', opts.title);
    shell.body.appendChild(levelsWrap);

    const levels: SortLevelControls[] = [];
    let selectedLevelIndex = 0;

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
        levels.length === 0 ? opts.columnLabel : opts.thenByLabel,
        columnValue,
        directionValue,
      );
      level.row.addEventListener('click', () => selectLevel(levels.indexOf(level)));
      level.row.addEventListener('focusin', () => selectLevel(levels.indexOf(level)));
      levels.push(level);
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

    const dataHeaderRow = document.createElement('label');
    dataHeaderRow.className = 'fc-fmtdlg__row fc-tb__dlg__label';
    const hasHeader = document.createElement('input');
    hasHeader.type = 'checkbox';
    hasHeader.checked = opts.initialHasHeader;
    dataHeaderRow.append(hasHeader, document.createTextNode(` ${opts.headerLabel}`));
    shell.body.appendChild(dataHeaderRow);

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
    });

    const lifecycle = installDialogLifecycle<SortDialogResult | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => lifecycle.finish(buildResult()),
    });
    okBtn.addEventListener('click', () => lifecycle.finish(buildResult()));
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, () => levels[0]?.column.focus({ preventScroll: true }));
  });
