import {
  applyMerge,
  applyUnmerge,
  autoSum,
  bumpDecimals,
  clearFilter,
  cycleBorders,
  cycleCurrency,
  cyclePercent,
  deleteCols,
  deleteRows,
  hideCols,
  hideRows,
  hiddenInSelection,
  insertCols,
  insertRows,
  mutators,
  recordFormatChange,
  type SpreadsheetInstance,
  setAlign,
  setAutoFilter,
  setFillColor,
  setFreezePanes,
  setFont,
  setFontColor,
  showCols,
  showRows,
  sortRange,
  toggleBold,
  toggleItalic,
  toggleStrike,
  toggleUnderline,
  toggleWrap,
} from '@libraz/formulon-cell';
import { type ReactElement, useCallback, useEffect, useState } from 'react';

interface Props {
  instance: SpreadsheetInstance | null;
}

interface ActiveState {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  alignLeft: boolean;
  alignCenter: boolean;
  alignRight: boolean;
  currency: boolean;
  percent: boolean;
  frozen: boolean;
  filterOn: boolean;
  rowsHidden: boolean;
  colsHidden: boolean;
  fontFamily: string;
  fontSize: number;
  fontColor: string;
  fillColor: string;
}

const EMPTY: ActiveState = {
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  alignLeft: false,
  alignCenter: false,
  alignRight: false,
  currency: false,
  percent: false,
  frozen: false,
  filterOn: false,
  rowsHidden: false,
  colsHidden: false,
  fontFamily: 'Aptos',
  fontSize: 11,
  fontColor: '#201f1e',
  fillColor: '#ffffff',
};

const FONT_FAMILIES = ['Aptos', 'Calibri', 'Arial', 'Segoe UI', 'Times New Roman', 'Consolas'];
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36];

const project = (inst: SpreadsheetInstance): ActiveState => {
  const s = inst.store.getState();
  const a = s.selection.active;
  const r = s.selection.range;
  const f = s.format.formats.get(`${a.sheet}:${a.row}:${a.col}`);
  return {
    bold: !!f?.bold,
    italic: !!f?.italic,
    underline: !!f?.underline,
    strike: !!f?.strike,
    alignLeft: f?.align === 'left',
    alignCenter: f?.align === 'center',
    alignRight: f?.align === 'right',
    currency: f?.numFmt?.kind === 'currency',
    percent: f?.numFmt?.kind === 'percent',
    frozen: s.layout.freezeRows > 0 || s.layout.freezeCols > 0,
    filterOn: s.ui.filterRange != null,
    rowsHidden: hiddenInSelection(s.layout, 'row', r.r0, r.r1).length > 0,
    colsHidden: hiddenInSelection(s.layout, 'col', r.c0, r.c1).length > 0,
    fontFamily: f?.fontFamily ?? 'Aptos',
    fontSize: f?.fontSize ?? 11,
    fontColor: f?.color ?? '#201f1e',
    fillColor: f?.fill ?? '#ffffff',
  };
};

type IconName =
  | 'undo'
  | 'redo'
  | 'decDown'
  | 'decUp'
  | 'autosum'
  | 'alignLeft'
  | 'alignCenter'
  | 'alignRight'
  | 'borders'
  | 'merge'
  | 'wrap'
  | 'freeze'
  | 'insertRows'
  | 'deleteRows'
  | 'insertCols'
  | 'deleteCols'
  | 'filter'
  | 'sortAsc'
  | 'sortDesc';

const Icon = ({ name }: { name: IconName }): ReactElement => {
  const common = {
    className: 'demo__rb-icon',
    viewBox: '0 0 20 20',
    fill: 'none',
    stroke: 'currentColor',
    strokeWidth: 1.45,
    strokeLinecap: 'round' as const,
    strokeLinejoin: 'round' as const,
    'aria-hidden': true,
  };
  switch (name) {
    case 'undo':
      return (
        <svg {...common}>
          <path d="M7.2 5.2H3.8v-3.4" />
          <path d="M4 5.2c2.2-2.1 5.7-2.3 8.1-.5 2.7 2.1 3 6.1.7 8.6-1.8 1.9-4.8 2.4-7.1 1.2" />
        </svg>
      );
    case 'redo':
      return (
        <svg {...common}>
          <path d="M12.8 5.2h3.4v-3.4" />
          <path d="M16 5.2c-2.2-2.1-5.7-2.3-8.1-.5-2.7 2.1-3 6.1-.7 8.6 1.8 1.9 4.8 2.4 7.1 1.2" />
        </svg>
      );
    case 'decDown':
      return (
        <svg {...common}>
          <path d="M3 14.5h5" />
          <path d="M11 5.5h6" />
          <path d="M11 9.5h4" />
          <path d="M11 13.5h2" />
          <path d="M5.5 5.8v6.5" />
          <path d="M3.8 10.5l1.7 1.8 1.7-1.8" />
        </svg>
      );
    case 'decUp':
      return (
        <svg {...common}>
          <path d="M3 14.5h5" />
          <path d="M11 5.5h2" />
          <path d="M11 9.5h4" />
          <path d="M11 13.5h6" />
          <path d="M5.5 12.2V5.7" />
          <path d="M3.8 7.5l1.7-1.8 1.7 1.8" />
        </svg>
      );
    case 'autosum':
      return (
        <svg {...common}>
          <path d="M15.5 4.5H5.2l5 5.5-5 5.5h10.3" />
        </svg>
      );
    case 'alignLeft':
      return (
        <svg {...common}>
          <path d="M4 5h12" />
          <path d="M4 8.5h8" />
          <path d="M4 12h12" />
          <path d="M4 15.5h7" />
        </svg>
      );
    case 'alignCenter':
      return (
        <svg {...common}>
          <path d="M4 5h12" />
          <path d="M6 8.5h8" />
          <path d="M4 12h12" />
          <path d="M6.5 15.5h7" />
        </svg>
      );
    case 'alignRight':
      return (
        <svg {...common}>
          <path d="M4 5h12" />
          <path d="M8 8.5h8" />
          <path d="M4 12h12" />
          <path d="M9 15.5h7" />
        </svg>
      );
    case 'borders':
      return (
        <svg {...common}>
          <path d="M4 4h12v12H4z" />
          <path d="M10 4v12" />
          <path d="M4 10h12" />
        </svg>
      );
    case 'merge':
      return (
        <svg {...common}>
          <path d="M4 5h12v10H4z" />
          <path d="M8 5v10" />
          <path d="M12 5v10" />
          <path d="M7 10h6" />
          <path d="M11.5 8.5L13 10l-1.5 1.5" />
          <path d="M8.5 8.5L7 10l1.5 1.5" />
        </svg>
      );
    case 'wrap':
      return (
        <svg {...common}>
          <path d="M4 5h12" />
          <path d="M4 9h9a3 3 0 0 1 0 6H8" />
          <path d="M9.8 12.8L7.6 15l2.2 2.2" />
        </svg>
      );
    case 'freeze':
      return (
        <svg {...common}>
          <path d="M4 4h12v12H4z" />
          <path d="M4 8h12" />
          <path d="M8 4v12" />
          <path d="M8 8h8v8H8z" />
        </svg>
      );
    case 'insertRows':
      return (
        <svg {...common}>
          <path d="M4 5h12" />
          <path d="M4 10h12" />
          <path d="M4 15h12" />
          <path d="M10 7.5v5" />
          <path d="M7.5 10h5" />
        </svg>
      );
    case 'deleteRows':
      return (
        <svg {...common}>
          <path d="M4 5h12" />
          <path d="M4 10h12" />
          <path d="M4 15h12" />
          <path d="M7.8 7.8l4.4 4.4" />
          <path d="M12.2 7.8l-4.4 4.4" />
        </svg>
      );
    case 'insertCols':
      return (
        <svg {...common}>
          <path d="M5 4v12" />
          <path d="M10 4v12" />
          <path d="M15 4v12" />
          <path d="M7.5 10h5" />
          <path d="M10 7.5v5" />
        </svg>
      );
    case 'deleteCols':
      return (
        <svg {...common}>
          <path d="M5 4v12" />
          <path d="M10 4v12" />
          <path d="M15 4v12" />
          <path d="M7.8 7.8l4.4 4.4" />
          <path d="M12.2 7.8l-4.4 4.4" />
        </svg>
      );
    case 'filter':
      return (
        <svg {...common}>
          <path d="M4 5h12l-4.8 5.3v4.1L8.8 16v-5.7z" />
        </svg>
      );
    case 'sortAsc':
      return (
        <svg {...common}>
          <path d="M6 15V5" />
          <path d="M3.8 7.2L6 5l2.2 2.2" />
          <path d="M11 6h4" />
          <path d="M11 10h3" />
          <path d="M11 14h2" />
        </svg>
      );
    case 'sortDesc':
      return (
        <svg {...common}>
          <path d="M6 5v10" />
          <path d="M3.8 12.8L6 15l2.2-2.2" />
          <path d="M11 6h2" />
          <path d="M11 10h3" />
          <path d="M11 14h4" />
        </svg>
      );
  }
};

export const Toolbar = ({ instance }: Props): ReactElement => {
  const [active, setActive] = useState<ActiveState>(EMPTY);

  useEffect(() => {
    if (!instance) return;
    setActive(project(instance));
    return instance.store.subscribe(() => setActive(project(instance)));
  }, [instance]);

  const wrapFormat = useCallback(
    (
      fn: (
        state: ReturnType<SpreadsheetInstance['store']['getState']>,
        store: SpreadsheetInstance['store'],
      ) => void,
    ) => {
      if (!instance) return;
      recordFormatChange(instance.history, instance.store, () =>
        fn(instance.store.getState(), instance.store),
      );
    },
    [instance],
  );

  const onUndo = useCallback(() => instance?.undo(), [instance]);
  const onRedo = useCallback(() => instance?.redo(), [instance]);

  const onAutoSum = useCallback(() => {
    if (!instance) return;
    const result = autoSum(instance.store.getState(), instance.workbook);
    if (!result) return;
    mutators.replaceCells(instance.store, instance.workbook.cells(result.addr.sheet));
    mutators.setActive(instance.store, result.addr);
  }, [instance]);

  const onMerge = useCallback(() => {
    if (!instance) return;
    const s = instance.store.getState();
    const r = s.selection.range;
    const anchor = s.merges.byAnchor.get(`${r.sheet}:${r.r0}:${r.c0}`);
    const isExact =
      anchor &&
      r.r0 === anchor.r0 &&
      r.c0 === anchor.c0 &&
      r.r1 === anchor.r1 &&
      r.c1 === anchor.c1;
    if (isExact) applyUnmerge(instance.store, instance.workbook, instance.history, r);
    else applyMerge(instance.store, instance.workbook, instance.history, r);
  }, [instance]);

  const onFreezeToggle = useCallback(() => {
    if (!instance) return;
    const s = instance.store.getState();
    if (s.layout.freezeRows > 0 || s.layout.freezeCols > 0) {
      setFreezePanes(instance.store, instance.history, 0, 0, instance.workbook);
    } else {
      // Freeze rows/cols up to active cell, or first row if at A1.
      const a = s.selection.active;
      const rows = a.row === 0 && a.col === 0 ? 1 : a.row;
      const cols = a.row === 0 && a.col === 0 ? 0 : a.col;
      setFreezePanes(instance.store, instance.history, rows, cols, instance.workbook);
    }
  }, [instance]);

  const onInsertRows = useCallback(() => {
    if (!instance) return;
    const r = instance.store.getState().selection.range;
    insertRows(instance.store, instance.workbook, instance.history, r.r0, r.r1 - r.r0 + 1);
  }, [instance]);

  const onDeleteRows = useCallback(() => {
    if (!instance) return;
    const r = instance.store.getState().selection.range;
    deleteRows(instance.store, instance.workbook, instance.history, r.r0, r.r1 - r.r0 + 1);
  }, [instance]);

  const onInsertCols = useCallback(() => {
    if (!instance) return;
    const r = instance.store.getState().selection.range;
    insertCols(instance.store, instance.workbook, instance.history, r.c0, r.c1 - r.c0 + 1);
  }, [instance]);

  const onDeleteCols = useCallback(() => {
    if (!instance) return;
    const r = instance.store.getState().selection.range;
    deleteCols(instance.store, instance.workbook, instance.history, r.c0, r.c1 - r.c0 + 1);
  }, [instance]);

  const onToggleRowsHidden = useCallback(() => {
    if (!instance) return;
    const s = instance.store.getState();
    const r = s.selection.range;
    if (hiddenInSelection(s.layout, 'row', r.r0, r.r1).length > 0) {
      showRows(instance.store, instance.history, r.r0, r.r1, instance.workbook);
    } else {
      hideRows(instance.store, instance.history, r.r0, r.r1, instance.workbook);
    }
  }, [instance]);

  const onToggleColsHidden = useCallback(() => {
    if (!instance) return;
    const s = instance.store.getState();
    const r = s.selection.range;
    if (hiddenInSelection(s.layout, 'col', r.c0, r.c1).length > 0) {
      showCols(instance.store, instance.history, r.c0, r.c1, instance.workbook);
    } else {
      hideCols(instance.store, instance.history, r.c0, r.c1, instance.workbook);
    }
  }, [instance]);

  const onFilterToggle = useCallback(() => {
    if (!instance) return;
    const s = instance.store.getState();
    if (s.ui.filterRange) clearFilter(s, instance.store, s.ui.filterRange);
    else setAutoFilter(instance.store, s.selection.range);
  }, [instance]);

  const onSort = useCallback(
    (direction: 'asc' | 'desc') => {
      if (!instance) return;
      const s = instance.store.getState();
      const ok = sortRange(s, instance.store, instance.workbook, s.selection.range, {
        byCol: s.selection.active.col,
        direction,
        hasHeader: s.selection.range.r0 < s.selection.range.r1,
      });
      if (ok) mutators.replaceCells(instance.store, instance.workbook.cells(s.data.sheetIndex));
    },
    [instance],
  );

  const tool = (
    id: string,
    title: string,
    label: string | ReactElement,
    onClick: () => void,
    isActive = false,
    extra = '',
  ): ReactElement => (
    <button
      key={id}
      type="button"
      className={`demo__rb${extra}${isActive ? ' demo__rb--active' : ''}`}
      title={title}
      aria-label={title}
      onClick={onClick}
      disabled={!instance}
    >
      {label}
    </button>
  );

  const group = (title: string, children: ReactElement[]): ReactElement => (
    <section className="demo__ribbon-group" aria-label={title}>
      <div className="demo__ribbon-tools">{children}</div>
      <div className="demo__ribbon-label">{title}</div>
    </section>
  );

  const select = (
    id: string,
    title: string,
    value: string | number,
    values: readonly (string | number)[],
    onChange: (value: string) => void,
    extra = '',
  ): ReactElement => (
    <select
      key={id}
      className={`demo__rb-select${extra}`}
      title={title}
      aria-label={title}
      value={value}
      disabled={!instance}
      onChange={(e) => onChange(e.currentTarget.value)}
    >
      {values.map((v) => (
        <option key={v} value={v}>
          {v}
        </option>
      ))}
    </select>
  );

  const color = (
    id: string,
    title: string,
    value: string,
    onChange: (value: string) => void,
    label: string,
  ): ReactElement => (
    <label key={id} className="demo__rb-color" title={title} aria-label={title}>
      <span>{label}</span>
      <input
        type="color"
        value={value}
        disabled={!instance}
        onChange={(e) => onChange(e.currentTarget.value)}
      />
    </label>
  );

  return (
    <div className="demo__ribbon-shell">
      <div className="demo__ribbon-tabs" role="tablist" aria-label="Ribbon tabs">
        <button
          type="button"
          className="demo__ribbon-tab demo__ribbon-tab--file"
          role="tab"
          aria-selected="false"
        >
          File
        </button>
        <button
          type="button"
          className="demo__ribbon-tab demo__ribbon-tab--active"
          role="tab"
          aria-selected="true"
        >
          Home
        </button>
        <button
          type="button"
          className="demo__ribbon-tab"
          role="tab"
          aria-selected="false"
          disabled
        >
          Insert
        </button>
        <button
          type="button"
          className="demo__ribbon-tab"
          role="tab"
          aria-selected="false"
          disabled
        >
          Formulas
        </button>
        <button
          type="button"
          className="demo__ribbon-tab"
          role="tab"
          aria-selected="false"
          disabled
        >
          Data
        </button>
        <button
          type="button"
          className="demo__ribbon-tab"
          role="tab"
          aria-selected="false"
          disabled
        >
          View
        </button>
      </div>
      <div className="demo__ribbon" role="toolbar" aria-label="Home ribbon">
        {group('Clipboard', [
          tool('undo', 'Undo (⌘Z)', <Icon name="undo" />, onUndo),
          tool('redo', 'Redo (⌘⇧Z)', <Icon name="redo" />, onRedo),
        ])}
        {group('Number', [
          tool(
            'currency',
            'Currency',
            '$',
            () => wrapFormat(cycleCurrency),
            active.currency,
            ' demo__rb--mono',
          ),
          tool(
            'percent',
            'Percent',
            '%',
            () => wrapFormat(cyclePercent),
            active.percent,
            ' demo__rb--mono',
          ),
          tool('decDown', 'Decrease decimals', <Icon name="decDown" />, () =>
            wrapFormat((s, st) => bumpDecimals(s, st, -1)),
          ),
          tool('decUp', 'Increase decimals', <Icon name="decUp" />, () =>
            wrapFormat((s, st) => bumpDecimals(s, st, 1)),
          ),
          tool('autosum', 'AutoSum (Σ)', <Icon name="autosum" />, onAutoSum),
        ])}
        {group('Font', [
          select('fontFamily', 'Font', active.fontFamily, FONT_FAMILIES, (value) =>
            wrapFormat((s, st) => setFont(s, st, { fontFamily: value })),
          ' demo__rb-select--font'),
          select('fontSize', 'Font size', active.fontSize, FONT_SIZES, (value) =>
            wrapFormat((s, st) => setFont(s, st, { fontSize: Number(value) })),
          ),
          tool(
            'bold',
            'Bold (⌘B)',
            'B',
            () => wrapFormat(toggleBold),
            active.bold,
            ' demo__rb--bold',
          ),
          tool(
            'italic',
            'Italic (⌘I)',
            'I',
            () => wrapFormat(toggleItalic),
            active.italic,
            ' demo__rb--italic',
          ),
          tool(
            'underline',
            'Underline (⌘U)',
            'U',
            () => wrapFormat(toggleUnderline),
            active.underline,
            ' demo__rb--underline',
          ),
          tool(
            'strike',
            'Strikethrough',
            'S',
            () => wrapFormat(toggleStrike),
            active.strike,
            ' demo__rb--strike',
          ),
          tool('borders', 'Borders', <Icon name="borders" />, () => wrapFormat(cycleBorders)),
          color('fontColor', 'Font color', active.fontColor, (value) =>
            wrapFormat((s, st) => setFontColor(s, st, value)),
          'A'),
          color('fillColor', 'Fill color', active.fillColor, (value) =>
            wrapFormat((s, st) => setFillColor(s, st, value)),
          '▾'),
        ])}
        {group('Alignment', [
          tool(
            'alignL',
            'Align left',
            <Icon name="alignLeft" />,
            () => wrapFormat((s, st) => setAlign(s, st, 'left')),
            active.alignLeft,
          ),
          tool(
            'alignC',
            'Align center',
            <Icon name="alignCenter" />,
            () => wrapFormat((s, st) => setAlign(s, st, 'center')),
            active.alignCenter,
          ),
          tool(
            'alignR',
            'Align right',
            <Icon name="alignRight" />,
            () => wrapFormat((s, st) => setAlign(s, st, 'right')),
            active.alignRight,
          ),
          tool('merge', 'Merge cells', <Icon name="merge" />, onMerge),
          tool('wrap', 'Wrap text', <Icon name="wrap" />, () => wrapFormat(toggleWrap)),
        ])}
        {group('Cells', [
          tool('insertRows', 'Insert selected rows', <Icon name="insertRows" />, onInsertRows),
          tool('deleteRows', 'Delete selected rows', <Icon name="deleteRows" />, onDeleteRows),
          tool('insertCols', 'Insert selected columns', <Icon name="insertCols" />, onInsertCols),
          tool('deleteCols', 'Delete selected columns', <Icon name="deleteCols" />, onDeleteCols),
          tool('hideRows', active.rowsHidden ? 'Show selected rows' : 'Hide selected rows', 'R', onToggleRowsHidden, active.rowsHidden, ' demo__rb--mono'),
          tool('hideCols', active.colsHidden ? 'Show selected columns' : 'Hide selected columns', 'C', onToggleColsHidden, active.colsHidden, ' demo__rb--mono'),
        ])}
        {group('Data', [
          tool('filter', 'Filter', <Icon name="filter" />, onFilterToggle, active.filterOn),
          tool('sortAsc', 'Sort ascending', <Icon name="sortAsc" />, () => onSort('asc')),
          tool('sortDesc', 'Sort descending', <Icon name="sortDesc" />, () => onSort('desc')),
        ])}
        {group('View', [
          tool('freeze', 'Freeze panes', <Icon name="freeze" />, onFreezeToggle, active.frozen),
        ])}
      </div>
    </div>
  );
};
