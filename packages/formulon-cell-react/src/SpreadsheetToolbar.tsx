import {
  applyMerge,
  applyUnmerge,
  autoSum,
  type CellBorderStyle,
  clearFilter,
  deleteCols,
  deleteRows,
  formatAsTable,
  hiddenInSelection,
  hideCols,
  hideRows,
  insertCols,
  insertRows,
  type MarginPreset,
  mutators,
  type PageOrientation,
  type PaperSize,
  recordFormatChange,
  recordPageSetupChange,
  removeDuplicates,
  type SpreadsheetInstance,
  setAutoFilter,
  setBorderPreset,
  setFreezePanes,
  setMarginPreset,
  setPageOrientation,
  setPaperSize,
  setSheetZoom,
  showCols,
  showRows,
  sortRange,
} from '@libraz/formulon-cell';
import {
  type KeyboardEvent,
  type ReactElement,
  useCallback,
  useEffect,
  useRef,
  useState,
} from 'react';
import { Dropdown } from './toolbar/Dropdown.js';
import { buildRibbonGroups } from './toolbar/groups.js';
import { Icon, type IconName } from './toolbar/icons.js';
import {
  type ActiveState,
  BORDER_PRESETS,
  BORDER_STYLES,
  type BorderPreset,
  EMPTY_ACTIVE_STATE,
  projectActiveState,
  RIBBON_KEYSHORTCUTS,
  RIBBON_TAB_LABELS,
  type RibbonTab,
  type SpreadsheetToolbarProps,
} from './toolbar/model.js';
import { toolbarText } from './toolbar/translations.js';

export type { RibbonTab, SpreadsheetToolbarProps } from './toolbar/model.js';

export const SpreadsheetToolbar = ({
  instance,
  activeTab,
  onTabChange,
  locale,
  onSpellingReview,
  onAccessibilityCheck,
  onRunScript,
  onDrawPen,
  onDrawEraser,
  onTranslate,
  onAddIn,
}: SpreadsheetToolbarProps): ReactElement => {
  const [active, setActive] = useState<ActiveState>(EMPTY_ACTIVE_STATE);
  const [borderStyle, setBorderStyle] = useState<CellBorderStyle>('thin');
  const tablistRef = useRef<HTMLDivElement | null>(null);
  const lang = locale === 'ja' ? 'ja' : 'en';
  const tr = toolbarText(lang);
  const borderPresets = BORDER_PRESETS.map((preset) => ({
    ...preset,
    label:
      preset.value === 'none'
        ? tr.noBorder
        : preset.value === 'outline'
          ? tr.outsideBorders
          : preset.value === 'all'
            ? tr.allBorders
            : preset.value === 'top'
              ? tr.topBorder
              : preset.value === 'bottom'
                ? tr.bottomBorder
                : preset.value === 'left'
                  ? tr.leftBorder
                  : preset.value === 'right'
                    ? tr.rightBorder
                    : tr.doubleBottomBorder,
  }));
  const borderStyles = BORDER_STYLES.map((style) => ({
    ...style,
    label:
      style.value === 'thin'
        ? tr.thin
        : style.value === 'medium'
          ? tr.medium
          : style.value === 'thick'
            ? tr.thick
            : style.value === 'dashed'
              ? tr.dashed
              : style.value === 'dotted'
                ? tr.dotted
                : tr.double,
  }));
  const ribbonTabs = (Object.keys(RIBBON_TAB_LABELS) as RibbonTab[])
    .filter((id) => id !== 'file')
    .map((id) => ({
      id,
      label: RIBBON_TAB_LABELS[id][lang],
    }));

  const focusRibbonTab = useCallback((tab: RibbonTab) => {
    requestAnimationFrame(() => {
      tablistRef.current
        ?.querySelector<HTMLButtonElement>(`[data-ribbon-tab="${tab}"]`)
        ?.focus({ preventScroll: true });
    });
  }, []);

  const onRibbonTabKeyDown = useCallback(
    (event: KeyboardEvent<HTMLDivElement>) => {
      const target = (event.target as Element | null)?.closest<HTMLButtonElement>(
        '[data-ribbon-tab]',
      );
      if (!target) return;
      const currentId = (target.dataset.ribbonTab as RibbonTab | undefined) ?? activeTab;
      const current = Math.max(
        0,
        ribbonTabs.findIndex((tab) => tab.id === currentId),
      );
      let next = current;
      if (event.key === 'ArrowRight') next = (current + 1) % ribbonTabs.length;
      else if (event.key === 'ArrowLeft')
        next = (current - 1 + ribbonTabs.length) % ribbonTabs.length;
      else if (event.key === 'Home') next = 0;
      else if (event.key === 'End') next = ribbonTabs.length - 1;
      else return;
      event.preventDefault();
      const nextTab = ribbonTabs[next]?.id;
      if (!nextTab) return;
      onTabChange(nextTab);
      focusRibbonTab(nextTab);
    },
    [activeTab, focusRibbonTab, onTabChange, ribbonTabs],
  );

  useEffect(() => {
    if (!instance) return;
    setActive(projectActiveState(instance));
    return instance.store.subscribe(() => setActive(projectActiveState(instance)));
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
  // Re-focus the host (canvas region) before delegating to the system
  // clipboard so the host-bound copy/cut/paste listeners run with a real
  // selection. document.execCommand still works on Safari/Chrome for copy
  // and cut; paste falls back to the same listener as Ctrl/⌘+V.
  const dispatchClipboard = useCallback(
    (kind: 'copy' | 'cut' | 'paste') => {
      if (!instance) return;
      instance.host.focus();
      try {
        document.execCommand(kind);
      } catch {
        // execCommand can throw on some browsers — swallow so the button
        // still feels like a hint rather than blowing up the chrome.
      }
    },
    [instance],
  );
  const onCopy = useCallback(() => dispatchClipboard('copy'), [dispatchClipboard]);
  const onCut = useCallback(() => dispatchClipboard('cut'), [dispatchClipboard]);
  const onPaste = useCallback(() => dispatchClipboard('paste'), [dispatchClipboard]);
  const onFormatPainter = useCallback(() => {
    instance?.formatPainter?.activate(false);
  }, [instance]);

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

  const onBorderPreset = useCallback(
    (preset: BorderPreset) => {
      wrapFormat((s, st) => {
        setBorderPreset(s, st, preset, borderStyle);
      });
    },
    [borderStyle, wrapFormat],
  );

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

  const onRemoveDuplicates = useCallback(() => {
    if (!instance) return;
    const s = instance.store.getState();
    const removed = removeDuplicates(s, instance.store, instance.workbook, s.selection.range);
    if (removed > 0) {
      mutators.replaceCells(instance.store, instance.workbook.cells(s.data.sheetIndex));
    }
  }, [instance]);

  const onZoom = useCallback(
    (zoom: number) => {
      if (!instance) return;
      setSheetZoom(instance.store, zoom, instance.workbook);
    },
    [instance],
  );

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

  const onPageOrientation = useCallback(
    (next: PageOrientation) => {
      if (!instance) return;
      const sheet = instance.store.getState().data.sheetIndex;
      recordPageSetupChange(instance.history, instance.store, () => {
        setPageOrientation(instance.store, sheet, next);
      });
    },
    [instance],
  );

  const onPaperSize = useCallback(
    (next: PaperSize) => {
      if (!instance) return;
      const sheet = instance.store.getState().data.sheetIndex;
      recordPageSetupChange(instance.history, instance.store, () => {
        setPaperSize(instance.store, sheet, next);
      });
    },
    [instance],
  );

  const onMarginPreset = useCallback(
    (next: MarginPreset) => {
      if (!instance) return;
      const sheet = instance.store.getState().data.sheetIndex;
      recordPageSetupChange(instance.history, instance.store, () => {
        setMarginPreset(instance.store, sheet, next);
      });
    },
    [instance],
  );

  // Insert tab > Format as Table — applies the default session table overlay
  // to the active range. Excel opens a style picker first; ours ships a
  // single default style today, so calling the command directly is honest
  // and skips a one-option dropdown.
  const onFormatAsTable = useCallback(() => {
    if (!instance) return;
    const r = instance.store.getState().selection.range;
    formatAsTable(instance.store, r);
  }, [instance]);

  const tool = (
    id: string,
    title: string,
    label: string | ReactElement,
    onClick: () => void,
    isActive = false,
    extra = '',
    disabled = false,
  ): ReactElement => (
    <button
      key={id}
      type="button"
      className={`demo__rb${extra}${isActive ? ' demo__rb--active' : ''}`}
      title={title}
      aria-label={title}
      aria-keyshortcuts={RIBBON_KEYSHORTCUTS[id]}
      onClick={onClick}
      disabled={disabled || !instance}
    >
      {label}
    </button>
  );

  const iconLabel = (icon: IconName, text: string): ReactElement => (
    <>
      <Icon name={icon} />
      <span>{text}</span>
    </>
  );

  const group = (title: string, children: ReactElement[], variant = ''): ReactElement => (
    <section
      key={`${title}-${variant || 'group'}`}
      className={`demo__ribbon-group${variant ? ` demo__ribbon-group--${variant}` : ''}`}
      aria-label={title}
    >
      <div className="demo__ribbon-tools">{children}</div>
      <div className="demo__ribbon-label">{title}</div>
    </section>
  );

  const rowBreak = (id: string): ReactElement => (
    <span key={id} className="demo__rb-break" aria-hidden="true" />
  );

  const select = (
    id: string,
    title: string,
    value: string | number,
    values: readonly (string | number)[],
    onChange: (value: string) => void,
    extra = '',
  ): ReactElement => (
    <Dropdown
      key={id}
      title={title}
      ariaKeyshortcuts={RIBBON_KEYSHORTCUTS[id]}
      value={value}
      options={values.map((v) => ({ value: v, label: String(v) }))}
      onChange={(v) => onChange(String(v))}
      disabled={!instance}
      className={extra.trim()}
      display={String(value)}
    />
  );

  const optionSelect = <T extends string>(
    id: string,
    title: string,
    value: T,
    options: readonly { value: T; label: string }[],
    onChange: (value: T) => void,
    extra = '',
  ): ReactElement => (
    <Dropdown<T>
      key={id}
      title={title}
      ariaKeyshortcuts={RIBBON_KEYSHORTCUTS[id]}
      value={value}
      options={options}
      onChange={onChange}
      disabled={!instance}
      className={extra.trim()}
    />
  );

  const color = (
    id: string,
    title: string,
    value: string,
    onChange: (value: string) => void,
    label: ReactElement,
  ): ReactElement => (
    <label
      key={id}
      className="demo__rb-color"
      title={title}
      aria-label={title}
      aria-keyshortcuts={RIBBON_KEYSHORTCUTS[id]}
    >
      <span>{label}</span>
      <input
        type="color"
        value={value}
        disabled={!instance}
        onChange={(e) => onChange(e.currentTarget.value)}
      />
    </label>
  );

  const ribbonGroups = buildRibbonGroups({
    active,
    borderPresets,
    borderStyle,
    borderStyles,
    color,
    group,
    iconLabel,
    instance,
    lang,
    onAutoSum,
    onBorderPreset,
    onCopy,
    onCut,
    onDeleteCols,
    onDeleteRows,
    onAddIn,
    onFilterToggle,
    onFormatAsTable,
    onFormatPainter,
    onFreezeToggle,
    onDrawEraser,
    onDrawPen,
    onInsertCols,
    onInsertRows,
    onMarginPreset,
    onMerge,
    onPageOrientation,
    onPaperSize,
    onPaste,
    onRedo,
    onRemoveDuplicates,
    onAccessibilityCheck,
    onRunScript,
    onSort,
    onSpellingReview,
    onTranslate,
    onToggleColsHidden,
    onToggleRowsHidden,
    onUndo,
    onZoom,
    optionSelect,
    rowBreak,
    select,
    setBorderStyle,
    tool,
    tr,
    wrapFormat,
  });

  return (
    <div className="demo__ribbon-shell">
      <div
        ref={tablistRef}
        className="demo__ribbon-tabs"
        role="tablist"
        aria-label={tr.ribbonTabs}
        onKeyDown={onRibbonTabKeyDown}
      >
        {ribbonTabs.map((tab) => (
          <button
            key={tab.id}
            type="button"
            className={`demo__ribbon-tab${activeTab === tab.id ? ' demo__ribbon-tab--active' : ''}`}
            role="tab"
            data-ribbon-tab={tab.id}
            aria-selected={activeTab === tab.id}
            tabIndex={activeTab === tab.id ? 0 : -1}
            onClick={() => onTabChange(tab.id)}
          >
            {tab.label}
          </button>
        ))}
      </div>
      <div
        className="demo__ribbon"
        role="toolbar"
        aria-label={`${RIBBON_TAB_LABELS[activeTab][lang]} ${tr.ribbon}`}
      >
        {ribbonGroups[activeTab]}
      </div>
    </div>
  );
};

export const Toolbar = SpreadsheetToolbar;
