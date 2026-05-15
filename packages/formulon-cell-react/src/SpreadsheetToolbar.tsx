import { ChevronDown12Regular } from '@fluentui/react-icons';
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
  type NumFmt,
  type PageOrientation,
  type PaperSize,
  recordFormatChange,
  recordMergesChangeWithEngine,
  recordPageSetupChange,
  removeDuplicates,
  type SpreadsheetInstance,
  setAlign,
  setAutoFilter,
  setBorderPreset,
  setFreezePanes,
  setMarginPreset,
  setNumFmt,
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

type MergeAction = 'mergeCenter' | 'mergeAcross' | 'mergeCells' | 'unmergeCells';
type NumberFormatAction =
  | 'general'
  | 'fixed'
  | 'currency'
  | 'accounting'
  | 'shortDate'
  | 'longDate'
  | 'time'
  | 'percent'
  | 'fraction'
  | 'scientific'
  | 'text'
  | 'more';

const numberFormatForAction = (action: NumberFormatAction, lang: 'ja' | 'en'): NumFmt | null => {
  const symbol = lang === 'ja' ? '¥' : '$';
  switch (action) {
    case 'general':
      return { kind: 'general' };
    case 'fixed':
      return { kind: 'fixed', decimals: 0 };
    case 'currency':
      return { kind: 'currency', decimals: 0, symbol };
    case 'accounting':
      return { kind: 'accounting', decimals: 0, symbol };
    case 'shortDate':
      return { kind: 'date', pattern: lang === 'ja' ? 'yyyy/m/d' : 'm/d/yyyy' };
    case 'longDate':
      return { kind: 'date', pattern: lang === 'ja' ? 'yyyy"年"m"月"d"日' : 'mmmm d, yyyy' };
    case 'time':
      return { kind: 'time', pattern: lang === 'ja' ? 'H:MM' : 'h:MM AM/PM' };
    case 'percent':
      return { kind: 'percent', decimals: 0 };
    case 'fraction':
      return { kind: 'custom', pattern: '# ?/?' };
    case 'scientific':
      return { kind: 'scientific', decimals: 2 };
    case 'text':
      return { kind: 'text' };
    case 'more':
      return null;
  }
};

const THEME_COLORS = [
  '#ffffff',
  '#000000',
  '#e7e6e6',
  '#44546a',
  '#5b9bd5',
  '#ed7d31',
  '#70ad47',
  '#4472c4',
  '#a64d79',
  '#70ad47',
  '#f2f2f2',
  '#7f7f7f',
  '#d9e2f3',
  '#d9eaf7',
  '#fce4d6',
  '#e2f0d9',
  '#d9e2f3',
  '#eadcf8',
  '#e2f0d9',
  '#d9d9d9',
  '#595959',
  '#b4c6e7',
  '#bdd7ee',
  '#f8cbad',
  '#c6e0b4',
  '#b4c6e7',
  '#d9bce3',
  '#c6e0b4',
  '#bfbfbf',
  '#404040',
  '#8eaadb',
  '#9dc3e6',
  '#f4b183',
  '#a9d18e',
  '#8eaadb',
  '#c27ba0',
  '#a9d18e',
  '#a6a6a6',
  '#262626',
  '#2f5597',
  '#2e75b6',
  '#c65911',
  '#548235',
  '#2f5597',
  '#741b47',
  '#548235',
] as const;

const STANDARD_COLORS = [
  '#c00000',
  '#ff0000',
  '#ffc000',
  '#ffff00',
  '#92d050',
  '#00b050',
  '#00b0f0',
  '#0070c0',
  '#002060',
  '#7030a0',
] as const;

const THEME_SWATCHES = THEME_COLORS.map((color, index) => ({
  color,
  id: `theme-${index}-${color}`,
}));

interface ColorDropdownProps {
  id: string;
  title: string;
  value: string;
  labels: {
    automatic: string;
    highContrastOnly: string;
    moreColors: string;
    standardColors: string;
    themeColors: string;
  };
  label: ReactElement;
  disabled: boolean;
  onChange: (value: string) => void;
}

function ColorDropdown({
  id,
  title,
  value,
  labels,
  label,
  disabled,
  onChange,
}: ColorDropdownProps): ReactElement {
  const [open, setOpen] = useState(false);
  const wrapRef = useRef<HTMLDivElement | null>(null);
  const inputRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    if (!open) return;
    const onDocDown = (e: MouseEvent): void => {
      if (e.target instanceof Node && wrapRef.current?.contains(e.target)) return;
      setOpen(false);
    };
    const onKey = (e: globalThis.KeyboardEvent): void => {
      if (e.key === 'Escape') setOpen(false);
    };
    document.addEventListener('mousedown', onDocDown, true);
    document.addEventListener('keydown', onKey, true);
    return () => {
      document.removeEventListener('mousedown', onDocDown, true);
      document.removeEventListener('keydown', onKey, true);
    };
  }, [open]);

  const choose = (next: string): void => {
    onChange(next);
    setOpen(false);
  };

  return (
    <div
      key={id}
      ref={wrapRef}
      className={`demo__rb-color${open ? ' demo__rb-color--open' : ''}`}
      title={title}
    >
      <button
        type="button"
        className="demo__rb-color__btn"
        aria-keyshortcuts={RIBBON_KEYSHORTCUTS[id]}
        aria-haspopup="menu"
        aria-expanded={open}
        disabled={disabled}
        onClick={() => setOpen((next) => !next)}
      >
        <span className="demo__rb-color__icon">{label}</span>
        <span className="demo__rb-color__swatch" style={{ backgroundColor: value }} />
        <ChevronDown12Regular className="demo__rb-color__chev" aria-hidden="true" />
      </button>
      {open ? (
        <div className="demo__color-menu" role="menu" aria-label={title}>
          <label className="demo__color-menu__check">
            <input type="checkbox" disabled />
            <span>{labels.highContrastOnly}</span>
          </label>
          <button
            className="demo__color-menu__auto"
            type="button"
            role="menuitem"
            onClick={() => choose('#000000')}
          >
            {labels.automatic}
          </button>
          <div className="demo__color-menu__section">{labels.themeColors}</div>
          <div className="demo__color-menu__grid demo__color-menu__grid--theme">
            {THEME_SWATCHES.map((swatch) => (
              <button
                key={swatch.id}
                type="button"
                className="demo__color-menu__chip"
                style={{ backgroundColor: swatch.color }}
                aria-label={swatch.color}
                onClick={() => choose(swatch.color)}
              />
            ))}
          </div>
          <div className="demo__color-menu__section">{labels.standardColors}</div>
          <div className="demo__color-menu__grid demo__color-menu__grid--standard">
            {STANDARD_COLORS.map((color) => (
              <button
                key={color}
                type="button"
                className="demo__color-menu__chip"
                style={{ backgroundColor: color }}
                aria-label={color}
                onClick={() => choose(color)}
              />
            ))}
          </div>
          <button
            type="button"
            className="demo__color-menu__more"
            role="menuitem"
            onClick={() => inputRef.current?.click()}
          >
            <span className="demo__color-menu__wheel" aria-hidden="true" />
            {labels.moreColors}
          </button>
          <input
            ref={inputRef}
            className="demo__color-menu__native"
            type="color"
            value={value}
            onChange={(e) => choose(e.currentTarget.value)}
          />
        </div>
      ) : null}
    </div>
  );
}

interface MergeMenuProps {
  disabled: boolean;
  labels: {
    mergeAndCenter: string;
    mergeAcross: string;
    mergeCells: string;
    unmergeCells: string;
  };
  onPick: (action: MergeAction) => void;
}

function MergeMenu({ disabled, labels, onPick }: MergeMenuProps): ReactElement {
  const [open, setOpen] = useState(false);
  const wrapRef = useRef<HTMLDivElement | null>(null);
  const options: readonly { action: MergeAction; label: string }[] = [
    { action: 'mergeCenter', label: labels.mergeAndCenter },
    { action: 'mergeAcross', label: labels.mergeAcross },
    { action: 'mergeCells', label: labels.mergeCells },
    { action: 'unmergeCells', label: labels.unmergeCells },
  ];

  useEffect(() => {
    if (!open) return;
    const onDocDown = (e: MouseEvent): void => {
      if (e.target instanceof Node && wrapRef.current?.contains(e.target)) return;
      setOpen(false);
    };
    const onKey = (e: globalThis.KeyboardEvent): void => {
      if (e.key === 'Escape') setOpen(false);
    };
    document.addEventListener('mousedown', onDocDown, true);
    document.addEventListener('keydown', onKey, true);
    return () => {
      document.removeEventListener('mousedown', onDocDown, true);
      document.removeEventListener('keydown', onKey, true);
    };
  }, [open]);

  return (
    <div ref={wrapRef} className={`demo__rb-menu${open ? ' demo__rb-menu--open' : ''}`}>
      <button
        type="button"
        className="demo__rb demo__rb-menu__btn"
        title={labels.mergeCells}
        aria-label={labels.mergeCells}
        aria-haspopup="menu"
        aria-expanded={open}
        disabled={disabled}
        onClick={() => setOpen((next) => !next)}
      >
        <Icon name="merge" />
        <ChevronDown12Regular className="demo__rb-menu__chev" aria-hidden="true" />
      </button>
      {open ? (
        <div className="demo__merge-menu" role="menu" aria-label={labels.mergeCells}>
          {options.map((option) => (
            <button
              key={option.action}
              type="button"
              className="demo__merge-menu__item"
              role="menuitem"
              onClick={() => {
                onPick(option.action);
                setOpen(false);
              }}
            >
              <Icon name="merge" />
              <span>{option.label}</span>
            </button>
          ))}
        </div>
      ) : null}
    </div>
  );
}

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

  const onMergeAction = useCallback(
    (action: MergeAction) => {
      if (!instance) return;
      const s = instance.store.getState();
      const r = s.selection.range;
      if (action === 'unmergeCells') {
        applyUnmerge(instance.store, instance.workbook, instance.history, r);
        return;
      }
      if (action === 'mergeAcross') {
        recordMergesChangeWithEngine(
          instance.history,
          instance.store,
          instance.workbook,
          r.sheet,
          () => {
            for (let row = r.r0; row <= r.r1; row += 1) {
              if (r.c0 === r.c1) continue;
              mutators.mergeRange(instance.store, {
                sheet: r.sheet,
                r0: row,
                c0: r.c0,
                r1: row,
                c1: r.c1,
              });
            }
          },
        );
        return;
      }
      applyMerge(instance.store, instance.workbook, instance.history, r);
      if (action === 'mergeCenter') {
        recordFormatChange(instance.history, instance.store, () =>
          setAlign(instance.store.getState(), instance.store, 'center'),
        );
      }
    },
    [instance],
  );

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

  const onNumberFormat = useCallback(
    (next: string) => {
      if (!instance) return;
      const action = next as NumberFormatAction;
      if (action === 'more') {
        instance.openFormatDialog();
        return;
      }
      const fmt = numberFormatForAction(action, lang);
      if (!fmt) return;
      wrapFormat((s, st) => setNumFmt(s, st, fmt));
    },
    [instance, lang, wrapFormat],
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
    <ColorDropdown
      key={id}
      id={id}
      title={title}
      value={value}
      labels={{
        automatic: tr.automatic,
        highContrastOnly: tr.highContrastOnly,
        moreColors: tr.moreColors,
        standardColors: tr.standardColors,
        themeColors: tr.themeColors,
      }}
      label={label}
      disabled={!instance}
      onChange={onChange}
    />
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
    mergeMenu: (
      <MergeMenu
        key="merge"
        disabled={!instance}
        labels={{
          mergeAndCenter: tr.mergeAndCenter,
          mergeAcross: tr.mergeAcross,
          mergeCells: tr.mergeCells,
          unmergeCells: tr.unmergeCells,
        }}
        onPick={onMergeAction}
      />
    ),
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
    onNumberFormat,
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
