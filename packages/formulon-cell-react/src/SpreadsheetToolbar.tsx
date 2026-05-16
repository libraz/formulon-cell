import { ChevronDown12Regular } from '@fluentui/react-icons';
import {
  type AutoSumFunction,
  addSheet,
  applyAdvancedFilter,
  applyCellStyle,
  applyConditionalPresetAction,
  applyMerge,
  applyTextScriptToRange,
  applyUnmerge,
  autofitColsWidth,
  autofitRowsHeight,
  autoSum,
  buildRibbonAddInReport,
  buildSpreadsheetCompatibilityReport,
  CELL_STYLE_GROUPS,
  CELL_STYLES,
  type CellBorderStyle,
  type CellDeleteAction,
  type CellInsertAction,
  type CellStyleId,
  type ConditionalMenuAction,
  type ConditionalPresetAction,
  circleInvalidValidationDataInSheet,
  clearPrintArea,
  clearPrintTitles,
  clearSheetBackgroundImage,
  clearTraceArrows,
  clearTraceArrowsByKind,
  clearValidationCircles,
  clearValidationInRangeWithEngine,
  clearWatchedCells,
  copyAdvancedFilterResult,
  createColorPalette,
  createDefinedNamesFromSelection,
  createRibbonChartFromSelection,
  deleteCells,
  deleteCols,
  deleteRows,
  deleteSelectedCols,
  deleteSelectedRows,
  dispatchHostClipboard,
  executeRibbonClearAction,
  executeRibbonCommentAction,
  executeRibbonFillAction,
  executeRibbonFilterDataAction,
  executeRibbonFindAction,
  executeRibbonFormulaAuditingAction,
  executeRibbonHyperlinkAction,
  executeRibbonPivotTableAction,
  executeRibbonProtectionAction,
  type FreezeAction,
  formatAsTable,
  groupCols,
  groupRows,
  handleAutoSum,
  handleAutoSumAction,
  handleConditionalAction,
  handleDeleteCellsAction,
  handleFreezeAction,
  handleInsertCellsAction,
  handleMergeAction,
  handlePasteAction,
  handleWindowAction,
  hiddenInSelection,
  hideCols,
  hideRows,
  inferSortHasHeader,
  insertCells,
  insertCols,
  insertDefinedNameFormula,
  insertManualPageBreak,
  insertRows,
  insertSelectedCols,
  insertSelectedRows,
  isCellWritable,
  isWorkbookStructureProtected,
  listDefinedNames,
  type MarginPreset,
  type MergeAction,
  makeRangeResolver,
  moveSheet,
  mutators,
  type NumberFormatAction,
  numberFormatForAction,
  type PageOrientation,
  type PaperSize,
  type PasteAction,
  recordConditionalRulesChange,
  recordDefinedNamesChange,
  recordFilterChange,
  recordFormatChange,
  recordLayoutChange,
  recordMergesChangeWithEngine,
  recordPageSetupChange,
  recordTablesChange,
  recordValidationCirclesChange,
  recordWatchesChange,
  removeDuplicates,
  removeManualPageBreak,
  removeSheet,
  renameSheet,
  resetManualPageBreaks,
  resolveRibbonPdfAction,
  type ScriptCommand,
  type SpreadsheetInstance,
  type Strings,
  setAlign,
  setBorderPreset,
  setColsWidth,
  setFitToPages,
  setFreezePanes,
  setMarginPreset,
  setNumFmt,
  setPageOrientation,
  setPageScale,
  setPaperSize,
  setPrintArea,
  setPrintTitleCols,
  setPrintTitleRows,
  setRotation,
  setRowsHeight,
  setSheetBackgroundImage,
  setSheetHidden,
  setSheetZoom,
  setWorkbookStructureProtected,
  showColsAroundSelection,
  showRowsAroundSelection,
  sortActiveColumnAuto,
  sortRangeWithHistory,
  textToColumns,
  toggleAutoFilterFromSelection,
  toggleSelectedColsHidden,
  toggleSelectedRowsHidden,
  ungroupCols,
  ungroupRows,
  unwatchCell,
  type WindowAction,
  warnProtected,
  watchRange,
} from '@libraz/formulon-cell';
import {
  type ChangeEvent,
  Fragment,
  type KeyboardEvent,
  type ReactElement,
  useCallback,
  useEffect,
  useRef,
  useState,
} from 'react';
import { useI18n } from './hooks.js';
import { Dropdown } from './toolbar/Dropdown.js';
import { buildRibbonGroups } from './toolbar/groups.js';
import { Icon, type IconName } from './toolbar/icons.js';
import {
  type ActiveState,
  type BorderPreset,
  EMPTY_ACTIVE_STATE,
  localizeBorderPresets,
  localizeBorderStyles,
  projectActiveState,
  RIBBON_KEYSHORTCUTS,
  RIBBON_TABS,
  type RibbonTab,
  type SpreadsheetToolbarProps,
} from './toolbar/model.js';
import { dictionaries, dictionaryLocaleFor } from './toolbar/translations.js';

export type { RibbonTab, SpreadsheetToolbarProps } from './toolbar/model.js';

import {
  type AddInAction,
  type AdvancedFilterDialogDraft,
  type AutomationRunDraft,
  type AutoSumAction,
  type CalculationAction,
  CELL_STYLE_SECTION_ACTION_PREFIX,
  type CellFormatAction,
  type CellStyleAction,
  type ChartAction,
  type ClearAction,
  type ClearArrowsAction,
  type CommentAction,
  cellLabel,
  colLetter,
  type DataValidationAction,
  type DefinedNameAction,
  type DimensionDialogDraft,
  type FillAction,
  type FilterDataAction,
  type FindAction,
  type FormatTableAction,
  type FormulaAuditingAction,
  type FunctionAction,
  formatA1Range,
  type HyperlinkAction,
  MORE_SYMBOL_ACTION,
  type OutlineAxisAction,
  type PageBreakAction,
  type PdfAction,
  type PictureAction,
  type PivotTableAction,
  type PrintAreaAction,
  type PrintTitleAction,
  type ProtectionAction,
  parseA1Range,
  type RemoveDuplicatesDialogDraft,
  type RibbonReportDialogDraft,
  type ScreenshotAction,
  type ScriptDialogDraft,
  SHEET_TAB_COLOR_ACTIONS,
  type ShapeAction,
  type SheetBackgroundAction,
  type SheetCell,
  type SheetRange,
  type SheetRenameDialogDraft,
  type SortAction,
  type SortDialogDraft,
  type SymbolAction,
  TEXT_TO_COLUMNS_DIALOG_KEYS,
  type TextOrientationAction,
  type TextToColumnsAction,
  type TextToColumnsDialogDraft,
  type ThemeAction,
  type WatchAction,
} from '@libraz/formulon-cell';

import { CellMenu } from './toolbar/menus/CellMenu.js';
import { ColorDropdown } from './toolbar/menus/ColorDropdown.js';
import { ConditionalMenu } from './toolbar/menus/ConditionalMenu.js';
import { MergeMenu } from './toolbar/menus/MergeMenu.js';
export const SpreadsheetToolbar = ({
  instance,
  features,
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
  onNewWorkbook,
  onOpenWorkbook,
  onSaveWorkbook,
  onSaveWorkbookAs,
}: SpreadsheetToolbarProps): ReactElement => {
  const [active, setActive] = useState<ActiveState>(EMPTY_ACTIVE_STATE);
  const [borderStyle, setBorderStyle] = useState<CellBorderStyle>('thin');
  const [borderColor, setBorderColor] = useState('#000000');
  const [ribbonCollapsed, setRibbonCollapsed] = useState(false);
  const [ribbonDisplayMenuOpen, setRibbonDisplayMenuOpen] = useState(false);
  const tablistRef = useRef<HTMLDivElement | null>(null);
  const ribbonDisplayRef = useRef<HTMLDivElement | null>(null);
  const sheetBackgroundInputRef = useRef<HTMLInputElement | null>(null);
  const previousNonFileTabRef = useRef<RibbonTab>('home');
  const i18n = useI18n(instance);
  const lang = dictionaryLocaleFor(locale);
  const liveLang = dictionaryLocaleFor(i18n.locale);
  const strings =
    i18n.strings && (i18n.locale === locale || liveLang === lang)
      ? i18n.strings
      : dictionaries[lang];
  const tr = strings.ribbon;
  const cellMenuText = strings.ribbonMenu;
  const viewToolbarText = strings.viewToolbar;
  const borderPresets = localizeBorderPresets(tr);
  const borderStyles = localizeBorderStyles(tr);
  const ribbonTabs = RIBBON_TABS.map((id) => ({
    id,
    label: strings.ribbon.tabs[id],
  }));

  const focusRibbonTab = useCallback((tab: RibbonTab) => {
    requestAnimationFrame(() => {
      tablistRef.current
        ?.querySelector<HTMLButtonElement>(`[data-ribbon-tab="${tab}"]`)
        ?.focus({ preventScroll: true });
    });
  }, []);

  useEffect(() => {
    const onGlobalKeyDown = (event: globalThis.KeyboardEvent): void => {
      if (event.key !== 'F1' || (!event.ctrlKey && !event.metaKey)) return;
      event.preventDefault();
      setRibbonDisplayMenuOpen(false);
      setRibbonCollapsed((value) => !value);
    };
    window.addEventListener('keydown', onGlobalKeyDown);
    return () => window.removeEventListener('keydown', onGlobalKeyDown);
  }, []);

  useEffect(() => {
    if (activeTab !== 'file') previousNonFileTabRef.current = activeTab;
  }, [activeTab]);

  const closeBackstage = useCallback(() => {
    onTabChange(previousNonFileTabRef.current);
    focusRibbonTab(previousNonFileTabRef.current);
  }, [focusRibbonTab, onTabChange]);

  useEffect(() => {
    if (activeTab !== 'file') return;
    const onEscape = (event: globalThis.KeyboardEvent): void => {
      if (event.key !== 'Escape') return;
      event.preventDefault();
      closeBackstage();
    };
    window.addEventListener('keydown', onEscape);
    return () => window.removeEventListener('keydown', onEscape);
  }, [activeTab, closeBackstage]);

  useEffect(() => {
    if (!ribbonDisplayMenuOpen) return;
    const onEscape = (event: globalThis.KeyboardEvent): void => {
      if (event.key !== 'Escape') return;
      event.preventDefault();
      setRibbonDisplayMenuOpen(false);
    };
    const onPointerDown = (event: PointerEvent): void => {
      const target = event.target as Node | null;
      if (target && ribbonDisplayRef.current?.contains(target)) return;
      setRibbonDisplayMenuOpen(false);
    };
    window.addEventListener('keydown', onEscape);
    document.addEventListener('pointerdown', onPointerDown, true);
    return () => {
      window.removeEventListener('keydown', onEscape);
      document.removeEventListener('pointerdown', onPointerDown, true);
    };
  }, [ribbonDisplayMenuOpen]);

  const onRibbonTabClick = useCallback(
    (tab: RibbonTab) => {
      setRibbonDisplayMenuOpen(false);
      if (tab !== 'file') previousNonFileTabRef.current = tab;
      onTabChange(tab);
    },
    [onTabChange],
  );

  const onRibbonTabKeyDown = useCallback(
    (event: KeyboardEvent<HTMLElement>) => {
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
      if (nextTab !== 'file') previousNonFileTabRef.current = nextTab;
      onTabChange(nextTab);
      focusRibbonTab(nextTab);
    },
    [activeTab, focusRibbonTab, onTabChange, ribbonTabs],
  );

  const focusRibbonDisplayOption = useCallback((index: number): void => {
    requestAnimationFrame(() => {
      const options = Array.from(
        ribbonDisplayRef.current?.querySelectorAll<HTMLButtonElement>(
          '.demo__ribbon-display-option',
        ) ?? [],
      );
      options[(index + options.length) % options.length]?.focus({ preventScroll: true });
    });
  }, []);

  const onRibbonDisplayKeyDown = useCallback(
    (event: KeyboardEvent<HTMLElement>) => {
      const options = Array.from(
        ribbonDisplayRef.current?.querySelectorAll<HTMLButtonElement>(
          '.demo__ribbon-display-option',
        ) ?? [],
      );
      const activeIndex = Math.max(0, options.indexOf(document.activeElement as HTMLButtonElement));
      if (event.key === 'ArrowDown') {
        event.preventDefault();
        if (!ribbonDisplayMenuOpen) {
          setRibbonDisplayMenuOpen(true);
          focusRibbonDisplayOption(0);
          return;
        }
        focusRibbonDisplayOption(activeIndex + 1);
      } else if (event.key === 'ArrowUp') {
        event.preventDefault();
        if (!ribbonDisplayMenuOpen) {
          setRibbonDisplayMenuOpen(true);
          focusRibbonDisplayOption(-1);
          return;
        }
        focusRibbonDisplayOption(activeIndex - 1);
      } else if (event.key === 'Home' && options.length) {
        event.preventDefault();
        focusRibbonDisplayOption(0);
      } else if (event.key === 'End' && options.length) {
        event.preventDefault();
        focusRibbonDisplayOption(options.length - 1);
      }
    },
    [focusRibbonDisplayOption, ribbonDisplayMenuOpen],
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
  // Clipboard actions delegate to the host element so the copy/cut/paste
  // listeners run with a real selection. `dispatchHostClipboard` handles
  // browser quirks (execCommand can throw on some engines).
  const onCopy = useCallback(() => dispatchHostClipboard(instance, 'copy'), [instance]);
  const onCut = useCallback(() => dispatchHostClipboard(instance, 'cut'), [instance]);
  const onPaste = useCallback(() => dispatchHostClipboard(instance, 'paste'), [instance]);
  const onPasteAction = useCallback(
    (action: PasteAction) => handlePasteAction(instance, action),
    [instance],
  );
  const onFormatPainter = useCallback(() => {
    instance?.formatPainter?.activate(false);
  }, [instance]);

  const onAutoSum = useCallback(
    (functionName: AutoSumFunction = 'SUM') => {
      handleAutoSum(instance, functionName);
    },
    [instance],
  );

  const onAutoSumAction = useCallback(
    (action: AutoSumAction) => {
      handleAutoSumAction(instance, action);
    },
    [instance],
  );

  const onMergeAction = useCallback(
    (action: MergeAction) => handleMergeAction(instance, action),
    [instance],
  );

  const onBorderPreset = useCallback(
    (preset: BorderPreset) => {
      wrapFormat((s, st) => {
        setBorderPreset(s, st, preset, borderStyle, borderColor);
      });
    },
    [borderColor, borderStyle, wrapFormat],
  );

  const onBorderStyleChange = useCallback(
    (next: CellBorderStyle) => {
      setBorderStyle(next);
      instance?.borderDraw?.setStyle(next);
    },
    [instance],
  );

  const onBorderColorChange = useCallback(
    (next: string) => {
      setBorderColor(next);
      instance?.borderDraw?.setColor(next);
    },
    [instance],
  );

  const onFreezeAction = useCallback(
    (action: FreezeAction) => handleFreezeAction(instance, action),
    [instance],
  );

  const onInsertRows = useCallback(() => insertSelectedRows(instance), [instance]);
  const onDeleteRows = useCallback(() => deleteSelectedRows(instance), [instance]);
  const onInsertCols = useCallback(() => insertSelectedCols(instance), [instance]);
  const onDeleteCols = useCallback(() => deleteSelectedCols(instance), [instance]);

  const onInsertCellsAction = useCallback(
    (action: CellInsertAction) => handleInsertCellsAction(instance, action),
    [instance],
  );
  const onDeleteCellsAction = useCallback(
    (action: CellDeleteAction) => handleDeleteCellsAction(instance, action),
    [instance],
  );

  const onToggleRowsHidden = useCallback(() => toggleSelectedRowsHidden(instance), [instance]);
  const onToggleColsHidden = useCallback(() => toggleSelectedColsHidden(instance), [instance]);
  const onWindowAction = useCallback(
    (action: WindowAction) => handleWindowAction(instance, action),
    [instance],
  );

  const [dimensionDialog, setDimensionDialog] = useState<DimensionDialogDraft | null>(null);
  const [sheetRenameDialog, setSheetRenameDialog] = useState<SheetRenameDialogDraft | null>(null);
  const sheetRenameInputRef = useRef<HTMLInputElement | null>(null);
  const [advancedFilterDialog, setAdvancedFilterDialog] =
    useState<AdvancedFilterDialogDraft | null>(null);
  const [zoomDialog, setZoomDialog] = useState<string | null>(null);
  const [ribbonReportDialog, setRibbonReportDialog] = useState<RibbonReportDialogDraft | null>(
    null,
  );

  const onCellFormatAction = useCallback(
    (action: CellFormatAction) => {
      if (!instance) return;
      const state = instance.store.getState();
      const r = state.selection.range;
      if (action === 'dialog') instance.openFormatDialog();
      else if (action === 'rowHeight') {
        const current = state.layout.rowHeights.get(r.r0) ?? state.layout.defaultRowHeight;
        setDimensionDialog({ kind: 'rowHeight', value: String(current) });
      } else if (action === 'colWidth') {
        const current = state.layout.colWidths.get(r.c0) ?? state.layout.defaultColWidth;
        setDimensionDialog({ kind: 'colWidth', value: String(current) });
      } else if (action === 'autoFitRowHeight') {
        autofitRowsHeight(instance.store, instance.history, r.r0, r.r1, instance.workbook);
      } else if (action === 'autoFitColWidth') {
        autofitColsWidth(instance.store, instance.history, r.c0, r.c1, instance.workbook);
      } else if (action === 'protectSheet') instance.toggleSheetProtection();
      else if (action === 'hideRows')
        hideRows(instance.store, instance.history, r.r0, r.r1, instance.workbook);
      else if (action === 'showRows')
        showRowsAroundSelection(instance.store, instance.history, r.r0, r.r1, instance.workbook);
      else if (action === 'hideCols')
        hideCols(instance.store, instance.history, r.c0, r.c1, instance.workbook);
      else if (action === 'showCols')
        showColsAroundSelection(instance.store, instance.history, r.c0, r.c1, instance.workbook);
      else if (action === 'renameSheet') {
        setSheetRenameDialog({ value: instance.workbook.sheetName(state.data.sheetIndex) });
      } else if (action === 'hideSheet') {
        setSheetHidden(
          instance.store,
          instance.workbook,
          instance.history,
          state.data.sheetIndex,
          true,
        );
      } else if (action === 'unhideSheet') {
        const firstHidden = [...state.layout.hiddenSheets].sort((a, b) => a - b)[0];
        if (firstHidden != null) {
          setSheetHidden(instance.store, instance.workbook, instance.history, firstHidden, false);
        }
      } else if (action === 'moveSheetLeft') {
        const sheet = state.data.sheetIndex;
        if (sheet > 0)
          moveSheet(instance.store, instance.workbook, sheet, sheet - 1, instance.history);
      } else if (action === 'moveSheetRight') {
        const sheet = state.data.sheetIndex;
        if (sheet < instance.workbook.sheetCount - 1) {
          moveSheet(instance.store, instance.workbook, sheet, sheet + 1, instance.history);
        }
      } else if (action === 'tabColorNone') {
        recordLayoutChange(instance.history, instance.store, () => {
          mutators.setSheetTabColor(instance.store, state.data.sheetIndex, null);
        });
      } else if (action.startsWith('tabColor')) {
        const entry = SHEET_TAB_COLOR_ACTIONS.find((item) => item.action === action);
        if (entry) {
          recordLayoutChange(instance.history, instance.store, () => {
            mutators.setSheetTabColor(instance.store, state.data.sheetIndex, entry.color);
          });
        }
      }
    },
    [instance],
  );

  const applyDimensionDialog = useCallback(() => {
    if (!instance || !dimensionDialog) return;
    const px = Number.parseFloat(dimensionDialog.value);
    if (!Number.isFinite(px) || px <= 0) return;
    const range = instance.store.getState().selection.range;
    if (dimensionDialog.kind === 'rowHeight') {
      setRowsHeight(instance.store, instance.history, range.r0, range.r1, px, instance.workbook);
    } else {
      setColsWidth(instance.store, instance.history, range.c0, range.c1, px, instance.workbook);
    }
    setDimensionDialog(null);
  }, [dimensionDialog, instance]);

  const applySheetRenameDialog = useCallback(() => {
    if (!instance || !sheetRenameDialog) return;
    const name = (sheetRenameInputRef.current?.value ?? sheetRenameDialog.value).trim();
    if (!name) return;
    renameSheet(
      instance.workbook,
      instance.store.getState().data.sheetIndex,
      name,
      instance.store,
      instance.history,
    );
    setSheetRenameDialog(null);
  }, [instance, sheetRenameDialog]);

  const onFillAction = useCallback(
    (action: FillAction) => {
      if (!instance) return;
      executeRibbonFillAction({
        store: instance.store,
        workbook: instance.workbook,
        history: instance.history,
        action,
      });
    },
    [instance],
  );

  const onClearAction = useCallback(
    (action: ClearAction) => {
      if (!instance) return;
      executeRibbonClearAction({
        store: instance.store,
        workbook: instance.workbook,
        history: instance.history,
        action,
      });
    },
    [instance],
  );

  const onFilterToggle = useCallback(() => {
    if (!instance) return;
    toggleAutoFilterFromSelection(instance.store, instance.history);
  }, [instance]);

  const [removeDuplicatesDialog, setRemoveDuplicatesDialog] =
    useState<RemoveDuplicatesDialogDraft | null>(null);

  const onRemoveDuplicates = useCallback(() => {
    if (!instance) return;
    const s = instance.store.getState();
    const range = s.selection.range;
    setRemoveDuplicatesDialog({
      columns: Array.from({ length: range.c1 - range.c0 + 1 }, (_, i) => range.c0 + i),
      hasHeader: inferSortHasHeader(s, range),
    });
  }, [instance]);

  const applyRemoveDuplicatesDialog = useCallback(() => {
    if (!instance || !removeDuplicatesDialog) return;
    const s = instance.store.getState();
    if (removeDuplicatesDialog.columns.length === 0) {
      setRibbonReportDialog({
        title: cellMenuText.removeDuplicatesDialogTitle,
        items: [
          {
            severity: 'warning',
            label: cellMenuText.removeDuplicatesNoColumns,
            detail: '',
          },
        ],
      });
      return;
    }
    instance.history.begin();
    let removed = 0;
    try {
      removed = removeDuplicates(s, instance.store, instance.workbook, s.selection.range, {
        columns: removeDuplicatesDialog.columns,
        hasHeader: removeDuplicatesDialog.hasHeader,
      });
    } finally {
      instance.history.end();
    }
    if (removed > 0) {
      mutators.replaceCells(instance.store, instance.workbook.cells(s.data.sheetIndex));
    }
    setRemoveDuplicatesDialog(null);
  }, [
    cellMenuText.removeDuplicatesDialogTitle,
    cellMenuText.removeDuplicatesNoColumns,
    instance,
    removeDuplicatesDialog,
  ]);

  const [sortDialog, setSortDialog] = useState<SortDialogDraft | null>(null);

  const onCustomSort = useCallback(() => {
    if (!instance) return;
    const s = instance.store.getState();
    const range = s.selection.range;
    setSortDialog({
      byCol:
        s.selection.active.col >= range.c0 && s.selection.active.col <= range.c1
          ? s.selection.active.col
          : range.c0,
      direction: 'asc',
      hasHeader: range.r0 < range.r1,
    });
  }, [instance]);

  const applyCustomSort = useCallback(() => {
    if (!instance || !sortDialog) return;
    sortRangeWithHistory({
      store: instance.store,
      workbook: instance.workbook,
      history: instance.history,
      range: instance.store.getState().selection.range,
      options: {
        byCol: sortDialog.byCol,
        direction: sortDialog.direction,
        hasHeader: sortDialog.hasHeader,
      },
    });
    setSortDialog(null);
  }, [instance, sortDialog]);

  const onSortMenuAction = useCallback(
    (action: SortAction) => {
      if (!instance) return;
      const s = instance.store.getState();
      if (action === 'asc' || action === 'desc') {
        sortActiveColumnAuto({
          store: instance.store,
          workbook: instance.workbook,
          history: instance.history,
          direction: action,
        });
      } else if (action === 'custom') onCustomSort();
      else if (action === 'filter') onFilterToggle();
      else if (action === 'filter-clear' && s.ui.filterRange)
        executeRibbonFilterDataAction({
          store: instance.store,
          history: instance.history,
          action: 'clear',
        });
      else if (action === 'filter-reapply')
        executeRibbonFilterDataAction({
          store: instance.store,
          history: instance.history,
          action: 'reapply',
        });
      else if (action === 'filter-by-selected')
        executeRibbonFilterDataAction({
          store: instance.store,
          history: instance.history,
          action: 'filter-by-selected',
        });
      else if (action === 'filter-advanced')
        setAdvancedFilterDialog({
          listRange: formatA1Range(s.selection.range),
          criteriaRange: '',
          copyTo: '',
          uniqueOnly: false,
        });
      else if (action === 'dedupe') onRemoveDuplicates();
      else if (action === 'conditional') instance.openCfRulesDialog();
      else if (action === 'named') instance.openNamedRangeDialog();
    },
    [instance, onCustomSort, onFilterToggle, onRemoveDuplicates],
  );

  const onFilterDataAction = useCallback(
    (action: FilterDataAction) => {
      if (!instance) return;
      const result = executeRibbonFilterDataAction({
        store: instance.store,
        history: instance.history,
        action,
      });
      if (result.kind === 'open-advanced') {
        setAdvancedFilterDialog({
          listRange: formatA1Range(result.range),
          criteriaRange: '',
          copyTo: '',
          uniqueOnly: false,
        });
      } else if (result.kind === 'open-filter-dropdown') {
        instance.openFilterDropdown(result.range, result.column);
      }
    },
    [instance],
  );

  const applyAdvancedFilterDialog = useCallback(() => {
    if (!instance || !advancedFilterDialog) return;
    const state = instance.store.getState();
    const sheet = state.data.sheetIndex;
    const sheetName = instance.workbook.sheetName(sheet);
    const listRange = parseA1Range(advancedFilterDialog.listRange, sheet, sheetName);
    const criteriaRange = parseA1Range(advancedFilterDialog.criteriaRange, sheet, sheetName);
    if (!listRange || !criteriaRange) return;
    const copyToRange = advancedFilterDialog.copyTo.trim()
      ? parseA1Range(advancedFilterDialog.copyTo, sheet, sheetName)
      : null;
    if (advancedFilterDialog.copyTo.trim()) {
      if (!copyToRange) return;
      instance.history.begin();
      let copied = 0;
      try {
        copied = copyAdvancedFilterResult(
          instance.store.getState(),
          instance.store,
          listRange,
          criteriaRange,
          { sheet, row: copyToRange.r0, col: copyToRange.c0 },
          { uniqueOnly: advancedFilterDialog.uniqueOnly },
          instance.workbook,
        );
      } finally {
        instance.history.end();
      }
      setRibbonReportDialog({
        title: cellMenuText.advancedFilterDialogTitle,
        items: [
          {
            severity: 'info',
            label: cellMenuText.filterAdvanced,
            detail: cellMenuText.advancedFilterCopiedStatus.replace('{count}', String(copied)),
          },
        ],
      });
    } else {
      recordFilterChange(instance.history, instance.store, () =>
        applyAdvancedFilter(instance.store.getState(), instance.store, listRange, criteriaRange),
      );
    }
    setAdvancedFilterDialog(null);
  }, [advancedFilterDialog, cellMenuText, instance]);

  const onTextOrientationAction = useCallback(
    (action: TextOrientationAction) => {
      if (action === 'formatAlignment') {
        instance?.openFormatDialog('align');
        return;
      }
      const rotation =
        action === 'angleCounterclockwise'
          ? 45
          : action === 'angleClockwise'
            ? -45
            : action === 'rotateTextUp' || action === 'verticalText'
              ? 90
              : action === 'rotateTextDown'
                ? -90
                : 0;
      wrapFormat((s, st) => setRotation(s, st, rotation));
    },
    [instance, wrapFormat],
  );

  const [textToColumnsDialog, setTextToColumnsDialog] = useState<TextToColumnsDialogDraft | null>(
    null,
  );

  const applyTextToColumns = useCallback(
    (delimiters: readonly string[], collapseConsecutive = false) => {
      if (!instance || delimiters.length === 0) return;
      const state = instance.store.getState();
      instance.history.begin();
      let max = 0;
      try {
        recordFormatChange(instance.history, instance.store, () => {
          max = textToColumns(
            state,
            instance.store,
            instance.workbook,
            state.selection.range,
            delimiters,
            { collapseConsecutiveDelimiters: collapseConsecutive },
          );
        });
      } finally {
        instance.history.end();
      }
      if (max > 0) {
        mutators.replaceCells(instance.store, instance.workbook.cells(state.data.sheetIndex));
      }
    },
    [instance],
  );

  const onTextToColumnsAction = useCallback(
    (action: TextToColumnsAction) => {
      if (!instance) return;
      if (action === 'custom') {
        setTextToColumnsDialog({
          comma: true,
          tab: false,
          semicolon: false,
          space: false,
          collapseConsecutive: false,
        });
        return;
      }
      const delimiter =
        action === 'tab' ? '\t' : action === 'semicolon' ? ';' : action === 'space' ? ' ' : ',';
      applyTextToColumns([delimiter]);
    },
    [applyTextToColumns, instance],
  );

  const applyTextToColumnsDialog = useCallback(() => {
    if (!textToColumnsDialog) return;
    const delimiters = [
      textToColumnsDialog.comma ? ',' : '',
      textToColumnsDialog.tab ? '\t' : '',
      textToColumnsDialog.semicolon ? ';' : '',
      textToColumnsDialog.space ? ' ' : '',
    ].filter(Boolean);
    applyTextToColumns(delimiters, textToColumnsDialog.collapseConsecutive);
    setTextToColumnsDialog(null);
  }, [applyTextToColumns, textToColumnsDialog]);

  const onDataValidationAction = useCallback(
    (action: DataValidationAction) => {
      if (!instance) return;
      if (action === 'settings') {
        instance.openDataValidationDialog();
        return;
      }
      if (action === 'clearValidation') {
        const state = instance.store.getState();
        clearValidationInRangeWithEngine(
          instance.store,
          instance.history,
          instance.workbook,
          state.selection.range,
        );
        return;
      }
      if (action === 'clearCircles') {
        recordValidationCirclesChange(instance.history, instance.store, () => {
          clearValidationCircles(instance.store);
        });
        return;
      }
      const state = instance.store.getState();
      recordValidationCirclesChange(instance.history, instance.store, () => {
        circleInvalidValidationDataInSheet(
          instance.store,
          state.selection.range.sheet,
          makeRangeResolver(instance.workbook, state.data.sheetIndex),
        );
      });
    },
    [instance],
  );

  const onFormulaAuditingAction = useCallback(
    (action: FormulaAuditingAction) => {
      if (!instance) return;
      const result = executeRibbonFormulaAuditingAction({
        store: instance.store,
        workbook: instance.workbook,
        history: instance.history,
        action,
        strings: { errorChecking: cellMenuText.errorChecking },
      });
      if (result.kind === 'trace-precedents') instance.tracePrecedents();
      else if (result.kind === 'report') setRibbonReportDialog(result.report);
    },
    [cellMenuText.errorChecking, instance],
  );

  const onClearArrowsAction = useCallback(
    (action: ClearArrowsAction) => {
      if (!instance) return;
      if (action === 'clear-precedents') {
        clearTraceArrowsByKind(instance.store, 'precedent', instance.history);
        return;
      }
      if (action === 'clear-dependents') {
        clearTraceArrowsByKind(instance.store, 'dependent', instance.history);
        return;
      }
      clearTraceArrows(instance.store, instance.history);
    },
    [instance],
  );

  const onCellStyleAction = useCallback(
    (action: CellStyleAction) => {
      if (!instance) return;
      const state = instance.store.getState();
      applyCellStyle(instance.store, instance.history, state.selection.range, action);
    },
    [instance],
  );

  const onFindAction = useCallback(
    (action: FindAction) => {
      if (!instance) return;
      const result = executeRibbonFindAction({
        store: instance.store,
        workbook: instance.workbook,
        action,
        strings: {
          findSelect: cellMenuText.findSelect,
          findNoMatches: cellMenuText.findNoMatches,
          commentNone: cellMenuText.commentNone,
        },
      });
      if (result.kind === 'open-find') instance.openFindReplace(result.mode);
      else if (result.kind === 'open-go-to') instance.openGoTo();
      else if (result.kind === 'open-go-to-special') instance.openGoToSpecial();
      else if (result.kind === 'report') setRibbonReportDialog(result.report);
    },
    [cellMenuText.commentNone, cellMenuText.findNoMatches, cellMenuText.findSelect, instance],
  );

  const onCommentAction = useCallback(
    (action: CommentAction) => {
      if (!instance) return;
      executeRibbonCommentAction({
        store: instance.store,
        workbook: instance.workbook,
        history: instance.history,
        action,
      });
    },
    [instance],
  );

  const onProtectionAction = useCallback(
    (action: ProtectionAction) => {
      if (!instance) return;
      const report = executeRibbonProtectionAction({
        store: instance.store,
        action,
        strings: cellMenuText,
      });
      setRibbonReportDialog(report);
    },
    [cellMenuText, instance],
  );

  const onHyperlinkAction = useCallback(
    (action: HyperlinkAction) => {
      if (!instance) return;
      const result = executeRibbonHyperlinkAction({
        store: instance.store,
        workbook: instance.workbook,
        history: instance.history,
        action,
        strings: cellMenuText,
      });
      if (result.kind === 'open-hyperlink-dialog') instance.openHyperlinkDialog();
      else if (result.kind === 'open-external-dialog') instance.openExternalLinksDialog();
      else if (result.kind === 'open-url') window.open(result.url, '_blank', 'noopener,noreferrer');
      else if (result.kind === 'report') setRibbonReportDialog(result.report);
    },
    [cellMenuText, instance],
  );

  const onFunctionAction = useCallback(
    (action: FunctionAction) => {
      instance?.openFunctionArguments(action);
    },
    [instance],
  );

  const onOutlineGroupAction = useCallback(
    (axis: OutlineAxisAction) => {
      if (!instance) return;
      const range = instance.store.getState().selection.range;
      if (axis === 'rows') {
        groupRows(instance.store, instance.history, range.r0, range.r1, instance.workbook);
      } else {
        groupCols(instance.store, instance.history, range.c0, range.c1, instance.workbook);
      }
    },
    [instance],
  );

  const onOutlineUngroupAction = useCallback(
    (axis: OutlineAxisAction) => {
      if (!instance) return;
      const range = instance.store.getState().selection.range;
      if (axis === 'rows') {
        ungroupRows(instance.store, instance.history, range.r0, range.r1, instance.workbook);
      } else {
        ungroupCols(instance.store, instance.history, range.c0, range.c1, instance.workbook);
      }
    },
    [instance],
  );

  const onChartAction = useCallback(
    (action: ChartAction) => {
      if (!instance) return;
      createRibbonChartFromSelection({
        store: instance.store,
        history: instance.history,
        range: instance.store.getState().selection.range,
        action,
        idPrefix: 'react-ribbon-chart',
      });
    },
    [instance],
  );

  const onSymbolAction = useCallback(
    (symbol: SymbolAction) => {
      if (!instance) return;
      const addr = instance.store.getState().selection.active;
      if (instance.workbook.cellFormula(addr)) return;
      if (!isCellWritable(instance.store.getState(), addr)) {
        warnProtected(addr);
        return;
      }
      const text =
        symbol === MORE_SYMBOL_ACTION
          ? typeof window.prompt === 'function'
            ? (window.prompt(cellMenuText.symbolPrompt, '')?.trim() ?? '')
            : ''
          : symbol;
      if (text.length === 0) {
        if (symbol === MORE_SYMBOL_ACTION) {
          setRibbonReportDialog({
            title: cellMenuText.symbol,
            items: [
              {
                severity: 'warning',
                label: cellMenuText.symbolMore,
                detail: cellMenuText.symbolInvalid,
              },
            ],
          });
        }
        return;
      }
      const value = instance.workbook.getValue(addr);
      const current = value.kind === 'text' ? value.value : '';
      instance.history.begin();
      try {
        instance.workbook.setText(addr, `${current}${text}`);
      } finally {
        instance.history.end();
      }
      mutators.replaceCells(instance.store, instance.workbook.cells(addr.sheet));
    },
    [
      cellMenuText.symbol,
      cellMenuText.symbolInvalid,
      cellMenuText.symbolMore,
      cellMenuText.symbolPrompt,
      instance,
    ],
  );

  const onPrintAreaAction = useCallback(
    (action: PrintAreaAction) => {
      if (!instance) return;
      const state = instance.store.getState();
      const sheet = state.data.sheetIndex;
      const range = state.selection.range;
      recordPageSetupChange(instance.history, instance.store, () => {
        if (action === 'clear') {
          clearPrintArea(instance.store, sheet);
          return;
        }
        const start = `${colLetter(range.c0)}${range.r0 + 1}`;
        const end = `${colLetter(range.c1)}${range.r1 + 1}`;
        setPrintArea(instance.store, sheet, start === end ? start : `${start}:${end}`);
      });
    },
    [instance],
  );

  const onPrintTitleAction = useCallback(
    (action: PrintTitleAction) => {
      if (!instance) return;
      const state = instance.store.getState();
      const sheet = state.data.sheetIndex;
      const range = state.selection.range;
      recordPageSetupChange(instance.history, instance.store, () => {
        if (action === 'clear') {
          clearPrintTitles(instance.store, sheet);
        } else if (action === 'rows') {
          const rows =
            range.r0 === range.r1 ? `${range.r0 + 1}` : `${range.r0 + 1}:${range.r1 + 1}`;
          setPrintTitleRows(instance.store, sheet, rows);
        } else {
          const cols =
            range.c0 === range.c1
              ? colLetter(range.c0)
              : `${colLetter(range.c0)}:${colLetter(range.c1)}`;
          setPrintTitleCols(instance.store, sheet, cols);
        }
      });
    },
    [instance],
  );

  const onPageBreakAction = useCallback(
    (action: PageBreakAction) => {
      if (!instance) return;
      const state = instance.store.getState();
      const sheet = state.data.sheetIndex;
      const activeCell = state.selection.active;
      recordPageSetupChange(instance.history, instance.store, () => {
        if (action === 'insert-row') {
          insertManualPageBreak(instance.store, sheet, 'row', activeCell.row);
        } else if (action === 'insert-col') {
          insertManualPageBreak(instance.store, sheet, 'col', activeCell.col);
        } else if (action === 'remove-row') {
          removeManualPageBreak(instance.store, sheet, 'row', activeCell.row);
        } else if (action === 'remove-col') {
          removeManualPageBreak(instance.store, sheet, 'col', activeCell.col);
        } else {
          resetManualPageBreaks(instance.store, sheet);
        }
      });
    },
    [instance],
  );

  const onSheetBackgroundAction = useCallback(
    (action: SheetBackgroundAction) => {
      if (!instance) return;
      const sheet = instance.store.getState().data.sheetIndex;
      if (action === 'clear') {
        clearSheetBackgroundImage(instance.store, sheet, instance.history);
        return;
      }
      sheetBackgroundInputRef.current?.click();
    },
    [instance],
  );

  const onSheetBackgroundFileChange = useCallback(
    (event: ChangeEvent<HTMLInputElement>) => {
      if (!instance) return;
      const file = event.currentTarget.files?.[0];
      event.currentTarget.value = '';
      if (!file?.type.startsWith('image/')) return;
      const sheet = instance.store.getState().data.sheetIndex;
      const reader = new FileReader();
      reader.onload = () => {
        if (typeof reader.result === 'string') {
          setSheetBackgroundImage(instance.store, sheet, reader.result, instance.history);
        }
      };
      reader.readAsDataURL(file);
    },
    [instance],
  );

  const onThemeAction = useCallback(
    (action: ThemeAction) => {
      if (!instance) return;
      instance.setTheme(action);
      setActive(projectActiveState(instance));
    },
    [instance],
  );

  const onDefinedNameAction = useCallback(
    (action: DefinedNameAction) => {
      if (!instance) return;
      if (action === 'manager' || action === 'define') {
        instance.openNamedRangeDialog();
        return;
      }
      if (
        action === 'createTopRow' ||
        action === 'createBottomRow' ||
        action === 'createLeftColumn' ||
        action === 'createRightColumn'
      ) {
        const result = recordDefinedNamesChange(instance.history, instance.workbook, () =>
          createDefinedNamesFromSelection(
            instance.store.getState(),
            instance.workbook,
            action === 'createTopRow'
              ? 'top-row'
              : action === 'createBottomRow'
                ? 'bottom-row'
                : action === 'createLeftColumn'
                  ? 'left-column'
                  : 'right-column',
          ),
        );
        const sheet = instance.store.getState().data.sheetIndex;
        if (result.ok) mutators.replaceCells(instance.store, instance.workbook.cells(sheet));
        return;
      }
      if (action.startsWith('use:')) {
        const result = insertDefinedNameFormula(
          instance.store.getState(),
          instance.workbook,
          action.slice('use:'.length),
          instance.store,
        );
        if (!result) return;
        mutators.replaceCells(instance.store, instance.workbook.cells(result.addr.sheet));
        mutators.setActive(instance.store, result.addr);
      }
    },
    [instance],
  );

  const onCalculationAction = useCallback(
    (action: CalculationAction) => {
      if (!instance) return;
      if (action === 'iterative') {
        instance.openIterativeDialog();
        return;
      }
      const mode = action === 'auto' ? 0 : action === 'manual' ? 1 : 2;
      instance.workbook.setCalcMode(mode);
      setActive(projectActiveState(instance));
    },
    [instance],
  );

  const onWatchAction = useCallback(
    (action: WatchAction) => {
      if (!instance) return;
      const state = instance.store.getState();
      if (action === 'add') {
        recordWatchesChange(instance.history, instance.store, () => {
          watchRange(instance.store, state.selection.range);
        });
      } else if (action === 'delete') {
        recordWatchesChange(instance.history, instance.store, () => {
          unwatchCell(instance.store, state.selection.active);
        });
      } else if (action === 'delete-all') {
        recordWatchesChange(instance.history, instance.store, () => {
          clearWatchedCells(instance.store);
        });
      }
      instance.openWatchWindow();
      setActive(projectActiveState(instance));
    },
    [instance],
  );

  const [scriptDialog, setScriptDialog] = useState<ScriptDialogDraft | null>(null);
  const [automationRunCount, setAutomationRunCount] = useState(0);
  const [lastAutomationRun, setLastAutomationRun] = useState<AutomationRunDraft | null>(null);
  const openScriptDialog = useCallback(() => {
    if (onRunScript) {
      onRunScript();
      return;
    }
    if (!instance) return;
    setScriptDialog({ command: 'uppercase' });
  }, [instance, onRunScript]);

  const applyScriptDialog = useCallback(() => {
    if (!instance || !scriptDialog) return;
    const state = instance.store.getState();
    const range = state.selection.range;
    instance.history.begin();
    try {
      const changed = applyTextScriptToRange(state, instance.workbook, range, scriptDialog.command);
      if (changed > 0) {
        mutators.replaceCells(instance.store, instance.workbook.cells(state.data.sheetIndex));
      }
      setAutomationRunCount((count) => count + 1);
      setLastAutomationRun({ command: scriptDialog.command, range: formatA1Range(range), changed });
    } finally {
      instance.history.end();
    }
    setScriptDialog(null);
  }, [instance, scriptDialog]);
  const automationCommandLabel = useCallback(
    (command: ScriptCommand): string => {
      switch (command) {
        case 'uppercase':
          return cellMenuText.scriptCommandUppercase;
        case 'lowercase':
          return cellMenuText.scriptCommandLowercase;
        case 'trim':
          return cellMenuText.scriptCommandTrim;
        case 'clear':
          return cellMenuText.scriptCommandClear;
      }
    },
    [
      cellMenuText.scriptCommandClear,
      cellMenuText.scriptCommandLowercase,
      cellMenuText.scriptCommandTrim,
      cellMenuText.scriptCommandUppercase,
    ],
  );
  const recordActions = useCallback(() => {
    if (!instance) return;
    const recordedDetail = lastAutomationRun
      ? cellMenuText.automationRunDetail
          .replace('{command}', automationCommandLabel(lastAutomationRun.command))
          .replace('{range}', lastAutomationRun.range)
          .replace('{count}', String(lastAutomationRun.changed))
      : cellMenuText.recordActionsEmpty;
    setRibbonReportDialog({
      title: tr.recordActions,
      items: [
        {
          severity: 'info',
          label: cellMenuText.recordActionsStatus,
          detail: recordedDetail,
        },
      ],
    });
  }, [
    automationCommandLabel,
    cellMenuText.automationRunDetail,
    cellMenuText.recordActionsEmpty,
    cellMenuText.recordActionsStatus,
    instance,
    lastAutomationRun,
    tr.recordActions,
  ]);
  const openAllScripts = useCallback(() => {
    const runStatus =
      automationRunCount > 0
        ? cellMenuText.automationRunStatus.replace('{count}', String(automationRunCount))
        : cellMenuText.automationNoRuns;
    const runDetail = lastAutomationRun
      ? cellMenuText.automationRunDetail
          .replace('{command}', automationCommandLabel(lastAutomationRun.command))
          .replace('{range}', lastAutomationRun.range)
          .replace('{count}', String(lastAutomationRun.changed))
      : null;
    setRibbonReportDialog({
      title: cellMenuText.automationScriptsTitle,
      items: [
        {
          severity: 'info',
          label: cellMenuText.automationBuiltInScriptsLabel,
          detail: cellMenuText.automationBuiltInScriptsDetail,
        },
        {
          severity: 'info',
          label: cellMenuText.automationRecentRunsLabel,
          detail: runDetail ? `${runStatus}\n${runDetail}` : runStatus,
        },
      ],
    });
  }, [
    automationRunCount,
    automationCommandLabel,
    cellMenuText.automationBuiltInScriptsDetail,
    cellMenuText.automationBuiltInScriptsLabel,
    cellMenuText.automationNoRuns,
    cellMenuText.automationRecentRunsLabel,
    cellMenuText.automationRunDetail,
    cellMenuText.automationRunStatus,
    cellMenuText.automationScriptsTitle,
    lastAutomationRun,
  ]);
  const onAddInAction = useCallback(
    (action: AddInAction) => {
      const report = buildRibbonAddInReport(action, {
        cellMenu: cellMenuText,
        addInDefaultTitle: tr.addIn,
      });
      if (report) setRibbonReportDialog(report);
    },
    [cellMenuText, tr.addIn],
  );
  const onPivotTableAction = useCallback(
    (action: PivotTableAction) => {
      if (!instance) return;
      const result = executeRibbonPivotTableAction({
        store: instance.store,
        workbook: instance.workbook,
        history: instance.history,
        action,
        strings: {
          pivotTable: tr.pivotTable,
          pivotTableNewSheet: cellMenuText.pivotTableNewSheet,
          recommendedPivotTables: cellMenuText.recommendedPivotTables,
          pivotAuthoringDetail: strings.workbookObjects.compatibilityDetails.pivotAuthoring,
          workbookStructureProtectedBlocked: cellMenuText.workbookStructureProtectedBlocked,
        },
      });
      if (result.kind === 'open-dialog') instance.openPivotTableDialog();
      else if (result.kind === 'report') setRibbonReportDialog(result.report);
      else setActive(projectActiveState(instance));
    },
    [
      cellMenuText.pivotTableNewSheet,
      cellMenuText.recommendedPivotTables,
      cellMenuText.workbookStructureProtectedBlocked,
      instance,
      strings.workbookObjects.compatibilityDetails.pivotAuthoring,
      tr.pivotTable,
    ],
  );
  const onPdfAction = useCallback(
    (action: PdfAction) => {
      if (!instance) return;
      const result = resolveRibbonPdfAction(action, {
        cellMenu: cellMenuText,
        pdfTitle: tr.pdf,
      });
      if (result.kind === 'open-page-setup') instance.openPageSetup();
      else {
        instance.print('pdf');
        if (result.report) setRibbonReportDialog(result.report);
      }
    },
    [cellMenuText, instance, tr.pdf],
  );
  const onIllustrationAction = useCallback(
    (label: string) => {
      if (!instance) return;
      setRibbonReportDialog({
        title: tr.illustrations,
        items: [
          {
            severity: 'info',
            label,
            detail: strings.workbookObjects.compatibilityDetails.chartsDrawings,
          },
        ],
      });
    },
    [instance, strings.workbookObjects.compatibilityDetails.chartsDrawings, tr.illustrations],
  );
  const protectWorkbookFromBackstage = useCallback(() => {
    if (!instance) return;
    setWorkbookStructureProtected(
      instance.store,
      !isWorkbookStructureProtected(instance.store.getState()),
    );
    setActive(projectActiveState(instance));
  }, [instance]);
  const inspectWorkbookFromBackstage = useCallback(() => {
    if (!instance) return;
    setRibbonReportDialog({
      title: strings.backstage.inspect,
      items: buildSpreadsheetCompatibilityReport(instance.workbook, strings.workbookObjects),
    });
  }, [instance, strings.backstage.inspect, strings.workbookObjects]);

  const onZoom = useCallback(
    (zoom: number) => {
      if (!instance) return;
      setSheetZoom(instance.store, zoom, instance.workbook);
    },
    [instance],
  );
  const openZoomDialog = useCallback(() => {
    if (!instance) return;
    setZoomDialog(String(Math.round(instance.store.getState().viewport.zoom * 100)));
  }, [instance]);
  const applyZoomDialog = useCallback(() => {
    if (!instance || zoomDialog == null) return;
    const percent = Number.parseFloat(zoomDialog);
    if (!Number.isFinite(percent)) return;
    const clamped = Math.max(10, Math.min(400, percent));
    setSheetZoom(instance.store, clamped / 100, instance.workbook);
    setZoomDialog(null);
  }, [instance, zoomDialog]);
  const onZoomSelection = useCallback(() => {
    if (!instance) return;
    const state = instance.store.getState();
    const range = state.selection.range;
    const selectedRows = Math.max(1, range.r1 - range.r0 + 1);
    const selectedCols = Math.max(1, range.c1 - range.c0 + 1);
    const rowFit = state.viewport.rowCount / selectedRows;
    const colFit = state.viewport.colCount / selectedCols;
    const next = state.viewport.zoom * Math.min(rowFit, colFit);
    setSheetZoom(instance.store, next, instance.workbook);
  }, [instance]);

  const [formulaBarVisible, setFormulaBarVisible] = useState(() => features?.formulaBar !== false);
  useEffect(() => {
    setFormulaBarVisible(features?.formulaBar !== false);
  }, [features]);
  const onToggleFormulaBar = useCallback(() => {
    if (!instance) return;
    setFormulaBarVisible((current) => {
      const next = !current;
      instance.setFeatures({ ...(features ?? {}), formulaBar: next });
      return next;
    });
  }, [instance, features]);

  const onSort = useCallback(
    (direction: 'asc' | 'desc') => {
      if (!instance) return;
      sortActiveColumnAuto({
        store: instance.store,
        workbook: instance.workbook,
        history: instance.history,
        direction,
      });
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

  const onScaleFit = useCallback(
    (axis: 'width' | 'height', pages: string) => {
      if (!instance) return;
      const sheet = instance.store.getState().data.sheetIndex;
      const n = Number.parseInt(pages, 10);
      setFitToPages(instance.store, sheet, axis, n > 0 ? n : undefined, instance.history);
    },
    [instance],
  );

  const onScalePercent = useCallback(
    (percent: string) => {
      if (!instance) return;
      const sheet = instance.store.getState().data.sheetIndex;
      const scale = Number.parseInt(percent, 10) / 100;
      setPageScale(instance.store, sheet, Number.isFinite(scale) ? scale : 1, instance.history);
    },
    [instance],
  );

  const onNumberFormat = useCallback(
    (next: string) => {
      if (!instance) return;
      const action = next as NumberFormatAction;
      if (action === 'more') {
        instance.openFormatDialog('number');
        return;
      }
      const fmt = numberFormatForAction(action, lang);
      if (!fmt) return;
      wrapFormat((s, st) => setNumFmt(s, st, fmt));
    },
    [instance, lang, wrapFormat],
  );

  const onFormatAsTable = useCallback(
    (style: FormatTableAction = 'medium') => {
      if (!instance) return;
      const r = instance.store.getState().selection.range;
      recordTablesChange(instance.history, instance.store, () => {
        formatAsTable(instance.store, r, { style });
      });
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
    disabled = false,
    allowWithoutInstance = false,
  ): ReactElement => (
    <button
      key={id}
      type="button"
      className={`demo__rb${extra}${isActive ? ' demo__rb--active' : ''}`}
      data-ribbon-command={id}
      title={title}
      aria-label={title}
      aria-keyshortcuts={RIBBON_KEYSHORTCUTS[id]}
      onClick={onClick}
      disabled={disabled || (!allowWithoutInstance && !instance)}
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
    <span key={id} className="demo__rb-break" data-ribbon-command={id} aria-hidden="true" />
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
      commandId={id}
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
      commandId={id}
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
        moreColors: tr.moreColors,
        standardColors: tr.standardColors,
        themeColors: tr.themeColors,
      }}
      label={label}
      disabled={!instance}
      onChange={onChange}
    />
  );

  const definedNameOptions: readonly {
    action: DefinedNameAction;
    label: string;
    separatorBefore?: boolean;
  }[] = [
    { action: 'define', label: cellMenuText.defineName },
    ...(instance ? listDefinedNames(instance.workbook) : []).map((entry) => ({
      action: `use:${entry.name}` as const,
      label: `${cellMenuText.useInFormula}: ${entry.name}`,
    })),
    { action: 'createTopRow', label: cellMenuText.createFromSelectionTop, separatorBefore: true },
    { action: 'createBottomRow', label: cellMenuText.createFromSelectionBottom },
    { action: 'createLeftColumn', label: cellMenuText.createFromSelectionLeft },
    { action: 'createRightColumn', label: cellMenuText.createFromSelectionRight },
    { action: 'manager', label: cellMenuText.nameManager, separatorBefore: true },
  ];
  const cellStyleLabel = (id: CellStyleId): string => {
    switch (id) {
      case 'normal':
        return cellMenuText.cellStyleNormal;
      case 'title':
        return cellMenuText.cellStyleTitle;
      case 'heading1':
        return cellMenuText.cellStyleHeading1;
      case 'heading2':
        return cellMenuText.cellStyleHeading2;
      case 'heading3':
        return cellMenuText.cellStyleHeading3;
      case 'heading4':
        return cellMenuText.cellStyleHeading4;
      case 'good':
        return cellMenuText.cellStyleGood;
      case 'bad':
        return cellMenuText.cellStyleBad;
      case 'neutral':
        return cellMenuText.cellStyleNeutral;
      case 'note':
        return cellMenuText.cellStyleNote;
      case 'warning':
        return cellMenuText.cellStyleWarning;
      case 'inputCell':
        return cellMenuText.cellStyleInputCell;
      case 'outputCell':
        return cellMenuText.cellStyleOutputCell;
      case 'calculation':
        return cellMenuText.cellStyleCalculation;
      case 'linkedCell':
        return cellMenuText.cellStyleLinkedCell;
      case 'totalCell':
        return cellMenuText.cellStyleTotalCell;
      case 'currency':
        return cellMenuText.cellStyleCurrency;
      case 'currency0':
        return cellMenuText.cellStyleCurrency0;
      case 'percent':
        return cellMenuText.cellStylePercent;
      case 'comma':
        return cellMenuText.cellStyleComma;
      case 'comma0':
        return cellMenuText.cellStyleComma0;
      default:
        return id;
    }
  };
  const cellStyleGroupLabel = (id: (typeof CELL_STYLE_GROUPS)[number]['id']): string =>
    strings.cellStylesGallery.groups[id];
  const cellStyleOptions = CELL_STYLE_GROUPS.flatMap((group) => [
    {
      action: `${CELL_STYLE_SECTION_ACTION_PREFIX}${group.id}` as CellStyleAction,
      label: cellStyleGroupLabel(group.id),
      section: true,
    },
    ...group.styleIds.map((id) => ({
      action: id,
      label: cellStyleLabel(id),
    })),
  ]);
  const sheetTabColorLabel = (action: CellFormatAction): string => {
    const sheetTabs = strings.sheetTabs;
    switch (action) {
      case 'tabColorRed':
        return sheetTabs.tabColorRed;
      case 'tabColorOrange':
        return sheetTabs.tabColorOrange;
      case 'tabColorYellow':
        return sheetTabs.tabColorYellow;
      case 'tabColorGreen':
        return sheetTabs.tabColorGreen;
      case 'tabColorBlue':
        return sheetTabs.tabColorBlue;
      case 'tabColorPurple':
        return sheetTabs.tabColorPurple;
      case 'tabColorGray':
        return sheetTabs.tabColorGray;
      default:
        return sheetTabs.tabColor;
    }
  };
  const currentFreezeAction = (() => {
    if (!instance) return null;
    const { freezeRows, freezeCols } = instance.store.getState().layout;
    if (freezeRows === 0 && freezeCols === 0) return 'none';
    if (freezeRows === 1 && freezeCols === 0) return 'topRow';
    if (freezeRows === 0 && freezeCols === 1) return 'firstColumn';
    return 'panes';
  })() satisfies FreezeAction | null;
  const workbookStructureProtected =
    !!instance && isWorkbookStructureProtected(instance.store.getState());
  const currentTheme = (instance?.store.getState().ui.theme ?? 'paper') as ThemeAction;

  const ribbonGroups = buildRibbonGroups({
    active,
    addInMenu: (
      <CellMenu<AddInAction>
        key="addIn"
        command="addIn"
        disabled={!instance && !onAddIn}
        icon="addIn"
        label={tr.addIn}
        options={[
          { action: 'get', label: cellMenuText.addInGet },
          { action: 'my', label: cellMenuText.addInMy },
          { action: 'manage', label: cellMenuText.addInManage, separatorBefore: true },
        ]}
        onPick={onAddInAction}
      />
    ),
    pivotTableMenu: (
      <CellMenu<PivotTableAction>
        key="pivotTableInsert"
        command="pivotTableInsert"
        disabled={!instance}
        icon="table"
        label={tr.pivotTable}
        options={[
          { action: 'dialog', label: cellMenuText.pivotTableFromRange },
          { action: 'recommended', label: cellMenuText.recommendedPivotTables },
          { action: 'new-sheet', label: cellMenuText.pivotTableNewSheet, separatorBefore: true },
          { action: 'existing-sheet', label: cellMenuText.pivotTableExistingSheet },
        ]}
        onPick={onPivotTableAction}
      />
    ),
    pictureInsertMenu: (
      <CellMenu<PictureAction>
        key="pictureInsert"
        command="pictureInsert"
        disabled={!instance}
        icon="page"
        label={tr.pictures}
        options={[
          { action: 'device', label: cellMenuText.pictureThisDevice },
          { action: 'online', label: cellMenuText.pictureOnline },
        ]}
        onPick={(action) =>
          onIllustrationAction(
            action === 'device' ? cellMenuText.pictureThisDevice : cellMenuText.pictureOnline,
          )
        }
      />
    ),
    shapesInsertMenu: (
      <CellMenu<ShapeAction>
        key="shapesInsert"
        command="shapesInsert"
        disabled={!instance}
        icon="options"
        label={tr.shapes}
        options={[
          { action: 'rectangle', label: cellMenuText.shapeRectangle },
          { action: 'rounded-rectangle', label: cellMenuText.shapeRoundedRectangle },
          { action: 'oval', label: cellMenuText.shapeOval },
          { action: 'line', label: cellMenuText.shapeLine, separatorBefore: true },
          { action: 'arrow', label: cellMenuText.shapeArrow },
        ]}
        onPick={(action) => {
          const labels: Record<ShapeAction, string> = {
            rectangle: cellMenuText.shapeRectangle,
            'rounded-rectangle': cellMenuText.shapeRoundedRectangle,
            oval: cellMenuText.shapeOval,
            line: cellMenuText.shapeLine,
            arrow: cellMenuText.shapeArrow,
          };
          onIllustrationAction(labels[action]);
        }}
      />
    ),
    screenshotInsertMenu: (
      <CellMenu<ScreenshotAction>
        key="screenshotInsert"
        command="screenshotInsert"
        disabled={!instance}
        icon="page"
        label={tr.screenshot}
        options={[
          { action: 'current-view', label: cellMenuText.screenshotCurrentView },
          { action: 'screen-clipping', label: cellMenuText.screenshotScreenClipping },
        ]}
        onPick={(action) =>
          onIllustrationAction(
            action === 'current-view'
              ? cellMenuText.screenshotCurrentView
              : cellMenuText.screenshotScreenClipping,
          )
        }
      />
    ),
    autosumFormulaMenu: (
      <CellMenu<AutoSumAction>
        key="autosumFormula"
        command="autosumFormula"
        disabled={!instance}
        icon="autosum"
        label={tr.autoSum}
        options={[
          { action: 'SUM', label: cellMenuText.autosumSum },
          { action: 'AVERAGE', label: cellMenuText.autosumAverage },
          { action: 'COUNT', label: cellMenuText.autosumCount },
          { action: 'MAX', label: cellMenuText.autosumMax },
          { action: 'MIN', label: cellMenuText.autosumMin },
          { action: 'MORE', label: cellMenuText.autosumMoreFunctions, separatorBefore: true },
        ]}
        onPick={onAutoSumAction}
      />
    ),
    autosumMenu: (
      <CellMenu<AutoSumAction>
        key="autosum"
        command="autosum"
        disabled={!instance}
        icon="autosum"
        label={tr.autoSum}
        options={[
          { action: 'SUM', label: cellMenuText.autosumSum },
          { action: 'AVERAGE', label: cellMenuText.autosumAverage },
          { action: 'COUNT', label: cellMenuText.autosumCount },
          { action: 'MAX', label: cellMenuText.autosumMax },
          { action: 'MIN', label: cellMenuText.autosumMin },
          { action: 'MORE', label: cellMenuText.autosumMoreFunctions, separatorBefore: true },
        ]}
        onPick={onAutoSumAction}
      />
    ),
    calcOptionsMenu: (
      <CellMenu<CalculationAction>
        key="calcOptions"
        command="calcOptions"
        disabled={!instance}
        icon="options"
        label={tr.options}
        options={[
          { action: 'auto', label: cellMenuText.calcAutomatic },
          { action: 'autoNoTable', label: cellMenuText.calcAutoNoTable },
          { action: 'manual', label: cellMenuText.calcManual },
          { action: 'iterative', label: cellMenuText.calcIterative, separatorBefore: true },
        ]}
        activeAction={
          active.calcMode == null
            ? null
            : active.calcMode === 0
              ? 'auto'
              : active.calcMode === 1
                ? 'manual'
                : 'autoNoTable'
        }
        onPick={onCalculationAction}
      />
    ),
    watchMenu: (
      <CellMenu<WatchAction>
        key="watch"
        command="watch"
        disabled={!instance}
        icon="watch"
        label={tr.watch}
        options={[
          { action: 'open', label: cellMenuText.watchWindow },
          { action: 'add', label: cellMenuText.watchAdd },
          { action: 'delete', label: cellMenuText.watchDelete },
          { action: 'delete-all', label: cellMenuText.watchDeleteAll, separatorBefore: true },
        ]}
        onPick={onWatchAction}
      />
    ),
    watchViewMenu: (
      <CellMenu<WatchAction>
        key="watchView"
        command="watchView"
        disabled={!instance}
        icon="watch"
        label={tr.watch}
        options={[
          { action: 'open', label: cellMenuText.watchWindow },
          { action: 'add', label: cellMenuText.watchAdd },
          { action: 'delete', label: cellMenuText.watchDelete },
          { action: 'delete-all', label: cellMenuText.watchDeleteAll, separatorBefore: true },
        ]}
        onPick={onWatchAction}
      />
    ),
    borderPresets,
    borderColor,
    borderStyle,
    borderStyles,
    cellInsertMenu: (
      <CellMenu<CellInsertAction>
        key="insertRows"
        command="insertRows"
        disabled={!instance}
        icon="insertRows"
        label={cellMenuText.insert}
        options={[
          { action: 'shiftDown', label: cellMenuText.insertShiftDown },
          { action: 'shiftRight', label: cellMenuText.insertShiftRight },
          { action: 'rows', label: cellMenuText.insertRows, separatorBefore: true },
          { action: 'cols', label: cellMenuText.insertCols },
          { action: 'sheet', label: strings.sheetTabs.insertSheet, separatorBefore: true },
        ]}
        onPick={onInsertCellsAction}
      />
    ),
    cellDeleteMenu: (
      <CellMenu<CellDeleteAction>
        key="deleteRows"
        command="deleteRows"
        disabled={!instance}
        icon="deleteRows"
        label={cellMenuText.delete}
        options={[
          { action: 'shiftUp', label: cellMenuText.deleteShiftUp },
          { action: 'shiftLeft', label: cellMenuText.deleteShiftLeft },
          { action: 'rows', label: cellMenuText.deleteRows, separatorBefore: true },
          { action: 'cols', label: cellMenuText.deleteCols },
          { action: 'sheet', label: strings.sheetTabs.deleteSheet, separatorBefore: true },
        ]}
        onPick={onDeleteCellsAction}
      />
    ),
    cellFormatMenu: (
      <CellMenu<CellFormatAction>
        key="formatCellsHome"
        command="formatCellsHome"
        disabled={!instance}
        icon="formatCells"
        label={cellMenuText.format}
        options={[
          { action: 'dialog', label: cellMenuText.formatCells },
          { action: 'rowHeight', label: cellMenuText.rowHeight, separatorBefore: true },
          { action: 'autoFitRowHeight', label: cellMenuText.autoFitRowHeight },
          { action: 'colWidth', label: cellMenuText.colWidth },
          { action: 'autoFitColWidth', label: cellMenuText.autoFitColWidth },
          { action: 'hideRows', label: cellMenuText.hideRows, separatorBefore: true },
          { action: 'showRows', label: cellMenuText.showRows },
          { action: 'hideCols', label: cellMenuText.hideCols },
          { action: 'showCols', label: cellMenuText.showCols },
          { action: 'renameSheet', label: strings.sheetTabs.rename, separatorBefore: true },
          { action: 'moveSheetLeft', label: strings.sheetTabs.moveLeft },
          { action: 'moveSheetRight', label: strings.sheetTabs.moveRight },
          { action: 'hideSheet', label: strings.sheetTabs.hideSheet },
          { action: 'unhideSheet', label: strings.sheetTabs.unhideSheet },
          {
            action: 'tabColorNone',
            label: `${strings.sheetTabs.tabColor}: ${strings.sheetTabs.noColor}`,
            separatorBefore: true,
          },
          ...SHEET_TAB_COLOR_ACTIONS.map((entry) => ({
            action: entry.action,
            label: `${strings.sheetTabs.tabColor}: ${sheetTabColorLabel(entry.action)}`,
          })),
          {
            action: 'protectSheet',
            label: cellMenuText.protectSheet,
            separatorBefore: true,
            active: active.protected,
          },
        ]}
        onPick={onCellFormatAction}
      />
    ),
    cellStylesMenu: (
      <CellMenu<CellStyleAction>
        key="cellStyles"
        command="cellStyles"
        disabled={!instance}
        icon="tableStyle"
        label={tr.cellStyles}
        options={cellStyleOptions}
        activeAction={
          CELL_STYLES.some((style) => style.id === active.cellStyle)
            ? (active.cellStyle as CellStyleAction)
            : null
        }
        activeButton={active.cellStyle != null}
        onPick={onCellStyleAction}
      />
    ),
    definedNamesMenu: (
      <CellMenu<DefinedNameAction>
        key="namedRanges"
        command="namedRanges"
        disabled={!instance}
        icon="names"
        label={cellMenuText.nameManager}
        options={definedNameOptions}
        onPick={onDefinedNameAction}
      />
    ),
    definedNamesInsertMenu: (
      <CellMenu<DefinedNameAction>
        key="namedRangesInsert"
        command="namedRangesInsert"
        disabled={!instance}
        icon="names"
        label={tr.names}
        options={definedNameOptions}
        onPick={onDefinedNameAction}
      />
    ),
    functionLogicalMenu: (
      <CellMenu<FunctionAction>
        key="ifFormula"
        command="ifFormula"
        disabled={!instance}
        icon="function"
        label={tr.functionLogical}
        options={[
          { action: 'IF', label: 'IF' },
          { action: 'IFS', label: 'IFS' },
          { action: 'AND', label: 'AND' },
          { action: 'OR', label: 'OR' },
        ]}
        onPick={onFunctionAction}
      />
    ),
    functionLookupMenu: (
      <CellMenu<FunctionAction>
        key="xlookupFormula"
        command="xlookupFormula"
        disabled={!instance}
        icon="function"
        label={tr.functionLookupReference}
        options={[
          { action: 'XLOOKUP', label: 'XLOOKUP' },
          { action: 'VLOOKUP', label: 'VLOOKUP' },
          { action: 'INDEX', label: 'INDEX' },
          { action: 'MATCH', label: 'MATCH' },
        ]}
        onPick={onFunctionAction}
      />
    ),
    functionTextMenu: (
      <CellMenu<FunctionAction>
        key="concatFormula"
        command="concatFormula"
        disabled={!instance}
        icon="function"
        label={tr.functionText}
        options={[
          { action: 'CONCAT', label: 'CONCAT' },
          { action: 'TEXT', label: 'TEXT' },
          { action: 'LEFT', label: 'LEFT' },
          { action: 'RIGHT', label: 'RIGHT' },
        ]}
        onPick={onFunctionAction}
      />
    ),
    functionDateTimeMenu: (
      <CellMenu<FunctionAction>
        key="todayFormula"
        command="todayFormula"
        disabled={!instance}
        icon="function"
        label={tr.functionDateTime}
        options={[
          { action: 'TODAY', label: 'TODAY' },
          { action: 'NOW', label: 'NOW' },
          { action: 'DATE', label: 'DATE' },
          { action: 'YEAR', label: 'YEAR' },
        ]}
        onPick={onFunctionAction}
      />
    ),
    functionFinancialMenu: (
      <CellMenu<FunctionAction>
        key="pmtFormula"
        command="pmtFormula"
        disabled={!instance}
        icon="function"
        label={tr.functionFinancial}
        options={[
          { action: 'PMT', label: 'PMT' },
          { action: 'NPV', label: 'NPV' },
          { action: 'IRR', label: 'IRR' },
          { action: 'RATE', label: 'RATE' },
        ]}
        onPick={onFunctionAction}
      />
    ),
    functionMathTrigMenu: (
      <CellMenu<FunctionAction>
        key="roundFormula"
        command="roundFormula"
        disabled={!instance}
        icon="function"
        label={tr.functionMathTrig}
        options={[
          { action: 'ROUND', label: 'ROUND' },
          { action: 'SUMIF', label: 'SUMIF' },
          { action: 'COUNTIF', label: 'COUNTIF' },
          { action: 'ABS', label: 'ABS' },
        ]}
        onPick={onFunctionAction}
      />
    ),
    hyperlinkMenu: (
      <CellMenu<HyperlinkAction>
        key="hyperlinkInsert"
        command="hyperlinkInsert"
        disabled={!instance}
        icon="link"
        label={tr.hyperlink}
        options={[
          { action: 'edit', label: cellMenuText.linkInsertOrEdit },
          { action: 'open', label: cellMenuText.linkOpen },
          { action: 'clear', label: cellMenuText.linkClear },
          { action: 'external', label: cellMenuText.linkExternalLinks, separatorBefore: true },
        ]}
        onPick={onHyperlinkAction}
      />
    ),
    outlineGroupMenu: (
      <CellMenu<OutlineAxisAction>
        key="outlineGroup"
        command="outlineGroup"
        disabled={!instance}
        icon="table"
        label={tr.groupOutline}
        options={[
          { action: 'rows', label: strings.contextMenu.rowGroup },
          { action: 'cols', label: strings.contextMenu.colGroup },
        ]}
        onPick={onOutlineGroupAction}
      />
    ),
    outlineUngroupMenu: (
      <CellMenu<OutlineAxisAction>
        key="outlineUngroup"
        command="outlineUngroup"
        disabled={!instance}
        icon="table"
        label={tr.ungroupOutline}
        options={[
          { action: 'rows', label: strings.contextMenu.rowUngroup },
          { action: 'cols', label: strings.contextMenu.colUngroup },
        ]}
        onPick={onOutlineUngroupAction}
      />
    ),
    dataFilterMenu: (
      <CellMenu<FilterDataAction>
        key="filter"
        command="filter"
        disabled={!instance}
        icon="filter"
        label={tr.filter}
        options={[
          { action: 'toggle', label: cellMenuText.filterToggle },
          { action: 'clear', label: cellMenuText.filterClearAll },
          { action: 'reapply', label: cellMenuText.filterReapply },
          { action: 'filter-by-selected', label: cellMenuText.filterBySelectedCellValue },
          { action: 'advanced', label: cellMenuText.filterAdvanced, separatorBefore: true },
        ]}
        onPick={onFilterDataAction}
      />
    ),
    dataSortMenu: (
      <CellMenu<SortAction>
        key="sortData"
        command="sortData"
        disabled={!instance}
        icon="sortAsc"
        label={cellMenuText.sortCustom}
        options={[
          { action: 'custom', label: cellMenuText.sortCustom },
          { action: 'asc', label: cellMenuText.sortAsc, separatorBefore: true },
          { action: 'desc', label: cellMenuText.sortDesc },
        ]}
        onPick={onSortMenuAction}
      />
    ),
    deleteCommentMenu: (
      <CellMenu<CommentAction>
        key="deleteCommentReview"
        command="deleteCommentReview"
        disabled={!instance}
        icon="clear"
        label={tr.deleteComment}
        options={[
          { action: 'delete-active', label: cellMenuText.commentDelete },
          { action: 'delete-all', label: cellMenuText.commentDeleteAll },
        ]}
        onPick={onCommentAction}
      />
    ),
    protectionMenu: (
      <CellMenu<ProtectionAction>
        key="protectionReview"
        command="protectionReview"
        disabled={!instance}
        icon="protect"
        label={cellMenuText.allowEditRangesCommand}
        options={[
          { action: 'allow-edit-range', label: cellMenuText.allowEditRangesCommand },
          {
            action: 'clear-allowed-edit-ranges',
            label: cellMenuText.allowEditRangesClearCommand,
          },
        ]}
        onPick={onProtectionAction}
      />
    ),
    chartMenu: (
      <CellMenu<ChartAction>
        key="chartInsert"
        command="chartInsert"
        disabled={!instance}
        icon="chart"
        label={cellMenuText.chart}
        options={[
          { action: 'column', label: cellMenuText.chartColumn },
          { action: 'bar', label: cellMenuText.chartBar },
          { action: 'line', label: cellMenuText.chartLine },
          { action: 'area', label: cellMenuText.chartArea },
          { action: 'pie', label: cellMenuText.chartPie },
          { action: 'scatter', label: cellMenuText.chartScatter },
          { action: 'recommended', label: cellMenuText.recommendedCharts, separatorBefore: true },
        ]}
        onPick={onChartAction}
      />
    ),
    fillMenu: (
      <CellMenu<FillAction>
        key="fillHome"
        command="fillHome"
        disabled={!instance}
        icon="fillColor"
        label={cellMenuText.fill}
        options={[
          { action: 'down', label: cellMenuText.fillDown },
          { action: 'right', label: cellMenuText.fillRight },
          { action: 'up', label: cellMenuText.fillUp },
          { action: 'left', label: cellMenuText.fillLeft },
          { action: 'flash', label: cellMenuText.flashFill, separatorBefore: true },
          { action: 'series', label: cellMenuText.series, separatorBefore: true },
          { action: 'days', label: cellMenuText.fillDays },
          { action: 'weekdays', label: cellMenuText.fillWeekdays },
          { action: 'months', label: cellMenuText.fillMonths },
          { action: 'years', label: cellMenuText.fillYears },
        ]}
        onPick={onFillAction}
      />
    ),
    formulaBarVisible,
    clearMenu: (
      <CellMenu<ClearAction>
        key="clearFormat"
        command="clearFormat"
        disabled={!instance}
        icon="clear"
        label={cellMenuText.clear}
        options={[
          { action: 'all', label: cellMenuText.clearAll },
          { action: 'formats', label: cellMenuText.clearFormats },
          { action: 'contents', label: cellMenuText.clearContents },
          { action: 'comments', label: cellMenuText.clearComments },
          { action: 'hyperlinks', label: cellMenuText.clearHyperlinks },
          { action: 'conditional', label: cellMenuText.clearConditional },
        ]}
        onPick={onClearAction}
      />
    ),
    color,
    conditionalMenu: (
      <ConditionalMenu
        key="conditional"
        disabled={!instance}
        active={active.conditionalFormatting}
        instance={instance}
        strings={strings}
      />
    ),
    sortMenu: (
      <CellMenu<SortAction>
        key="sortFilterHome"
        command="sortFilterHome"
        disabled={!instance}
        icon="sortAsc"
        label={cellMenuText.sortFilter}
        options={[
          { action: 'asc', label: cellMenuText.sortAsc },
          { action: 'desc', label: cellMenuText.sortDesc },
          { action: 'custom', label: cellMenuText.sortCustom },
          { action: 'filter', label: cellMenuText.filter },
          { action: 'filter-clear', label: cellMenuText.clearFilter },
          { action: 'filter-reapply', label: cellMenuText.filterReapply },
          { action: 'filter-by-selected', label: cellMenuText.filterBySelectedCellValue },
          { action: 'filter-advanced', label: cellMenuText.filterAdvanced },
          { action: 'dedupe', label: cellMenuText.removeDuplicates, separatorBefore: true },
          { action: 'conditional', label: cellMenuText.conditional },
          { action: 'named', label: cellMenuText.namedRanges },
        ]}
        onPick={onSortMenuAction}
      />
    ),
    findMenu: (
      <CellMenu<FindAction>
        key="findHome"
        command="findHome"
        disabled={!instance}
        icon="find"
        label={cellMenuText.findSelect}
        options={[
          { action: 'find', label: cellMenuText.find },
          { action: 'replace', label: cellMenuText.replace },
          { action: 'go-to', label: cellMenuText.goTo },
          { action: 'go-to-special', label: cellMenuText.goToSpecial },
          { action: 'formulas', label: cellMenuText.findFormulas, separatorBefore: true },
          { action: 'constants', label: cellMenuText.findConstants },
          { action: 'numbers', label: strings.goToDialog.kindNumbers },
          { action: 'text', label: strings.goToDialog.kindText },
          { action: 'errors', label: strings.goToDialog.kindErrors },
          {
            action: 'conditional-format',
            label: cellMenuText.findConditionalFormatting,
            separatorBefore: true,
          },
          { action: 'data-validation', label: cellMenuText.findDataValidation },
          { action: 'comments', label: cellMenuText.comments },
        ]}
        onPick={onFindAction}
      />
    ),
    printAreaMenu: (
      <CellMenu<PrintAreaAction>
        key="printArea"
        command="printArea"
        disabled={!instance}
        icon="table"
        label={tr.printArea}
        options={[
          { action: 'set', label: cellMenuText.printAreaSet },
          { action: 'clear', label: cellMenuText.printAreaClear },
        ]}
        onPick={onPrintAreaAction}
      />
    ),
    pageBreaksMenu: (
      <CellMenu<PageBreakAction>
        key="pageBreaks"
        command="pageBreaks"
        disabled={!instance}
        icon="page"
        label={tr.breaks}
        options={[
          { action: 'insert-row', label: cellMenuText.pageBreakInsertRow },
          { action: 'insert-col', label: cellMenuText.pageBreakInsertCol },
          { action: 'remove-row', label: cellMenuText.pageBreakRemoveRow, separatorBefore: true },
          { action: 'remove-col', label: cellMenuText.pageBreakRemoveCol },
          { action: 'reset', label: cellMenuText.pageBreakResetAll, separatorBefore: true },
        ]}
        onPick={onPageBreakAction}
      />
    ),
    sheetBackgroundMenu: (
      <CellMenu<SheetBackgroundAction>
        key="sheetBackground"
        command="sheetBackground"
        disabled={!instance}
        icon="page"
        label={tr.background}
        options={[
          { action: 'set', label: cellMenuText.sheetBackgroundSet },
          { action: 'clear', label: cellMenuText.sheetBackgroundClear },
        ]}
        onPick={onSheetBackgroundAction}
      />
    ),
    printTitlesMenu: (
      <CellMenu<PrintTitleAction>
        key="printTitles"
        command="printTitles"
        disabled={!instance}
        icon="table"
        label={tr.printTitles}
        options={[
          { action: 'rows', label: cellMenuText.printTitleRowsSet },
          { action: 'cols', label: cellMenuText.printTitleColsSet },
          { action: 'clear', label: cellMenuText.printTitlesClear, separatorBefore: true },
        ]}
        onPick={onPrintTitleAction}
      />
    ),
    textToColumnsMenu: (
      <CellMenu<TextToColumnsAction>
        key="textToColumns"
        command="textToColumns"
        disabled={!instance}
        icon="table"
        label={cellMenuText.textToColumns}
        options={[
          { action: 'comma', label: cellMenuText.textToColumnsComma },
          { action: 'tab', label: cellMenuText.textToColumnsTab },
          { action: 'semicolon', label: cellMenuText.textToColumnsSemicolon },
          { action: 'space', label: cellMenuText.textToColumnsSpace },
          { action: 'custom', label: cellMenuText.textToColumnsCustom, separatorBefore: true },
        ]}
        onPick={onTextToColumnsAction}
      />
    ),
    dataValidationMenu: (
      <CellMenu<DataValidationAction>
        key="dataValidation"
        command="dataValidation"
        disabled={!instance}
        icon="options"
        label={tr.dataValidation}
        options={[
          { action: 'settings', label: cellMenuText.validationSettings },
          { action: 'circleInvalid', label: cellMenuText.validationCircleInvalid },
          { action: 'clearCircles', label: cellMenuText.validationClearCircles },
          {
            action: 'clearValidation',
            label: cellMenuText.validationClearRules,
            separatorBefore: true,
          },
        ]}
        onPick={onDataValidationAction}
      />
    ),
    errorCheckingMenu: (
      <CellMenu<FormulaAuditingAction>
        key="errorChecking"
        command="errorChecking"
        disabled={!instance}
        icon="options"
        label={tr.errorChecking}
        options={[
          { action: 'errorChecking', label: cellMenuText.errorChecking },
          { action: 'traceError', label: cellMenuText.traceError },
          { action: 'ignoreError', label: cellMenuText.ignoreError, separatorBefore: true },
          {
            action: 'circleInvalid',
            label: cellMenuText.validationCircleInvalid,
          },
          { action: 'clearCircles', label: cellMenuText.validationClearCircles },
        ]}
        onPick={onFormulaAuditingAction}
      />
    ),
    clearArrowsMenu: (
      <CellMenu<ClearArrowsAction>
        key="clearArrows"
        command="clearArrows"
        disabled={!instance}
        icon="clearArrows"
        label={tr.removeArrows}
        options={[
          { action: 'clear-all', label: cellMenuText.removeArrowsAll },
          { action: 'clear-precedents', label: cellMenuText.removePrecedentArrows },
          { action: 'clear-dependents', label: cellMenuText.removeDependentArrows },
        ]}
        onPick={onClearArrowsAction}
      />
    ),
    freezeMenu: (
      <CellMenu<FreezeAction>
        key="freeze"
        command="freeze"
        disabled={!instance}
        icon="freeze"
        label={tr.freeze}
        activeAction={currentFreezeAction}
        options={[
          { action: 'none', label: viewToolbarText.freezeNone },
          { action: 'panes', label: viewToolbarText.freezePanes },
          { action: 'topRow', label: viewToolbarText.freezeTopRow },
          { action: 'firstColumn', label: viewToolbarText.freezeFirstColumn },
        ]}
        onPick={onFreezeAction}
      />
    ),
    windowMenu: (
      <CellMenu<WindowAction>
        key="windowVisibility"
        command="windowVisibility"
        disabled={!instance}
        icon="table"
        label={cellMenuText.format}
        options={[
          { action: 'hideRows', label: cellMenuText.hideRows },
          { action: 'showRows', label: cellMenuText.showRows },
          { action: 'hideCols', label: cellMenuText.hideCols, separatorBefore: true },
          { action: 'showCols', label: cellMenuText.showCols },
        ]}
        onPick={onWindowAction}
      />
    ),
    formatTableHomeMenu: (
      <CellMenu<FormatTableAction>
        key="formatTableHome"
        command="formatTableHome"
        disabled={!instance}
        icon="tableStyle"
        label={tr.formatTable}
        options={[
          { action: 'light', label: cellMenuText.tableStyleLight },
          { action: 'medium', label: cellMenuText.tableStyleMedium },
          { action: 'dark', label: cellMenuText.tableStyleDark },
        ]}
        activeButton={active.formatAsTable}
        onPick={onFormatAsTable}
      />
    ),
    formatTableInsertMenu: (
      <CellMenu<FormatTableAction>
        key="formatTableInsert"
        command="formatTableInsert"
        disabled={!instance}
        icon="tableStyle"
        label={tr.formatTable}
        options={[
          { action: 'light', label: cellMenuText.tableStyleLight },
          { action: 'medium', label: cellMenuText.tableStyleMedium },
          { action: 'dark', label: cellMenuText.tableStyleDark },
        ]}
        activeButton={active.formatAsTable}
        onPick={onFormatAsTable}
      />
    ),
    symbolMenu: (
      <CellMenu<SymbolAction>
        key="symbolInsert"
        command="symbolInsert"
        disabled={!instance}
        icon="function"
        label={cellMenuText.symbol}
        options={[
          ...['±', '×', '÷', '≤', '≥', '≠', '≈', '∞', '√', '∑', '∫', 'π'].map((symbol) => ({
            action: symbol,
            label: symbol,
          })),
          ...['Α', 'Β', 'Γ', 'Δ', 'Θ', 'Λ', 'Ξ', 'Π', 'Σ', 'Φ', 'Ψ', 'Ω'].map((symbol, index) => ({
            action: symbol,
            label: symbol,
            separatorBefore: index === 0,
          })),
          ...['$', '€', '¥', '£', '¢', '₩', '₹', '₽'].map((symbol, index) => ({
            action: symbol,
            label: symbol,
            separatorBefore: index === 0,
          })),
          ...['©', '®', '™', '§', '¶', '†', '‡', '•'].map((symbol, index) => ({
            action: symbol,
            label: symbol,
            separatorBefore: index === 0,
          })),
          {
            action: MORE_SYMBOL_ACTION,
            label: cellMenuText.symbolMore,
            separatorBefore: true,
          },
        ]}
        onPick={onSymbolAction}
      />
    ),
    textOrientationMenu: (
      <CellMenu<TextOrientationAction>
        key="textOrientation"
        command="textOrientation"
        disabled={!instance}
        icon="textOrientation"
        label={tr.textOrientation}
        activeAction={active.textOrientation}
        activeButton={active.textOrientation !== 'horizontalText'}
        options={[
          {
            action: 'angleCounterclockwise',
            label: cellMenuText.orientationAngleCounterclockwise,
          },
          { action: 'angleClockwise', label: cellMenuText.orientationAngleClockwise },
          { action: 'verticalText', label: cellMenuText.orientationVerticalText },
          { action: 'rotateTextUp', label: cellMenuText.orientationRotateTextUp },
          { action: 'rotateTextDown', label: cellMenuText.orientationRotateTextDown },
          {
            action: 'horizontalText',
            label: cellMenuText.orientationHorizontalText,
            separatorBefore: true,
          },
          {
            action: 'formatAlignment',
            label: cellMenuText.orientationFormatAlignment,
            separatorBefore: true,
          },
        ]}
        onPick={onTextOrientationAction}
      />
    ),
    themeMenu: (
      <CellMenu<ThemeAction>
        key="pageTheme"
        command="pageTheme"
        disabled={!instance}
        icon="options"
        label={cellMenuText.theme}
        activeAction={currentTheme}
        options={[
          { action: 'paper', label: cellMenuText.themePaper },
          { action: 'ink', label: cellMenuText.themeInk },
          { action: 'contrast', label: cellMenuText.themeContrast },
        ]}
        onPick={onThemeAction}
      />
    ),
    group,
    iconLabel,
    instance,
    lang,
    locale,
    strings,
    workbookStructureProtected,
    mergeMenu: (
      <MergeMenu
        key="merge"
        disabled={!instance}
        activeAction={active.mergeCenter ? 'mergeCenter' : active.merged ? 'mergeCells' : null}
        labels={{
          mergeAndCenter: tr.mergeAndCenter,
          mergeAcross: tr.mergeAcross,
          mergeCells: tr.mergeCells,
          unmergeCells: tr.unmergeCells,
        }}
        onPick={onMergeAction}
      />
    ),
    onBorderPreset,
    onCopy,
    onCut,
    onDeleteCols,
    onDeleteRows,
    onAddIn,
    onFilterToggle,
    onFormatPainter,
    onDrawEraser,
    onDrawPen,
    onInsertCols,
    onInsertRows,
    onMarginPreset,
    onNumberFormat,
    onPageOrientation,
    onPaperSize,
    onPaste,
    pasteMenu: (
      <CellMenu<PasteAction>
        key="paste"
        command="paste"
        disabled={!instance}
        icon="paste"
        label={tr.paste}
        options={[
          { action: 'paste', label: tr.paste },
          {
            action: 'pasteFormulas',
            label: strings.contextMenu.pasteFormulas,
            separatorBefore: true,
          },
          { action: 'pasteFormulasNumFmt', label: strings.contextMenu.pasteFormulasNumFmt },
          { action: 'pasteValues', label: strings.contextMenu.pasteValues, separatorBefore: true },
          { action: 'pasteValuesNumFmt', label: strings.contextMenu.pasteValuesNumFmt },
          { action: 'pasteFormatsOnly', label: strings.contextMenu.pasteFormatsOnly },
          { action: 'pasteTranspose', label: strings.contextMenu.pasteTranspose },
          {
            action: 'insertCopiedCells',
            label: strings.contextMenu.insertCopiedCells,
            separatorBefore: true,
          },
          {
            action: 'pasteSpecial',
            label: strings.contextMenu.pasteSpecial,
            separatorBefore: true,
          },
        ]}
        onPick={onPasteAction}
      />
    ),
    onProtectWorkbook: protectWorkbookFromBackstage,
    onInspectWorkbook: inspectWorkbookFromBackstage,
    onRedo,
    onRemoveDuplicates,
    onScaleFit,
    onScalePercent,
    onAccessibilityCheck,
    onBuiltInReview: (title, items) => setRibbonReportDialog({ title, items }),
    onRunScript: instance || onRunScript ? openScriptDialog : undefined,
    onRecordActions: instance ? recordActions : undefined,
    onAllScripts: openAllScripts,
    onSort,
    onSpellingReview,
    onTranslate,
    onToggleColsHidden,
    onToggleFormulaBar,
    onToggleRowsHidden,
    onUndo,
    onZoom,
    onZoomDialog: openZoomDialog,
    onZoomSelection,
    pdfMenu: (
      <CellMenu<PdfAction>
        key="pdf"
        command="pdf"
        disabled={!instance}
        icon="pdf"
        label={tr.pdf}
        options={[
          { action: 'create', label: cellMenuText.pdfCreate },
          { action: 'share', label: cellMenuText.pdfShare },
          { action: 'preferences', label: cellMenuText.pdfPreferences, separatorBefore: true },
        ]}
        onPick={onPdfAction}
      />
    ),
    optionSelect,
    rowBreak,
    select,
    setBorderStyle: onBorderStyleChange,
    setBorderColor: onBorderColorChange,
    tool,
    tr,
    wrapFormat,
  });
  const fileLabel = strings.ribbon.tabs.file;
  const backstageCopy = strings.backstage;
  const backstageCard = (
    command: string | null,
    title: string,
    body: string,
    onClick?: () => void,
    isDisabled = false,
  ): ReactElement => (
    <button
      type="button"
      className="demo__backstage-card"
      data-ribbon-command={command ?? undefined}
      disabled={isDisabled || !onClick}
      onClick={onClick}
    >
      <strong>{title}</strong>
      <span>{body}</span>
    </button>
  );
  const backstageCommand = (
    command: string | null,
    title: string,
    body: string,
    icon: string,
    onClick?: () => void,
    isDisabled = false,
    isPressed = false,
  ): ReactElement => (
    <button
      type="button"
      className={`demo__backstage-command${isPressed ? ' demo__backstage-command--active' : ''}`}
      data-ribbon-command={command ?? undefined}
      aria-pressed={isPressed ? 'true' : undefined}
      disabled={isDisabled || !onClick}
      onClick={onClick}
    >
      <span className="demo__backstage-command-icon">{icon}</span>
      <span>
        <strong>{title}</strong>
        <span>{body}</span>
      </span>
    </button>
  );
  const ribbonCopy = strings.ribbonDisplay;
  const ribbonDisplayOptions = [
    { id: 'expanded' as const, label: ribbonCopy.expanded },
    { id: 'collapsed' as const, label: ribbonCopy.collapsed },
  ];
  const sortColumnOptions = (() => {
    if (!instance) return [];
    const state = instance.store.getState();
    const range = state.selection.range;
    const headerRow = range.r0;
    const options: { value: number; label: string }[] = [];
    for (let col = range.c0; col <= range.c1; col += 1) {
      const header = cellLabel(
        state.data.cells.get(`${state.data.sheetIndex}:${headerRow}:${col}`) as
          | SheetCell
          | undefined,
      );
      const columnName = colLetter(col);
      options.push({
        value: col,
        label: sortDialog?.hasHeader && header ? header : columnName,
      });
    }
    return options;
  })();
  const removeDuplicateColumnOptions = (() => {
    if (!instance) return [];
    const state = instance.store.getState();
    const range = state.selection.range;
    const headerRow = range.r0;
    const options: { value: number; label: string }[] = [];
    for (let col = range.c0; col <= range.c1; col += 1) {
      const header = cellLabel(
        state.data.cells.get(`${state.data.sheetIndex}:${headerRow}:${col}`) as
          | SheetCell
          | undefined,
      );
      options.push({
        value: col,
        label: removeDuplicatesDialog?.hasHeader && header ? header : colLetter(col),
      });
    }
    return options;
  })();
  const scriptOptions: { value: ScriptCommand; label: string }[] = [
    { value: 'uppercase', label: cellMenuText.scriptCommandUppercase },
    { value: 'lowercase', label: cellMenuText.scriptCommandLowercase },
    { value: 'trim', label: cellMenuText.scriptCommandTrim },
    { value: 'clear', label: cellMenuText.scriptCommandClear },
  ];

  return (
    <div className={`demo__ribbon-shell${ribbonCollapsed ? ' demo__ribbon-shell--collapsed' : ''}`}>
      <div
        ref={tablistRef}
        className="demo__ribbon-tabs"
        role="tablist"
        aria-label={tr.ribbonTabs}
        data-ribbon-collapsed={ribbonCollapsed ? 'true' : 'false'}
        onKeyDown={onRibbonTabKeyDown}
      >
        {ribbonTabs.map((tab) => (
          <button
            key={tab.id}
            type="button"
            className={`demo__ribbon-tab${tab.id === 'file' ? ' demo__ribbon-tab--file' : ''}${
              activeTab === tab.id ? ' demo__ribbon-tab--active' : ''
            }`}
            role="tab"
            data-ribbon-tab={tab.id}
            aria-selected={activeTab === tab.id}
            tabIndex={activeTab === tab.id ? 0 : -1}
            onClick={() => onRibbonTabClick(tab.id)}
            onDoubleClick={() => setRibbonCollapsed((value) => !value)}
          >
            {tab.label}
          </button>
        ))}
      </div>
      {activeTab !== 'file' ? (
        <div className="demo__ribbon-display" ref={ribbonDisplayRef}>
          <button
            type="button"
            className="demo__ribbon-toggle"
            aria-label={ribbonCopy.label}
            aria-haspopup="menu"
            aria-expanded={ribbonDisplayMenuOpen}
            title={ribbonCopy.label}
            onClick={() => setRibbonDisplayMenuOpen((value) => !value)}
            onKeyDown={onRibbonDisplayKeyDown}
          />
          {ribbonDisplayMenuOpen ? (
            <div
              className="demo__ribbon-display-menu"
              role="menu"
              onKeyDown={onRibbonDisplayKeyDown}
            >
              {ribbonDisplayOptions.map((option) => (
                <button
                  key={option.id}
                  type="button"
                  className="demo__ribbon-display-option"
                  role="menuitemradio"
                  aria-checked={(option.id === 'collapsed') === ribbonCollapsed}
                  onClick={() => {
                    setRibbonCollapsed(option.id === 'collapsed');
                    setRibbonDisplayMenuOpen(false);
                  }}
                >
                  {option.label}
                </button>
              ))}
            </div>
          ) : null}
        </div>
      ) : null}
      {activeTab === 'file' ? (
        <div className="demo__backstage" role="dialog" aria-modal="true" aria-label={fileLabel}>
          <nav className="demo__backstage-nav" aria-label={fileLabel}>
            <button
              type="button"
              className="demo__backstage-navitem"
              aria-label={backstageCopy.back}
              onClick={closeBackstage}
            >
              ←
            </button>
            <strong>{fileLabel}</strong>
            <button
              type="button"
              className="demo__backstage-navitem demo__backstage-navitem--active"
            >
              {backstageCopy.info}
            </button>
            <button
              type="button"
              className="demo__backstage-navitem"
              disabled={!onNewWorkbook}
              onClick={onNewWorkbook}
            >
              {backstageCopy.newLabel}
            </button>
            <button
              type="button"
              className="demo__backstage-navitem"
              disabled={!onOpenWorkbook}
              onClick={onOpenWorkbook}
            >
              {backstageCopy.open}
            </button>
            <button
              type="button"
              className="demo__backstage-navitem"
              disabled={!onSaveWorkbook}
              onClick={onSaveWorkbook}
            >
              {backstageCopy.save}
            </button>
            <button
              type="button"
              className="demo__backstage-navitem"
              disabled={!onSaveWorkbookAs}
              onClick={onSaveWorkbookAs}
            >
              {backstageCopy.saveAs}
            </button>
            <button
              type="button"
              className="demo__backstage-navitem"
              data-ribbon-command="print"
              disabled={!instance}
              onClick={() => instance?.print('print')}
            >
              {tr.print}
            </button>
            <button
              type="button"
              className="demo__backstage-navitem"
              data-ribbon-command="pageSetup"
              disabled={!instance}
              onClick={() => instance?.openPageSetup()}
            >
              {backstageCopy.options}
            </button>
          </nav>
          <main className="demo__backstage-main">
            <div className="demo__backstage-title">
              <span className="demo__backstage-xl">X</span>
              <div>
                <h1>{backstageCopy.title}</h1>
                <p>{backstageCopy.subtitle}</p>
              </div>
            </div>
            <section className="demo__backstage-info">
              <div>
                <h2 className="demo__backstage-section-title">{backstageCopy.workbookInfo}</h2>
                <div className="demo__backstage-command-list">
                  {backstageCommand(
                    'protect',
                    backstageCopy.protect,
                    backstageCopy.protectBody,
                    'P',
                    protectWorkbookFromBackstage,
                    !instance,
                    workbookStructureProtected,
                  )}
                  {backstageCommand(
                    'inspect',
                    backstageCopy.inspect,
                    backstageCopy.inspectBody,
                    '!',
                    inspectWorkbookFromBackstage,
                    !instance,
                  )}
                  {backstageCommand(
                    null,
                    backstageCopy.manage,
                    backstageCopy.manageBody,
                    'S',
                    onSaveWorkbookAs,
                  )}
                </div>
              </div>
              <aside className="demo__backstage-properties">
                <h2 className="demo__backstage-section-title">{backstageCopy.properties}</h2>
                <div className="demo__backstage-preview">X</div>
                <dl className="demo__backstage-prop-list">
                  <dt>{backstageCopy.name}</dt>
                  <dd>{backstageCopy.title}</dd>
                  <dt>{backstageCopy.type}</dt>
                  <dd>{backstageCopy.typeValue}</dd>
                  <dt>{backstageCopy.status}</dt>
                  <dd>{backstageCopy.statusValue}</dd>
                  <dt>{backstageCopy.location}</dt>
                  <dd>{backstageCopy.locationValue}</dd>
                </dl>
              </aside>
            </section>
            <div className="demo__backstage-grid">
              {backstageCard(null, backstageCopy.newLabel, backstageCopy.newBody, onNewWorkbook)}
              {backstageCard(null, backstageCopy.open, backstageCopy.openBody, onOpenWorkbook)}
              {backstageCard(null, backstageCopy.save, backstageCopy.saveBody, onSaveWorkbook)}
              {backstageCard(
                null,
                backstageCopy.saveAs,
                backstageCopy.saveAsBody,
                onSaveWorkbookAs,
              )}
              {backstageCard(
                'print',
                tr.print,
                backstageCopy.printBody,
                () => instance?.print('print'),
                !instance,
              )}
              {backstageCard(
                'pageSetup',
                backstageCopy.options,
                backstageCopy.optionsBody,
                () => instance?.openPageSetup(),
                !instance,
              )}
            </div>
          </main>
        </div>
      ) : (
        <div
          className={`demo__ribbon${activeTab === 'home' ? ' demo__ribbon--office365-home' : ''}`}
          role="toolbar"
          aria-label={`${strings.ribbon.tabs[activeTab]} ${tr.ribbon}`}
        >
          {ribbonGroups[activeTab]}
        </div>
      )}
      <input
        ref={sheetBackgroundInputRef}
        type="file"
        accept="image/*"
        hidden
        data-ribbon-file-input="sheetBackground"
        onChange={onSheetBackgroundFileChange}
      />
      {ribbonReportDialog ? (
        <div className="demo__modal" role="presentation">
          <div
            className="demo__modal-panel demo__modal-panel--narrow"
            role="dialog"
            aria-modal="true"
            aria-label={ribbonReportDialog.title}
          >
            <header className="demo__modal-header">
              <h2>{ribbonReportDialog.title}</h2>
              <button
                type="button"
                className="demo__modal-x"
                aria-label={cellMenuText.sortDialogCancel}
                onClick={() => setRibbonReportDialog(null)}
              >
                ×
              </button>
            </header>
            <div className="demo__modal-body demo__sort-dialog">
              {ribbonReportDialog.items.length === 0 ? (
                <p className="demo__modal-note">{strings.reviewReports.noIssues}</p>
              ) : (
                <div className="demo__report-list">
                  {ribbonReportDialog.items.map((item) => (
                    <div
                      key={`${item.severity}-${item.label}-${item.detail}`}
                      className="demo__report-item"
                    >
                      <strong>
                        {item.severity === 'warning'
                          ? strings.reviewReports.warning
                          : strings.reviewReports.info}
                        {' - '}
                        {item.label}
                      </strong>
                      <span>{item.detail}</span>
                    </div>
                  ))}
                </div>
              )}
            </div>
            <footer className="demo__modal-footer">
              <button
                type="button"
                className="demo__btn demo__btn--primary"
                onClick={() => setRibbonReportDialog(null)}
              >
                {cellMenuText.sortDialogApply}
              </button>
            </footer>
          </div>
        </div>
      ) : null}
      {sortDialog ? (
        <div className="demo__modal" role="presentation">
          <div
            className="demo__modal-panel demo__modal-panel--narrow"
            role="dialog"
            aria-modal="true"
            aria-label={cellMenuText.sortDialogTitle}
          >
            <header className="demo__modal-header">
              <h2>{cellMenuText.sortDialogTitle}</h2>
              <button
                type="button"
                className="demo__modal-x"
                aria-label={cellMenuText.sortDialogCancel}
                onClick={() => setSortDialog(null)}
              >
                ×
              </button>
            </header>
            <div className="demo__modal-body demo__sort-dialog">
              <label className="demo__modal-field">
                <span>{cellMenuText.sortDialogColumn}</span>
                <select
                  value={sortDialog.byCol}
                  onChange={(event) => {
                    const byCol = Number(event.currentTarget.value);
                    setSortDialog((draft) => (draft ? { ...draft, byCol } : draft));
                  }}
                >
                  {sortColumnOptions.map((option) => (
                    <option key={option.value} value={option.value}>
                      {option.label}
                    </option>
                  ))}
                </select>
              </label>
              <label className="demo__modal-field">
                <span>{cellMenuText.sortDialogOrder}</span>
                <select
                  value={sortDialog.direction}
                  onChange={(event) => {
                    const direction = event.currentTarget.value as 'asc' | 'desc';
                    setSortDialog((draft) => (draft ? { ...draft, direction } : draft));
                  }}
                >
                  <option value="asc">{cellMenuText.sortDialogAscending}</option>
                  <option value="desc">{cellMenuText.sortDialogDescending}</option>
                </select>
              </label>
              <label className="demo__sort-dialog__check">
                <input
                  type="checkbox"
                  checked={sortDialog.hasHeader}
                  onChange={(event) => {
                    const hasHeader = event.currentTarget.checked;
                    setSortDialog((draft) => (draft ? { ...draft, hasHeader } : draft));
                  }}
                />
                <span>{cellMenuText.sortDialogHeader}</span>
              </label>
            </div>
            <footer className="demo__modal-footer">
              <button type="button" className="demo__btn" onClick={() => setSortDialog(null)}>
                {cellMenuText.sortDialogCancel}
              </button>
              <button
                type="button"
                className="demo__btn demo__btn--primary"
                onClick={applyCustomSort}
              >
                {cellMenuText.sortDialogApply}
              </button>
            </footer>
          </div>
        </div>
      ) : null}
      {removeDuplicatesDialog ? (
        <div className="demo__modal" role="presentation">
          <div
            className="demo__modal-panel demo__modal-panel--narrow"
            role="dialog"
            aria-modal="true"
            aria-label={cellMenuText.removeDuplicatesDialogTitle}
          >
            <header className="demo__modal-header">
              <h2>{cellMenuText.removeDuplicatesDialogTitle}</h2>
              <button
                type="button"
                className="demo__modal-x"
                aria-label={cellMenuText.sortDialogCancel}
                onClick={() => setRemoveDuplicatesDialog(null)}
              >
                ×
              </button>
            </header>
            <div className="demo__modal-body demo__sort-dialog">
              <label className="demo__sort-dialog__check">
                <input
                  type="checkbox"
                  checked={removeDuplicatesDialog.hasHeader}
                  onChange={(event) => {
                    const hasHeader = event.currentTarget.checked;
                    setRemoveDuplicatesDialog((draft) => (draft ? { ...draft, hasHeader } : draft));
                  }}
                />
                <span>{cellMenuText.sortDialogHeader}</span>
              </label>
              <fieldset className="demo__modal-field">
                <legend>{cellMenuText.removeDuplicatesColumns}</legend>
                <div className="demo__modal-actions">
                  <button
                    type="button"
                    className="demo__btn"
                    onClick={() =>
                      setRemoveDuplicatesDialog((draft) =>
                        draft
                          ? { ...draft, columns: removeDuplicateColumnOptions.map((o) => o.value) }
                          : draft,
                      )
                    }
                  >
                    {cellMenuText.removeDuplicatesSelectAll}
                  </button>
                  <button
                    type="button"
                    className="demo__btn"
                    onClick={() =>
                      setRemoveDuplicatesDialog((draft) =>
                        draft ? { ...draft, columns: [] } : draft,
                      )
                    }
                  >
                    {cellMenuText.removeDuplicatesUnselectAll}
                  </button>
                </div>
                {removeDuplicateColumnOptions.map((option) => (
                  <label key={option.value} className="demo__sort-dialog__check">
                    <input
                      type="checkbox"
                      checked={removeDuplicatesDialog.columns.includes(option.value)}
                      onChange={(event) => {
                        const checked = event.currentTarget.checked;
                        setRemoveDuplicatesDialog((draft) => {
                          if (!draft) return draft;
                          const columns = checked
                            ? [...draft.columns, option.value].sort((a, b) => a - b)
                            : draft.columns.filter((col) => col !== option.value);
                          return { ...draft, columns };
                        });
                      }}
                    />
                    <span>{option.label}</span>
                  </label>
                ))}
              </fieldset>
            </div>
            <footer className="demo__modal-footer">
              <button
                type="button"
                className="demo__btn"
                onClick={() => setRemoveDuplicatesDialog(null)}
              >
                {cellMenuText.sortDialogCancel}
              </button>
              <button
                type="button"
                className="demo__btn demo__btn--primary"
                onClick={applyRemoveDuplicatesDialog}
              >
                {cellMenuText.sortDialogApply}
              </button>
            </footer>
          </div>
        </div>
      ) : null}
      {zoomDialog != null ? (
        <div className="demo__modal" role="presentation">
          <div
            className="demo__modal-panel demo__modal-panel--narrow"
            role="dialog"
            aria-modal="true"
            aria-label={tr.zoomDialogTitle}
          >
            <header className="demo__modal-header">
              <h2>{tr.zoomDialogTitle}</h2>
              <button
                type="button"
                className="demo__modal-x"
                aria-label={cellMenuText.sortDialogCancel}
                onClick={() => setZoomDialog(null)}
              >
                ×
              </button>
            </header>
            <div className="demo__modal-body demo__sort-dialog">
              <label className="demo__modal-field">
                <span>{tr.zoomDialogPercent}</span>
                <input
                  type="number"
                  min={10}
                  max={400}
                  value={zoomDialog}
                  onChange={(event) => setZoomDialog(event.currentTarget.value)}
                />
              </label>
            </div>
            <footer className="demo__modal-footer">
              <button type="button" className="demo__btn" onClick={() => setZoomDialog(null)}>
                {cellMenuText.sortDialogCancel}
              </button>
              <button
                type="button"
                className="demo__btn demo__btn--primary"
                onClick={applyZoomDialog}
              >
                {cellMenuText.sortDialogApply}
              </button>
            </footer>
          </div>
        </div>
      ) : null}
      {advancedFilterDialog ? (
        <div className="demo__modal" role="presentation">
          <div
            className="demo__modal-panel demo__modal-panel--narrow"
            role="dialog"
            aria-modal="true"
            aria-label={cellMenuText.advancedFilterDialogTitle}
          >
            <header className="demo__modal-header">
              <h2>{cellMenuText.advancedFilterDialogTitle}</h2>
              <button
                type="button"
                className="demo__modal-x"
                aria-label={cellMenuText.sortDialogCancel}
                onClick={() => setAdvancedFilterDialog(null)}
              >
                ×
              </button>
            </header>
            <div className="demo__modal-body demo__sort-dialog">
              <label className="demo__modal-field">
                <span>{cellMenuText.advancedFilterListRange}</span>
                <input
                  value={advancedFilterDialog.listRange}
                  onChange={(event) => {
                    const listRange = event.currentTarget.value;
                    setAdvancedFilterDialog((draft) => (draft ? { ...draft, listRange } : draft));
                  }}
                />
              </label>
              <label className="demo__modal-field">
                <span>{cellMenuText.advancedFilterCriteriaRange}</span>
                <input
                  value={advancedFilterDialog.criteriaRange}
                  onChange={(event) => {
                    const criteriaRange = event.currentTarget.value;
                    setAdvancedFilterDialog((draft) =>
                      draft ? { ...draft, criteriaRange } : draft,
                    );
                  }}
                />
              </label>
              <label className="demo__modal-field">
                <span>{cellMenuText.advancedFilterCopyTo}</span>
                <input
                  value={advancedFilterDialog.copyTo}
                  onChange={(event) => {
                    const copyTo = event.currentTarget.value;
                    setAdvancedFilterDialog((draft) => (draft ? { ...draft, copyTo } : draft));
                  }}
                />
              </label>
              <label className="demo__sort-dialog__check">
                <input
                  type="checkbox"
                  checked={advancedFilterDialog.uniqueOnly}
                  onChange={(event) => {
                    const uniqueOnly = event.currentTarget.checked;
                    setAdvancedFilterDialog((draft) => (draft ? { ...draft, uniqueOnly } : draft));
                  }}
                />
                <span>{cellMenuText.advancedFilterUniqueOnly}</span>
              </label>
            </div>
            <footer className="demo__modal-footer">
              <button
                type="button"
                className="demo__btn"
                onClick={() => setAdvancedFilterDialog(null)}
              >
                {cellMenuText.sortDialogCancel}
              </button>
              <button
                type="button"
                className="demo__btn demo__btn--primary"
                onClick={applyAdvancedFilterDialog}
              >
                {cellMenuText.sortDialogApply}
              </button>
            </footer>
          </div>
        </div>
      ) : null}
      {dimensionDialog ? (
        <div className="demo__modal" role="presentation">
          <div
            className="demo__modal-panel demo__modal-panel--narrow"
            role="dialog"
            aria-modal="true"
            aria-label={
              dimensionDialog.kind === 'rowHeight' ? cellMenuText.rowHeight : cellMenuText.colWidth
            }
          >
            <header className="demo__modal-header">
              <h2>
                {dimensionDialog.kind === 'rowHeight'
                  ? cellMenuText.rowHeight
                  : cellMenuText.colWidth}
              </h2>
              <button
                type="button"
                className="demo__modal-x"
                aria-label={cellMenuText.sortDialogCancel}
                onClick={() => setDimensionDialog(null)}
              >
                ×
              </button>
            </header>
            <div className="demo__modal-body demo__sort-dialog">
              <label className="demo__modal-field">
                <span>
                  {dimensionDialog.kind === 'rowHeight'
                    ? cellMenuText.rowHeightPrompt
                    : cellMenuText.colWidthPrompt}
                </span>
                <input
                  type="number"
                  min="1"
                  step="1"
                  value={dimensionDialog.value}
                  onChange={(event) => {
                    const value = event.currentTarget.value;
                    setDimensionDialog((draft) => (draft ? { ...draft, value } : draft));
                  }}
                />
              </label>
            </div>
            <footer className="demo__modal-footer">
              <button type="button" className="demo__btn" onClick={() => setDimensionDialog(null)}>
                {cellMenuText.sortDialogCancel}
              </button>
              <button
                type="button"
                className="demo__btn demo__btn--primary"
                onClick={applyDimensionDialog}
              >
                {cellMenuText.sortDialogApply}
              </button>
            </footer>
          </div>
        </div>
      ) : null}
      {sheetRenameDialog ? (
        <div className="demo__modal" role="presentation">
          <div
            className="demo__modal-panel demo__modal-panel--narrow"
            role="dialog"
            aria-modal="true"
            aria-label={strings.sheetTabs.rename}
          >
            <header className="demo__modal-header">
              <h2>{strings.sheetTabs.rename}</h2>
              <button
                type="button"
                className="demo__modal-x"
                aria-label={cellMenuText.sortDialogCancel}
                onClick={() => setSheetRenameDialog(null)}
              >
                ×
              </button>
            </header>
            <div className="demo__modal-body demo__sort-dialog">
              <label className="demo__modal-field">
                <span>
                  {strings.sheetTabs.renameSheet.replace('{name}', sheetRenameDialog.value)}
                </span>
                <input ref={sheetRenameInputRef} defaultValue={sheetRenameDialog.value} />
              </label>
            </div>
            <footer className="demo__modal-footer">
              <button
                type="button"
                className="demo__btn"
                onClick={() => setSheetRenameDialog(null)}
              >
                {cellMenuText.sortDialogCancel}
              </button>
              <button
                type="button"
                className="demo__btn demo__btn--primary"
                onClick={applySheetRenameDialog}
              >
                {cellMenuText.sortDialogApply}
              </button>
            </footer>
          </div>
        </div>
      ) : null}
      {scriptDialog ? (
        <div className="demo__modal" role="presentation">
          <div
            className="demo__modal-panel demo__modal-panel--narrow"
            role="dialog"
            aria-modal="true"
            aria-label={cellMenuText.scriptDialogTitle}
          >
            <header className="demo__modal-header">
              <h2>{cellMenuText.scriptDialogTitle}</h2>
              <button
                type="button"
                className="demo__modal-x"
                aria-label={cellMenuText.sortDialogCancel}
                onClick={() => setScriptDialog(null)}
              >
                ×
              </button>
            </header>
            <div className="demo__modal-body demo__sort-dialog">
              <label className="demo__modal-field">
                <span>{cellMenuText.scriptDialogCommand}</span>
                <select
                  value={scriptDialog.command}
                  onChange={(event) => {
                    const command = event.currentTarget.value as ScriptCommand;
                    setScriptDialog((draft) => (draft ? { ...draft, command } : draft));
                  }}
                >
                  {scriptOptions.map((option) => (
                    <option key={option.value} value={option.value}>
                      {option.label}
                    </option>
                  ))}
                </select>
              </label>
            </div>
            <footer className="demo__modal-footer">
              <button type="button" className="demo__btn" onClick={() => setScriptDialog(null)}>
                {cellMenuText.sortDialogCancel}
              </button>
              <button
                type="button"
                className="demo__btn demo__btn--primary"
                onClick={applyScriptDialog}
              >
                {cellMenuText.sortDialogApply}
              </button>
            </footer>
          </div>
        </div>
      ) : null}
      {textToColumnsDialog ? (
        <div className="demo__modal" role="presentation">
          <div
            className="demo__modal-panel demo__modal-panel--narrow"
            role="dialog"
            aria-modal="true"
            aria-label={cellMenuText.textToColumnsDialogTitle}
          >
            <header className="demo__modal-header">
              <h2>{cellMenuText.textToColumnsDialogTitle}</h2>
              <button
                type="button"
                className="demo__modal-x"
                aria-label={cellMenuText.sortDialogCancel}
                onClick={() => setTextToColumnsDialog(null)}
              >
                ×
              </button>
            </header>
            <div className="demo__modal-body demo__sort-dialog">
              <fieldset className="demo__modal-field">
                <legend>{cellMenuText.textToColumnsDialogDelimiters}</legend>
                {TEXT_TO_COLUMNS_DIALOG_KEYS.map((key) => {
                  const label =
                    key === 'comma'
                      ? cellMenuText.textToColumnsComma
                      : key === 'tab'
                        ? cellMenuText.textToColumnsTab
                        : key === 'semicolon'
                          ? cellMenuText.textToColumnsSemicolon
                          : cellMenuText.textToColumnsSpace;
                  return (
                    <label key={key} className="demo__sort-dialog__check">
                      <input
                        type="checkbox"
                        checked={!!textToColumnsDialog[key]}
                        onChange={(event) => {
                          const checked = event.currentTarget.checked;
                          setTextToColumnsDialog((draft) =>
                            draft ? { ...draft, [key]: checked } : draft,
                          );
                        }}
                      />
                      <span>{label}</span>
                    </label>
                  );
                })}
              </fieldset>
              <label className="demo__sort-dialog__check">
                <input
                  type="checkbox"
                  checked={textToColumnsDialog.collapseConsecutive}
                  onChange={(event) => {
                    const collapseConsecutive = event.currentTarget.checked;
                    setTextToColumnsDialog((draft) =>
                      draft ? { ...draft, collapseConsecutive } : draft,
                    );
                  }}
                />
                <span>{cellMenuText.textToColumnsTreatConsecutive}</span>
              </label>
            </div>
            <footer className="demo__modal-footer">
              <button
                type="button"
                className="demo__btn"
                onClick={() => setTextToColumnsDialog(null)}
              >
                {cellMenuText.sortDialogCancel}
              </button>
              <button
                type="button"
                className="demo__btn demo__btn--primary"
                onClick={applyTextToColumnsDialog}
              >
                {cellMenuText.sortDialogApply}
              </button>
            </footer>
          </div>
        </div>
      ) : null}
    </div>
  );
};

export const Toolbar = SpreadsheetToolbar;
