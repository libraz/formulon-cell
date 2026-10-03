import { parseAddrKey } from '../../engine/address.js';
import {
  applyAdvancedFilter,
  colLetter,
  conditionalRulesForRange,
  copyAdvancedFilterResult,
  executeRibbonClearAction,
  executeRibbonFillAction,
  executeRibbonFilterDataAction,
  executeRibbonFindAction,
  fillRange,
  formatA1Cell,
  formatA1Range,
  inferAutoFilterRange,
  inferFillSeriesDirection,
  inferSortHasHeader,
  mutators,
  parseA1Range,
  type Range,
  type RibbonFillAction,
  type RibbonFillSeriesMode,
  recordFormatChange,
  removeDuplicates,
  type SpreadsheetInstance,
  sortActiveColumnAuto,
  sortRangeWithHistory,
} from '../../index.js';
import { selectionContainsAddr } from '../../store/selection-geometry.js';
import { showAdvancedFilterDialog } from '../../toolbar/dialogs/advanced-filter.js';
import { showRemoveDuplicatesDialog } from '../../toolbar/dialogs/remove-duplicates.js';
import { type SortDialogColumn, showSortDialog } from '../../toolbar/dialogs/sort.js';
import type { DynamicDropdownsCtx } from '../../toolbar/ribbon/dynamic-dropdowns.js';
import { fillSeriesSourceRange, showFillSeriesDialog } from '../../toolbar/ribbon/fill-series.js';
import { setMenuControlDisabled, showInstanceReport } from './menu-feedback.js';
import { normalizedSelectionRange } from './selection.js';

const buildFillDirection =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyFillDirection'] =>
  (direction) => {
    executeRibbonFillAction({
      store: instance.store,
      workbook: instance.workbook,
      history: instance.history,
      action: direction satisfies RibbonFillAction,
    });
    instance.host.focus();
  };

const updateFillMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateFillMenu'] =>
  (menu) => {
    const range = normalizedSelectionRange(instance);
    const hasMultipleRows = range.r1 > range.r0;
    const hasMultipleCols = range.c1 > range.c0;
    const strings = instance.i18n.strings.ribbonMenu;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-fill]')) {
      const action = button.dataset.fill;
      const disabled =
        action === 'group' ||
        action === 'justify' ||
        ((action === 'down' || action === 'up') && !hasMultipleRows) ||
        ((action === 'right' || action === 'left') && !hasMultipleCols);
      const reason =
        action === 'group' || action === 'justify'
          ? button.textContent?.trim()
          : action === 'down' || action === 'up'
            ? strings.fillRequiresMultipleRows
            : action === 'right' || action === 'left'
              ? strings.fillRequiresMultipleCols
              : undefined;
      setMenuControlDisabled(button, disabled, reason);
    }
  };

const buildFillSeries =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyFillSeries'] =>
  async (mode) => {
    const range = normalizedSelectionRange(instance);
    const fillSeriesDialogStrings = (
      instance.i18n.strings as typeof instance.i18n.strings & {
        fillSeriesDialog: {
          title: string;
          seriesIn: string;
          columns: string;
          rows: string;
          up: string;
          left: string;
          type: string;
          autoFill: string;
          copy: string;
          day: string;
          weekday: string;
          month: string;
          year: string;
          ok: string;
          cancel: string;
        };
      }
    ).fillSeriesDialog;
    const choice = mode
      ? { direction: inferFillSeriesDirection(range), mode }
      : await showFillSeriesDialog(range, fillSeriesDialogStrings);
    if (!choice) return;
    const src = fillSeriesSourceRange(range, choice.direction);
    if (src.r0 === range.r0 && src.r1 === range.r1 && src.c0 === range.c0 && src.c1 === range.c1) {
      return;
    }
    const dateUnit: RibbonFillSeriesMode | undefined =
      choice.mode === 'days' ||
      choice.mode === 'weekdays' ||
      choice.mode === 'months' ||
      choice.mode === 'years'
        ? choice.mode
        : undefined;
    instance.history.begin();
    try {
      recordFormatChange(instance.history, instance.store, () => {
        fillRange(instance.store.getState(), instance.workbook, src, range, {
          copyOnly: choice.mode === 'copy',
          dateUnit,
          formatting: 'with',
          store: instance.store,
        });
      });
    } finally {
      instance.history.end();
    }
    instance.host.focus();
  };

const visualClearFormatKeys = new Set([
  'cellStyle',
  'numFmt',
  'bold',
  'italic',
  'underline',
  'strike',
  'align',
  'vAlign',
  'wrap',
  'shrinkToFit',
  'indent',
  'rotation',
  'textDirection',
  'borders',
  'color',
  'fill',
  'fillPattern',
  'fillPatternColor',
  'fontFamily',
  'fontSize',
]);

const hasClearableVisualFormat = (format: object): boolean =>
  Object.keys(format).some((key) => visualClearFormatKeys.has(key));

const hasContentsInSelection = (
  state: ReturnType<SpreadsheetInstance['store']['getState']>,
): boolean => {
  const selection = state.selection;
  for (const [key, cell] of state.data.cells) {
    const addr = parseAddrKey(key);
    if (
      addr &&
      addr.sheet === selection.range.sheet &&
      (cell.formula !== null || cell.value.kind !== 'blank') &&
      selectionContainsAddr(selection, addr)
    ) {
      return true;
    }
  }
  return false;
};

const buildClearAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyClearAction'] =>
  (action) => {
    const clearAction = action === 'remove-hyperlinks' ? 'hyperlinks' : action;
    if (
      clearAction !== 'all' &&
      clearAction !== 'formats' &&
      clearAction !== 'contents' &&
      clearAction !== 'comments' &&
      clearAction !== 'hyperlinks' &&
      clearAction !== 'conditional'
    ) {
      return;
    }
    executeRibbonClearAction({
      store: instance.store,
      workbook: instance.workbook,
      history: instance.history,
      action: clearAction,
      commands: instance.commands,
    });
    instance.host.focus();
  };

const updateClearMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateClearMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const range = state.selection.range;
    const hasContentsSelection = hasContentsInSelection(state);
    let hasFormats = false;
    let hasComments = false;
    let hasHyperlinks = false;
    const pending = (
      state.ui as typeof state.ui & {
        pendingFormat?: {
          addr: { sheet: number; row: number; col: number };
          format: object;
        } | null;
      }
    ).pendingFormat;
    if (
      pending &&
      pending.addr.sheet === range.sheet &&
      selectionContainsAddr(state.selection, pending.addr) &&
      Object.keys(pending.format).length > 0
    ) {
      hasFormats = true;
    }
    for (const [key, format] of state.format.formats) {
      const addr = parseAddrKey(key);
      if (!addr || addr.sheet !== range.sheet || !selectionContainsAddr(state.selection, addr)) {
        continue;
      }
      if (typeof format.comment === 'string' && format.comment.length > 0) hasComments = true;
      if (typeof format.hyperlink === 'string' && format.hyperlink.length > 0) {
        hasHyperlinks = true;
      }
      if (hasClearableVisualFormat(format)) hasFormats = true;
      if (hasFormats && hasComments && hasHyperlinks) break;
    }
    const hasConditional = [range, ...(state.selection.extraRanges ?? [])]
      .filter((selected) => selected.sheet === range.sheet)
      .some((selected) => conditionalRulesForRange(state, selected).length > 0);
    const hasAny =
      hasContentsSelection || hasFormats || hasComments || hasHyperlinks || hasConditional;
    const disabledReason = instance.i18n.strings.ribbon.clearRequiresTarget;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-clear]')) {
      const action = button.dataset.clear;
      const disabled =
        (action === 'all' && !hasAny) ||
        (action === 'contents' && !hasContentsSelection) ||
        (action === 'formats' && !hasFormats) ||
        (action === 'comments' && !hasComments) ||
        ((action === 'hyperlinks' || action === 'remove-hyperlinks') && !hasHyperlinks) ||
        (action === 'conditional' && !hasConditional);
      setMenuControlDisabled(button, disabled, disabledReason);
    }
  };

const buildFindSelectAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyFindSelectAction'] =>
  async (action) => {
    const strings = instance.i18n.strings;
    const result = executeRibbonFindAction({
      store: instance.store,
      workbook: instance.workbook,
      action: action as Parameters<typeof executeRibbonFindAction>[0]['action'],
      strings: {
        findSelect: strings.ribbon.findSelect,
        findNoMatches: strings.ribbonMenu.findNoMatches,
        commentNone: strings.ribbonMenu.commentNone,
      },
    });
    if (result.kind === 'open-find') {
      instance.openFindReplace(result.mode);
      return;
    }
    if (result.kind === 'open-go-to') {
      instance.openGoTo();
      return;
    }
    if (result.kind === 'open-go-to-special') {
      instance.openGoToSpecial();
      return;
    }
    if (result.kind === 'open-objects') {
      instance.openWorkbookObjects();
      return;
    }
    if (result.kind === 'report') {
      await showInstanceReport(instance, result.report.title, result.report.items);
      return;
    }
    if (result.kind === 'selected') instance.host.focus();
  };

const sortColumnsForRange = (range: Range): SortDialogColumn[] =>
  Array.from({ length: range.c1 - range.c0 + 1 }, (_, i) => {
    const col = range.c0 + i;
    const letter = colLetter(col);
    return { value: String(col), label: letter };
  });

const buildSortMenuAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applySortMenuAction'] =>
  async (action) => {
    if (action === 'asc' || action === 'desc') {
      sortActiveColumnAuto({
        store: instance.store,
        workbook: instance.workbook,
        history: instance.history,
        direction: action,
      });
      instance.host.focus();
      return;
    }
    if (action === 'custom') {
      const strings = instance.i18n.strings.ribbonMenu;
      const state = instance.store.getState();
      const range = inferAutoFilterRange(state);
      const columns = sortColumnsForRange(range);
      const result = await showSortDialog({
        title: strings.sortDialogTitle,
        columnLabel: strings.sortDialogColumn,
        thenByLabel: strings.sortThenBy,
        noThenByLabel: strings.sortNoThenBy,
        orderLabel: strings.sortDialogOrder,
        headerLabel: strings.sortDialogHeader,
        addLevelLabel: strings.sortAddLevel,
        deleteLevelLabel: strings.sortDeleteLevel,
        copyLevelLabel: strings.sortCopyLevel,
        levelUnavailableLabel: strings.sortLevelUnavailable,
        ascendingLabel: strings.sortDialogAscending,
        descendingLabel: strings.sortDialogDescending,
        columns,
        initialColumn: String(Math.min(Math.max(state.selection.active.col, range.c0), range.c1)),
        initialDirection: 'asc',
        initialHasHeader: inferSortHasHeader(state, range),
        okLabel: strings.sortDialogApply,
        cancelLabel: strings.sortDialogCancel,
      });
      if (!result) {
        instance.host.focus();
        return;
      }
      sortRangeWithHistory({
        store: instance.store,
        workbook: instance.workbook,
        history: instance.history,
        range,
        options: {
          byCol: Number(result.column),
          direction: result.direction,
          keys: result.levels.map((level) => ({
            byCol: Number(level.column),
            direction: level.direction,
          })),
          hasHeader: result.hasHeader,
        },
      });
      instance.host.focus();
      return;
    }
    if (action === 'dedupe') {
      const strings = instance.i18n.strings.ribbonMenu;
      const state = instance.store.getState();
      const range = inferAutoFilterRange(state);
      const columns = sortColumnsForRange(range);
      const result = await showRemoveDuplicatesDialog({
        title: strings.removeDuplicatesDialogTitle,
        columnsLabel: strings.removeDuplicatesColumns,
        headerLabel: strings.sortDialogHeader,
        selectAllLabel: strings.removeDuplicatesSelectAll,
        unselectAllLabel: strings.removeDuplicatesUnselectAll,
        noColumnsLabel: strings.removeDuplicatesNoColumns,
        columns,
        initialColumns: columns.map((column) => column.value),
        initialHasHeader: inferSortHasHeader(state, range),
        okLabel: strings.sortDialogApply,
        cancelLabel: strings.sortDialogCancel,
      });
      if (!result) {
        instance.host.focus();
        return;
      }
      instance.history.begin();
      try {
        removeDuplicates(instance.store.getState(), instance.store, instance.workbook, range, {
          columns: result.columns.map(Number),
          hasHeader: result.hasHeader,
        });
      } finally {
        instance.history.end();
      }
      mutators.replaceCells(instance.store, instance.workbook.cells(range.sheet));
      instance.host.focus();
      return;
    }
    if (action === 'conditional') {
      instance.openConditionalDialog();
      return;
    }
    if (action === 'named') {
      instance.openNamedRangeDialog();
      return;
    }
    if (action === 'filter-advanced') {
      const strings = instance.i18n.strings
        .ribbonMenu as typeof instance.i18n.strings.ribbonMenu & {
        advancedFilterInvalidRange: string;
        advancedFilterRangePicker: string;
      };
      const range = inferAutoFilterRange(instance.store.getState());
      const sheetName = instance.workbook.sheetName(range.sheet);
      const invalidRange = strings.advancedFilterInvalidRange;
      const validateRange = (value: string): string | null =>
        parseA1Range(value.trim(), range.sheet, sheetName) ? null : invalidRange;
      const validateAddress = (value: string): string | null =>
        parseA1Range(value.trim(), range.sheet, sheetName) ? null : invalidRange;
      const result = await showAdvancedFilterDialog({
        title: strings.advancedFilterDialogTitle,
        listRangeLabel: strings.advancedFilterListRange,
        criteriaRangeLabel: strings.advancedFilterCriteriaRange,
        copyToLabel: strings.advancedFilterCopyTo,
        uniqueOnlyLabel: strings.advancedFilterUniqueOnly,
        initialListRange: formatA1Range(range),
        okLabel: strings.sortDialogApply,
        cancelLabel: strings.sortDialogCancel,
        rangePickerLabel: strings.advancedFilterRangePicker,
        pickRange: () => formatA1Range(normalizedSelectionRange(instance)),
        pickAddress: () => {
          const active = instance.store.getState().selection.active;
          return formatA1Cell(active.row, active.col);
        },
        subscribeToRangeChanges: (listener) => instance.store.subscribe(listener),
        validateRange,
        validateAddress,
      });
      if (!result) {
        instance.host.focus();
        return;
      }
      const listRange = parseA1Range(result.listRange, range.sheet, sheetName);
      const criteriaRange = parseA1Range(result.criteriaRange, range.sheet, sheetName);
      const copyToRange = result.copyTo
        ? parseA1Range(result.copyTo, range.sheet, sheetName)
        : null;
      if (!listRange || !criteriaRange || (result.copyTo && !copyToRange)) {
        instance.host.focus();
        return;
      }
      instance.history.begin();
      try {
        if (copyToRange) {
          const copied = copyAdvancedFilterResult(
            instance.store.getState(),
            instance.store,
            listRange,
            criteriaRange,
            { sheet: copyToRange.sheet, row: copyToRange.r0, col: copyToRange.c0 },
            { uniqueOnly: result.uniqueOnly },
            instance.workbook,
          );
          mutators.replaceCells(instance.store, instance.workbook.cells(listRange.sheet));
          await showInstanceReport(instance, strings.advancedFilterDialogTitle, [
            {
              severity: 'info',
              label: strings.filterAdvanced,
              detail: strings.advancedFilterCopiedStatus.replace('{count}', String(copied)),
            },
          ]);
        } else {
          applyAdvancedFilter(instance.store.getState(), instance.store, listRange, criteriaRange);
        }
      } finally {
        instance.history.end();
      }
      instance.host.focus();
      return;
    }
    const filterAction =
      action === 'filter'
        ? 'toggle'
        : action === 'filter-clear'
          ? 'clear'
          : action === 'filter-reapply'
            ? 'reapply'
            : action === 'filter-by-value'
              ? 'filter-by-selected'
              : null;
    if (filterAction) {
      executeRibbonFilterDataAction({
        store: instance.store,
        history: instance.history,
        action: filterAction,
      });
      instance.host.focus();
    }
  };

const updateSortMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateSortMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const hasFilterRange = state.ui.filterRange !== null;
    const hasFilterCriteria = state.ui.filterCriteria.length > 0;
    const strings = instance.i18n.strings.ribbonMenu;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-sort]')) {
      const action = button.dataset.sort;
      const disabled =
        (action === 'filter-clear' && !hasFilterRange) ||
        (action === 'filter-reapply' && !hasFilterCriteria);
      const reason =
        action === 'filter-clear'
          ? strings.filterClearRequiresRange
          : action === 'filter-reapply'
            ? strings.filterReapplyRequiresCriteria
            : undefined;
      setMenuControlDisabled(button, disabled, reason);
      if (action === 'filter') {
        button.setAttribute('aria-pressed', String(hasFilterRange));
        button.classList.toggle('fc-tb__menu-item--active', hasFilterRange);
      }
    }
  };

type EditingDropdownDefaults = Pick<
  DynamicDropdownsCtx,
  | 'applyFillSeries'
  | 'updateFillMenu'
  | 'applyFillDirection'
  | 'applyClearAction'
  | 'updateClearMenu'
  | 'applyFindSelectAction'
  | 'applySortMenuAction'
  | 'updateSortMenu'
>;

export function createEditingDropdownDefaults(
  instance: SpreadsheetInstance,
): EditingDropdownDefaults {
  return {
    applyFillSeries: buildFillSeries(instance),
    updateFillMenu: updateFillMenu(instance),
    applyFillDirection: buildFillDirection(instance),
    applyClearAction: buildClearAction(instance),
    updateClearMenu: updateClearMenu(instance),
    applyFindSelectAction: buildFindSelectAction(instance),
    applySortMenuAction: buildSortMenuAction(instance),
    updateSortMenu: updateSortMenu(instance),
  };
}
