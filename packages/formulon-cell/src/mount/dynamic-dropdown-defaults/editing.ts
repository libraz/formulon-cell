import { conditionalRulesForRange } from '../../commands/conditional-format.js';
import {
  executeRibbonFillAction,
  fillDirectionRange,
  fillOutputRange,
  fillRange,
  type RibbonFillAction,
} from '../../commands/fill.js';
import {
  applyAdvancedFilter,
  copyAdvancedFilterResult,
  inferAutoFilterRange,
} from '../../commands/filter.js';
import { activeFilterRange } from '../../commands/filter-ranges.js';
import { executeRibbonFindAction } from '../../commands/goto-special.js';
import { runAfterAllowedRangeUnlock, selectionEditRanges } from '../../commands/protection.js';
import { executeRibbonClearAction } from '../../commands/ribbon-clear.js';
import { recordFormatChange } from '../../commands/slice-history.js';
import {
  inferSortHasHeader,
  removeDuplicates,
  sortActiveColumnAuto,
  sortRangeWithHistory,
} from '../../commands/sort.js';
import {
  addrKey,
  colLetter,
  formatA1Range,
  MAX_COL,
  MAX_ROW,
  parseAddrKey,
} from '../../engine/address.js';
import type { Range } from '../../engine/types.js';
import { formatCell } from '../../engine/value.js';
import { createRangePickFormatter } from '../../interact/range-pick-format.js';
import { executeRibbonFilterDataAction } from '../../interact/ribbon-filter-data.js';
import type { SpreadsheetInstance } from '../../mount.js';
import { selectionContainsAddr } from '../../store/selection-geometry.js';
import { mutators } from '../../store/store.js';
import { showAdvancedFilterDialog } from '../../toolbar/dialogs/advanced-filter.js';
import { showRemoveDuplicatesDialog } from '../../toolbar/dialogs/remove-duplicates.js';
import { type SortDialogColumn, showSortDialog } from '../../toolbar/dialogs/sort.js';
import type { DynamicDropdownsCtx } from '../../toolbar/ribbon/dynamic-dropdowns.js';
import {
  fillSeriesSourceRange,
  inferFillSeriesDirection,
  type RibbonFillSeriesMode,
  showFillSeriesDialog,
} from '../../toolbar/ribbon/fill-series.js';
import { parseA1Range } from '../../wrappers/toolbar-a1.js';
import { setMenuControlDisabled, showInstanceReport } from './menu-feedback.js';
import { normalizedSelectionRange } from './selection.js';

const buildFillDirection =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyFillDirection'] =>
  (direction) => {
    const range = normalizedSelectionRange(instance);
    // A fill writes everything but its source row or column; Flash Fill may reach the whole selection.
    const wholeSelection = direction === 'flash' || direction === 'justify';
    const worked = wholeSelection ? range : fillDirectionRange(range, direction);
    const target = wholeSelection
      ? range
      : fillOutputRange(fillSeriesSourceRange(worked, direction), worked);
    runAfterAllowedRangeUnlock(instance.store, [target], () => {
      executeRibbonFillAction({
        store: instance.store,
        workbook: instance.workbook,
        history: instance.history,
        action: direction satisfies RibbonFillAction,
      });
      instance.host.focus();
    });
  };

const updateFillMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateFillMenu'] =>
  (menu) => {
    const range = normalizedSelectionRange(instance);
    // A single row or column copies from its neighbour, so only the sheet edge leaves nothing to copy.
    const canFillDown = range.r1 > range.r0 || range.r0 > 0;
    const canFillUp = range.r1 > range.r0 || range.r1 < MAX_ROW;
    const canFillRight = range.c1 > range.c0 || range.c0 > 0;
    const canFillLeft = range.c1 > range.c0 || range.c1 < MAX_COL;
    const strings = instance.i18n.strings.ribbonMenu;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-fill]')) {
      const action = button.dataset.fill;
      const disabled =
        action === 'group' ||
        (action === 'down' && !canFillDown) ||
        (action === 'up' && !canFillUp) ||
        (action === 'right' && !canFillRight) ||
        (action === 'left' && !canFillLeft);
      const reason =
        action === 'group'
          ? button.textContent?.trim()
          : action === 'down' || action === 'up'
            ? strings.fillRequiresMultipleRows
            : action === 'right' || action === 'left'
              ? strings.fillRequiresMultipleCols
              : undefined;
      setMenuControlDisabled(button, disabled, reason);
    }
  };

const buildFillSeries = (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyFillSeries'] => {
  const apply: DynamicDropdownsCtx['applyFillSeries'] = async (mode) => {
    const range = normalizedSelectionRange(instance);
    const fillSeriesDialogStrings = instance.i18n.strings.fillSeriesDialog;
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
    runAfterAllowedRangeUnlock(instance.store, [fillOutputRange(src, range)], () => {
      instance.history.begin();
      try {
        recordFormatChange(instance.history, instance.store, () => {
          fillRange(instance.store.getState(), instance.workbook, src, range, {
            copyOnly: choice.mode === 'copy',
            series: choice.mode === 'linear' || undefined,
            growth: choice.mode === 'growth' || undefined,
            dateUnit,
            formatting: 'with',
            store: instance.store,
          });
        });
      } finally {
        instance.history.end();
      }
      instance.host.focus();
    });
  };
  return apply;
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
  'fillGradient',
  'fontOutline',
  'fontShadow',
  'fontCondense',
  'fontExtend',
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
    runAfterAllowedRangeUnlock(
      instance.store,
      selectionEditRanges(instance.store.getState()),
      () => {
        executeRibbonClearAction({
          store: instance.store,
          workbook: instance.workbook,
          history: instance.history,
          action: clearAction,
          commands: instance.commands,
        });
        instance.host.focus();
      },
    );
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
    const disabledReason = instance.i18n.strings.ribbon.clearRequiresTarget;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-clear]')) {
      const action = button.dataset.clear;
      const disabled =
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
        selectionTooLarge: strings.goToDialog.selectionTooLarge,
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

/** Text a header cell shows, or undefined when it is blank. */
const headerText = (
  instance: SpreadsheetInstance,
  sheet: number,
  row: number,
  col: number,
): string | undefined => {
  const shown = instance.workbook.getDisplayText(sheet, row, col);
  if (shown) return shown;
  const cell = instance.store.getState().data.cells.get(addrKey({ sheet, row, col }));
  return cell && cell.value.kind !== 'blank' ? formatCell(cell.value) : undefined;
};

/** The range's key columns by letter, each with its first-row text as the
 *  heading the dialogs show while that row counts as headers. */
const sortColumnsForRange = (instance: SpreadsheetInstance, range: Range): SortDialogColumn[] =>
  Array.from({ length: range.c1 - range.c0 + 1 }, (_, i) => {
    const col = range.c0 + i;
    return {
      value: String(col),
      label: colLetter(col),
      heading: headerText(instance, range.sheet, range.r0, col),
    };
  });

const buildSortMenuAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applySortMenuAction'] =>
  async (action) => {
    if (action === 'asc' || action === 'desc') {
      runAfterAllowedRangeUnlock(
        instance.store,
        [inferAutoFilterRange(instance.store.getState())],
        () => {
          sortActiveColumnAuto({
            store: instance.store,
            workbook: instance.workbook,
            history: instance.history,
            direction: action,
          });
          instance.host.focus();
        },
      );
      return;
    }
    if (action === 'custom') {
      const mac = instance.host.dataset.fcPlatform === 'mac';
      // The dialog's strings, with the Mac wording laid over them on the Mac host.
      const strings = mac
        ? { ...instance.i18n.strings.ribbonMenu, ...instance.i18n.strings.macSortDialog }
        : instance.i18n.strings.ribbonMenu;
      const state = instance.store.getState();
      const range = inferAutoFilterRange(state);
      const columns = sortColumnsForRange(instance, range);
      const result = await showSortDialog({
        host: instance.host,
        mac: mac ? instance.i18n.strings.macSortDialog : undefined,
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
      runAfterAllowedRangeUnlock(instance.store, [range], () => {
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
      });
      return;
    }
    if (action === 'dedupe') {
      const mac = instance.host.dataset.fcPlatform === 'mac';
      // The dialog's strings, with the Mac wording laid over them on the Mac host.
      const strings = mac
        ? { ...instance.i18n.strings.ribbonMenu, ...instance.i18n.strings.macRemoveDuplicates }
        : instance.i18n.strings.ribbonMenu;
      const state = instance.store.getState();
      const range = inferAutoFilterRange(state);
      const columns = sortColumnsForRange(instance, range);
      const result = await showRemoveDuplicatesDialog({
        host: instance.host,
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
      runAfterAllowedRangeUnlock(instance.store, [range], () => {
        instance.history.begin();
        try {
          removeDuplicates(
            instance.store.getState(),
            instance.store,
            instance.workbook,
            range,
            { columns: result.columns.map(Number), hasHeader: result.hasHeader },
            instance.history,
          );
        } finally {
          instance.history.end();
        }
        mutators.replaceCells(instance.store, instance.workbook.cells(range.sheet));
        instance.host.focus();
      });
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
      const mac =
        instance.host.dataset.fcPlatform === 'mac' ? instance.i18n.strings.macAdvancedFilter : null;
      const strings = {
        ...instance.i18n.strings.ribbonMenu,
        ...mac,
      } as typeof instance.i18n.strings.ribbonMenu & {
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
      const pickFormat = createRangePickFormatter({
        store: instance.store,
        sheetName: (index) => instance.workbook.sheetName(index),
      });
      const result = await showAdvancedFilterDialog({
        title: strings.advancedFilterDialogTitle,
        listRangeLabel: strings.advancedFilterListRange,
        criteriaRangeLabel: strings.advancedFilterCriteriaRange,
        copyToLabel: strings.advancedFilterCopyTo,
        uniqueOnlyLabel: strings.advancedFilterUniqueOnly,
        ...(mac?.filterInPlace && mac.copyToLocation
          ? { actionLabels: { inPlace: mac.filterInPlace, copy: mac.copyToLocation } }
          : {}),
        // The Mac dialog seeds the list range pinned, as `$A$1:$D$7`.
        initialListRange: formatA1Range(range, { absolute: mac !== null }),
        okLabel: strings.sortDialogApply,
        cancelLabel: strings.sortDialogCancel,
        rangePickerLabel: strings.advancedFilterRangePicker,
        // The list and criteria ranges are sheet-qualified from the first pick.
        pickRange: () => pickFormat.range({ prefix: 'always' }),
        pickAddress: () => pickFormat.cell(),
        subscribeToRangeChanges: pickFormat.subscribe,
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
          applyAdvancedFilter(
            instance.store.getState(),
            instance.store,
            listRange,
            criteriaRange,
            instance.workbook,
          );
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
        getWb: () => instance.workbook,
      });
      instance.host.focus();
    }
  };

const updateSortMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateSortMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const hasFilterRange = activeFilterRange(state) !== null;
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
