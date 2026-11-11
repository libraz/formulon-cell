import { canExecuteBuiltInRequest } from '../../commands/built-in-command-policy.js';
import {
  activeCellStyleId,
  applyCellStyleToSelection,
  cellStyleFallbackProfileForPlatform,
  createCellStyleFromActiveFormat,
  mergeCellStylesFromWorkbook,
} from '../../commands/cell-styles.js';
import { addConditionalRule } from '../../commands/conditional-format.js';
import {
  applyPivotTableStyleById,
  type CustomTableStyle,
  createPivotTableStyleFromActivePivot,
  createTableStyleFromActiveTable,
  DEFAULT_TABLE_COLOR,
  defaultTableStyleId,
  formatAsTableByStyleId,
  inferTableHasHeaders,
  nextTableName,
  overlayStyleName,
  pivotTableStyleAssignment,
  tableNameError,
  tableOverlayAt,
  tableVariantFromOptions,
} from '../../commands/format-as-table.js';
import { runAfterAllowedRangeUnlock, selectionEditRanges } from '../../commands/protection.js';
import { recordConditionalRulesChange, recordTablesChange } from '../../commands/slice-history.js';
import { tableSourceForSelection, tableWrite } from '../../commands/table-actions.js';
import { applyTableStyle, clearTableStyle } from '../../commands/table-style-actions.js';
import { addrKey, formatA1Range } from '../../engine/address.js';
import { findPivotTableAtCell } from '../../engine/passthrough-sync.js';
import type { Range } from '../../engine/types.js';
import { dictionaryLocaleFor } from '../../i18n/strings.js';
import { type DialogEntry, openFormatDialogFrom } from '../../interact/dialog-entry.js';
import { createRangePickFormatter } from '../../interact/range-pick-format.js';
import type { SpreadsheetInstance } from '../../mount.js';
import { type ConditionalRule, mutators } from '../../store/store.js';
import { describeCellStyleFormat, showCellStyleDialog } from '../../toolbar/dialogs/cell-style.js';
import { showChoiceDialog } from '../../toolbar/dialogs/choice.js';
import {
  showConditionalFormatNumberDialog,
  showConditionalFormatTextDialog,
} from '../../toolbar/dialogs/conditional-format.js';
import { showFormatAsTableDialog } from '../../toolbar/dialogs/format-as-table.js';
import { showMessage } from '../../toolbar/dialogs/prompt.js';
import { showTableStyleDialog } from '../../toolbar/dialogs/table-style.js';
import { applyConditionalMenuAction } from '../../toolbar/ribbon/conditional-menu-action.js';
import type { DynamicDropdownsCtx } from '../../toolbar/ribbon/dynamic-dropdowns.js';
import { formatSheetAbsoluteRange, parseA1Range } from '../../wrappers/toolbar-a1.js';
import { showInstanceReport } from './menu-feedback.js';
import { normalizedSelectionRange } from './selection.js';

const buildConditionalMenuAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyConditionalMenuAction'] =>
  async (action, panel) => {
    const strings = instance.i18n.strings;
    const conditionalMenuStrings = strings.conditionalMenu as typeof strings.conditionalMenu & {
      ok: string;
      cancel: string;
      formatWith: string;
      formatPreview: string;
      customFormat: string;
      customFormatTitle: string;
      customFillColor: string;
      customTextColor: string;
      customBold: string;
      customItalic: string;
      customUnderline: string;
      customStrike: string;
      formatLightRed: string;
      formatYellow: string;
      formatGreen: string;
      formatLightRedFill: string;
      formatRedText: string;
      formatRedBorder: string;
      formatRedFill: string;
      formatRedTextFill: string;
      invalidNumber: string;
      invalidText: string;
    };
    const conditionalDialogStrings = {
      ok: conditionalMenuStrings.ok,
      cancel: conditionalMenuStrings.cancel,
      formatWith: conditionalMenuStrings.formatWith,
      formatPreview: conditionalMenuStrings.formatPreview,
      customFormat: conditionalMenuStrings.customFormat,
      customFormatTitle: conditionalMenuStrings.customFormatTitle,
      customFillColor: conditionalMenuStrings.customFillColor,
      customTextColor: conditionalMenuStrings.customTextColor,
      customBold: conditionalMenuStrings.customBold,
      customItalic: conditionalMenuStrings.customItalic,
      customUnderline: conditionalMenuStrings.customUnderline,
      customStrike: conditionalMenuStrings.customStrike,
      formatLightRed: conditionalMenuStrings.formatLightRed,
      formatYellow: conditionalMenuStrings.formatYellow,
      formatGreen: conditionalMenuStrings.formatGreen,
      formatLightRedFill: conditionalMenuStrings.formatLightRedFill,
      formatRedText: conditionalMenuStrings.formatRedText,
      formatRedBorder: conditionalMenuStrings.formatRedBorder,
      formatRedFill: conditionalMenuStrings.formatRedFill,
      formatRedTextFill: conditionalMenuStrings.formatRedTextFill,
      invalidNumber: conditionalMenuStrings.invalidNumber,
      invalidText: conditionalMenuStrings.invalidText,
    };
    const refreshWorkbookCells = (): void => {
      mutators.replaceCells(
        instance.store,
        instance.workbook.cells(instance.store.getState().data.sheetIndex),
      );
    };
    const pickFormat = createRangePickFormatter({
      store: instance.store,
      sheetName: (index) => instance.workbook.sheetName(index),
    });
    await applyConditionalMenuAction(
      {
        inst: instance,
        ribbonLang: dictionaryLocaleFor(instance.i18n.locale),
        range: normalizedSelectionRange(instance),
        cfFill: { fill: '#ffc7ce', color: '#9c0006' },
        promptCfNumber: ({ allowReference, ...spec }) =>
          showConditionalFormatNumberDialog({
            ...spec,
            locale: instance.locale,
            strings: conditionalDialogStrings,
            ...(allowReference
              ? {
                  reference: {
                    pickerLabel: instance.i18n.strings.pivotTableDialog.rangePickerSelect,
                    getValue: () => `=${pickFormat.range()}`,
                    subscribeToRangeChanges: pickFormat.subscribe,
                  },
                }
              : {}),
          }),
        promptCfText: (spec) =>
          showConditionalFormatTextDialog({
            ...spec,
            strings: conditionalDialogStrings,
          }),
        showChoiceDialog,
        showMessage,
        refreshWorkbookCells,
        addConditionalRuleFromRibbon: (rule: ConditionalRule) => {
          recordConditionalRulesChange(instance.history, instance.store, () => {
            addConditionalRule(instance.store, rule);
          });
          refreshWorkbookCells();
        },
      },
      action === 'clear' ? 'clear-selection' : action,
      panel,
    );
    instance.host.focus();
  };

const buildTableStyleAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['createTableFromSelection'] =>
  (picked, color, variant = 'banded') => {
    const state = instance.store.getState();
    const style = picked ?? defaultTableStyleId(state, instance.workbook);
    const active = state.selection.active;
    const table = tableOverlayAt(state, active.sheet, active.row, active.col);
    const entry = { origin: 'ribbon', commandId: 'formatTableHome' } as const;
    if (!table) {
      void showCreateTableDialog(instance, { style, color, variant }, entry);
      return;
    }
    // Inside a table the gallery restyles it; a refused restyle is silent.
    applyTableStyle({ store: instance.store, history: instance.history, ...entry }, table, style);
    instance.host.focus();
  };

const updateTableStylesMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateTableStylesMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const active = state.selection.active;
    const activeTable = tableOverlayAt(state, active.sheet, active.row, active.col);
    const activeStyleName = activeTable ? overlayStyleName(activeTable) : null;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-table-style]')) {
      const activeStyle = activeStyleName !== null && button.dataset.tableStyle === activeStyleName;
      button.setAttribute('role', 'menuitemradio');
      button.setAttribute('aria-checked', String(activeStyle));
      button.classList.toggle('fc-tb__menu-item--active', activeStyle);
    }
    const clear = menu.querySelector<HTMLButtonElement>(
      '[data-table-style-footer="clear-table-style"]',
    );
    if (clear) clear.hidden = !activeTable;

    const activePivot = findPivotTableAtCell(
      instance.workbook as unknown as Parameters<typeof findPivotTableAtCell>[0],
      active,
    );
    const activePivotStyle = activePivot
      ? pivotTableStyleAssignment(state, activePivot.sheetIndex, activePivot.pivotIndex)?.styleId
      : null;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-pivot-table-style]')) {
      const activeStyle = button.dataset.pivotTableStyle === activePivotStyle;
      button.setAttribute('role', 'menuitemradio');
      button.setAttribute('aria-checked', String(activeStyle));
      button.classList.toggle('fc-tb__menu-item--active', activeStyle);
    }
  };

/** The table-creation dialog over the selection (a single cell proposes its current region),
 *  from the entry point `entry` names: Insert > Table by default. A refused entry opens nothing,
 *  and the table is created only on OK. */
export const showCreateTableDialog = async (
  instance: SpreadsheetInstance,
  options: {
    style?: string;
    color?: string;
    variant?: CustomTableStyle['variant'];
  } = {},
  entry: DialogEntry = { origin: 'ribbon', commandId: 'formatTableInsert' },
): Promise<void> => {
  if (!canExecuteBuiltInRequest(instance.store, entry.commandId, entry.origin).allowed) return;
  const selection = tableSourceForSelection(instance.store.getState());
  const sheetName = instance.workbook.sheetName(selection.sheet);
  const pivotDialogStrings = instance.i18n.strings
    .pivotTableDialog as typeof instance.i18n.strings.pivotTableDialog & {
    rangePickerSelect: string;
    createTableTitle: string;
    createTableRangeLabel: string;
    createTableHeadersLabel: string;
    createTableInvalidRange: string;
  };
  const parsedRange = (value: string): Range | null =>
    parseA1Range(value.trim().replace(/^=/, ''), selection.sheet, sheetName) as Range | null;
  // The Mac dialog writes the range as a formula reference on the active
  // sheet, `=$A$1:$D$7`; elsewhere it carries the sheet name.
  const mac = instance.host.dataset.fcPlatform === 'mac';
  const formatRange = (range: Range): string =>
    mac
      ? `=${formatA1Range(range, { absolute: true })}`
      : formatSheetAbsoluteRange(instance.workbook.sheetName(range.sheet), range);
  const pickFormat = createRangePickFormatter({
    store: instance.store,
    sheetName: (index) => instance.workbook.sheetName(index),
  });
  const result = await showFormatAsTableDialog({
    title: pivotDialogStrings.createTableTitle,
    rangeLabel: pivotDialogStrings.createTableRangeLabel,
    headersLabel: pivotDialogStrings.createTableHeadersLabel,
    initialRange: formatRange(selection),
    initialHasHeaders: inferTableHasHeaders(instance.workbook, selection),
    okLabel: pivotDialogStrings.ok,
    cancelLabel: pivotDialogStrings.cancel,
    rangePickerLabel: pivotDialogStrings.rangePickerSelect,
    pickRange: () => (mac ? `=${pickFormat.range()}` : pickFormat.range({ prefix: 'always' })),
    subscribeToRangeChanges: pickFormat.subscribe,
    validateRange: (value) =>
      parsedRange(value) ? null : pivotDialogStrings.createTableInvalidRange,
    // The Mac dialog names the table up front; only an engine table carries a name.
    ...(mac && instance.workbook.capabilities.tableMutate
      ? {
          nameLabel: instance.i18n.strings.mac.createTable.nameLabel,
          initialName: nextTableName(instance.workbook, instance.store.getState().tables.tables),
          validateName: (value: string) => {
            const error = tableNameError(instance.store, '', value);
            const ribbon = instance.i18n.strings.ribbon;
            return error === 'invalid'
              ? ribbon.tableNameInvalid
              : error === 'duplicate'
                ? ribbon.tableNameDuplicate
                : null;
          },
        }
      : {}),
  });
  if (!result) {
    instance.host.focus();
    return;
  }
  const range = parsedRange(result.range);
  if (!range) return;
  tableWrite(
    { store: instance.store, history: instance.history, ...entry },
    range,
    undefined,
    () => {
      recordTablesChange(instance.history, instance.store, () => {
        formatAsTableByStyleId(
          instance.store,
          range,
          options.style ?? defaultTableStyleId(instance.store.getState(), instance.workbook),
          options.color,
          options.variant ?? 'banded',
          {
            showHeader: result.hasHeaders,
            workbook: instance.workbook,
            ...(result.name ? { name: result.name } : {}),
          },
        );
      });
    },
  );
  instance.host.focus();
};

const buildCellStyleAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyCellStyleFromRibbon'] =>
  (id) => {
    runAfterAllowedRangeUnlock(
      instance.store,
      selectionEditRanges(instance.store.getState()),
      () => {
        applyCellStyleToSelection(instance.store, instance.history, id, {
          origin: 'ribbon',
          commandId: 'cellStyles',
          getWorkbook: () => instance.workbook,
          getFallbackProfile: () =>
            cellStyleFallbackProfileForPlatform(instance.host.dataset.fcPlatform),
        });
        instance.host.focus();
      },
    );
  };

const updateCellStylesMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateCellStylesMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const current = activeCellStyleId(state);
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-cell-style]')) {
      const active = button.dataset.cellStyle === current;
      button.setAttribute('role', 'menuitemradio');
      button.setAttribute('aria-checked', String(active));
      button.classList.toggle('fc-tb__menu-item--active', active);
    }
  };

const buildTableStyleFooterAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['openTableStyleFooterAction'] =>
  async (action) => {
    const strings = instance.i18n.strings;
    const activePivot = findPivotTableAtCell(
      instance.workbook as unknown as Parameters<typeof findPivotTableAtCell>[0],
      instance.store.getState().selection.active,
    );
    const state = instance.store.getState();
    const active = state.selection.active;
    const activeTable = tableOverlayAt(state, active.sheet, active.row, active.col);
    if (action === 'clear-table-style') {
      if (activeTable) {
        clearTableStyle(
          {
            store: instance.store,
            history: instance.history,
            origin: 'ribbon',
            commandId: 'formatTableHome',
          },
          activeTable,
        );
      }
      instance.host.focus();
      return;
    }
    const initial = {
      name: strings.ribbonMenu.tableStyleMedium,
      style: activeTable?.style ?? 'medium',
      color: activeTable?.color ?? DEFAULT_TABLE_COLOR,
      variant: tableVariantFromOptions({
        banded: activeTable?.banded ?? true,
        firstCol: activeTable?.firstCol ?? false,
      }),
    } as const;
    if (action === 'new-table-style') {
      const value = await showTableStyleDialog({
        title: strings.ribbonMenu.tableStyleNew,
        strings,
        initial,
      });
      if (value) {
        createTableStyleFromActiveTable(
          instance.store as unknown as Parameters<typeof createTableStyleFromActiveTable>[0],
          instance.history as unknown as Parameters<typeof createTableStyleFromActiveTable>[1],
          value.name,
          {
            style: value.style,
            color: value.color,
            variant: value.variant,
          },
        );
      }
      instance.host.focus();
      return;
    }
    if (action === 'new-pivot-style') {
      const value = await showTableStyleDialog({
        title: strings.ribbonMenu.tableStyleNewPivot,
        strings,
        initial,
      });
      if (value) {
        createPivotTableStyleFromActivePivot(
          instance.store as unknown as Parameters<typeof createPivotTableStyleFromActivePivot>[0],
          instance.history as unknown as Parameters<typeof createPivotTableStyleFromActivePivot>[1],
          value.name,
          activePivot
            ? { sheetIndex: activePivot.sheetIndex, pivotIndex: activePivot.pivotIndex }
            : null,
          {
            style: value.style,
            color: value.color,
            variant: value.variant,
          },
        );
      }
      instance.host.focus();
      return;
    }
    const label =
      action === 'new-pivot-style'
        ? strings.ribbonMenu.tableStyleNewPivot
        : strings.ribbonMenu.tableStyleNew;
    const detail =
      action === 'new-pivot-style'
        ? strings.workbookObjects.compatibilityDetails.pivotAuthoring
        : strings.workbookObjects.compatibilityDetails.formatAsTable;
    await showInstanceReport(instance, label, [{ severity: 'info', label, detail }]);
    instance.host.focus();
  };

const buildPivotTableStyleAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyPivotTableStyleFromRibbon'] =>
  (styleId) => {
    const pivot = findPivotTableAtCell(
      instance.workbook as unknown as Parameters<typeof findPivotTableAtCell>[0],
      instance.store.getState().selection.active,
    );
    if (pivot) {
      applyPivotTableStyleById(
        instance.store as unknown as Parameters<typeof applyPivotTableStyleById>[0],
        instance.history as unknown as Parameters<typeof applyPivotTableStyleById>[1],
        { sheetIndex: pivot.sheetIndex, pivotIndex: pivot.pivotIndex },
        styleId,
      );
    }
    instance.host.focus();
  };

const buildCellStyleFooterAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['openCellStyleFooterAction'] =>
  async (action) => {
    const strings = instance.i18n.strings;
    if (action === 'new-cell-style') {
      // The Mac host shows the desktop Style dialog.
      const macStyle =
        instance.host.dataset.fcPlatform === 'mac' ? strings.mac.styleDialog : undefined;
      const active = instance.store.getState().selection.active;
      const activeFormat = instance.store.getState().format.formats.get(addrKey(active));
      const value = await showCellStyleDialog({
        title: strings.ribbonMenu.cellStyleNew,
        strings,
        initialName: strings.ribbonMenu.cellStyleNormal,
        mac: macStyle && {
          strings: macStyle,
          summaries: describeCellStyleFormat(activeFormat, macStyle.summary),
          preview: {
            fill: activeFormat?.fill,
            color: activeFormat?.color,
            bold: activeFormat?.bold,
            italic: activeFormat?.italic,
          },
          onFormat: () =>
            openFormatDialogFrom(instance.store, undefined, {
              origin: 'ribbon',
              commandId: 'cellStyles',
            }),
        },
      });
      if (value) {
        createCellStyleFromActiveFormat(
          instance.store as unknown as Parameters<typeof createCellStyleFromActiveFormat>[0],
          instance.history as unknown as Parameters<typeof createCellStyleFromActiveFormat>[1],
          normalizedSelectionRange(instance),
          value.name,
          { include: value.include },
        );
      }
      instance.host.focus();
      return;
    }
    if (action === 'merge-cell-style') {
      const result = mergeCellStylesFromWorkbook(
        instance.store as unknown as Parameters<typeof mergeCellStylesFromWorkbook>[0],
        instance.history as unknown as Parameters<typeof mergeCellStylesFromWorkbook>[1],
        instance.workbook as unknown as Parameters<typeof mergeCellStylesFromWorkbook>[2],
      );
      const detail =
        result.imported > 0
          ? strings.ribbonMenu.cellStyleMergeImported.replace('{count}', String(result.imported))
          : strings.workbookObjects.compatibilityDetails.cellFormatting;
      await showInstanceReport(instance, strings.ribbonMenu.cellStyleMerge, [
        {
          severity: result.imported > 0 ? 'info' : 'warning',
          label: strings.ribbonMenu.cellStyleMerge,
          detail,
        },
      ]);
      instance.host.focus();
      return;
    }
    const label = strings.ribbonMenu.cellStyleNew;
    await showInstanceReport(instance, label, [
      {
        severity: 'info',
        label,
        detail: strings.workbookObjects.compatibilityDetails.cellFormatting,
      },
    ]);
    instance.host.focus();
  };

type StylesDropdownDefaults = Pick<
  DynamicDropdownsCtx,
  | 'applyConditionalMenuAction'
  | 'createTableFromSelection'
  | 'openTableStyleFooterAction'
  | 'updateTableStylesMenu'
  | 'applyPivotTableStyleFromRibbon'
  | 'applyCellStyleFromRibbon'
  | 'updateCellStylesMenu'
  | 'openCellStyleFooterAction'
>;

export function createStylesDropdownDefaults(
  instance: SpreadsheetInstance,
): StylesDropdownDefaults {
  return {
    applyConditionalMenuAction: buildConditionalMenuAction(instance),
    createTableFromSelection: buildTableStyleAction(instance),
    openTableStyleFooterAction: buildTableStyleFooterAction(instance),
    updateTableStylesMenu: updateTableStylesMenu(instance),
    applyPivotTableStyleFromRibbon: buildPivotTableStyleAction(instance),
    applyCellStyleFromRibbon: buildCellStyleAction(instance),
    updateCellStylesMenu: updateCellStylesMenu(instance),
    openCellStyleFooterAction: buildCellStyleFooterAction(instance),
  };
}
