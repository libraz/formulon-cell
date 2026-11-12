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
  customPivotTableStyleById,
  customStyleNameError,
  customTableStyleById,
  defaultTableStyleId,
  formatAsTableByStyleId,
  inferTableHasHeaders,
  nextTableName,
  sameStyleName,
  type TableVariant,
  tableNameError,
} from '../../commands/format-as-table.js';
import { runAfterAllowedRangeUnlock, selectionEditRanges } from '../../commands/protection.js';
import { recordConditionalRulesChange, recordTablesChange } from '../../commands/slice-history.js';
import {
  overlapsTable,
  tableChangeAllowed,
  tableSourceForSelection,
  tableWrite,
} from '../../commands/table-actions.js';
import {
  afterClearingFormatting,
  createCustomStyle,
  modifyCustomStyle,
  ribbonStyleTarget,
  type TableStyleKind,
} from '../../commands/table-style-actions.js';
import { addrKey, formatA1Range } from '../../engine/address.js';
import { builtinTableStyleElements } from '../../engine/builtin-table-styles.js';
import { builtinPivotStyleElements } from '../../engine/pivot-style.js';
import type { Range } from '../../engine/types.js';
import { dictionaryLocaleFor } from '../../i18n/strings.js';
import { type DialogEntry, openFormatDialogFrom } from '../../interact/dialog-entry.js';
import { attachFormatDialog, type FormatDialogHandle } from '../../interact/format-dialog.js';
import { createRangePickFormatter } from '../../interact/range-pick-format.js';
import type { SpreadsheetInstance } from '../../mount.js';
import { type ConditionalRule, mutators } from '../../store/store.js';
import type { CellFormat } from '../../store/types.js';
import { describeCellStyleFormat, showCellStyleDialog } from '../../toolbar/dialogs/cell-style.js';
import { showChoiceDialog } from '../../toolbar/dialogs/choice.js';
import {
  showConditionalFormatNumberDialog,
  showConditionalFormatTextDialog,
} from '../../toolbar/dialogs/conditional-format.js';
import { showFormatAsTableDialog } from '../../toolbar/dialogs/format-as-table.js';
import { showMessage } from '../../toolbar/dialogs/prompt.js';
import {
  showTableStyleDialog,
  type TableStyleDialogOptions,
  type TableStyleDialogValue,
} from '../../toolbar/dialogs/table-style.js';
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
      mutators.refreshActiveSheetCells(instance.store, instance.workbook);
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
    const style = picked ?? defaultTableStyleId(instance.store.getState(), instance.workbook);
    const entry = { origin: 'ribbon', commandId: 'formatTableHome' } as const;
    const table = ribbonStyleTarget(instance, 'table', entry.commandId);
    if (!table) {
      void showCreateTableDialog(instance, { style, color, variant }, entry);
      return;
    }
    // Inside a table the gallery restyles it; a refused restyle is silent.
    table.apply(style);
    instance.host.focus();
  };

/** Each kind's gallery menu: the item attribute that names a style, and its Clear footer. */
const STYLES_MENU = {
  table: {
    selector: '[data-table-style]',
    item: 'tableStyle',
    clear: 'clear-table-style',
    commandId: 'formatTableHome',
  },
  pivot: {
    selector: '[data-pivot-table-style]',
    item: 'pivotTableStyle',
    clear: 'clear-pivot-style',
    commandId: 'pivotQuickStyles',
  },
} as const;

/** Marks the current style of the active `kind` object and shows Clear only inside one. */
const updateStylesMenu =
  (instance: SpreadsheetInstance, kind: TableStyleKind) =>
  (menu: HTMLElement): void => {
    const spec = STYLES_MENU[kind];
    const target = ribbonStyleTarget(instance, kind, spec.commandId);
    for (const button of menu.querySelectorAll<HTMLButtonElement>(spec.selector)) {
      const active = target !== null && button.dataset[spec.item] === target.currentId;
      button.setAttribute('role', 'menuitemradio');
      button.setAttribute('aria-checked', String(active));
      button.classList.toggle('fc-tb__menu-item--active', active);
    }
    const clear = menu.querySelector<HTMLButtonElement>(
      `[data-table-style-footer="${spec.clear}"]`,
    );
    if (clear) clear.hidden = !target;
  };

/** The table-creation dialog over the selection (a single cell proposes its current region),
 *  from the entry point `entry` names: Insert > Table by default. A refused entry opens nothing,
 *  and the table is created only on OK. */
export const showCreateTableDialog = async (
  instance: SpreadsheetInstance,
  options: {
    style?: string;
    color?: string;
    variant?: TableVariant;
    /** Clear the direct formatting of the range first, as Apply and Clear Formatting does. */
    clearFormatting?: boolean;
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
  const ctx = { store: instance.store, history: instance.history, ...entry };
  if (options.clearFormatting && !tableChangeAllowed(ctx, range)) return;
  // A range over another table creates nothing, so it clears nothing either.
  if (overlapsTable(instance.store.getState(), range)) {
    instance.host.focus();
    return;
  }
  const create = (): boolean =>
    tableWrite(ctx, range, false, () => {
      let table: unknown = null;
      recordTablesChange(instance.history, instance.store, () => {
        table = formatAsTableByStyleId(
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
      return table !== null;
    });
  // The engine can still refuse the table; its formatting then stays as it was.
  if (options.clearFormatting) afterClearingFormatting(ctx, range, create);
  else create();
  instance.host.focus();
};

/** The first `template` name, numbered from 1, that no custom table or PivotTable style, and no
 *  built-in style, uses. */
const freeStyleName = (instance: SpreadsheetInstance, template: string, first = 1): string => {
  const state = instance.store.getState();
  for (let n = first; ; n += 1) {
    const name = template.replace('{n}', String(n));
    if (!customStyleNameError(state, name)) return name;
  }
};

/** Runs the table style dialog with the instance's differential Format Cells behind Format.
 *  `exceptLabel` is the style being modified, which may keep its own name. */
const runTableStyleDialog = async (
  instance: SpreadsheetInstance,
  opts: Pick<TableStyleDialogOptions, 'kind' | 'mode' | 'initial'>,
  exceptLabel?: string,
): Promise<TableStyleDialogValue | null> => {
  const strings = instance.i18n.strings;
  const format: { dialog: FormatDialogHandle | null } = { dialog: null };
  const summary = strings.mac.styleDialog.summary;
  try {
    return await showTableStyleDialog({
      ...opts,
      mac: instance.host.dataset.fcPlatform === 'mac',
      strings,
      validateName: (name) =>
        customStyleNameError(instance.store.getState(), name, { exceptLabel })
          ? strings.ribbonMenu.tableStyleNameExists
          : null,
      editFormat: (initial, apply) => {
        format.dialog ??= attachFormatDialog({
          host: instance.host,
          store: instance.store,
          strings,
          getWb: () => instance.workbook,
          getLocale: () => instance.i18n.locale,
        });
        format.dialog.open('font', {
          mode: 'dxf',
          dxfForm: 'styleElement',
          initialFormat: initial,
          onApplyDxf: apply,
        });
      },
      // The element line names what the element sets: its font, border and fill.
      describe: (format) => {
        const parts = describeCellStyleFormat(format as CellFormat, summary);
        const font =
          format.bold || format.italic || format.underline || format.strike || format.color;
        return [
          ...(font ? [parts.font] : []),
          ...(format.borders ? [parts.border] : []),
          ...(format.fill ? [parts.fill] : []),
        ].join('; ');
      },
    });
  } finally {
    format.dialog?.detach();
  }
};

/** Writes a dialog result as a new custom style of `kind`, and its default when asked. */
const saveNewStyle = (
  instance: SpreadsheetInstance,
  kind: TableStyleKind,
  value: TableStyleDialogValue,
): void => {
  createCustomStyle(
    { store: instance.store, history: instance.history },
    instance.workbook,
    kind,
    value,
  );
};

/** New Table Style / New PivotTable Style from a gallery footer. */
export const openNewTableStyle = async (
  instance: SpreadsheetInstance,
  kind: TableStyleKind,
): Promise<void> => {
  const t = instance.i18n.strings.ribbonMenu;
  const value = await runTableStyleDialog(instance, {
    kind,
    mode: 'new',
    initial: {
      name: freeStyleName(
        instance,
        kind === 'table' ? t.tableStyleDefaultName : t.pivotStyleDefaultName,
      ),
      elements: {},
      stripeSizes: {},
    },
  });
  if (value) saveNewStyle(instance, kind, value);
  instance.host.focus();
};

/** Modify or Duplicate from a style gallery chip's context menu, for a table or PivotTable style.
 *  Duplicate opens under the Modify title and names the copy after its source's style name (a
 *  built-in's engine name, such as `TableStyleLight20`) with the first free number from 2. */
export const openTableStyleEditor = async (
  instance: SpreadsheetInstance,
  kind: TableStyleKind,
  mode: 'modify' | 'duplicate',
  styleId: string,
): Promise<void> => {
  const state = instance.store.getState();
  const custom =
    kind === 'table'
      ? customTableStyleById(state, styleId)
      : customPivotTableStyleById(state, styleId);
  const themeColors = instance.workbook.getTheme?.()?.colors;
  const builtinElements = custom
    ? null
    : kind === 'table'
      ? builtinTableStyleElements(styleId, themeColors)
      : styleId
        ? builtinPivotStyleElements(styleId, themeColors)
        : null;
  const label = custom ? custom.label : builtinElements ? styleId : null;
  if (label === null || (mode === 'modify' && !custom)) return;
  const defaults = instance.workbook.getTableStyleDefaults();
  const currentDefault = kind === 'table' ? defaults?.tableStyle : defaults?.pivotStyle;
  const value = await runTableStyleDialog(
    instance,
    {
      kind,
      mode,
      initial: {
        name: mode === 'modify' ? label : freeStyleName(instance, `${label} {n}`, 2),
        elements: structuredClone(custom?.elements ?? builtinElements ?? {}),
        stripeSizes: { ...custom?.stripeSizes },
        setAsDefault:
          mode === 'modify' && currentDefault !== undefined && sameStyleName(currentDefault, label),
      },
    },
    mode === 'modify' ? label : undefined,
  );
  if (value && mode === 'modify') {
    modifyCustomStyle(
      { store: instance.store, history: instance.history },
      label,
      { label: value.name, elements: value.elements, stripeSizes: value.stripeSizes },
      { workbook: instance.workbook, kind, setAsDefault: value.setAsDefault },
    );
  } else if (value) {
    saveNewStyle(instance, kind, value);
  }
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
    if (action === 'clear-table-style' || action === 'clear-pivot-style') {
      const kind = action === 'clear-table-style' ? 'table' : 'pivot';
      ribbonStyleTarget(instance, kind, STYLES_MENU[kind].commandId)?.clear();
      instance.host.focus();
      return;
    }
    if (action === 'new-table-style' || action === 'new-pivot-style') {
      await openNewTableStyle(instance, action === 'new-table-style' ? 'table' : 'pivot');
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
    ribbonStyleTarget(instance, 'pivot', STYLES_MENU.pivot.commandId)?.apply(styleId);
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
  | 'updatePivotStylesMenu'
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
    updateTableStylesMenu: updateStylesMenu(instance, 'table'),
    updatePivotStylesMenu: updateStylesMenu(instance, 'pivot'),
    applyPivotTableStyleFromRibbon: buildPivotTableStyleAction(instance),
    applyCellStyleFromRibbon: buildCellStyleAction(instance),
    updateCellStylesMenu: updateCellStylesMenu(instance),
    openCellStyleFooterAction: buildCellStyleFooterAction(instance),
  };
}
