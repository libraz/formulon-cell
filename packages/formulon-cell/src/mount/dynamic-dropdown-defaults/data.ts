import { cellValueViolatesValidation } from '../../commands/validate.js';
import { parseAddrKey } from '../../engine/address.js';
import {
  clearValidationInRangeWithEngine,
  mutators,
  type SpreadsheetInstance,
  textToColumns,
} from '../../index.js';
import { rangeContainsAddr } from '../../store/selection-geometry.js';
import { showTextToColumnsDialog } from '../../toolbar/dialogs/text-to-columns.js';
import type { DynamicDropdownsCtx } from '../../toolbar/ribbon/dynamic-dropdowns.js';
import { setMenuControlDisabled } from './menu-feedback.js';
import { normalizedSelectionRange } from './selection.js';

const buildTextToColumnsAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['splitTextToColumns'] =>
  (delimiter) => {
    const range = normalizedSelectionRange(instance);
    instance.history.begin();
    try {
      textToColumns(instance.store.getState(), instance.store, instance.workbook, range, delimiter);
      mutators.replaceCells(instance.store, instance.workbook.cells(range.sheet));
    } finally {
      instance.history.end();
    }
    instance.host.focus();
  };

const buildTextToColumnsCustom =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['splitTextToColumnsCustom'] =>
  async () => {
    const strings = instance.i18n.strings;
    const range = normalizedSelectionRange(instance);
    const previewRows: string[] = [];
    const cells = instance.store.getState().data.cells;
    for (let row = range.r0; row <= Math.min(range.r1, range.r0 + 6); row += 1) {
      const value = cells.get(`${range.sheet}:${row}:${range.c0}`)?.value;
      if (value?.kind === 'text') previewRows.push(value.value);
    }
    const textToColumnsStrings = strings.ribbonMenu as typeof strings.ribbonMenu & {
      textToColumnsDataType: string;
      textToColumnsDelimited: string;
      textToColumnsFixedWidth: string;
      textToColumnsFixedWidthUnavailable: string;
      textToColumnsOther: string;
      textToColumnsPreview: string;
    };
    const result = await showTextToColumnsDialog({
      strings: {
        title: strings.ribbonMenu.textToColumnsDialogTitle,
        dataType: textToColumnsStrings.textToColumnsDataType,
        delimited: textToColumnsStrings.textToColumnsDelimited,
        fixedWidth: textToColumnsStrings.textToColumnsFixedWidth,
        fixedWidthUnavailable: textToColumnsStrings.textToColumnsFixedWidthUnavailable,
        delimiters: strings.ribbonMenu.textToColumnsDialogDelimiters,
        tab: strings.ribbonMenu.textToColumnsTab,
        semicolon: strings.ribbonMenu.textToColumnsSemicolon,
        comma: strings.ribbonMenu.textToColumnsComma,
        space: strings.ribbonMenu.textToColumnsSpace,
        other: textToColumnsStrings.textToColumnsOther,
        treatConsecutive: strings.ribbonMenu.textToColumnsTreatConsecutive,
        preview: textToColumnsStrings.textToColumnsPreview,
        noDelimited: strings.ribbonMenu.textToColumnsNoDelimited,
        ok: strings.hyperlinkDialog.ok,
        cancel: strings.hyperlinkDialog.cancel,
      },
      initialDelimiters: [','],
      previewRows,
    });
    if (result === null) {
      instance.host.focus();
      return;
    }
    const selected = normalizedSelectionRange(instance);
    instance.history.begin();
    try {
      textToColumns(
        instance.store.getState(),
        instance.store,
        instance.workbook,
        selected,
        result.delimiters,
        { collapseConsecutiveDelimiters: result.collapseConsecutiveDelimiters },
      );
      mutators.replaceCells(instance.store, instance.workbook.cells(selected.sheet));
    } finally {
      instance.history.end();
    }
    instance.host.focus();
  };

const buildDataValidationAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyDataValidationAction'] =>
  (action) => {
    if (action === 'manage' || action === 'more' || action === 'open' || action === 'settings') {
      instance.openDataValidationDialog();
      return;
    }
    if (action === 'clear-circles') {
      mutators.clearValidationCircles(instance.store);
      instance.host.focus();
      return;
    }
    if (action === 'clear-rules') {
      clearValidationInRangeWithEngine(
        instance.store,
        instance.history,
        instance.workbook,
        normalizedSelectionRange(instance),
      );
      mutators.clearValidationCircles(instance.store);
      instance.host.focus();
      return;
    }
    if (action !== 'circle-invalid') return;
    const state = instance.store.getState();
    const range = normalizedSelectionRange(instance);
    const invalid = new Set<string>();
    for (const [key, format] of state.format.formats) {
      if (!format.validation) continue;
      const addr = parseAddrKey(key);
      if (!addr || !rangeContainsAddr(range, addr)) continue;
      const value = instance.workbook.getValue(addr);
      if (cellValueViolatesValidation(value, format.validation)) invalid.add(key);
    }
    mutators.setValidationCircles(instance.store, invalid);
    instance.host.focus();
  };

const updateDataValidationMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateDataValidationMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const range = normalizedSelectionRange(instance);
    let hasValidation = false;
    for (const [key, format] of state.format.formats) {
      if (!format.validation) continue;
      const addr = parseAddrKey(key);
      if (addr && rangeContainsAddr(range, addr)) {
        hasValidation = true;
        break;
      }
    }
    const hasValidationCircles = state.errorIndicators.validationCircles.size > 0;
    const strings = instance.i18n.strings.ribbonMenu;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-validation-action]')) {
      const action = button.dataset.validationAction;
      const disabled =
        ((action === 'circle-invalid' || action === 'clear-rules') && !hasValidation) ||
        (action === 'clear-circles' && !hasValidationCircles);
      const reason =
        action === 'clear-circles'
          ? strings.validationClearCirclesRequiresAny
          : strings.validationRequiresRules;
      setMenuControlDisabled(button, disabled, reason);
    }
  };

type DataDropdownDefaults = Pick<
  DynamicDropdownsCtx,
  | 'splitTextToColumns'
  | 'splitTextToColumnsCustom'
  | 'applyDataValidationAction'
  | 'updateDataValidationMenu'
>;

export function createDataDropdownDefaults(instance: SpreadsheetInstance): DataDropdownDefaults {
  return {
    splitTextToColumns: buildTextToColumnsAction(instance),
    splitTextToColumnsCustom: buildTextToColumnsCustom(instance),
    applyDataValidationAction: buildDataValidationAction(instance),
    updateDataValidationMenu: updateDataValidationMenu(instance),
  };
}
