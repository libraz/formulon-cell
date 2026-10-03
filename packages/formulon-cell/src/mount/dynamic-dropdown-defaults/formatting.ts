import { applySelectionFormatPatch, withSelectionFormatOrigin } from '../../commands/format.js';
import {
  recordRepeatableFormatChange,
  type SpreadsheetInstance,
  setNumFmt,
  setRotation,
} from '../../index.js';
import { formatWithPending } from '../../store/pending-format.js';
import { applyMergeAction } from '../../toolbar/merge-action.js';
import type { DynamicDropdownsCtx } from '../../toolbar/ribbon/dynamic-dropdowns.js';

const buildUnderlineAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyUnderlineAction'] =>
  async (action) => {
    recordRepeatableFormatChange(instance.history, instance.store, () => {
      applySelectionFormatPatch(
        instance.store.getState(),
        instance.store,
        { underline: action === 'single' ? true : 'double' },
        { origin: 'ribbon', commandId: 'underline' },
      );
    });
    instance.host.focus();
  };

const buildWrapAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyWrapAction'] =>
  (action) => {
    const patch =
      action === 'shrinkToFit'
        ? { shrinkToFit: true, wrap: false }
        : { wrap: true, shrinkToFit: false };
    recordRepeatableFormatChange(instance.history, instance.store, () => {
      applySelectionFormatPatch(instance.store.getState(), instance.store, patch, {
        origin: 'ribbon',
        commandId: 'wrap',
      });
    });
    instance.host.focus();
  };

const buildTextOrientation =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyTextOrientationAction'] =>
  (action) => {
    if (action === 'format') {
      instance.openFormatDialog();
      return;
    }
    const rotations: Record<string, number> = {
      horizontal: 0,
      ccw: 45,
      cw: -45,
      vertical: 90,
      up: 90,
      down: -90,
    };
    const rotation = rotations[action];
    if (typeof rotation !== 'number') return;
    recordRepeatableFormatChange(instance.history, instance.store, () => {
      withSelectionFormatOrigin(
        instance.store,
        'ribbon',
        () => setRotation(instance.store.getState(), instance.store, rotation),
        'textOrientation',
      );
    });
    instance.host.focus();
  };

const buildMergeAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyMergeAction'] =>
  (action) => {
    if (
      action !== 'unmergeCells' &&
      action !== 'mergeAcross' &&
      action !== 'mergeCenter' &&
      action !== 'mergeCells'
    ) {
      return;
    }
    void applyMergeAction(
      {
        store: instance.store,
        workbook: instance.workbook,
        history: instance.history,
        strings: instance.i18n.strings,
      },
      action,
    ).then((applied) => {
      if (applied) instance.host.focus();
    });
  };

const updateTextOrientationMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateTextOrientationMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const active = state.selection.active;
    const rotation = formatWithPending(state, active)?.rotation ?? 0;
    const current =
      rotation === 45
        ? 'ccw'
        : rotation === -45
          ? 'cw'
          : rotation === 90
            ? 'up'
            : rotation === -90
              ? 'down'
              : 'horizontal';
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-text-orientation]')) {
      const action = button.dataset.textOrientation;
      if (action === 'format') continue;
      const activeItem = action === current;
      button.setAttribute('role', 'menuitemradio');
      button.setAttribute('aria-checked', String(activeItem));
      button.classList.toggle('fc-tb__menu-item--active', activeItem);
    }
  };

const buildCurrencyPresetAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyCurrencyPreset'] =>
  (symbol) => {
    recordRepeatableFormatChange(instance.history, instance.store, () => {
      withSelectionFormatOrigin(
        instance.store,
        'ribbon',
        () =>
          setNumFmt(instance.store.getState(), instance.store, {
            kind: 'currency',
            decimals: 2,
            symbol,
          }),
        'currency',
      );
    });
    instance.host.focus();
  };

const buildCurrencyFooterAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['openCurrencyFooterAction'] =>
  (action) => {
    if (action === 'more') instance.openFormatDialog();
  };

const updateCurrencyMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateCurrencyMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const current = formatWithPending(state, state.selection.active)?.numFmt;
    const activeSymbol = current?.kind === 'currency' ? current.symbol : null;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-currency-preset]')) {
      const active = button.dataset.currencyPreset === activeSymbol;
      button.setAttribute('role', 'menuitemradio');
      button.setAttribute('aria-checked', String(active));
      button.classList.toggle('fc-tb__menu-item--active', active);
    }
  };

type FormattingDropdownDefaults = Pick<
  DynamicDropdownsCtx,
  | 'applyUnderlineAction'
  | 'applyWrapAction'
  | 'applyMergeAction'
  | 'applyTextOrientationAction'
  | 'updateTextOrientationMenu'
  | 'applyCurrencyPreset'
  | 'openCurrencyFooterAction'
  | 'updateCurrencyMenu'
>;

export function createFormattingDropdownDefaults(
  instance: SpreadsheetInstance,
): FormattingDropdownDefaults {
  return {
    applyUnderlineAction: buildUnderlineAction(instance),
    applyWrapAction: buildWrapAction(instance),
    applyMergeAction: buildMergeAction(instance),
    applyTextOrientationAction: buildTextOrientation(instance),
    updateTextOrientationMenu: updateTextOrientationMenu(instance),
    applyCurrencyPreset: buildCurrencyPresetAction(instance),
    openCurrencyFooterAction: buildCurrencyFooterAction(instance),
    updateCurrencyMenu: updateCurrencyMenu(instance),
  };
}
