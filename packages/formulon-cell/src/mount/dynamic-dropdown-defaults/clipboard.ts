import {
  dispatchHostClipboard,
  handlePasteAction,
  type PasteAction,
  type SpreadsheetInstance,
} from '../../index.js';
import type { DynamicDropdownsCtx } from '../../toolbar/ribbon/dynamic-dropdowns.js';
import { setMenuControlDisabled, showInstanceReport } from './menu-feedback.js';

const buildCopyAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyCopyAction'] =>
  async (action) => {
    if (action === 'picture') {
      const ribbonMenu = instance.i18n.strings
        .ribbonMenu as typeof instance.i18n.strings.ribbonMenu & {
        copyAsPicture: string;
      };
      const title = ribbonMenu.copyAsPicture;
      await showInstanceReport(instance, title, [
        {
          severity: 'warning',
          label: title,
          detail: instance.i18n.strings.workbookObjects.compatibilityDetails.cellFormatting,
        },
      ]);
      instance.host.focus();
      return;
    }
    dispatchHostClipboard(instance, 'copy');
    instance.host.focus();
  };

const buildPasteAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyRibbonPasteAction'] =>
  (action) => {
    if (action === 'dialog') {
      instance.openPasteSpecial();
      return;
    }
    // PASTE_SPECIAL_PRESETS is private to handlePasteAction; the action ids
    // accepted here are the ribbon "paste-action" attribute values:
    //   all | formulas | formulas-and-numfmt | values | values-and-numfmt |
    //   formats | transpose | dialog
    // Map them onto handlePasteAction's PasteAction string so we route
    // through the same `instance.pasteSpecial` / clipboard glue already
    // wired up — that takes care of snapshot fallback for `all` / `values`.
    const map: Record<string, PasteAction> = {
      all: 'paste',
      formulas: 'pasteFormulas',
      'formulas-and-numfmt': 'pasteFormulasNumFmt',
      values: 'pasteValues',
      'values-and-numfmt': 'pasteValuesNumFmt',
      formats: 'pasteFormatsOnly',
      transpose: 'pasteTranspose',
    };
    const mapped = map[action];
    if (mapped === 'paste') {
      dispatchHostClipboard(instance, 'paste');
      return;
    }
    if (mapped) handlePasteAction(instance, mapped);
  };

const updatePasteMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updatePasteMenu'] =>
  (menu) => {
    const hasSnapshot = instance.clipboard?.getSnapshot() != null;
    const disabledReason = instance.i18n.strings.ribbon.pasteRequiresClipboard;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-paste-action]')) {
      const action = button.dataset.pasteAction;
      button.hidden = !hasSnapshot && action !== 'all' && action !== 'dialog';
      const disabled = action !== 'all' && !hasSnapshot;
      setMenuControlDisabled(button, disabled, disabledReason);
    }
    for (const separator of menu.querySelectorAll<HTMLElement>('.fc-tb__menu-sep')) {
      separator.hidden = !hasSnapshot;
    }
  };

type ClipboardDropdownDefaults = Pick<
  DynamicDropdownsCtx,
  'applyCopyAction' | 'applyRibbonPasteAction' | 'updatePasteMenu'
>;

export function createClipboardDropdownDefaults(
  instance: SpreadsheetInstance,
): ClipboardDropdownDefaults {
  return {
    applyCopyAction: buildCopyAction(instance),
    applyRibbonPasteAction: buildPasteAction(instance),
    updatePasteMenu: updatePasteMenu(instance),
  };
}
