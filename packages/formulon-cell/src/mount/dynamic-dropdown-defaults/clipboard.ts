import { openPasteSpecialFrom } from '../../interact/dialog-entry.js';
import { focusSheet } from '../../interact/dialog-focus.js';
import type { SpreadsheetInstance } from '../../mount.js';
import type { DynamicDropdownsCtx } from '../../toolbar/ribbon/dynamic-dropdowns.js';
import {
  dispatchHostClipboard,
  handlePasteAction,
  type PasteAction,
} from '../../wrappers/toolbar-actions.js';
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
      focusSheet(instance.host);
      return;
    }
    dispatchHostClipboard(instance, 'copy');
    focusSheet(instance.host);
  };

const buildPasteAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyRibbonPasteAction'] =>
  (action) => {
    if (action === 'dialog') {
      openPasteSpecialFrom(instance.store, { origin: 'ribbon', commandId: 'pasteSpecial' });
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
    // Keep Text Only pastes the clipboard's text as the host's ordinary paste does.
    if (action === 'text-only') {
      dispatchHostClipboard(instance, 'paste');
      return;
    }
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
    // On the Mac, cells on the clipboard and plain text give two different menus.
    const textMenu = menu.querySelector('[data-paste-action="text-only"]') !== null;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-paste-action]')) {
      const action = button.dataset.pasteAction;
      const textRow = action === 'text-only' || action === 'text-wizard';
      if (textRow) {
        button.hidden = hasSnapshot;
        // The wizard needs the host to import text into columns, which it does not do.
        setMenuControlDisabled(
          button,
          action === 'text-wizard',
          instance.i18n.strings.ribbon.disabled,
        );
        continue;
      }
      button.hidden = !hasSnapshot && action !== 'dialog' && !(action === 'all' && !textMenu);
      const disabled = action !== 'all' && action !== 'dialog' && !hasSnapshot;
      setMenuControlDisabled(button, disabled, disabledReason);
    }
    for (const separator of menu.querySelectorAll<HTMLElement>('.fc-tb__menu-sep')) {
      separator.hidden = separator.dataset.pasteText === 'true' ? hasSnapshot : !hasSnapshot;
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
