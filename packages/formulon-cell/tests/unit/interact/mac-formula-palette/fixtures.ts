import { vi } from 'vitest';
import type { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { defaultStrings } from '../../../../src/i18n/strings.js';
import type {
  FormulaEditLease,
  FormulaEditLeaseContext,
} from '../../../../src/interact/formula-edit-lease.js';
import {
  attachMacFormulaPalette,
  type MacFormulaArgumentHelp,
} from '../../../../src/interact/mac-formula-palette.js';
import {
  attachFormulaBarController,
  type FormulaBarController,
} from '../../../../src/mount/formula-bar.js';
import type { MountedStubSheet } from '../../../test-utils/index.js';

export const attachFormulaBarHarness = (
  sheet: MountedStubSheet,
): {
  controller: FormulaBarController;
  input: HTMLTextAreaElement;
  detach: () => void;
} => {
  const formulabar = document.createElement('div');
  const input = document.createElement('textarea');
  const cancel = document.createElement('button');
  const accept = document.createElement('button');
  formulabar.append(cancel, accept, input);
  sheet.host.appendChild(formulabar);
  const autocomplete = {
    isOpen: () => false,
    move: () => {},
    acceptHighlighted: () => false,
    close: () => {},
    refresh: () => {},
  };
  const argHelper = { close: () => {}, refresh: () => {} };
  const controller = attachFormulaBarController({
    formulabar,
    fxAccept: accept,
    fxCancel: cancel,
    fxInput: input,
    getArgHelper: () => argHelper,
    getAutocomplete: () => autocomplete,
    getStrings: () => defaultStrings,
    cancelBindingEditor: () => {},
    host: sheet.host,
    onValidation: vi.fn(),
    store: sheet.instance.store,
    updateChrome: () => {},
    wb: () => sheet.workbook,
  });
  return { controller, input, detach: controller.detach };
};

export type PaletteSetupArgs = [
  getWorkbook?: () => WorkbookHandle,
  getPaletteStrings?: () => typeof defaultStrings,
  getPaletteLocale?: () => string,
  getArgumentHelp?: (name: string, index: number, locale: string) => MacFormulaArgumentHelp | null,
  suspend?: (
    formulaBar: FormulaBarController,
    context: FormulaEditLeaseContext,
  ) => FormulaEditLease | null,
];

/** Mounts a palette docked in `sheet` with a live formula bar behind it. */
export const setupPalette = (
  sheet: MountedStubSheet,
  ...[
    getWorkbook = () => sheet.workbook,
    getPaletteStrings = () => defaultStrings,
    getPaletteLocale = () => 'en-US',
    getArgumentHelp,
    suspend,
  ]: PaletteSetupArgs
) => {
  const opener = document.createElement('button');
  opener.textContent = 'open';
  sheet.host.appendChild(opener);
  opener.focus();
  const dock = document.createElement('div');
  dock.className = 'fc-host__taskpane-dock';
  sheet.host.appendChild(dock);
  const formulaBar = attachFormulaBarHarness(sheet);
  const beginDraft = vi.fn((...args: Parameters<FormulaBarController['beginExternalDraft']>) =>
    formulaBar.controller.beginExternalDraft(...args),
  );
  const mirror = vi.fn();
  const anchor = { sheet: 0, row: 0, col: 0 };
  const palette = attachMacFormulaPalette({
    host: sheet.host,
    dock,
    store: sheet.instance.store,
    getWb: getWorkbook,
    getLocale: getPaletteLocale,
    getStrings: getPaletteStrings,
    getAnchor: () => anchor,
    beginDraft,
    projectMirror: mirror,
    ...(getArgumentHelp ? { getArgumentHelp } : {}),
    ...(suspend ? { suspendActiveEdit: (context) => suspend(formulaBar.controller, context) } : {}),
  });
  return { anchor, beginDraft, dock, formulaBar, mirror, opener, palette };
};

export const paletteRoot = (palette: { isOpen(): boolean }): HTMLElement => {
  if (!palette.isOpen()) throw new Error('palette is not open');
  const root = document.querySelector<HTMLElement>('.fc-mac-formula-palette');
  if (!root) throw new Error('palette root is missing');
  return root;
};
