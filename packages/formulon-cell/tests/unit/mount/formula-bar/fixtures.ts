import type { Addr } from '../../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { defaultStrings } from '../../../../src/i18n/strings.js';
import type { FormulaEditLease } from '../../../../src/interact/formula-edit-lease.js';
import {
  attachFormulaBarController,
  type FormulaBarController,
} from '../../../../src/mount/formula-bar.js';
import { mutators, type SpreadsheetStore } from '../../../../src/store/store.js';
import type { MountedStubSheet } from '../../../test-utils/index.js';

export const selectRange = (
  sheet: MountedStubSheet,
  active: { sheet: number; row: number; col: number },
  range: { sheet: number; r0: number; c0: number; r1: number; c1: number },
  extraRanges: { sheet: number; r0: number; c0: number; r1: number; c1: number }[] = [],
): void => {
  mutators.setActive(sheet.instance.store, active);
  mutators.setRange(sheet.instance.store, range);
  sheet.instance.store.setState((state) => ({
    ...state,
    selection: {
      ...state.selection,
      extraRanges: extraRanges.map((extra) => ({ ...extra })),
    },
  }));
};

export const attachFormulaBarHarness = (
  sheet: MountedStubSheet,
  onValidation: (outcome: {
    severity: 'stop' | 'warning' | 'information';
    title?: string;
    message: string;
  }) => void,
  store: SpreadsheetStore = sheet.instance.store,
  getWorkbook: () => WorkbookHandle = () => sheet.workbook,
  cancelBindingEditor: () => void = () => {},
) => {
  const formulabar = document.createElement('div');
  const fxInput = document.createElement('textarea');
  const fxCancel = document.createElement('button');
  const fxAccept = document.createElement('button');
  formulabar.append(fxCancel, fxAccept, fxInput);
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
    fxAccept,
    fxCancel,
    fxInput,
    getArgHelper: () => argHelper,
    getAutocomplete: () => autocomplete,
    getStrings: () => defaultStrings,
    cancelBindingEditor,
    host: sheet.host,
    onValidation,
    store,
    updateChrome: () => {},
    wb: getWorkbook,
  });
  return { controller, fxInput, detach: controller.detach };
};

export interface ExternalDraftTestHandle {
  readonly anchor: Addr;
  value(): string;
  setValue(raw: string, caret?: number): void;
  commit(): boolean;
  cancel(): void;
  discard(): void;
  subscribe(fn: (raw: string) => void): () => void;
}

export type ExternalDraftTestController = FormulaBarController & {
  beginExternalDraft(
    anchor: Addr,
    seed: string,
    hooks: {
      onFinish(outcome: 'committed' | 'cancelled', restoredFocusTarget?: HTMLElement | null): void;
    },
    options?: { lease?: FormulaEditLease },
  ): ExternalDraftTestHandle | null;
};

export const asExternalDraftController = (
  controller: FormulaBarController,
): ExternalDraftTestController => controller as ExternalDraftTestController;
