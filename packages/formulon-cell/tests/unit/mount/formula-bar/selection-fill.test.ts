import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  type InteractionController,
  registerInteractionController,
} from '../../../../src/commands/interaction-controller.js';
import type { ChangeBatchResult } from '../../../../src/commands/interaction-policy.js';
import { addrKey } from '../../../../src/engine/address.js';
import { defaultStrings } from '../../../../src/i18n/strings.js';
import { sheetHasFocus } from '../../../../src/interact/dialog-focus.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/index.js';
import { attachFormulaBarHarness } from './fixtures.js';

const refused = (
  rejections: { code: string; reason: string; reasonKey?: string }[],
): ChangeBatchResult =>
  ({ status: 'refused', applied: [], rejections, range: null }) as unknown as ChangeBatchResult;

describe('mount/formula-bar — selection fill (Ctrl+Enter) outcomes', () => {
  let sheet: MountedStubSheet;
  let store: ReturnType<typeof createSpreadsheetStore>;
  const execute = vi.fn<(command: unknown) => ChangeBatchResult>();
  const active = { sheet: 0, row: 0, col: 0 };
  const strings = defaultStrings.permissionReasons;

  beforeEach(async () => {
    sheet = await mountStubSheet();
    store = createSpreadsheetStore();
    mutators.setActive(store, active);
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    execute.mockReset();
  });

  afterEach(() => {
    vi.restoreAllMocks();
    sheet.dispose();
  });

  /** Mounts a bar on the fresh store, types `value` and presses Ctrl+Enter. */
  const fill = (
    value: string,
    opts: { controller: boolean; onValidation?: Parameters<typeof attachFormulaBarHarness>[1] },
  ) => {
    const unregister = opts.controller
      ? registerInteractionController(store, { execute } as unknown as InteractionController)
      : () => {};
    const harness = attachFormulaBarHarness(sheet, opts.onValidation, store);
    harness.fxInput.focus();
    harness.fxInput.dispatchEvent(new FocusEvent('focus'));
    harness.fxInput.value = value;
    harness.fxInput.dispatchEvent(new Event('input'));
    harness.fxInput.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, cancelable: true }),
    );
    return {
      harness,
      cleanup: () => {
        harness.detach();
        unregister();
      },
    };
  };

  it('controller execute throwing warns and shows the localized cellsNotFilled reason', () => {
    const err = new Error('boom');
    execute.mockImplementation(() => {
      throw err;
    });
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const onValidation = vi.fn();
    const { harness, cleanup } = fill('x', { controller: true, onValidation });

    expect(warn).toHaveBeenCalledWith('formulon-cell: formula-bar selection fill failed', err);
    expect(onValidation).toHaveBeenCalledTimes(1);
    expect(onValidation).toHaveBeenCalledWith(
      expect.objectContaining({ severity: 'stop', message: strings.cellsNotFilled }),
    );
    expect(harness.controller.isEditing()).toBe(true);
    expect(sheet.workbook.getValue(active)).toEqual({ kind: 'blank' });
    cleanup();
  });

  it('refused with a built-in rejection shows its localized reason and keeps editing', () => {
    execute.mockReturnValue(
      refused([{ code: 'protected', reason: 'cell is protected', reasonKey: 'cellProtected' }]),
    );
    const onValidation = vi.fn();
    const { harness, cleanup } = fill('x', { controller: true, onValidation });

    expect(onValidation).toHaveBeenCalledTimes(1);
    expect(onValidation).toHaveBeenCalledWith(
      expect.objectContaining({
        severity: 'stop',
        message: defaultStrings.protection.protectedCellEdit,
      }),
    );
    expect(harness.controller.isEditing()).toBe(true);
    expect(sheetHasFocus(sheet.host)).toBe(false);
    cleanup();
  });

  it('refused with a host-text rejection shows only the first rejection unchanged', () => {
    execute.mockReturnValue(
      refused([
        { code: 'operationDenied', reason: 'Host says no' },
        { code: 'protected', reason: 'second', reasonKey: 'cellProtected' },
      ]),
    );
    const onValidation = vi.fn();
    const { cleanup } = fill('x', { controller: true, onValidation });

    expect(onValidation).toHaveBeenCalledTimes(1);
    expect(onValidation).toHaveBeenCalledWith(expect.objectContaining({ message: 'Host says no' }));
    cleanup();
  });

  it('refused with no rejections shows the operationDenied reason', () => {
    execute.mockReturnValue(refused([]));
    const onValidation = vi.fn();
    const { harness, cleanup } = fill('x', { controller: true, onValidation });

    expect(onValidation).toHaveBeenCalledTimes(1);
    expect(onValidation).toHaveBeenCalledWith(
      expect.objectContaining({ severity: 'stop', message: strings.operationDenied }),
    );
    expect(harness.controller.isEditing()).toBe(true);
    cleanup();
  });

  it('refusal without an onValidation callback falls back to console.warn', () => {
    execute.mockReturnValue(refused([]));
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const { cleanup } = fill('x', { controller: true });

    expect(warn).toHaveBeenCalledWith(
      `formulon-cell: formula-bar commit rejected: ${strings.operationDenied}`,
    );
    cleanup();
  });

  it('applied result hands the batch to the controller, ends editing and focuses the sheet', () => {
    execute.mockImplementation(() => {
      sheet.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'stub');
      return { status: 'applied', applied: [], rejections: [], range: null };
    });
    const onValidation = vi.fn();
    const { harness, cleanup } = fill('x', { controller: true, onValidation });

    expect(execute).toHaveBeenCalledTimes(1);
    expect(execute).toHaveBeenCalledWith({
      type: 'cellBatch',
      operation: 'valueEdit',
      origin: 'formulaBar',
      changes: [
        { addr: { sheet: 0, row: 0, col: 0 }, input: 'x' },
        { addr: { sheet: 0, row: 0, col: 1 }, input: 'x' },
      ],
    });
    expect(onValidation).not.toHaveBeenCalled();
    expect(harness.controller.isEditing()).toBe(false);
    expect(sheetHasFocus(sheet.host)).toBe(true);
    expect(store.getState().data.cells.get(addrKey(active))).toBeDefined();
    cleanup();
  });

  it('direct path (no controller) rejects on a stop validation with its title and message', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 1 },
      {
        validation: {
          kind: 'list',
          source: ['accepted'],
          errorTitle: 'Bad value',
          errorMessage: 'Choose an accepted value.',
          showErrorMessage: true,
        },
      },
    );
    const onValidation = vi.fn();
    const { harness, cleanup } = fill('nope', { controller: false, onValidation });

    expect(onValidation).toHaveBeenCalledTimes(1);
    expect(onValidation).toHaveBeenCalledWith(
      expect.objectContaining({
        severity: 'stop',
        title: 'Bad value',
        message: 'Choose an accepted value.',
      }),
    );
    expect(sheet.workbook.getValue(active)).toEqual({ kind: 'blank' });
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'blank' });
    expect(harness.controller.isEditing()).toBe(true);
    cleanup();
  });

  it('direct path success refreshes the store cells, ends editing and focuses the sheet', () => {
    const { harness, cleanup } = fill('both', { controller: false });

    expect(sheet.workbook.getValue(active)).toEqual({ kind: 'text', value: 'both' });
    expect(store.getState().data.cells.get(addrKey({ sheet: 0, row: 0, col: 1 }))).toBeDefined();
    expect(harness.controller.isEditing()).toBe(false);
    expect(sheetHasFocus(sheet.host)).toBe(true);
    cleanup();
  });

  it('selection over the unique-cell limit refuses before any write on either path', () => {
    for (const controller of [false, true]) {
      execute.mockReset();
      store.setState((s) => ({
        ...s,
        selection: {
          ...s.selection,
          active,
          anchor: active,
          range: { sheet: 0, r0: 0, c0: 0, r1: 100_000, c1: 0 },
          extraRanges: [],
        },
      }));
      const onValidation = vi.fn();
      const { harness, cleanup } = fill('x', { controller, onValidation });

      expect(execute).not.toHaveBeenCalled();
      expect(onValidation).toHaveBeenCalledWith(
        expect.objectContaining({ severity: 'stop', message: strings.tooManyCellsChanged }),
      );
      expect(harness.controller.isEditing()).toBe(true);
      expect(sheet.workbook.getValue(active)).toEqual({ kind: 'blank' });
      cleanup();
    }
  });
});
