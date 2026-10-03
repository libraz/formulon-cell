import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { commitCellInput, inputOperation } from '../../../src/commands/cell-input-commit.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { addrKey } from '../../../src/engine/address.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

describe('commands/cell-input-commit', () => {
  const addr = { sheet: 0, row: 0, col: 0 };
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await WorkbookHandle.createDefault({ preferStub: true });
  });

  const withController = (): { onChanged: ReturnType<typeof vi.fn>; dispose: () => void } => {
    const onChanged = vi.fn();
    const controller = new InteractionController({
      store,
      getWb: () => wb,
      history: new History(),
      onChanged,
    });
    const unregister = registerInteractionController(store, controller);
    return {
      onChanged,
      dispose: () => {
        unregister();
        controller.dispose();
      },
    };
  };

  let dispose: (() => void) | null = null;
  afterEach(() => {
    dispose?.();
    dispose = null;
  });

  it('classifies formula entries unless the cell is Text-formatted', () => {
    expect(inputOperation(' =A1')).toBe('formulaEdit');
    expect(inputOperation('=A1', true)).toBe('valueEdit');
    expect(inputOperation('12')).toBe('valueEdit');
  });

  it('writes directly and applies the pending format when no controller is registered', () => {
    mutators.setPendingFormat(store, { addr, format: { bold: true } });
    const result = commitCellInput({ store, wb, addr, raw: '7', origin: 'editor' });

    expect(result).toEqual({ status: 'applied', notice: null });
    expect(wb.getValue(addr)).toEqual({ kind: 'number', value: 7 });
    expect(store.getState().format.formats.get(addrKey(addr))?.bold).toBe(true);
    expect(store.getState().ui.pendingFormat).toBeNull();
  });

  it('hands the write and pending format to a registered controller without a policy', () => {
    const controller = withController();
    dispose = controller.dispose;
    mutators.setPendingFormat(store, { addr, format: { bold: true } });
    const result = commitCellInput({ store, wb, addr, raw: '7', origin: 'formulaBar' });

    expect(result).toEqual({ status: 'applied', notice: null });
    expect(controller.onChanged).toHaveBeenCalledWith(
      expect.objectContaining({ status: 'applied' }),
    );
    expect(wb.getValue(addr)).toEqual({ kind: 'number', value: 7 });
    expect(store.getState().format.formats.get(addrKey(addr))?.bold).toBe(true);
    expect(store.getState().ui.pendingFormat).toBeNull();
  });

  it('reports a throwing write as failed', () => {
    wb.setNumber = () => {
      throw new Error('boom');
    };
    const result = commitCellInput({ store, wb, addr, raw: '7', origin: 'editor' });

    expect(result).toMatchObject({ status: 'failed', operation: 'valueEdit' });
  });
});
