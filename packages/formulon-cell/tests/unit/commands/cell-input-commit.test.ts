import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { commitCellInput, inputOperation } from '../../../src/commands/cell-input-commit.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { setCellLocked, setProtectedSheet } from '../../../src/commands/protection.js';
import { addrKey } from '../../../src/engine/address.js';
import type { Range } from '../../../src/engine/types.js';
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

  const createRealWorkbook = async (): Promise<WorkbookHandle> => {
    const realWb = await WorkbookHandle.createDefault();
    expect(realWb.isStub).toBe(false);
    return realWb;
  };

  const withController = (
    targetWb = wb,
    options: { getBounds?: () => Range } = {},
  ): {
    onChanged: ReturnType<typeof vi.fn>;
    history: History;
    controller: InteractionController;
    dispose: () => void;
  } => {
    const onChanged = vi.fn();
    const history = new History();
    const controller = new InteractionController({
      store,
      getWb: () => targetWb,
      history,
      getBounds: options.getBounds,
      onChanged,
    });
    const unregister = registerInteractionController(store, controller);
    return {
      onChanged,
      history,
      controller,
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

  it('blocks standalone stop validation before a workbook write or history entry', async () => {
    const realWb = await createRealWorkbook();
    const pending = { addr, format: { bold: true } };
    try {
      mutators.setCellFormat(store, addr, {
        validation: {
          kind: 'whole',
          op: '=',
          a: 5,
          errorTitle: 'Five only',
          errorMessage: 'Enter five.',
        },
      });
      mutators.setPendingFormat(store, pending);

      const result = commitCellInput({ store, wb: realWb, addr, raw: '4', origin: 'editor' });

      expect(result).toEqual({
        status: 'rejected',
        operation: 'valueEdit',
        alert: { severity: 'stop', title: 'Five only', message: 'Enter five.' },
      });
      expect(realWb.getValue(addr)).toEqual({ kind: 'blank' });
      expect(realWb.canUndo()).toBe(false);
      expect(store.getState().ui.pendingFormat).toEqual(pending);
      expect(store.getState().format.formats.get(addrKey(addr))?.bold).toBeUndefined();
    } finally {
      realWb.dispose();
    }
  });

  it.each(['warning', 'information'] as const)(
    'writes a real workbook value and returns a %s validation notice',
    async (severity) => {
      const realWb = await createRealWorkbook();
      try {
        mutators.setCellFormat(store, addr, {
          validation: { kind: 'whole', op: 'between', a: 1, b: 5, errorStyle: severity },
        });

        const result = commitCellInput({ store, wb: realWb, addr, raw: '9', origin: 'editor' });

        expect(result).toEqual({
          status: 'applied',
          notice: {
            severity,
            message: 'number must be between 1 and 5',
          },
        });
        expect(realWb.getValue(addr)).toEqual({ kind: 'number', value: 9 });
        expect(realWb.canUndo()).toBe(true);
      } finally {
        realWb.dispose();
      }
    },
  );

  it('rejects a custom rule against a real workbook without writing', async () => {
    const realWb = await createRealWorkbook();
    try {
      mutators.setCellFormat(store, addr, {
        validation: {
          kind: 'custom',
          formula: '=A1>0',
          errorTitle: 'Positive only',
        },
      });

      const result = commitCellInput({ store, wb: realWb, addr, raw: '-5', origin: 'formulaBar' });

      expect(result).toMatchObject({
        status: 'rejected',
        operation: 'valueEdit',
        alert: { severity: 'stop', title: 'Positive only' },
      });
      expect(realWb.getValue(addr)).toEqual({ kind: 'blank' });
      expect(realWb.canUndo()).toBe(false);
    } finally {
      realWb.dispose();
    }
  });

  it('writes a silent stop rule to a real workbook without a notice', async () => {
    const realWb = await createRealWorkbook();
    try {
      mutators.setCellFormat(store, addr, {
        validation: {
          kind: 'whole',
          op: 'between',
          a: 1,
          b: 5,
          errorStyle: 'stop',
          showErrorMessage: false,
        },
      });

      const result = commitCellInput({ store, wb: realWb, addr, raw: '9', origin: 'editor' });

      expect(result).toEqual({ status: 'applied', notice: null });
      expect(realWb.getValue(addr)).toEqual({ kind: 'number', value: 9 });
    } finally {
      realWb.dispose();
    }
  });

  it('classifies and writes merged body input at the anchor with validation and pending style', async () => {
    const realWb = await createRealWorkbook();
    const anchor = { sheet: 0, row: 0, col: 0 };
    const body = { sheet: 0, row: 0, col: 1 };
    const pending = { addr: body, format: { bold: true } };
    try {
      mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
      mutators.setCellFormat(store, anchor, {
        numFmt: { kind: 'text' },
        validation: { kind: 'list', source: ['=A1'] },
      });
      mutators.setPendingFormat(store, pending);

      const result = commitCellInput({
        store,
        wb: realWb,
        addr: body,
        raw: '=A1',
        origin: 'formulaBar',
      });

      expect(result).toEqual({ status: 'applied', notice: null });
      expect(realWb.getValue(anchor)).toEqual({ kind: 'text', value: '=A1' });
      expect(realWb.getValue(body)).toEqual({ kind: 'blank' });
      expect(store.getState().format.formats.get(addrKey(anchor))).toMatchObject({
        numFmt: { kind: 'text' },
        bold: true,
      });
      expect(store.getState().ui.pendingFormat).toBeNull();
    } finally {
      realWb.dispose();
    }
  });

  it.each([false, true] as const)(
    'resolves pending-only merged Text and validation through the %s path',
    async (registered) => {
      const realWb = await createRealWorkbook();
      const handle = registered ? withController(realWb) : null;
      const anchor = { sheet: 0, row: 0, col: 0 };
      const pendingAddr = { sheet: 0, row: 0, col: 1 };
      const submittedAddr = { sheet: 0, row: 0, col: 2 };
      const pending = {
        addr: pendingAddr,
        format: {
          numFmt: { kind: 'text' as const },
          validation: { kind: 'list' as const, source: ['=A1'] },
          bold: true,
        },
      };
      try {
        mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
        mutators.setPendingFormat(store, pending);

        const result = commitCellInput({
          store,
          wb: realWb,
          addr: submittedAddr,
          raw: '=A1',
          origin: 'formulaBar',
        });

        expect(result).toEqual({ status: 'applied', notice: null });
        expect(realWb.getValue(anchor)).toEqual({ kind: 'text', value: '=A1' });
        expect(realWb.getValue(pendingAddr)).toEqual({ kind: 'blank' });
        expect(realWb.getValue(submittedAddr)).toEqual({ kind: 'blank' });
        expect(store.getState().format.formats.get(addrKey(anchor))).toEqual(pending.format);
        expect(store.getState().format.formats.get(addrKey(pendingAddr))).toBeUndefined();
        expect(store.getState().format.formats.get(addrKey(submittedAddr))).toBeUndefined();
        expect(store.getState().ui.pendingFormat).toBeNull();

        if (handle) {
          expect(handle.history.canUndo()).toBe(true);
          expect(handle.history.undo()).toBe(true);
          expect(realWb.getValue(anchor)).toEqual({ kind: 'blank' });
          expect(store.getState().format.formats.get(addrKey(anchor))).toBeUndefined();
          expect(handle.history.redo()).toBe(true);
          expect(realWb.getValue(anchor)).toEqual({ kind: 'text', value: '=A1' });
          expect(store.getState().format.formats.get(addrKey(anchor))).toEqual(pending.format);
        }
      } finally {
        handle?.dispose();
        realWb.dispose();
      }
    },
  );

  it.each([false, true] as const)(
    'retains pending-only merged Stop validation on the %s path',
    async (registered) => {
      const realWb = await createRealWorkbook();
      const handle = registered ? withController(realWb) : null;
      const anchor = { sheet: 0, row: 0, col: 0 };
      const pendingAddr = { sheet: 0, row: 0, col: 1 };
      const submittedAddr = { sheet: 0, row: 0, col: 2 };
      const pending = {
        addr: pendingAddr,
        format: {
          numFmt: { kind: 'text' as const },
          validation: {
            kind: 'list' as const,
            source: ['allowed'],
            errorTitle: 'Allowed only',
          },
          bold: true,
        },
      };
      try {
        mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
        mutators.setPendingFormat(store, pending);

        const result = commitCellInput({
          store,
          wb: realWb,
          addr: submittedAddr,
          raw: '=A1',
          origin: 'editor',
        });

        expect(result).toMatchObject({
          status: 'rejected',
          operation: 'valueEdit',
          alert: { severity: 'stop', title: 'Allowed only' },
        });
        expect(realWb.getValue(anchor)).toEqual({ kind: 'blank' });
        expect(realWb.getValue(pendingAddr)).toEqual({ kind: 'blank' });
        expect(realWb.getValue(submittedAddr)).toEqual({ kind: 'blank' });
        expect(store.getState().format.formats.get(addrKey(anchor))).toBeUndefined();
        expect(store.getState().ui.pendingFormat).toEqual(pending);
        if (handle) expect(handle.history.canUndo()).toBe(false);
      } finally {
        handle?.dispose();
        realWb.dispose();
      }
    },
  );

  it('retains merged pending format when the standalone write fails', async () => {
    const realWb = await createRealWorkbook();
    const anchor = { sheet: 0, row: 0, col: 0 };
    const pendingAddr = { sheet: 0, row: 0, col: 1 };
    const submittedAddr = { sheet: 0, row: 0, col: 2 };
    const pending = {
      addr: pendingAddr,
      format: { numFmt: { kind: 'text' as const }, bold: true },
    };
    try {
      mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
      mutators.setPendingFormat(store, pending);
      realWb.setText = () => {
        throw new Error('boom');
      };

      const result = commitCellInput({
        store,
        wb: realWb,
        addr: submittedAddr,
        raw: '=A1',
        origin: 'editor',
      });

      expect(result).toMatchObject({ status: 'failed', operation: 'valueEdit' });
      expect(store.getState().ui.pendingFormat).toEqual(pending);
      expect(store.getState().format.formats.get(addrKey(anchor))).toBeUndefined();
    } finally {
      realWb.dispose();
    }
  });

  it('ignores a pending format on another cell when committing a merged body', async () => {
    const realWb = await createRealWorkbook();
    const anchor = { sheet: 0, row: 0, col: 0 };
    const submittedAddr = { sheet: 0, row: 0, col: 2 };
    const unrelated = { sheet: 0, row: 1, col: 0 };
    const pending = {
      addr: unrelated,
      format: { numFmt: { kind: 'text' as const }, bold: true },
    };
    try {
      mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
      mutators.setPendingFormat(store, pending);

      const result = commitCellInput({
        store,
        wb: realWb,
        addr: submittedAddr,
        raw: '=A1',
        origin: 'editor',
      });

      expect(result).toEqual({ status: 'applied', notice: null });
      expect(realWb.cellFormula(anchor)).toBe('=A1');
      expect(realWb.getValue(submittedAddr)).toEqual({ kind: 'blank' });
      expect(store.getState().format.formats.get(addrKey(anchor))).toBeUndefined();
      expect(store.getState().ui.pendingFormat).toEqual(pending);
    } finally {
      realWb.dispose();
    }
  });

  it('keeps the requested merged body address for controller bounds rejection', async () => {
    const realWb = await createRealWorkbook();
    const submittedAddr = { sheet: 0, row: 0, col: 2 };
    const pending = {
      addr: { sheet: 0, row: 0, col: 1 },
      format: { numFmt: { kind: 'text' as const }, bold: true },
    };
    const handle = withController(realWb, {
      getBounds: () => ({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }),
    });
    try {
      mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
      mutators.setPendingFormat(store, pending);

      const result = commitCellInput({
        store,
        wb: realWb,
        addr: submittedAddr,
        raw: '=A1',
        origin: 'formulaBar',
      });

      expect(result).toMatchObject({ status: 'rejected', operation: 'valueEdit' });
      expect(handle.history.canUndo()).toBe(false);
      expect((result as { alert: { message: string } | null }).alert?.message).toContain(
        'outside the configured bounds',
      );
      expect(store.getState().ui.pendingFormat).toEqual(pending);
      expect(realWb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });
    } finally {
      handle.dispose();
      realWb.dispose();
    }
  });

  it('keeps the merged pending format when policy rejects an expanded body', async () => {
    const realWb = await createRealWorkbook();
    const submittedAddr = { sheet: 0, row: 0, col: 2 };
    const pending = {
      addr: { sheet: 0, row: 0, col: 1 },
      format: { numFmt: { kind: 'text' as const }, bold: true },
    };
    const handle = withController(realWb);
    try {
      mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
      mutators.setPendingFormat(store, pending);
      handle.controller.setPolicy({
        operations: { valueEdit: true, format: true },
        editable: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }],
      });

      const result = commitCellInput({
        store,
        wb: realWb,
        addr: submittedAddr,
        raw: '=A1',
        origin: 'formulaBar',
      });

      expect(result).toMatchObject({ status: 'rejected', operation: 'valueEdit' });
      expect(handle.history.canUndo()).toBe(false);
      expect(store.getState().ui.pendingFormat).toEqual(pending);
      expect(realWb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });
    } finally {
      handle.dispose();
      realWb.dispose();
    }
  });

  it('records one controller history entry for value and pending style with undo and redo', async () => {
    const realWb = await createRealWorkbook();
    const handle = withController(realWb);
    dispose = handle.dispose;
    try {
      mutators.setPendingFormat(store, { addr, format: { italic: true } });
      const result = commitCellInput({ store, wb: realWb, addr, raw: '7', origin: 'editor' });

      expect(result).toEqual({ status: 'applied', notice: null });
      expect(realWb.getValue(addr)).toEqual({ kind: 'number', value: 7 });
      expect(store.getState().format.formats.get(addrKey(addr))?.italic).toBe(true);
      expect(handle.history.canUndo()).toBe(true);
      expect(handle.history.undo()).toBe(true);
      expect(realWb.getValue(addr)).toEqual({ kind: 'blank' });
      expect(store.getState().format.formats.get(addrKey(addr))).toBeUndefined();
      expect(handle.history.canUndo()).toBe(false);
      expect(handle.history.redo()).toBe(true);
      expect(realWb.getValue(addr)).toEqual({ kind: 'number', value: 7 });
      expect(store.getState().format.formats.get(addrKey(addr))?.italic).toBe(true);
    } finally {
      handle.dispose();
      dispose = null;
      realWb.dispose();
    }
  });

  it('retains matching pending format when controller authorization denies the format', async () => {
    const realWb = await createRealWorkbook();
    const handle = withController(realWb);
    dispose = handle.dispose;
    try {
      const pending = { addr, format: { underline: true } };
      mutators.setPendingFormat(store, pending);
      handle.controller.setPolicy({ operations: { valueEdit: true, format: false } });

      const result = commitCellInput({ store, wb: realWb, addr, raw: '7', origin: 'formulaBar' });

      expect(result).toMatchObject({ status: 'rejected', operation: 'valueEdit' });
      expect(realWb.getValue(addr)).toEqual({ kind: 'blank' });
      expect(store.getState().format.formats.get(addrKey(addr))).toBeUndefined();
      expect(store.getState().ui.pendingFormat).toEqual(pending);
      expect(handle.history.canUndo()).toBe(false);
    } finally {
      handle.dispose();
      dispose = null;
      realWb.dispose();
    }
  });

  it.each([false, true] as const)(
    'rejects a mixed-lock merged edit through the %s path without writing or clearing pending',
    async (registered) => {
      const realWb = await createRealWorkbook();
      const handle = registered ? withController(realWb) : null;
      const anchor = { sheet: 0, row: 0, col: 0 };
      const pendingAddr = { sheet: 0, row: 0, col: 1 };
      const submittedAddr = { sheet: 0, row: 0, col: 2 };
      const pending = { addr: pendingAddr, format: { bold: true } };
      try {
        mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
        setCellLocked(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, false);
        mutators.setPendingFormat(store, pending);
        setProtectedSheet(store, 0, true);
        const beforeFormats = new Map(store.getState().format.formats);
        const result = commitCellInput({
          store,
          wb: realWb,
          addr: submittedAddr,
          raw: '7',
          origin: 'editor',
        });

        expect(result).toMatchObject({
          status: 'rejected',
          operation: 'valueEdit',
          alert: { severity: 'stop', message: 'cell is protected' },
        });
        expect(realWb.getValue(anchor)).toEqual({ kind: 'blank' });
        expect(realWb.getValue(pendingAddr)).toEqual({ kind: 'blank' });
        expect(realWb.getValue(submittedAddr)).toEqual({ kind: 'blank' });
        expect(store.getState().format.formats).toEqual(beforeFormats);
        expect(store.getState().ui.pendingFormat).toEqual(pending);
        if (handle) expect(handle.history.canUndo()).toBe(false);
        else expect(realWb.canUndo()).toBe(false);
      } finally {
        handle?.dispose();
        realWb.dispose();
      }
    },
  );
});
