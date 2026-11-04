import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { commitCellInput, inputOperation } from '../../../src/commands/cell-input-commit.js';
import { setFont } from '../../../src/commands/format.js';
import { defaultObjectNames, formatAsTable } from '../../../src/commands/format-as-table.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { setCellLocked, setProtectedSheet } from '../../../src/commands/protection.js';
import { addrKey, MAX_COL, MAX_ROW } from '../../../src/engine/address.js';
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

  it('writes directly and keeps the cell format when no controller is registered', () => {
    mutators.setCellFormat(store, addr, { bold: true });
    const result = commitCellInput({ store, wb, addr, raw: '7', origin: 'editor' });

    expect(result).toEqual({ status: 'applied', notice: null });
    expect(wb.getValue(addr)).toEqual({ kind: 'number', value: 7 });
    expect(store.getState().format.formats.get(addrKey(addr))?.bold).toBe(true);
  });

  it('keeps a saved worksheet font on controller-less input after reload', async () => {
    const source = await createRealWorkbook();
    const sourceStore = createSpreadsheetStore();
    source.attachStore(sourceStore);
    sourceStore.setState((s) => ({
      ...s,
      selection: {
        ...s.selection,
        range: { sheet: 0, r0: 0, c0: 0, r1: MAX_ROW, c1: MAX_COL },
      },
    }));
    expect(setFont(sourceStore.getState(), sourceStore, { fontSize: 24 })).toBe(true);
    const bytes = source.save();
    source.detachStore(sourceStore);
    source.dispose();

    const loaded = await WorkbookHandle.loadBytes(bytes);
    const loadedStore = createSpreadsheetStore();
    loaded.attachStore(loadedStore);
    const target = { sheet: 0, row: 0, col: 2 } as const;
    const result = commitCellInput({
      store: loadedStore,
      wb: loaded,
      addr: target,
      raw: 'after reload',
      origin: 'editor',
    });
    expect(result).toEqual({ status: 'applied', notice: null });
    expect(loadedStore.getState().format.formats.has(addrKey(target))).toBe(false);
    const effective = loaded.getEffectiveStyle(target.sheet, target.row, target.col);
    expect(effective?.source).toBe(0);
    const xf = effective ? loaded.getCellXf(effective.xfIndex) : null;
    const font = xf ? loaded.getFontRecord(xf.fontIndex) : null;
    expect(font?.size).toBeCloseTo(18, 10);

    const roundTripped = await WorkbookHandle.loadBytes(loaded.save());
    try {
      const roundTripEffective = roundTripped.getEffectiveStyle(
        target.sheet,
        target.row,
        target.col,
      );
      const roundTripXf = roundTripEffective
        ? roundTripped.getCellXf(roundTripEffective.xfIndex)
        : null;
      const roundTripFont = roundTripXf ? roundTripped.getFontRecord(roundTripXf.fontIndex) : null;
      expect(roundTripFont?.size).toBeCloseTo(18, 10);
    } finally {
      loaded.detachStore(loadedStore);
      loaded.dispose();
      roundTripped.dispose();
    }
  });

  it('hands the write to a registered controller without a policy and keeps the cell format', () => {
    const controller = withController();
    dispose = controller.dispose;
    mutators.setCellFormat(store, addr, { bold: true });
    const result = commitCellInput({ store, wb, addr, raw: '7', origin: 'formulaBar' });

    expect(result).toEqual({ status: 'applied', notice: null });
    expect(controller.onChanged).toHaveBeenCalledWith(
      expect.objectContaining({ status: 'applied' }),
    );
    expect(wb.getValue(addr)).toEqual({ kind: 'number', value: 7 });
    expect(store.getState().format.formats.get(addrKey(addr))?.bold).toBe(true);
  });

  it('uses the live workbook profile when committing localized input', async () => {
    const realWb = await createRealWorkbook();
    try {
      expect(realWb.setSpreadsheetProfileId('windows-de_DE')).toBe(true);

      const result = commitCellInput({ store, wb: realWb, addr, raw: '1,25', origin: 'editor' });

      expect(result).toEqual({ status: 'applied', notice: null });
      expect(realWb.getValue(addr)).toEqual({ kind: 'number', value: 1.25 });
    } finally {
      realWb.dispose();
    }
  });

  it('classifies inherited native Text number formats before committing input', async () => {
    const realWb = await createRealWorkbook();
    try {
      const baseXf = realWb.getCellXf(0);
      if (!baseXf) throw new Error('missing workbook default XF');
      const textNumFmt = realWb.addNumFmtCode('@');
      const inheritedTextXf = realWb.addXfRecord({
        ...baseXf,
        numFmtId: textNumFmt,
        applyNumberFormat: true,
      });
      expect(realWb.setColumnStyle(0, 0, MAX_COL, inheritedTextXf)).toBe(true);

      expect(commitCellInput({ store, wb: realWb, addr, raw: '00123', origin: 'editor' })).toEqual({
        status: 'applied',
        notice: null,
      });
      expect(realWb.getValue(addr)).toEqual({ kind: 'text', value: '00123' });

      expect(
        commitCellInput({
          store,
          wb: realWb,
          addr: { ...addr, col: 1 },
          raw: '=1+1',
          origin: 'editor',
        }),
      ).toEqual({ status: 'applied', notice: null });
      expect(realWb.getValue({ ...addr, col: 1 })).toEqual({ kind: 'text', value: '=1+1' });

      const explicitGeneral = { ...addr, col: 2 };
      mutators.setCellFormat(store, explicitGeneral, { numFmt: { kind: 'general' } });
      expect(
        commitCellInput({
          store,
          wb: realWb,
          addr: explicitGeneral,
          raw: '00123',
          origin: 'editor',
        }),
      ).toEqual({ status: 'applied', notice: null });
      expect(realWb.getValue(explicitGeneral)).toEqual({ kind: 'number', value: 123 });
    } finally {
      realWb.dispose();
    }
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
      const result = commitCellInput({ store, wb: realWb, addr, raw: '4', origin: 'editor' });

      expect(result).toEqual({
        status: 'rejected',
        operation: 'valueEdit',
        alert: { severity: 'stop', title: 'Five only', message: 'Enter five.' },
      });
      expect(realWb.getValue(addr)).toEqual({ kind: 'blank' });
      expect(realWb.canUndo()).toBe(false);
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

  it('classifies and writes merged body input at the anchor with validation', async () => {
    const realWb = await createRealWorkbook();
    const anchor = { sheet: 0, row: 0, col: 0 };
    const body = { sheet: 0, row: 0, col: 1 };
    try {
      mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
      mutators.setCellFormat(store, anchor, {
        numFmt: { kind: 'text' },
        validation: { kind: 'list', source: ['=A1'] },
      });

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
      });
    } finally {
      realWb.dispose();
    }
  });

  it('keeps the requested merged body address for controller bounds rejection', async () => {
    const realWb = await createRealWorkbook();
    const submittedAddr = { sheet: 0, row: 0, col: 2 };
    const handle = withController(realWb, {
      getBounds: () => ({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }),
    });
    try {
      mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });

      const result = commitCellInput({
        store,
        wb: realWb,
        addr: submittedAddr,
        raw: '=A1',
        origin: 'formulaBar',
      });

      expect(result).toMatchObject({ status: 'rejected', operation: 'formulaEdit' });
      expect(handle.history.canUndo()).toBe(false);
      expect((result as { alert: { message: string } | null }).alert?.message).toContain(
        'outside the configured bounds',
      );
      expect(realWb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });
    } finally {
      handle.dispose();
      realWb.dispose();
    }
  });

  it.each([false, true] as const)(
    'rejects a mixed-lock merged edit through the %s path without writing',
    async (registered) => {
      const realWb = await createRealWorkbook();
      const handle = registered ? withController(realWb) : null;
      const anchor = { sheet: 0, row: 0, col: 0 };
      const bodyAddr = { sheet: 0, row: 0, col: 1 };
      const submittedAddr = { sheet: 0, row: 0, col: 2 };
      try {
        mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
        setCellLocked(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, false);
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
          alert: { severity: 'stop', code: 'protectedCell' },
        });
        expect(realWb.getValue(anchor)).toEqual({ kind: 'blank' });
        expect(realWb.getValue(bodyAddr)).toEqual({ kind: 'blank' });
        expect(realWb.getValue(submittedAddr)).toEqual({ kind: 'blank' });
        expect(store.getState().format.formats).toEqual(beforeFormats);
        if (handle) expect(handle.history.canUndo()).toBe(false);
        else expect(realWb.canUndo()).toBe(false);
      } finally {
        handle?.dispose();
        realWb.dispose();
      }
    },
  );

  describe('table auto-expansion', () => {
    const at = (row: number, col: number) => ({ sheet: 0, row, col });
    const tableRange = () => store.getState().tables.tables[0]?.range;
    const seedTable = (r0: number, c0: number, headers: string[], rows: number): void => {
      headers.forEach((h, i) => {
        wb.setText(at(r0, c0 + i), h);
        for (let r = 1; r <= rows; r += 1) wb.setNumber(at(r0 + r, c0 + i), r);
      });
      formatAsTable(store, { sheet: 0, r0, c0, r1: r0 + rows, c1: c0 + headers.length - 1 });
    };
    const type = (row: number, col: number, raw: string) =>
      commitCellInput({ store, wb, addr: at(row, col), raw, origin: 'editor' });

    it('grows by the row typed directly below, in any table column, value or formula', () => {
      seedTable(9, 0, ['Name', 'Qty'], 3);
      type(13, 0, 'd');
      expect(tableRange()).toEqual({ sheet: 0, r0: 9, c0: 0, r1: 13, c1: 1 });
      type(14, 1, '=1+1');
      expect(tableRange()).toEqual({ sheet: 0, r0: 9, c0: 0, r1: 14, c1: 1 });
    });

    it('grows by the column typed directly right and names its header', () => {
      seedTable(9, 0, ['Name', 'Qty'], 4);
      type(10, 2, '7');
      expect(tableRange()).toEqual({ sheet: 0, r0: 9, c0: 0, r1: 13, c1: 2 });
      expect(wb.getValue(at(9, 2))).toEqual({
        kind: 'text',
        value: `${defaultObjectNames(wb).column}1`,
      });
    });

    it('continues a trailing-number header pattern for the new column', () => {
      seedTable(0, 10, ['k1', 'k2'], 2);
      type(1, 12, '9');
      expect(tableRange()).toEqual({ sheet: 0, r0: 0, c0: 10, r1: 2, c1: 12 });
      expect(wb.getValue(at(0, 12))).toEqual({ kind: 'text', value: 'k3' });
    });

    it('leaves the table for a diagonal corner, a gap, an empty entry or a filled row', () => {
      seedTable(9, 0, ['Name', 'Qty'], 5);
      const before = { sheet: 0, r0: 9, c0: 0, r1: 14, c1: 1 };
      type(15, 2, '5');
      expect(tableRange()).toEqual(before);
      type(16, 0, '1');
      expect(tableRange()).toEqual(before);
      type(9, 4, '9');
      expect(tableRange()).toEqual(before);
      type(15, 0, '');
      expect(tableRange()).toEqual(before);
      type(15, 1, '4');
      expect(tableRange()).toEqual(before);
    });

    it('grows an engine-backed table and its engine columns', async () => {
      const realWb = await createRealWorkbook();
      const realStore = createSpreadsheetStore();
      realWb.attachStore(realStore);
      try {
        realWb.setText(at(0, 0), 'k1');
        realWb.setText(at(0, 1), 'k2');
        realWb.setNumber(at(1, 0), 1);
        realWb.setNumber(at(1, 1), 2);
        formatAsTable(realStore, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 }, { workbook: realWb });
        commitCellInput({
          store: realStore,
          wb: realWb,
          addr: at(1, 2),
          raw: '3',
          origin: 'editor',
        });
        expect(realStore.getState().tables.tables[0]?.range).toEqual({
          sheet: 0,
          r0: 0,
          c0: 0,
          r1: 1,
          c1: 2,
        });
        expect(realWb.getTables()[0]?.columns).toEqual(['k1', 'k2', 'k3']);
      } finally {
        realWb.detachStore(realStore);
        realWb.dispose();
      }
    });

    it('undoes the expansion first, keeping the typed value', () => {
      const { history, dispose: release } = withController();
      dispose = release;
      seedTable(0, 0, ['Name', 'Qty'], 2);
      type(1, 2, '7');
      expect(tableRange()?.c1).toBe(2);
      history.undo();
      expect(tableRange()?.c1).toBe(1);
      expect(wb.getValue(at(0, 2)).kind).toBe('blank');
      expect(wb.getValue(at(1, 2))).toEqual({ kind: 'number', value: 7 });
      history.undo();
      expect(wb.getValue(at(1, 2)).kind).toBe('blank');
    });
  });
});
