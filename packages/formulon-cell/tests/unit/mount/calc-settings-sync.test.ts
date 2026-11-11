import { afterEach, describe, expect, it, vi } from 'vitest';
import { solveMacGoalSeek } from '../../../src/commands/mac-data-tools.js';
import { setR1C1ReferenceStyle } from '../../../src/commands/view.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';

async function savedWith(settings: { r1c1?: boolean; date1904?: boolean }): Promise<Uint8Array> {
  const wb = await WorkbookHandle.createDefault();
  try {
    if (settings.r1c1) wb.setCalcSetting('r1c1', true);
    if (settings.date1904) wb.setCalcSetting('date1904', true);
    return wb.save();
  } finally {
    wb.dispose();
  }
}

describe.skipIf(!canLoadWasm())('calc settings follow the workbook', () => {
  let sheet: MountedStubSheet | null = null;
  afterEach(() => {
    sheet?.dispose();
    sheet = null;
  });

  it('hydrates R1C1 and the 1904 date system from a loaded file', async () => {
    const workbook = await WorkbookHandle.loadBytes(
      await savedWith({ r1c1: true, date1904: true }),
    );
    sheet = await mountStubSheet({ workbook });
    const ui = sheet.instance.store.getState().ui;
    expect(ui.r1c1).toBe(true);
    expect(ui.date1904).toBe(true);
  });

  it('writes the reference style to the engine so it saves', async () => {
    const workbook = await WorkbookHandle.createDefault();
    sheet = await mountStubSheet({ workbook });
    setR1C1ReferenceStyle(sheet.instance.store, true);
    expect(workbook.calcSettings()?.r1c1).toBe(true);
    mutators.setR1C1(sheet.instance.store, false);
    expect(workbook.calcSettings()?.r1c1).toBe(false);
    setR1C1ReferenceStyle(sheet.instance.store, true);
    const restored = await WorkbookHandle.loadBytes(workbook.save());
    try {
      expect(restored.calcSettings()?.r1c1).toBe(true);
    } finally {
      restored.dispose();
    }
  });

  it('takes the incoming workbook settings on setWorkbook', async () => {
    const workbook = await WorkbookHandle.loadBytes(
      await savedWith({ r1c1: true, date1904: true }),
    );
    sheet = await mountStubSheet({ workbook });
    const next = await WorkbookHandle.createDefault();
    await sheet.instance.setWorkbook(next);
    const ui = sheet.instance.store.getState().ui;
    expect(ui.r1c1).toBe(false);
    expect(ui.date1904).toBe(false);
    expect(next.calcSettings()?.r1c1).toBe(false);
  });

  it('shows the values the save-time recalculation wrote', async () => {
    const workbook = await WorkbookHandle.createDefault();
    sheet = await mountStubSheet({ workbook });
    workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    workbook.setFormula({ sheet: 0, row: 0, col: 1 }, '=A1*10');
    workbook.setCalcMode(1);
    workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 2);
    const cells = () => sheet?.instance.store.getState().data.cells;
    mutators.replaceCells(sheet.instance.store, workbook.cells(0));
    expect(cells()?.get('0:0:1')?.value).toEqual({ kind: 'number', value: 10 });
    workbook.save();
    expect(cells()?.get('0:0:1')?.value).toEqual({ kind: 'number', value: 20 });
  });

  it('leaves a Manual workbook stale when Goal Seek copies it', async () => {
    const workbook = await WorkbookHandle.createDefault();
    sheet = await mountStubSheet({ workbook });
    workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    workbook.setFormula({ sheet: 0, row: 0, col: 1 }, '=A1*10');
    workbook.setCalcMode(1);
    workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 2);
    const result = await solveMacGoalSeek(sheet.instance, {
      formulaCell: { sheet: 0, row: 0, col: 1 },
      changingCell: { sheet: 0, row: 0, col: 0 },
      targetValue: 50,
      maxIterations: 100,
    });
    expect(result.ok).toBe(true);
    expect(workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 10 });
    expect(workbook.calcSettings()?.calcOnSave).toBe(true);
  });

  it('takes a Manual snapshot without events, undo steps or recalculation', async () => {
    const workbook = await WorkbookHandle.createDefault();
    sheet = await mountStubSheet({ workbook });
    workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    workbook.setFormula({ sheet: 0, row: 0, col: 1 }, '=A1*10');
    workbook.setCalcMode(1);
    workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 2);
    const changes = vi.fn();
    const facadeChanges = vi.fn();
    const offChanges = workbook.subscribe(changes);
    const offFacade = workbook.subscribeFacadeChanges(facadeChanges);
    const { history } = sheet.instance;
    const undoable = history.canUndo();
    workbook.save({ snapshot: true });
    offChanges();
    offFacade();
    expect(changes).not.toHaveBeenCalled();
    expect(facadeChanges).not.toHaveBeenCalled();
    expect(history.canUndo()).toBe(undoable);
    expect(workbook.calcSettings()?.calcOnSave).toBe(true);
    // The pending recalculation is still pending: a user save runs it.
    expect(workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 10 });
    workbook.save();
    expect(workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 20 });
  });

  it('restores calculate-on-save when a snapshot save fails', async () => {
    const workbook = await WorkbookHandle.createDefault();
    workbook.setCalcMode(1);
    const engine = (workbook as unknown as { wb: { saveWithDiagnostics: () => unknown } }).wb;
    const failing = vi.spyOn(engine, 'saveWithDiagnostics').mockImplementation(() => {
      throw new Error('write failed');
    });
    try {
      expect(() => workbook.save({ snapshot: true })).toThrow('write failed');
    } finally {
      failing.mockRestore();
    }
    expect(workbook.calcSettings()?.calcOnSave).toBe(true);
    workbook.dispose();
  });

  it('hydrates Set precision as displayed so typed numbers round', async () => {
    const source = await WorkbookHandle.createDefault();
    source.setCalcSetting('fullPrecision', false);
    const workbook = await WorkbookHandle.loadBytes(source.save());
    source.dispose();
    sheet = await mountStubSheet({ workbook });
    expect(sheet.instance.store.getState().ui.precisionAsDisplayed).toBe(true);
  });
});
