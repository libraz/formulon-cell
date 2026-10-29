import { afterEach, describe, expect, it, vi } from 'vitest';
import { copy } from '../../../../src/commands/clipboard/copy.js';
import { pasteSpecial } from '../../../../src/commands/clipboard/paste-special.js';
import { captureSnapshotFromCopyResult } from '../../../../src/commands/clipboard/snapshot.js';
import type { CellValue, Range } from '../../../../src/engine/types.js';
import { type CellPatch, WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';

const MAX_ROW = 1_048_575;
const sourceColumn: Range = { sheet: 0, r0: 0, c0: 0, r1: MAX_ROW, c1: 0 };
const destinationColumn: Range = { sheet: 0, r0: 0, c0: 1, r1: MAX_ROW, c1: 1 };
const value = (n: number): CellValue => ({ kind: 'number', value: n });

interface MountedCopyBand {
  readonly workbook: WorkbookHandle;
  readonly sheet: MountedStubSheet;
  readonly snapshot: NonNullable<ReturnType<typeof captureSnapshotFromCopyResult>>;
  dispose(): void;
}

const mountCopyBand = async (
  sourcePatches: readonly CellPatch[],
  destinationPatches: readonly CellPatch[],
  calcMode?: 0 | 1 | 2,
): Promise<MountedCopyBand> => {
  const workbook = await WorkbookHandle.createDefault();
  if (workbook.isStub) throw new Error('native workbook required');
  workbook.applyCellPatchAtomic([...sourcePatches, ...destinationPatches]);
  if (calcMode !== undefined) expect(workbook.setCalcMode(calcMode)).toBe(true);

  const sheet = await mountStubSheet({ workbook });
  mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
  mutators.setRange(sheet.instance.store, sourceColumn);
  const copied = copy(sheet.instance.store.getState());
  if (!copied) throw new Error('whole-column copy failed');
  const snapshot = captureSnapshotFromCopyResult(sheet.instance.store.getState(), copied);
  if (!snapshot) throw new Error('whole-column snapshot failed');
  mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 1 });
  mutators.setRange(sheet.instance.store, destinationColumn);

  return {
    workbook,
    sheet,
    snapshot,
    dispose: () => {
      sheet.dispose();
      workbook.dispose();
    },
  };
};

const pasteAll = (mounted: MountedCopyBand, skipBlanks = false) => {
  const { instance } = mounted.sheet;
  const transaction = instance.history.begin();
  try {
    return pasteSpecial(
      instance.store.getState(),
      instance.store,
      mounted.workbook,
      mounted.snapshot,
      { what: 'all', operation: 'none', skipBlanks, transpose: false },
      instance.history,
    );
  } finally {
    instance.history.end(transaction);
  }
};

describe('whole-band Paste All value batching', () => {
  const mountedSheets: MountedCopyBand[] = [];
  afterEach(() => {
    for (const mounted of mountedSheets.splice(0)) mounted.dispose();
  });

  it('publishes a 1000-cell copied column once and keeps formula, stale clear, and undo atomic', async () => {
    const sourcePatches = Array.from(
      { length: 1_000 },
      (_, row): CellPatch =>
        row === 10
          ? { addr: { sheet: 0, row, col: 0 }, value: value(2), formula: '=A1+1' }
          : { addr: { sheet: 0, row, col: 0 }, value: value(row + 1), formula: null },
    );
    const destinationPatches = Array.from(
      { length: 1_000 },
      (_, row): CellPatch => ({
        addr: { sheet: 0, row, col: 1 },
        value: value(10_000 + row),
        formula: null,
      }),
    );
    destinationPatches.push({
      addr: { sheet: 0, row: 1_200, col: 1 },
      value: value(99),
      formula: '=B1*99',
    });
    const mounted = await mountCopyBand(sourcePatches, destinationPatches);
    mountedSheets.push(mounted);

    const { instance } = mounted.sheet;
    const store = instance.store;
    let cellMapPublications = 0;
    const unsubscribe = store.subscribe((next, previous) => {
      if (next.data.cells !== previous.data.cells) cellMapPublications += 1;
    });
    const callbackCellMaps: ReadonlyMap<string, unknown>[] = [];
    instance.on('cellChange', () => callbackCellMaps.push(store.getState().data.cells));
    const historyPush = vi.spyOn(instance.history, 'push');

    const result = pasteAll(mounted);
    unsubscribe();

    expect(result?.writtenRange).toEqual(destinationColumn);
    expect(cellMapPublications).toBe(1);
    expect(callbackCellMaps).toHaveLength(1_001);
    const publishedCells = store.getState().data.cells;
    expect(callbackCellMaps.every((cells) => cells === publishedCells)).toBe(true);
    expect(historyPush).toHaveBeenCalledTimes(1);
    expect(mounted.workbook.getValue({ sheet: 0, row: 10, col: 1 })).toEqual(value(2));
    expect(mounted.workbook.cellFormula({ sheet: 0, row: 10, col: 1 })).toBe('=B1+1');
    expect(mounted.workbook.getValue({ sheet: 0, row: 1_200, col: 1 })).toEqual({ kind: 'blank' });
    expect(mounted.workbook.cellFormula({ sheet: 0, row: 1_200, col: 1 })).toBeNull();

    expect(instance.history.undo()).toBe(true);
    expect(mounted.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual(value(10_000));
    expect(mounted.workbook.cellFormula({ sheet: 0, row: 1_200, col: 1 })).toBe('=B1*99');
    expect(instance.history.redo()).toBe(true);
    expect(mounted.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual(value(1));
    expect(mounted.workbook.cellFormula({ sheet: 0, row: 10, col: 1 })).toBe('=B1+1');
    expect(mounted.workbook.getValue({ sheet: 0, row: 1_200, col: 1 })).toEqual({ kind: 'blank' });
  });

  it('leaves blank source cells and the sparse tail untouched with Skip Blanks', async () => {
    const mounted = await mountCopyBand(
      [
        { addr: { sheet: 0, row: 0, col: 0 }, value: value(1), formula: null },
        { addr: { sheet: 0, row: 2, col: 0 }, value: value(3), formula: null },
      ],
      [
        { addr: { sheet: 0, row: 0, col: 1 }, value: value(10), formula: null },
        { addr: { sheet: 0, row: 1, col: 1 }, value: value(11), formula: null },
        { addr: { sheet: 0, row: 2, col: 1 }, value: value(12), formula: null },
        { addr: { sheet: 0, row: 50, col: 1 }, value: value(50), formula: null },
      ],
    );
    mountedSheets.push(mounted);

    expect(pasteAll(mounted, true)?.writtenRange).toEqual(destinationColumn);
    expect(mounted.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual(value(1));
    expect(mounted.workbook.getValue({ sheet: 0, row: 1, col: 1 })).toEqual(value(11));
    expect(mounted.workbook.getValue({ sheet: 0, row: 2, col: 1 })).toEqual(value(3));
    expect(mounted.workbook.getValue({ sheet: 0, row: 50, col: 1 })).toEqual(value(50));
  });

  it.each([
    { mode: 0 as const, recalcs: 1 },
    { mode: 1 as const, recalcs: 0 },
    { mode: 2 as const, recalcs: 1 },
  ])('preserves recalculation behavior for calc mode $mode', async ({ mode, recalcs }) => {
    const mounted = await mountCopyBand(
      [
        { addr: { sheet: 0, row: 0, col: 0 }, value: value(1), formula: null },
        { addr: { sheet: 0, row: 1, col: 0 }, value: value(2), formula: '=A1+1' },
      ],
      [],
      mode,
    );
    mountedSheets.push(mounted);
    let recalcEvents = 0;
    mounted.sheet.instance.on('recalc', () => {
      recalcEvents += 1;
    });

    expect(pasteAll(mounted)).not.toBeNull();
    expect(recalcEvents).toBe(recalcs);
    expect(mounted.workbook.cellFormula({ sheet: 0, row: 1, col: 1 })).toBe('=B1+1');
  });

  it('rolls back native writes and records no history when a setter fails mid-batch', async () => {
    const mounted = await mountCopyBand(
      [
        { addr: { sheet: 0, row: 0, col: 0 }, value: value(1), formula: null },
        { addr: { sheet: 0, row: 1, col: 0 }, value: value(2), formula: null },
        { addr: { sheet: 0, row: 2, col: 0 }, value: value(3), formula: null },
      ],
      [
        { addr: { sheet: 0, row: 0, col: 1 }, value: value(10), formula: null },
        { addr: { sheet: 0, row: 1, col: 1 }, value: value(11), formula: null },
        { addr: { sheet: 0, row: 2, col: 1 }, value: value(12), formula: null },
      ],
    );
    mountedSheets.push(mounted);
    const raw = (
      mounted.workbook as unknown as {
        wb: {
          setNumber: (
            sheet: number,
            row: number,
            col: number,
            input: number,
          ) => { ok: boolean; status: number; message: string; context: string };
        };
      }
    ).wb;
    const originalSetNumber = raw.setNumber.bind(raw);
    let setterCalls = 0;
    const failingSetter = vi
      .spyOn(raw, 'setNumber')
      .mockImplementation((sheet, row, col, input) => {
        setterCalls += 1;
        if (setterCalls === 2)
          return { ok: false, status: 1, message: 'injected mid-batch failure', context: '' };
        return originalSetNumber(sheet, row, col, input);
      });
    const historyPush = vi.spyOn(mounted.sheet.instance.history, 'push');
    const cellsBefore = mounted.sheet.instance.store.getState().data.cells;

    try {
      expect(() => pasteAll(mounted)).toThrow('injected mid-batch failure');
      expect(mounted.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual(value(10));
      expect(mounted.workbook.getValue({ sheet: 0, row: 1, col: 1 })).toEqual(value(11));
      expect(mounted.workbook.getValue({ sheet: 0, row: 2, col: 1 })).toEqual(value(12));
      expect(mounted.sheet.instance.store.getState().data.cells).toBe(cellsBefore);
      expect(historyPush).not.toHaveBeenCalled();
      expect(mounted.sheet.instance.history.canUndo()).toBe(false);
    } finally {
      failingSetter.mockRestore();
    }
  });
});
