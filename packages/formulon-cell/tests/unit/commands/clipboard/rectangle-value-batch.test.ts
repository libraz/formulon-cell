import { afterEach, describe, expect, it } from 'vitest';
import { pasteSpecial } from '../../../../src/commands/clipboard/paste-special.js';
import { captureSnapshot } from '../../../../src/commands/clipboard/snapshot.js';
import { addrKey } from '../../../../src/engine/address.js';
import type { CellValue, Range } from '../../../../src/engine/types.js';
import { type CellPatch, WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { mutators } from '../../../../src/store/store.js';
import type { MountedStubSheet } from '../../../test-utils/mount.js';
import { mountStubSheet } from '../../../test-utils/mount.js';

const sourceRange: Range = { sheet: 0, r0: 0, c0: 0, r1: 39, c1: 24 };
const destinationRange: Range = { sheet: 0, r0: 100, c0: 30, r1: 139, c1: 54 };
const value = (n: number): CellValue => ({ kind: 'number', value: n });

interface MountedRectangleCopy {
  readonly workbook: WorkbookHandle;
  readonly sheet: MountedStubSheet;
  readonly snapshot: NonNullable<ReturnType<typeof captureSnapshot>>;
  dispose(): void;
}

const mountRectangleCopy = async (): Promise<MountedRectangleCopy> => {
  const workbook = await WorkbookHandle.createDefault();
  if (workbook.isStub) throw new Error('native workbook required');

  const patches: CellPatch[] = [];
  for (let row = 0; row < 40; row += 1) {
    for (let col = 0; col < 25; col += 1) {
      const sourceAddr = { sheet: 0, row, col };
      patches.push({
        addr: sourceAddr,
        value: row === 1 && col === 1 ? value(3) : value(row * 25 + col + 1),
        formula: row === 1 && col === 1 ? '=A1+A2' : null,
      });
      patches.push({
        addr: { sheet: 0, row: row + 100, col: col + 30 },
        value: value(10_000 + row * 25 + col),
        formula: null,
      });
    }
  }
  workbook.applyCellPatchAtomic(patches);
  expect(workbook.setCommentEntry(0, 0, 0, 'Alice', 'copied note')).toBe(true);
  expect(workbook.setCommentEntry(0, 100, 30, 'Bob', 'old note')).toBe(true);

  const sheet = await mountStubSheet({ workbook });
  const { instance } = sheet;
  mutators.setCellFormat(
    instance.store,
    { sheet: 0, row: 0, col: 0 },
    {
      bold: true,
      comment: 'copied note',
      commentAuthor: 'Alice',
    },
  );
  mutators.setCellFormat(
    instance.store,
    { sheet: 0, row: 100, col: 30 },
    {
      italic: true,
      comment: 'old note',
      commentAuthor: 'Bob',
    },
  );
  mutators.setActive(instance.store, { sheet: 0, row: 0, col: 0 });
  mutators.setRange(instance.store, sourceRange);
  const snapshot = captureSnapshot(instance.store.getState(), sourceRange);
  if (!snapshot) throw new Error('rectangle snapshot failed');
  mutators.setActive(instance.store, { sheet: 0, row: 100, col: 30 });
  mutators.setRange(instance.store, destinationRange);

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

const pasteAll = (mounted: MountedRectangleCopy) => {
  const { instance } = mounted.sheet;
  const transaction = instance.history.begin();
  try {
    return pasteSpecial(
      instance.store.getState(),
      instance.store,
      mounted.workbook,
      mounted.snapshot,
      { what: 'all', operation: 'none', skipBlanks: false, transpose: false },
      instance.history,
    );
  } finally {
    instance.history.end(transaction);
  }
};

describe('rectangular Paste All value batching', () => {
  const mountedSheets: MountedRectangleCopy[] = [];
  afterEach(() => {
    for (const mounted of mountedSheets.splice(0)) mounted.dispose();
  });

  it('publishes a 1000-cell copied rectangle once and restores values, formula, format, and comment', async () => {
    const mounted = await mountRectangleCopy();
    mountedSheets.push(mounted);
    const { instance } = mounted.sheet;
    const store = instance.store;
    let cellMapPublications = 0;
    const unsubscribe = store.subscribe((next, previous) => {
      if (next.data.cells !== previous.data.cells) cellMapPublications += 1;
    });
    const callbackCellMaps: ReadonlyMap<string, unknown>[] = [];
    instance.on('cellChange', () => callbackCellMaps.push(store.getState().data.cells));

    const result = pasteAll(mounted);
    unsubscribe();

    expect(result?.writtenRange).toEqual(destinationRange);
    expect(cellMapPublications).toBe(1);
    expect(callbackCellMaps).toHaveLength(1_000);
    const publishedCells = store.getState().data.cells;
    expect(callbackCellMaps.every((cells) => cells === publishedCells)).toBe(true);
    expect(mounted.workbook.getValue({ sheet: 0, row: 100, col: 30 })).toEqual(value(1));
    expect(mounted.workbook.getValue({ sheet: 0, row: 101, col: 31 })).toEqual(value(27));
    expect(mounted.workbook.cellFormula({ sheet: 0, row: 101, col: 31 })).toBe('=AE101+AE102');
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 100, col: 30 }))).toEqual({
      bold: true,
      comment: 'copied note',
      commentAuthor: 'Alice',
    });
    expect(mounted.workbook.getComment(0, 100, 30)).toEqual({
      author: 'Alice',
      text: 'copied note',
    });

    expect(instance.history.undo()).toBe(true);
    expect(mounted.workbook.getValue({ sheet: 0, row: 100, col: 30 })).toEqual(value(10_000));
    expect(mounted.workbook.getValue({ sheet: 0, row: 101, col: 31 })).toEqual(value(10_026));
    expect(mounted.workbook.cellFormula({ sheet: 0, row: 101, col: 31 })).toBeNull();
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 100, col: 30 }))).toEqual({
      italic: true,
      comment: 'old note',
      commentAuthor: 'Bob',
    });
    expect(mounted.workbook.getComment(0, 100, 30)).toEqual({ author: 'Bob', text: 'old note' });

    expect(instance.history.redo()).toBe(true);
    expect(mounted.workbook.getValue({ sheet: 0, row: 100, col: 30 })).toEqual(value(1));
    expect(mounted.workbook.getValue({ sheet: 0, row: 101, col: 31 })).toEqual(value(27));
    expect(mounted.workbook.cellFormula({ sheet: 0, row: 101, col: 31 })).toBe('=AE101+AE102');
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 100, col: 30 }))).toEqual({
      bold: true,
      comment: 'copied note',
      commentAuthor: 'Alice',
    });
    expect(mounted.workbook.getComment(0, 100, 30)).toEqual({
      author: 'Alice',
      text: 'copied note',
    });
  });
});
