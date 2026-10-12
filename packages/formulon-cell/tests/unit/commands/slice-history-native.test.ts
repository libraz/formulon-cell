import { describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  recordFormatChange,
  recordRepeatableFormatChange,
} from '../../../src/commands/slice-history.js';
import { hydrateCellFormatsFromEngine } from '../../../src/engine/cell-format-sync.js';
import { formatHistoryBridgeFor } from '../../../src/engine/format-history-bridge.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import type { CellFormat } from '../../../src/store/types.js';

const target = { sheet: 0, row: 0, col: 0 } as const;
const outside = { sheet: 0, row: 0, col: 1 } as const;

async function mountedNativeFormatCase() {
  const workbook = await WorkbookHandle.createDefault();
  workbook.setText(target, 'target');
  const normal = workbook.getCellXf(0);
  if (!normal) throw new Error('missing default XF');
  const sourceXf = workbook.addXfRecord({
    ...normal,
    applyAlignment: false,
    applyFill: false,
    quotePrefix: true,
  });
  if (!workbook.setCellXfIndex(target.sheet, target.row, target.col, sourceXf)) {
    throw new Error('failed to install source XF');
  }
  const store = createSpreadsheetStore();
  hydrateCellFormatsFromEngine(workbook, store, 0);
  workbook.attachStore(store);
  mutators.setCellFormat(store, target, { bold: true });
  return { workbook, store };
}

async function mountedNativeTwoCellCase(seedFormats = true) {
  const workbook = await WorkbookHandle.createDefault();
  workbook.setText(target, 'target');
  workbook.setText(outside, 'outside');
  const normal = workbook.getCellXf(0);
  if (!normal) throw new Error('missing default XF');
  const sourceXf = workbook.addXfRecord({
    ...normal,
    applyAlignment: false,
    applyFill: false,
    quotePrefix: true,
  });
  for (const addr of [target, outside]) {
    if (!workbook.setCellXfIndex(addr.sheet, addr.row, addr.col, sourceXf)) {
      throw new Error(`failed to install source XF at ${addr.col}`);
    }
  }
  const store = createSpreadsheetStore();
  hydrateCellFormatsFromEngine(workbook, store, 0);
  workbook.attachStore(store);
  if (seedFormats) {
    mutators.setCellFormat(store, target, { bold: true });
    mutators.setCellFormat(store, outside, { italic: true });
  }
  return { workbook, store };
}

describe('generic format history with an attached native workbook', () => {
  it.each([
    ['alignment', { align: 'center' }],
    ['fill', { fill: '#ffeeaa' }],
  ] as const)('restores the exact native XF after %s undo/redo', async (_label, patch) => {
    const { workbook, store } = await mountedNativeFormatCase();
    const history = new History();
    try {
      const beforeIndex = workbook.getCellXfIndex(target.sheet, target.row, target.col);
      const beforeXf = workbook.getCellXf(beforeIndex ?? -1);
      expect(beforeXf).toMatchObject({ quotePrefix: true });

      recordFormatChange(history, store, () => {
        mutators.setCellFormat(store, target, patch as Partial<CellFormat>);
      });

      const afterIndex = workbook.getCellXfIndex(target.sheet, target.row, target.col);
      const afterXf = workbook.getCellXf(afterIndex ?? -1);
      expect(afterIndex).not.toBe(beforeIndex);
      expect(afterXf).toMatchObject({ quotePrefix: true });
      expect(store.getState().format.formats.get('0:0:0')).toMatchObject(patch);

      mutators.setCellFormat(store, outside, { italic: true });
      const outsideAfterEdit = workbook.getCellXfIndex(outside.sheet, outside.row, outside.col);
      expect(history.undo()).toBe(true);
      expect(workbook.getCellXfIndex(target.sheet, target.row, target.col)).toBe(beforeIndex);
      expect(workbook.getCellXf(beforeIndex ?? -1)).toEqual(beforeXf);
      expect(store.getState().format.formats.get('0:0:0')).toMatchObject({ bold: true });
      expect(store.getState().format.formats.get('0:0:0')?.align).toBeUndefined();
      expect(store.getState().format.formats.get('0:0:1')).toEqual({ italic: true });
      expect(workbook.getCellXfIndex(outside.sheet, outside.row, outside.col)).toBe(
        outsideAfterEdit,
      );

      expect(history.redo()).toBe(true);
      expect(workbook.getCellXfIndex(target.sheet, target.row, target.col)).toBe(afterIndex);
      expect(workbook.getCellXf(afterIndex ?? -1)).toEqual(afterXf);
    } finally {
      workbook.dispose();
    }
  });

  it('repeats a mounted format command against the moved selection', async () => {
    const { workbook, store } = await mountedNativeFormatCase();
    const history = new History();
    try {
      recordRepeatableFormatChange(history, store, () => {
        mutators.setCellFormat(store, store.getState().selection.active, { fill: '#d9ead3' });
      });
      mutators.setActive(store, outside);

      expect(history.repeatLast()).toBe(true);
      expect(store.getState().format.formats.get('0:0:1')).toEqual({ fill: '#d9ead3' });
      expect(workbook.getCellXfIndex(outside.sheet, outside.row, outside.col)).toBeGreaterThan(0);
      expect(
        workbook.getCellXf(workbook.getCellXfIndex(outside.sheet, outside.row, outside.col) ?? -1),
      ).toMatchObject({
        fillIndex: expect.any(Number),
      });
    } finally {
      workbook.dispose();
    }
  });

  it('keeps a newer store bridge when an older mirror detaches', async () => {
    const store = createSpreadsheetStore();
    const first = await WorkbookHandle.createDefault();
    const second = await WorkbookHandle.createDefault();
    try {
      first.attachStore(store);
      const firstBridge = formatHistoryBridgeFor(store);
      expect(firstBridge).toBeDefined();
      expect(firstBridge?.workbook).toBe(first);
      second.attachStore(store);
      const secondBridge = formatHistoryBridgeFor(store);
      expect(secondBridge).toBeDefined();
      expect(secondBridge).not.toBe(firstBridge);
      expect(secondBridge?.workbook).toBe(second);

      first.detachStore(store);
      expect(formatHistoryBridgeFor(store)).toBe(secondBridge);
      second.detachStore(store);
      expect(formatHistoryBridgeFor(store)).toBeUndefined();
    } finally {
      first.dispose();
      second.dispose();
    }
  });

  it('restores every native cell after a partial forward publication failure', async () => {
    const { workbook, store } = await mountedNativeTwoCellCase(false);
    const history = new History();
    const before = new Map([
      ['0:0:0', workbook.getCellXfIndex(target.sheet, target.row, target.col)],
      ['0:0:1', workbook.getCellXfIndex(outside.sheet, outside.row, outside.col)],
    ]);
    const beforeXfs = new Map(
      [...before].map(([key, index]) => [key, workbook.getCellXf(index ?? -1)]),
    );
    const originalSetter = workbook.setCellXfIndex.bind(workbook);
    let calls = 0;
    const setter = vi
      .spyOn(workbook, 'setCellXfIndex')
      .mockImplementation((sheet, row, col, xf) => {
        calls += 1;
        if (calls === 2 || calls === 3) return false;
        return originalSetter(sheet, row, col, xf);
      });
    let setterRestored = false;
    try {
      expect(() =>
        recordFormatChange(history, store, () => {
          mutators.setCellFormat(store, target, { align: 'center' });
          mutators.setCellFormat(store, outside, { align: 'right' });
        }),
      ).toThrow(AggregateError);
      expect(calls).toBeGreaterThanOrEqual(5);
      expect(history.canUndo()).toBe(false);
      expect(store.getState().format.formats).toEqual(new Map());
      for (const [key, index] of before) {
        const addr = key === '0:0:0' ? target : outside;
        expect(workbook.getCellXfIndex(addr.sheet, addr.row, addr.col)).toBe(index);
        expect(workbook.getCellXf(index ?? -1)).toEqual(beforeXfs.get(key));
      }

      setter.mockRestore();
      setterRestored = true;
      mutators.setCellFormat(store, { sheet: 0, row: 1, col: 0 }, { fill: '#d9ead3' });
      expect(store.getState().format.formats.get('0:1:0')).toEqual({ fill: '#d9ead3' });
      for (const [key, index] of before) {
        const addr = key === '0:0:0' ? target : outside;
        expect(workbook.getCellXfIndex(addr.sheet, addr.row, addr.col)).toBe(index);
        expect(workbook.getCellXf(index ?? -1)).toEqual(beforeXfs.get(key));
      }
    } finally {
      if (!setterRestored) setter.mockRestore();
      workbook.dispose();
    }
  });

  it('keeps the history position and native state after a partial replay failure', async () => {
    const { workbook, store } = await mountedNativeTwoCellCase();
    const history = new History();
    try {
      recordFormatChange(history, store, () => {
        mutators.setCellFormat(store, target, { align: 'center' });
        mutators.setCellFormat(store, outside, { align: 'right' });
      });
      const after = new Map([
        ['0:0:0', workbook.getCellXfIndex(target.sheet, target.row, target.col)],
        ['0:0:1', workbook.getCellXfIndex(outside.sheet, outside.row, outside.col)],
      ]);
      const afterXfs = new Map(
        [...after].map(([key, index]) => [key, workbook.getCellXf(index ?? -1)]),
      );
      const originalSetter = workbook.setCellXfIndex.bind(workbook);
      let calls = 0;
      const setter = vi
        .spyOn(workbook, 'setCellXfIndex')
        .mockImplementation((sheet, row, col, xf) => {
          calls += 1;
          if (calls === 2 || calls === 3) return false;
          return originalSetter(sheet, row, col, xf);
        });
      try {
        expect(() => history.undo()).toThrow(AggregateError);
        expect(calls).toBeGreaterThanOrEqual(5);
        expect(history.canUndo()).toBe(true);
        expect(history.canRedo()).toBe(false);
        for (const [key, index] of after) {
          const addr = key === '0:0:0' ? target : outside;
          expect(workbook.getCellXfIndex(addr.sheet, addr.row, addr.col)).toBe(index);
          expect(workbook.getCellXf(index ?? -1)).toEqual(afterXfs.get(key));
        }
      } finally {
        setter.mockRestore();
      }
      expect(history.undo()).toBe(true);
      expect(store.getState().format.formats.get('0:0:0')).toMatchObject({ bold: true });
      expect(store.getState().format.formats.get('0:0:1')).toMatchObject({ italic: true });
    } finally {
      workbook.dispose();
    }
  });
});
