import { FilterKind } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { applyColorFilter, setAutoFilter } from '../../../src/commands/filter.js';
import { sortRange } from '../../../src/commands/sort.js';
import {
  hydrateAutoFilterFromEngine,
  syncAutoFilterToEngine,
} from '../../../src/engine/auto-filter-sync.js';
import { hydrateLayoutFromEngine } from '../../../src/engine/layout-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';
const range = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 1 };

function seed(wb: WorkbookHandle): void {
  const rows = [
    ['Item', 'Qty'],
    ['b', '3'],
    ['a', '1'],
    ['c', '2'],
    ['d', '4'],
  ];
  rows.forEach((cols, row) => {
    cols.forEach((text, col) => {
      if (row > 0 && col === 1) wb.setNumber({ sheet: 0, row, col }, Number(text));
      else wb.setText({ sheet: 0, row, col }, text);
    });
  });
}

function storeFor(wb: WorkbookHandle) {
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  return store;
}

function rehydrate(wb: WorkbookHandle) {
  const store = storeFor(wb);
  hydrateLayoutFromEngine(wb, store, 0);
  hydrateAutoFilterFromEngine(wb, store, 0);
  return store;
}

describe.skipIf(!canLoadWasm())('AutoFilter colour criterion and sort state', () => {
  it('writes a cell-colour criterion through a dxf and reads it back', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      seed(wb);
      const store = storeFor(wb);
      store.setState((s) => ({
        ...s,
        format: {
          ...s.format,
          formats: new Map([
            ['0:1:0', { fill: '#ff0000' }],
            ['0:3:0', { fill: '#ff0000' }],
          ]),
        },
      }));
      setAutoFilter(store, range, wb);
      applyColorFilter(
        store.getState(),
        store,
        range,
        0,
        { kind: 'cellColor', color: '#FF0000' },
        wb,
      );

      const column = wb.getAutoFilter(0)?.columns[0];
      expect(column).toMatchObject({ colId: 0, kind: FilterKind.Color, cellColor: true });
      expect(wb.getDxf(column?.dxfId ?? -1)?.fill).toBeDefined();

      // Re-syncing the same criterion reuses its differential style.
      const dxfs = wb.dxfCount();
      syncAutoFilterToEngine(wb, store.getState(), 0, store);
      syncAutoFilterToEngine(wb, store.getState(), 0, store);
      expect(wb.dxfCount()).toBe(dxfs);

      const restored = await WorkbookHandle.loadBytes(wb.save());
      try {
        const next = rehydrate(restored);
        expect(next.getState().ui.filterCriteria).toEqual([
          { range, byCol: 0, hiddenValues: [], color: { kind: 'cellColor', color: '#ff0000' } },
        ]);
        expect(restored.getAutoFilter(0)?.columns[0]?.kind).toBe(FilterKind.Color);
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('writes a font-colour criterion', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      seed(wb);
      const store = storeFor(wb);
      store.setState((s) => ({
        ...s,
        format: { ...s.format, formats: new Map([['0:2:1', { color: '#0000ff' }]]) },
      }));
      setAutoFilter(store, range, wb);
      applyColorFilter(
        store.getState(),
        store,
        range,
        1,
        { kind: 'fontColor', color: '#0000ff' },
        wb,
      );
      const column = wb.getAutoFilter(0)?.columns[0];
      expect(column).toMatchObject({ colId: 1, kind: FilterKind.Color, cellColor: false });
      const restored = await WorkbookHandle.loadBytes(wb.save());
      try {
        const next = rehydrate(restored);
        expect(next.getState().ui.filterCriteria[0]?.color).toEqual({
          kind: 'fontColor',
          color: '#0000ff',
        });
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('writes sortState for a sort inside the filter range and restores the arrow', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      seed(wb);
      const store = storeFor(wb);
      setAutoFilter(store, range, wb);
      expect(
        sortRange(store.getState(), store, wb, range, {
          byCol: 1,
          direction: 'desc',
          hasHeader: true,
        }),
      ).toBe(true);
      const sort = wb.getAutoFilter(0)?.sort;
      expect(sort?.conditions).toHaveLength(1);
      expect(sort?.conditions[0]).toMatchObject({ descending: true, sortBy: 0 });
      expect(sort?.conditions[0]?.ref).toMatchObject({
        firstRow: 1,
        lastRow: 4,
        firstCol: 1,
        lastCol: 1,
      });

      const restored = await WorkbookHandle.loadBytes(wb.save());
      try {
        const next = rehydrate(restored);
        expect(next.getState().ui.filterSorts).toEqual([
          { range, keys: [{ byCol: 1, direction: 'desc' }] },
        ]);
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });
});
