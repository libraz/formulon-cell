import { describe, expect, it } from 'vitest';

import {
  hydrateTableOverlaysFromEngine,
  tableOverlaysFromEngine,
} from '../../../src/engine/table-sync.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const reader = {
  getTables: () => [
    {
      index: 0,
      name: 'Table1',
      displayName: 'Sales Table',
      ref: 'B2:D6',
      sheetIndex: 1,
      columns: ['Region', 'Sales', 'Margin'],
    },
    {
      index: 1,
      name: 'Broken',
      displayName: 'Broken',
      ref: 'not-a-range',
      sheetIndex: 0,
      columns: [],
    },
  ],
};

describe('table sync', () => {
  it('maps engine workbook tables to engine-backed overlays', () => {
    const overlays = tableOverlaysFromEngine(reader);

    expect(overlays).toEqual([
      {
        id: 'engine-table-Sales-Table-1-B2-D6',
        source: 'engine',
        name: 'Table1',
        range: { sheet: 1, r0: 1, c0: 1, r1: 5, c1: 3 },
        style: 'medium',
        showHeader: true,
        showTotal: false,
        banded: true,
      },
    ]);
  });

  it('replaces only engine overlays and preserves session Format-as-Table overlays', () => {
    const store = createSpreadsheetStore();
    mutators.upsertTableOverlay(store, {
      id: 'session',
      source: 'session',
      range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
      style: 'dark',
      showHeader: true,
      showTotal: false,
      banded: true,
    });
    hydrateTableOverlaysFromEngine(reader, store);
    hydrateTableOverlaysFromEngine({ getTables: () => [] }, store);

    expect(store.getState().tables.tables).toEqual([
      {
        id: 'session',
        source: 'session',
        range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
        style: 'dark',
        showHeader: true,
        showTotal: false,
        banded: true,
      },
    ]);
  });

  it('clears loaded workbook table overlays like session overlays', () => {
    const store = createSpreadsheetStore();
    hydrateTableOverlaysFromEngine(reader, store);
    mutators.upsertTableOverlay(store, {
      id: 'session',
      source: 'session',
      range: { sheet: 1, r0: 1, c0: 1, r1: 5, c1: 3 },
      style: 'dark',
      showHeader: true,
      showTotal: false,
      banded: true,
    });

    mutators.clearTableOverlaysInRange(store, { sheet: 1, r0: 1, c0: 1, r1: 5, c1: 3 });

    expect(store.getState().tables.tables).toEqual([]);
  });

  it('keeps the look of a known table across a rehydrate', () => {
    const store = createSpreadsheetStore();
    hydrateTableOverlaysFromEngine(reader, store);
    const [loaded] = store.getState().tables.tables;
    if (!loaded) throw new Error('expected a loaded table');
    mutators.upsertTableOverlay(store, {
      ...loaded,
      style: 'dark',
      color: '#70ad47',
      banded: false,
    });
    const moved = {
      getTables: () => [{ ...reader.getTables()[0], ref: 'B2:D8' } as never],
    };

    hydrateTableOverlaysFromEngine(moved, store);

    expect(store.getState().tables.tables).toEqual([
      {
        ...loaded,
        style: 'dark',
        color: '#70ad47',
        banded: false,
        range: { sheet: 1, r0: 1, c0: 1, r1: 7, c1: 3 },
      },
    ]);
  });
});
