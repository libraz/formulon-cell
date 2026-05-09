import { describe, expect, it } from 'vitest';

import {
  hydrateTableOverlaysFromEngine,
  tableOverlaysFromEngine,
} from '../../../src/engine/table-sync.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const reader = {
  getTables: () => [
    {
      name: 'Table1',
      displayName: 'Sales Table',
      ref: 'B2:D6',
      sheetIndex: 1,
      columns: ['Region', 'Sales', 'Margin'],
    },
    {
      name: 'Broken',
      displayName: 'Broken',
      ref: 'not-a-range',
      sheetIndex: 0,
      columns: [],
    },
  ],
};

describe('table sync', () => {
  it('maps engine workbook tables to read-only renderer overlays', () => {
    const overlays = tableOverlaysFromEngine(reader);

    expect(overlays).toEqual([
      {
        id: 'engine-table-Sales-Table-1-B2-D6',
        source: 'engine',
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

  it('keeps loaded workbook table overlays when clearing session overlays', () => {
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

    expect(store.getState().tables.tables).toEqual([
      {
        id: 'engine-table-Sales-Table-1-B2-D6',
        source: 'engine',
        range: { sheet: 1, r0: 1, c0: 1, r1: 5, c1: 3 },
        style: 'medium',
        showHeader: true,
        showTotal: false,
        banded: true,
      },
    ]);
  });
});
