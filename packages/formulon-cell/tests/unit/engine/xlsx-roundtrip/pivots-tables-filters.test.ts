import { describe, expect, it } from 'vitest';
import { applyValueFilter, clearFilter } from '../../../../src/commands/filter.js';
import { formatAsTable } from '../../../../src/commands/format-as-table.js';
import { createPivotTableFromRange } from '../../../../src/commands/pivot-table.js';
import {
  hydrateAutoFilterFromEngine,
  syncAutoFilterToEngine,
} from '../../../../src/engine/auto-filter-sync.js';
import { tableOverlaysFromEngine } from '../../../../src/engine/table-sync.js';
import { PivotAggregation, PivotReportLayout, type Range } from '../../../../src/engine/types.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';

import { canLoadWasm, seedStoreText, seedTableCells } from './fixtures.js';

describe.skipIf(!canLoadWasm())('real xlsx round-trip', () => {
  it('saves and reloads PivotTable source metadata and report layout when supported', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      expect(first.isStub).toBe(false);
      if (
        !first.capabilities.pivotTableMutate ||
        !first.capabilities.pivotCacheSource ||
        !first.capabilities.pivotReportLayout
      ) {
        return;
      }

      first.setText({ sheet: 0, row: 0, col: 0 }, 'Region');
      first.setText({ sheet: 0, row: 0, col: 1 }, 'Sales');
      first.setText({ sheet: 0, row: 1, col: 0 }, 'East');
      first.setNumber({ sheet: 0, row: 1, col: 1 }, 12);
      first.setText({ sheet: 0, row: 2, col: 0 }, 'West');
      first.setNumber({ sheet: 0, row: 2, col: 1 }, 8);

      const created = createPivotTableFromRange(first, {
        source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
        destination: { sheet: 0, row: 5, col: 0 },
        name: 'SalesPivot',
        rowField: 'Region',
        valueField: 'Sales',
        aggregation: PivotAggregation.Sum,
      });
      expect(created).toMatchObject({ ok: true });
      if (!created.ok) return;
      expect(first.setPivotReportLayout(0, created.pivotIndex, PivotReportLayout.Tabular)).toBe(
        true,
      );

      const bytes = first.save();
      expect(bytes.length).toBeGreaterThan(0);

      const reloaded = await WorkbookHandle.loadBytes(bytes);
      try {
        expect(reloaded.isStub).toBe(false);
        const pivots = reloaded.getPivotTables();
        expect(pivots).toHaveLength(1);
        expect(pivots[0]).toMatchObject({
          sheetIndex: 0,
          pivotIndex: 0,
          top: 5,
          left: 0,
        });
        const cacheId = reloaded.pivotCacheIds()[0] ?? -1;
        expect(cacheId).toBeGreaterThanOrEqual(0);
        expect(reloaded.getPivotCacheWorksheetSource(cacheId)).toMatchObject({
          present: true,
          ref: 'A1:B3',
          sheet: 'Sheet1',
        });
        expect(reloaded.getPivotReportLayout(0, 0)).toBe(PivotReportLayout.Tabular);
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('writes Format as Table as a real ListObject and reads it back as an overlay', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      if (!first.capabilities.tableMutate) return;
      const range: Range = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 };
      seedTableCells(first);

      const store = createSpreadsheetStore();
      const overlay = formatAsTable(store, range, { workbook: first, style: 'medium' });
      expect(overlay).not.toBeNull();

      const reloaded = await WorkbookHandle.loadBytes(first.save());
      try {
        const tables = reloaded.getTables();
        expect(tables).toHaveLength(1);
        expect(tables[0]).toMatchObject({ sheetIndex: 0, ref: 'A1:B3' });

        const overlays = tableOverlaysFromEngine(reloaded);
        expect(overlays).toHaveLength(1);
        expect(overlays[0]?.source).toBe('engine');
        expect(overlays[0]?.range).toEqual(range);
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('keeps a table across a load and save that never touches it', async () => {
    const authored = await WorkbookHandle.createDefault();
    let bytes: Uint8Array;
    try {
      if (!authored.capabilities.tableMutate) return;
      seedTableCells(authored);
      formatAsTable(
        createSpreadsheetStore(),
        { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
        {
          workbook: authored,
        },
      );
      bytes = authored.save();
    } finally {
      authored.dispose();
    }

    const passthrough = await WorkbookHandle.loadBytes(bytes);
    let resaved: Uint8Array;
    try {
      expect(passthrough.getTables()).toHaveLength(1);
      resaved = passthrough.save();
    } finally {
      passthrough.dispose();
    }

    const reloaded = await WorkbookHandle.loadBytes(resaved);
    try {
      expect(reloaded.getTables()).toHaveLength(1);
      expect(reloaded.getTables()[0]).toMatchObject({ sheetIndex: 0, ref: 'A1:B3' });
    } finally {
      reloaded.dispose();
    }
  });

  it('round-trips an AutoFilter definition and rehydrates its range', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      if (!first.capabilities.autoFilter) return;
      const range: Range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 };
      const store = createSpreadsheetStore();
      seedStoreText(store, 0, 0, 'Region');
      seedStoreText(store, 1, 0, 'East');
      seedStoreText(store, 2, 0, 'West');
      seedStoreText(store, 3, 0, 'North');
      applyValueFilter(store.getState(), store, range, 0, ['West']);

      expect(syncAutoFilterToEngine(first, store.getState(), 0)).toBe(true);

      const reloaded = await WorkbookHandle.loadBytes(first.save());
      try {
        const xml = reloaded.getSheetAutoFilterXml(0) ?? '';
        expect(xml).toContain('ref="A1:B4"');
        expect(xml).toContain('colId="0"');
        expect(xml).toContain('<filter val="East"/>');
        expect(xml).not.toContain('<filter val="West"/>');

        const restored = createSpreadsheetStore();
        hydrateAutoFilterFromEngine(reloaded, restored, 0);
        expect(restored.getState().ui.filterRange).toEqual(range);
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('clears a saved AutoFilter when the user removes the filter', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      if (!first.capabilities.autoFilter) return;
      const range: Range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 };
      const store = createSpreadsheetStore();
      seedStoreText(store, 1, 0, 'East');
      applyValueFilter(store.getState(), store, range, 0, []);
      syncAutoFilterToEngine(first, store.getState(), 0);

      clearFilter(store.getState(), store);
      expect(syncAutoFilterToEngine(first, store.getState(), 0)).toBe(true);

      const reloaded = await WorkbookHandle.loadBytes(first.save());
      try {
        expect(reloaded.getSheetAutoFilterXml(0) ?? '').toBe('');
        const restored = createSpreadsheetStore();
        hydrateAutoFilterFromEngine(reloaded, restored, 0);
        expect(restored.getState().ui.filterRange).toBeNull();
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });
});
