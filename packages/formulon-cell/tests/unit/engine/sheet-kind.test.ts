import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { moveSheet } from '../../../src/commands/sheet-mutate.js';
import { hydrateChartsFromEngine } from '../../../src/engine/charts-sync.js';
import { hydrateLayoutFromEngine } from '../../../src/engine/layout-sync.js';
import { isChartsheet, sheetHasCells, sheetZoomable } from '../../../src/engine/sheet-kind.js';
import { SheetKind } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const fixture = (): Uint8Array =>
  new Uint8Array(readFileSync(join(__dirname, '../../fixtures/charts/chart_chartsheet.xlsx')));

describe('sheet kind', () => {
  const opened: WorkbookHandle[] = [];
  afterEach(() => {
    for (const wb of opened.splice(0)) wb.dispose();
  });

  it('reports every sheet of the opt-in stub as a worksheet', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    opened.push(wb);
    expect(wb.sheetKind(0)).toBe(SheetKind.Worksheet);
    expect(wb.chartsheetChartId(0)).toBeNull();
  });

  it('zooms a worksheet and a chartsheet, and no other kind', () => {
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        sheetKinds: new Map([
          [1, SheetKind.Chartsheet],
          [2, SheetKind.Dialogsheet],
          [3, SheetKind.Macrosheet],
          [4, SheetKind.Other],
        ]),
      },
    }));
    const state = store.getState();
    expect([0, 1, 2, 3, 4].map((sheet) => sheetZoomable(state, sheet))).toEqual([
      true,
      true,
      false,
      false,
      false,
    ]);
  });

  describe('in a file with a chartsheet', () => {
    it('reads the kind of each sheet and the chartsheet chart', async () => {
      const wb = await WorkbookHandle.loadBytes(fixture());
      opened.push(wb);
      expect(wb.sheetKind(0)).toBe(SheetKind.Chartsheet);
      expect(wb.sheetKind(1)).toBe(SheetKind.Worksheet);
      expect(wb.chartsheetChartId(0)).not.toBeNull();
      expect(wb.chartsheetChartId(1)).toBeNull();
    });

    it('keeps each sheet kind in the store, only for sheets that are not worksheets', async () => {
      const wb = await WorkbookHandle.loadBytes(fixture());
      opened.push(wb);
      const store = createSpreadsheetStore();
      hydrateLayoutFromEngine(wb, store, 0);
      const state = store.getState();
      expect(isChartsheet(state, 0)).toBe(true);
      expect(sheetHasCells(state, 0)).toBe(false);
      expect(sheetHasCells(state, 1)).toBe(true);
      expect([...(state.layout.sheetKinds ?? [])]).toEqual([[0, SheetKind.Chartsheet]]);
    });

    it('moves the kind with its sheet', async () => {
      const wb = await WorkbookHandle.loadBytes(fixture());
      opened.push(wb);
      const store = createSpreadsheetStore();
      hydrateLayoutFromEngine(wb, store, 0);
      expect(moveSheet(store, wb, 0, 1)).toBe(true);
      expect([...(store.getState().layout.sheetKinds ?? [])]).toEqual([[1, SheetKind.Chartsheet]]);
    });

    it('projects the chartsheet chart through the ordinary chart path', async () => {
      const wb = await WorkbookHandle.loadBytes(fixture());
      opened.push(wb);
      const store = createSpreadsheetStore();
      hydrateChartsFromEngine(wb, store, 0);
      const charts = store.getState().charts.charts.filter((chart) => chart.sheet === 0);
      expect(charts).toHaveLength(1);
      expect(charts[0]?.nativeObjectId).toBe(wb.chartsheetChartId(0));
    });
  });
});
