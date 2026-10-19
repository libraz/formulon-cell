import { AnchorKind, ChartInsertKind, ChartPlotType, DrawingObjectKind } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';

describe.skipIf(!canLoadWasm())('chart handle wrapper', () => {
  it('detects drawing, image and chart capabilities', async () => {
    const h = await WorkbookHandle.createDefault();
    try {
      expect(h.isStub).toBe(false);
      expect(h.capabilities.drawings).toBe(true);
      expect(h.capabilities.images).toBe(true);
      expect(h.capabilities.charts).toBe(true);
    } finally {
      h.dispose();
    }
  });

  it('round-trips a chart through anchor, snapshot, remove and restore', async () => {
    const h = await WorkbookHandle.createDefault();
    try {
      const rows: [string, number][] = [
        ['a', 1],
        ['b', 2],
        ['c', 3],
      ];
      h.setText({ sheet: 0, row: 0, col: 0 }, 'k');
      h.setText({ sheet: 0, row: 0, col: 1 }, 'v');
      rows.forEach(([label, value], i) => {
        h.setText({ sheet: 0, row: i + 1, col: 0 }, label);
        h.setNumber({ sheet: 0, row: i + 1, col: 1 }, value);
      });

      const id = h.insertChart(0, { kind: ChartInsertKind.Column, source: 'Sheet1!$A$1:$B$4' });
      expect(id).not.toBeNull();
      const objectId = id ?? 0;

      const info = h.getChart(0, objectId);
      expect(info?.objectId).toBe(objectId);
      expect(info?.plots.length).toBeGreaterThan(0);
      expect(h.getChartData(0, objectId)?.series.length).toBeGreaterThan(0);
      expect(
        h
          .listDrawingObjects(0)
          .some((o) => o.objectId === objectId && o.kind === DrawingObjectKind.Chart),
      ).toBe(true);

      expect(
        h.setDrawingAnchor(0, objectId, { anchorKind: AnchorKind.OneCell, row: 7, col: 3 }),
      ).toBe(true);
      const moved = h.listDrawingObjects(0).find((o) => o.objectId === objectId);
      expect(moved).toMatchObject({ fromRow: 7, fromCol: 3 });

      const snapshot = h.snapshotDrawing(0, objectId);
      expect(snapshot?.length).toBeGreaterThan(0);
      expect(h.removeDrawing(0, objectId)).toBe(true);
      expect(h.getChart(0, objectId)).toBeNull();

      expect(h.restoreDrawing(0, snapshot ?? new Uint8Array())).toBe(objectId);
      expect(h.getChart(0, objectId)?.objectId).toBe(objectId);
    } finally {
      h.dispose();
    }
  });

  it('changes the chart type to a kind without grouping', async () => {
    const h = await WorkbookHandle.createDefault();
    try {
      h.setText({ sheet: 0, row: 0, col: 0 }, 'a');
      h.setNumber({ sheet: 0, row: 0, col: 1 }, 1);
      h.setText({ sheet: 0, row: 1, col: 0 }, 'b');
      h.setNumber({ sheet: 0, row: 1, col: 1 }, 2);
      const objectId =
        h.insertChart(0, { kind: ChartInsertKind.Column, source: 'Sheet1!$A$1:$B$2' }) ?? 0;
      expect(h.setChartType(0, objectId, 0, ChartInsertKind.Pie)).toBe(true);
      expect(h.getChart(0, objectId)?.plots[0]?.type).toBe(ChartPlotType.Pie);
    } finally {
      h.dispose();
    }
  });

  it('appends a series by reference and removes it by c:idx', async () => {
    const h = await WorkbookHandle.createDefault();
    try {
      h.setText({ sheet: 0, row: 0, col: 1 }, 'v');
      h.setText({ sheet: 0, row: 0, col: 2 }, 'w');
      ['a', 'b'].forEach((label, i) => {
        h.setText({ sheet: 0, row: i + 1, col: 0 }, label);
        h.setNumber({ sheet: 0, row: i + 1, col: 1 }, i + 1);
        h.setNumber({ sheet: 0, row: i + 1, col: 2 }, i + 10);
      });
      const objectId =
        h.insertChart(0, { kind: ChartInsertKind.Column, source: 'Sheet1!$A$1:$B$3' }) ?? 0;
      const seriesIdx = h.addChartSeries(0, objectId, 0, {
        name: 'Sheet1!$C$1',
        values: 'Sheet1!$C$2:$C$3',
      });
      expect(seriesIdx).not.toBeNull();
      expect(h.getChartData(0, objectId)?.series.map((series) => series.name)).toEqual(['v', 'w']);
      // A series needs its values.
      expect(h.addChartSeries(0, objectId, 0, { name: 'Sheet1!$C$1' })).toBeNull();
      expect(h.removeChartSeries(0, objectId, 0, seriesIdx ?? -1)).toBe(true);
      expect(h.removeChartSeries(0, objectId, 0, 99)).toBe(false);
      expect(h.getChart(0, objectId)?.plots[0]?.series).toHaveLength(1);
    } finally {
      h.dispose();
    }
  });

  it('reports failure as null or false', async () => {
    const h = await WorkbookHandle.createDefault();
    try {
      expect(h.getChart(0, 9999)).toBeNull();
      expect(h.getChartData(0, 9999)).toBeNull();
      expect(h.setChartTitle(0, 9999, 'x')).toBe(false);
      expect(h.removeDrawing(0, 9999)).toBe(false);
      expect(h.snapshotDrawing(0, 9999)).toBeNull();
    } finally {
      h.dispose();
    }
  });
});
