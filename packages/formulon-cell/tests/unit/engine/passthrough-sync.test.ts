import { describe, expect, it } from 'vitest';
import {
  classifyWorkbookObjectPath,
  findPivotTableAtCell,
  listWorkbookObjects,
  summarizePassthroughs,
  summarizePivotTables,
  WORKBOOK_OBJECT_KINDS,
  workbookObjectExtension,
  workbookObjectKindCounts,
  workbookObjectKindLabel,
  workbookObjectName,
  workbookObjectsByKind,
} from '../../../src/engine/passthrough-sync.js';
import { PivotAxis, PivotFilterType, PivotFilterValueKind } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const wb = (paths: readonly string[]) =>
  ({
    getPassthroughs: () => paths.map((path) => ({ path })),
  }) as unknown as WorkbookHandle;

const wbWithPivots = () =>
  ({
    getPivotTables: () => [
      {
        sheetIndex: 0,
        pivotIndex: 0,
        top: 4,
        left: 1,
        rows: 6,
        cols: 3,
        cells: 18,
        fields: ['Region', 'Sales'],
        pivotFilters: [
          {
            axis: PivotAxis.Page,
            fieldName: 'Sales',
            type: PivotFilterType.ValueGreaterThan,
            valueKind: PivotFilterValueKind.Double,
            valueDouble: 10,
          },
        ],
      },
      {
        sheetIndex: 1,
        pivotIndex: 0,
        top: 0,
        left: 0,
        rows: 2,
        cols: 2,
        cells: 4,
        fields: ['Channel'],
      },
    ],
  }) as unknown as WorkbookHandle;

describe('passthrough workbook objects', () => {
  it('classifies preserved workbook object paths', () => {
    expect(classifyWorkbookObjectPath('xl/charts/chart1.xml')).toBe('charts');
    expect(classifyWorkbookObjectPath('xl/drawings/drawing1.xml')).toBe('drawings');
    expect(classifyWorkbookObjectPath('xl/media/image1.png')).toBe('media');
    expect(classifyWorkbookObjectPath('xl/comments1.xml')).toBe('comments');
    expect(classifyWorkbookObjectPath('xl/threadedComments/threadedComment1.xml')).toBe(
      'threadedComments',
    );
    expect(classifyWorkbookObjectPath('xl/slicerCaches/slicerCache1.xml')).toBe('slicers');
    expect(classifyWorkbookObjectPath('xl/timelineCaches/timelineCache1.xml')).toBe('timelines');
    expect(classifyWorkbookObjectPath('xl/connections.xml')).toBe('connections');
    expect(classifyWorkbookObjectPath('xl/externalLinks/externalLink1.xml')).toBe('externalLinks');
    expect(classifyWorkbookObjectPath('xl/vbaProject.bin')).toBe('vbaProject');
    expect(classifyWorkbookObjectPath('xl/activeX/activeX1.xml')).toBe('controls');
    expect(classifyWorkbookObjectPath('xl/printerSettings/printerSettings1.bin')).toBe(
      'printerSettings',
    );
    expect(classifyWorkbookObjectPath('customXml/item1.xml')).toBe('customXml');
    expect(classifyWorkbookObjectPath('docProps/app.xml')).toBe('other');
  });

  it('extracts display metadata from object paths', () => {
    expect(workbookObjectName('xl/media/image1.png')).toBe('image1.png');
    expect(workbookObjectName('workbook.xml')).toBe('workbook.xml');
    expect(workbookObjectExtension('xl/media/image1.PNG')).toBe('png');
    expect(workbookObjectExtension('xl/drawings/drawing1')).toBe('');
  });

  it('lists sorted object records and groups them by kind', () => {
    const objects = listWorkbookObjects(
      wb(['xl/media/image2.png', 'xl/charts/chart1.xml', 'docProps/app.xml']),
    );

    expect(objects).toEqual([
      {
        kind: 'charts',
        path: 'xl/charts/chart1.xml',
        name: 'chart1.xml',
        extension: 'xml',
      },
      {
        kind: 'media',
        path: 'xl/media/image2.png',
        name: 'image2.png',
        extension: 'png',
      },
      {
        kind: 'other',
        path: 'docProps/app.xml',
        name: 'app.xml',
        extension: 'xml',
      },
    ]);

    const byKind = workbookObjectsByKind(objects);
    expect(byKind.charts.map((o) => o.name)).toEqual(['chart1.xml']);
    expect(byKind.media.map((o) => o.name)).toEqual(['image2.png']);
    expect(byKind.other.map((o) => o.name)).toEqual(['app.xml']);
    expect(byKind.drawings).toEqual([]);
    expect(WORKBOOK_OBJECT_KINDS[0]).toBe('charts');
    expect(workbookObjectKindLabel('vbaProject')).toBe('Macro project');
    expect(workbookObjectKindCounts(objects)).toMatchObject({
      charts: 1,
      media: 1,
      other: 1,
      drawings: 0,
    });
  });

  it('builds summaries from classified object records', () => {
    expect(
      summarizePassthroughs(
        wb(['xl/media/image2.png', 'xl/charts/chart1.xml', 'docProps/app.xml']),
      ),
    ).toEqual({
      count: 3,
      byCategory: { charts: 1, media: 1, other: 1 },
      paths: ['xl/charts/chart1.xml', 'xl/media/image2.png', 'docProps/app.xml'],
    });
  });

  it('finds the PivotTable containing an active cell', () => {
    const workbook = wbWithPivots();

    expect(findPivotTableAtCell(workbook, { sheet: 0, row: 4, col: 1 })?.pivotIndex).toBe(0);
    expect(findPivotTableAtCell(workbook, { sheet: 0, row: 9, col: 3 })?.pivotIndex).toBe(0);
    expect(findPivotTableAtCell(workbook, { sheet: 0, row: 10, col: 1 })).toBeNull();
    expect(findPivotTableAtCell(workbook, { sheet: 1, row: 1, col: 1 })?.sheetIndex).toBe(1);
  });

  it('preserves optional PivotTable filter specs in the summary surface', () => {
    expect(summarizePivotTables(wbWithPivots()).items[0]?.pivotFilters).toEqual([
      {
        axis: PivotAxis.Page,
        fieldName: 'Sales',
        type: PivotFilterType.ValueGreaterThan,
        valueKind: PivotFilterValueKind.Double,
        valueDouble: 10,
      },
    ]);
  });
});
