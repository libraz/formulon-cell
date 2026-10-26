import { describe, expect, it } from 'vitest';

import {
  isSpreadsheetFeatureAvailable,
  isSpreadsheetFeatureWritable,
  spreadsheetCompatibilityItem,
  spreadsheetCompatibilityStatus,
  summarizeSpreadsheetCompatibility,
} from '../../../src/engine/compatibility.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const baseCaps = {
  cellFormatting: false,
  conditionalFormat: false,
  conditionalFormatMutate: false,
  dataValidation: false,
  freeze: false,
  sheetZoom: false,
  hiddenRowsCols: false,
  outlines: false,
  pivotTables: false,
  pivotTableMutate: false,
  externalLinks: false,
  hyperlinks: false,
  comments: false,
  commentsEnumerable: false,
  definedNameMutate: false,
  definedNameScopes: false,
  formulaTextEvaluation: false,
  conditionalFormulaEvaluation: false,
  sheetProtectionRoundtrip: false,
};

describe('summarizeSpreadsheetCompatibility', () => {
  it('distinguishes writable, read-only, session-only, and unsupported surfaces', () => {
    const wb = {
      capabilities: {
        ...baseCaps,
        phonetic: true,
        autoFilter: true,
        cellFormatting: true,
        conditionalFormat: true,
        dataValidation: true,
        pivotTables: true,
        pivotTableMutate: true,
        externalLinks: true,
        hyperlinks: true,
        comments: true,
        commentsEnumerable: true,
        definedNameMutate: true,
        definedNameScopes: true,
        sheetProtectionRoundtrip: true,
      },
      getPassthroughs: () => [
        { path: 'xl/charts/chart1.xml' },
        { path: 'xl/drawings/drawing1.xml' },
        { path: 'xl/media/image1.png' },
      ],
      getTables: () => [
        { name: 'Table1', displayName: 'Sales', ref: 'A1:C5', sheetIndex: 0, columns: [] },
      ],
      getPivotTables: () => [
        {
          sheetIndex: 0,
          pivotIndex: 0,
          top: 0,
          left: 0,
          rows: 4,
          cols: 3,
          cells: 12,
          fields: ['Region'],
        },
      ],
    } as unknown as WorkbookHandle;

    const summary = summarizeSpreadsheetCompatibility(wb);

    expect(summary.items.find((i) => i.id === 'cell-formatting')?.status).toBe('writable');
    expect(summary.items.find((i) => i.id === 'phonetic')?.status).toBe('writable');
    expect(summary.items.find((i) => i.id === 'auto-filter')?.status).toBe('writable');
    expect(summary.items.find((i) => i.id === 'loaded-tables')).toMatchObject({
      status: 'read-only',
      count: 1,
    });
    expect(summary.items.find((i) => i.id === 'data-validation')?.reason).toContain(
      'dropdown visibility',
    );
    expect(summary.items.find((i) => i.id === 'format-as-table')?.status).toBe('session');
    expect(summary.items.find((i) => i.id === 'hyperlinks')?.status).toBe('writable');
    expect(summary.items.find((i) => i.id === 'comments')?.status).toBe('writable');
    expect(summary.items.find((i) => i.id === 'comments')?.reason).toContain('enumerated');
    expect(summary.byId.phonetic.status).toBe('writable');
    expect(summary.items.find((i) => i.id === 'defined-names')?.status).toBe('writable');
    expect(summary.items.find((i) => i.id === 'defined-names')?.reason).toContain(
      'Workbook- and sheet-scoped',
    );
    expect(summary.byId['named-cell-styles'].status).toBe('unsupported');
    expect(summary.byId['auto-filter'].reason).toContain('AutoFilter');
    expect(summary.items.find((i) => i.id === 'sheet-protection')?.status).toBe('writable');
    expect(summary.items.find((i) => i.id === 'sheet-views')?.status).toBe('session');
    expect(summary.items.find((i) => i.id === 'pivot-authoring')?.status).toBe('writable');
    expect(summary.items.find((i) => i.id === 'charts-drawings')).toMatchObject({
      status: 'read-only',
      count: 3,
    });
    expect(summary.items.find((i) => i.id === 'chart-authoring')?.status).toBe('unsupported');
    expect(summary.byId.hyperlinks.status).toBe('writable');
    expect(spreadsheetCompatibilityItem(summary, 'comments').label).toBe('Comments');
    expect(spreadsheetCompatibilityStatus(summary, 'sheet-protection')).toBe('writable');
    expect(isSpreadsheetFeatureWritable(summary, 'defined-names')).toBe(true);
    expect(isSpreadsheetFeatureAvailable(summary, 'chart-authoring')).toBe(false);
    expect(summary.byStatus.writable).toBeGreaterThan(0);
    expect(summary.byStatus['read-only']).toBeGreaterThan(0);
    expect(summary.byStatus.session).toBeGreaterThan(0);
    expect(summary.byStatus.unsupported).toBeGreaterThan(0);
  });

  it('marks charts writable once the engine models charts and pictures', () => {
    const wb = {
      capabilities: { ...baseCaps, drawings: true, images: true, charts: true },
      getPassthroughs: () => [{ path: 'xl/charts/chart1.xml' }],
      getTables: () => [],
      getPivotTables: () => [],
    } as unknown as WorkbookHandle;

    const summary = summarizeSpreadsheetCompatibility(wb);

    expect(summary.byId['charts-drawings']).toMatchObject({ status: 'writable', count: 1 });
    expect(summary.byId['chart-authoring'].status).toBe('writable');
    expect(summary.byId['chart-authoring'].reason).toContain('moved, resized and deleted');
  });

  it('marks shapes, sparklines and slicers writable once the engine models them', () => {
    const wb = (capabilities: Record<string, boolean>) =>
      ({
        capabilities: { ...baseCaps, ...capabilities },
        getPassthroughs: () => [],
        getTables: () => [],
        getPivotTables: () => [],
      }) as unknown as WorkbookHandle;

    const on = summarizeSpreadsheetCompatibility(
      wb({ shapes: true, sparklines: true, slicers: true }),
    );
    expect(on.byId.shapes.status).toBe('writable');
    expect(on.byId.shapes.reason).toContain('cannot be saved');
    expect(on.byId.shapes.reason).toContain('rounded rectangles');
    expect(on.byId.sparklines.status).toBe('writable');
    expect(on.byId.slicers.status).toBe('writable');

    const off = summarizeSpreadsheetCompatibility(wb({}));
    expect([off.byId.shapes, off.byId.sparklines, off.byId.slicers].map((i) => i.status)).toEqual([
      'session',
      'session',
      'session',
    ]);
  });

  it('surfaces scoped defined names as an engine capability', () => {
    const wb = {
      capabilities: {
        ...baseCaps,
        definedNameMutate: true,
        definedNameScopes: false,
      },
      getPassthroughs: () => [],
      getTables: () => [],
      getPivotTables: () => [],
    } as unknown as WorkbookHandle;

    const summary = summarizeSpreadsheetCompatibility(wb);

    expect(summary.byId['defined-names'].status).toBe('writable');
    expect(summary.byId['defined-names'].reason).toContain('sheet-scoped names need engine');
  });

  it('promotes tables and named cell styles to writable once the engine authors them', () => {
    const wb = {
      capabilities: {
        ...baseCaps,
        cellStyles: true,
        cellStyleMutate: true,
        tableMutate: true,
      },
      getPassthroughs: () => [],
      getTables: () => [
        { name: 'Table1', displayName: 'Sales', ref: 'A1:C5', sheetIndex: 0, columns: [] },
      ],
      getPivotTables: () => [],
    } as unknown as WorkbookHandle;

    const summary = summarizeSpreadsheetCompatibility(wb);

    expect(summary.byId['named-cell-styles'].status).toBe('writable');
    expect(summary.byId['named-cell-styles'].reason).toContain('named cellStyle');
    expect(summary.byId['loaded-tables'].status).toBe('writable');
    expect(summary.byId['format-as-table'].status).toBe('writable');
    expect(summary.byId['format-as-table'].reason).toContain('ListObject');
  });

  it('surfaces point-only comment support as missing blank-cell enumeration', () => {
    const wb = {
      capabilities: {
        ...baseCaps,
        comments: true,
        commentsEnumerable: false,
      },
      getPassthroughs: () => [],
      getTables: () => [],
      getPivotTables: () => [],
    } as unknown as WorkbookHandle;

    const summary = summarizeSpreadsheetCompatibility(wb);

    expect(summary.byId.comments.status).toBe('writable');
    expect(summary.byId.comments.reason).toContain('blank-cell comment hydration');
  });
});
