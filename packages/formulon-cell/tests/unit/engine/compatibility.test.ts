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
    expect(summary.items.find((i) => i.id === 'defined-names')?.status).toBe('writable');
    expect(summary.items.find((i) => i.id === 'defined-names')?.reason).toContain(
      'Workbook- and sheet-scoped',
    );
    expect(summary.items.find((i) => i.id === 'sheet-protection')?.status).toBe('writable');
    expect(summary.items.find((i) => i.id === 'sheet-views')?.status).toBe('session');
    expect(summary.items.find((i) => i.id === 'session-charts')?.status).toBe('session');
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
    expect(isSpreadsheetFeatureAvailable(summary, 'session-charts')).toBe(true);
    expect(isSpreadsheetFeatureAvailable(summary, 'chart-authoring')).toBe(false);
    expect(summary.byStatus.writable).toBeGreaterThan(0);
    expect(summary.byStatus['read-only']).toBeGreaterThan(0);
    expect(summary.byStatus.session).toBeGreaterThan(0);
    expect(summary.byStatus.unsupported).toBeGreaterThan(0);
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
