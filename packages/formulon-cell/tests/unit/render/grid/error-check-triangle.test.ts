import { describe, expect, it } from 'vitest';
import {
  createPivotTableFromRange,
  refreshPivotReportsAfterEdit,
} from '../../../../src/commands/pivot-table.js';
import type { CellValue } from '../../../../src/engine/types.js';
import { PivotAxis } from '../../../../src/engine/types.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createConditionalCache } from '../../../../src/render/conditional.js';
import { buildColLayout, buildRowLayout, layoutForView } from '../../../../src/render/geometry.js';
import { paintCells } from '../../../../src/render/grid/cells.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

const theme = {
  bg: '#fff',
  bgRail: '#eee',
  bgElev: '#fff',
  bgHeader: '#eee',
  fg: '#000',
  fgMute: '#666',
  fgFaint: '#aaa',
  fgStrong: '#000',
  rule: '#ccc',
  ruleStrong: '#999',
  freezeRule: '#9b9b9b',
  accent: '#0078d4',
  accentFg: '#fff',
  accentSoft: 'rgba(0,120,212,0.1)',
  cellErrorFg: '#c00',
  cellFormulaFg: '#333',
  cellBoolFg: '#060',
  cellNumFg: '#000',
  hoverStripe: 'rgba(0,0,0,0.04)',
  headerBg: '#eee',
  headerRule: '#ccc',
  headerFg: '#666',
  headerFgActive: '#000',
  pageBackdrop: '#aaa',
  pagePaper: '#fff',
  pageEdge: '#999',
  pageBandFg: '#666',
  pageBreakAuto: '#2f6fd0',
  pageBreakManual: '#1a4fa0',
  pageNumberFg: '#999',
  pageOutside: '#ddd',
  fontUi: 'sans-serif',
  fontMono: 'monospace',
  textCell: 13,
  textHeader: 11.5,
} as ResolvedTheme;

/** Paint the sheet and record every fillRect with the fill colour it used. */
function paint(store: SpreadsheetStore, mac: boolean, wb: WorkbookHandle | null = null) {
  const rects: { color: string; rect: number[] }[] = [];
  const target: Record<string, unknown> = {
    fillStyle: '',
    font: '',
    textAlign: 'left',
    measureText: () =>
      ({ width: 7, actualBoundingBoxAscent: 9, actualBoundingBoxDescent: 2 }) as TextMetrics,
    fillRect: (...rect: number[]) => rects.push({ color: String(target.fillStyle), rect }),
  };
  const ctx = new Proxy(target, {
    get: (o, p: string) => (p in o ? o[p] : () => undefined),
    set: (o, p: string, v) => {
      o[p] = v;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
  // The Mac profile draws one canonical px as 0.75 screen px (1 pt = 1 px).
  store.setState((s) => ({
    ...s,
    layout: mac ? { ...s.layout, screenScale: 0.75 } : s.layout,
    viewport: { ...s.viewport, widthPx: 640, rowCount: 14, colCount: 8 },
  }));
  const state = store.getState();
  const view = { ...state, layout: layoutForView(state) };
  const result = paintCells(
    {
      ctx,
      wb,
      locale: 'ja-JP',
      getDisplay: undefined,
      sheetBackgroundImage: null,
      fontEffects: mac,
      conditionalCache: createConditionalCache(),
    },
    view,
    theme,
    buildColLayout(view.layout, view.viewport),
    buildRowLayout(view.layout, view.viewport),
  );
  return { result, rects };
}

const at = (store: SpreadsheetStore, row: number, col: number, value: CellValue, formula: string) =>
  mutators.setCell(store, { sheet: 0, row, col }, value, formula);

const DIV0: CellValue = { kind: 'error', code: 1, text: '#DIV/0!' };
const NA: CellValue = { kind: 'error', code: 6, text: '#N/A' };

describe('error-check triangle (Mac desktop)', () => {
  it('marks a formula error with a 6 pt (0,129,0) triangle in the top-left corner', () => {
    const store = createSpreadsheetStore();
    at(store, 1, 1, DIV0, '=1/0');
    const { result, rects } = paint(store, true);
    expect(result.errorTriangles.map((t) => [t.addr.row, t.addr.col, t.rect.w])).toEqual([
      [1, 1, 6],
    ]);
    const triangle = rects.filter((r) => r.color === '#008100');
    expect(triangle.length).toBe(6);
    expect(triangle[0]?.rect.slice(2)).toEqual([6, 1]);
  });

  it('leaves every #N/A unmarked, whatever formula gave it', () => {
    const store = createSpreadsheetStore();
    at(store, 1, 1, NA, '=NA()');
    at(store, 2, 1, NA, '=XLOOKUP(9,J1:J3,J1:J3)');
    expect(paint(store, true).result.errorTriangles).toEqual([]);
  });

  it('marks the #SPILL! anchor of a blocked PivotTable', async () => {
    const wb = await WorkbookHandle.createDefault();
    const rows: (string | number)[][] = [
      ['Region', 'Product', 'Amt'],
      ['East', 'A', 500],
      ['North', 'B', 100],
    ];
    for (const [r, row] of rows.entries()) {
      for (const [c, v] of row.entries()) {
        const addr = { sheet: 0, row: r, col: 10 + c };
        if (typeof v === 'number') wb.setNumber(addr, v);
        else wb.setText(addr, v);
      }
    }
    createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 10, r1: 2, c1: 12 },
      destination: { sheet: 0, row: 2, col: 0 },
      rowField: 'Region',
      valueField: 'Amt',
    });
    const store = createSpreadsheetStore();
    mutators.replaceCells(store, wb.cells(0));
    wb.setText({ sheet: 0, row: 7, col: 0 }, 'X');
    wb.setPivotFieldAxis(0, 0, 1, PivotAxis.Row);
    refreshPivotReportsAfterEdit(store, wb);
    expect(wb.pivotSpills(0)).toHaveLength(1);
    const { result } = paint(store, true, wb);
    expect(result.errorTriangles.map((t) => [t.addr.row, t.addr.col])).toEqual([[2, 0]]);
    wb.dispose();
  });
});

describe('error-check triangle (other profiles)', () => {
  it('keeps the existing colour and rule', () => {
    const store = createSpreadsheetStore();
    at(store, 1, 1, NA, '=VLOOKUP("zz",A1:A2,1,FALSE)');
    const { result, rects } = paint(store, false);
    expect(result.errorTriangles).toHaveLength(1);
    expect(rects.some((r) => r.color === '#008000')).toBe(true);
  });
});

describe('number stored as text (Mac desktop)', () => {
  const flagged = ['123', '0123', ' 1', '1e3', '0123456789', '-5', '1,234', '¥100', '50%'];
  const plain = ['12/3', 'abc'];

  it('marks text that reads as a number, not date-like text, other text or a real number', () => {
    const store = createSpreadsheetStore();
    for (const [i, text] of [...flagged, ...plain].entries()) {
      mutators.setCell(store, { sheet: 0, row: i, col: 1 }, { kind: 'text', value: text }, null);
    }
    mutators.setCell(store, { sheet: 0, row: 12, col: 1 }, { kind: 'number', value: 123 }, null);
    const { result, rects } = paint(store, true);
    expect(result.errorTriangles.map((t) => t.addr.row)).toEqual(flagged.map((_, i) => i));
    expect(result.errorTriangles.every((t) => t.rect.w === 6)).toBe(true);
    expect(rects.some((r) => r.color === '#008100')).toBe(true);
  });

  it('scales with the zoom: 12 pt legs at 200%', () => {
    const store = createSpreadsheetStore();
    mutators.setCell(store, { sheet: 0, row: 1, col: 1 }, { kind: 'text', value: '123' }, null);
    store.setState((s) => ({ ...s, viewport: { ...s.viewport, zoom: 2 } }));
    expect(paint(store, true).result.errorTriangles[0]?.rect.w).toBe(12);
  });
});
