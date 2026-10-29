import { describe, expect, it, vi } from 'vitest';

import { addrKey } from '../../../../src/engine/address.js';
import type { CellValue, Range } from '../../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import {
  type AxisLayout,
  buildColLayout,
  buildRowLayout,
  layoutForView,
  type ViewState,
} from '../../../../src/render/geometry.js';
import { GridRenderer } from '../../../../src/render/grid.js';
import { createSpreadsheetStore, type State } from '../../../../src/store/store.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

type Cell = { value: CellValue; formula: string | null };
type SpillInfo = { anchorRow: number; anchorCol: number; rows: number; cols: number } | null;

class CountingCellMap extends Map<string, Cell> {
  iterations = 0;

  reads = 0;

  override [Symbol.iterator](): MapIterator<[string, Cell]> {
    this.iterations += 1;
    return super[Symbol.iterator]();
  }

  override get(key: string): Cell | undefined {
    this.reads += 1;
    return super.get(key);
  }
}

const theme = { bg: '#fff', accent: '#0078d4' } as unknown as ResolvedTheme;

interface CanvasSpy {
  ctx: CanvasRenderingContext2D;
  strokes: { rect: [number, number, number, number]; dash: number[] }[];
}

function canvasSpy(): CanvasSpy {
  const strokes: CanvasSpy['strokes'] = [];
  let dash: number[] = [];
  const ctx = {
    save(): void {},
    restore(): void {},
    setLineDash(value: number[]): void {
      dash = [...value];
    },
    strokeRect(x: number, y: number, width: number, height: number): void {
      strokes.push({ rect: [x, y, width, height], dash: [...dash] });
    },
  } as unknown as CanvasRenderingContext2D;
  return { ctx, strokes };
}

function stateWithCells(
  cells: Map<string, Cell>,
  options: {
    rowStart?: number;
    rowCount?: number;
    colStart?: number;
    colCount?: number;
    freezeRows?: number;
    freezeCols?: number;
    hiddenRows?: readonly number[];
    hiddenCols?: readonly number[];
  } = {},
): State {
  const base = createSpreadsheetStore().getState();
  return {
    ...base,
    layout: {
      ...base.layout,
      freezeRows: options.freezeRows ?? 0,
      freezeCols: options.freezeCols ?? 0,
      hiddenRows: new Set(options.hiddenRows ?? []),
      hiddenCols: new Set(options.hiddenCols ?? []),
    },
    viewport: {
      ...base.viewport,
      rowStart: options.rowStart ?? 10,
      rowCount: options.rowCount ?? 2,
      colStart: options.colStart ?? 5,
      colCount: options.colCount ?? 2,
      widthPx: 320,
    },
    data: { ...base.data, cells },
  };
}

function makeWorkbook(
  spillRanges: (sheet: number) => Range[] | null,
  spillInfo: (sheet: number, row: number, col: number) => SpillInfo,
  capability: boolean | null = true,
): WorkbookHandle {
  const capabilityFlags = capability === null ? {} : { capabilities: { spillInfo: capability } };
  return {
    ...capabilityFlags,
    spillRanges,
    spillInfo,
  } as unknown as WorkbookHandle;
}

function makeRenderer(
  getState: () => State,
  canvas: CanvasSpy,
  workbook: WorkbookHandle,
): GridRenderer {
  const host = { getBoundingClientRect: () => ({ width: 640, height: 480 }) } as HTMLElement;
  return new GridRenderer({
    host,
    canvas: { getContext: () => canvas.ctx, style: {} } as unknown as HTMLCanvasElement,
    getState,
    getTheme: () => theme,
    getWb: () => workbook,
  });
}

function paintSpills(renderer: GridRenderer, state: State): { rows: AxisLayout; cols: AxisLayout } {
  const view = { ...state, layout: layoutForView(state) } as ViewState;
  const rows = buildRowLayout(view.layout, view.viewport);
  const cols = buildColLayout(view.layout, view.viewport);
  (
    renderer as unknown as {
      paintSpills(state: ViewState, theme: ResolvedTheme, cols: AxisLayout, rows: AxisLayout): void;
    }
  ).paintSpills(view, theme, cols, rows);
  return { rows, cols };
}

const numberCell = (value: number): Cell => ({ value: { kind: 'number', value }, formula: null });
const blankCell = (): Cell => ({ value: { kind: 'blank' }, formula: null });

describe('grid spill painting', () => {
  it('queries each visible cell, including blank cached cells, for large spill-aware workbooks', () => {
    const cells = new CountingCellMap();
    for (let i = 0; i < 8; i += 1)
      cells.set(addrKey({ sheet: 0, row: 100 + i, col: 100 }), numberCell(i));
    const spillInfo = vi.fn((_sheet: number, row: number, col: number): SpillInfo => {
      if (row === 10 && (col === 5 || col === 6)) {
        return { anchorRow: 8, anchorCol: 4, rows: 4, cols: 3 };
      }
      if (row === 11 && col === 5) return { anchorRow: row, anchorCol: col, rows: 1, cols: 1 };
      return null;
    });
    const spillRanges = vi.fn(() => []);
    const state = stateWithCells(cells);
    const canvas = canvasSpy();
    const renderer = makeRenderer(() => state, canvas, makeWorkbook(spillRanges, spillInfo));

    const { rows, cols } = paintSpills(renderer, state);

    expect(spillRanges).not.toHaveBeenCalled();
    expect(spillInfo.mock.calls.map(([, row, col]) => [row, col])).toEqual(
      rows.visible.flatMap((row) => cols.visible.map((col) => [row, col])),
    );
    expect(canvas.strokes.filter(({ dash }) => dash.length === 0)).toHaveLength(1);
    expect(cells.iterations).toBe(1); // Only the #SPILL! candidate index is built.
  });

  it('checks only visible blockers for a huge offscreen #SPILL! target', () => {
    const cells = new CountingCellMap();
    cells.set(addrKey({ sheet: 0, row: 0, col: 0 }), {
      value: { kind: 'error', code: 8, text: '#SPILL!' },
      formula: '=SEQUENCE(100,100)',
    });
    cells.set(addrKey({ sheet: 0, row: 10, col: 5 }), {
      value: { kind: 'number', value: 4 },
      formula: '=1+3',
    });
    cells.set(addrKey({ sheet: 0, row: 10, col: 6 }), {
      value: { kind: 'text', value: 'blocked' },
      formula: null,
    });
    cells.set(addrKey({ sheet: 0, row: 11, col: 5 }), blankCell());
    for (let i = 0; i < 4; i += 1)
      cells.set(addrKey({ sheet: 0, row: 200 + i, col: 1 }), numberCell(i));

    const spillInfo = vi.fn(
      (_sheet: number, row: number, col: number): SpillInfo =>
        row === 0 && col === 0 ? { anchorRow: 0, anchorCol: 0, rows: 100, cols: 100 } : null,
    );
    const spillRanges = vi.fn(() => []);
    const state = stateWithCells(cells);
    const canvas = canvasSpy();
    const renderer = makeRenderer(() => state, canvas, makeWorkbook(spillRanges, spillInfo));

    paintSpills(renderer, state);

    expect(spillInfo).toHaveBeenCalledTimes(5); // Four visible cells plus the offscreen error anchor.
    expect(spillRanges).not.toHaveBeenCalled();
    expect(cells.reads).toBeLessThanOrEqual(4);
    expect(canvas.strokes.filter(({ dash }) => dash.join(',') === '3,3')).toHaveLength(2);
  });

  it('refreshes engine shapes every paint and reuses the error candidate index by map identity', () => {
    const cells = new CountingCellMap();
    cells.set(addrKey({ sheet: 0, row: 10, col: 5 }), {
      value: { kind: 'error', code: 8, text: '#SPILL!' },
      formula: '=SEQUENCE(10)',
    });
    for (let i = 0; i < 8; i += 1)
      cells.set(addrKey({ sheet: 0, row: 200 + i, col: 1 }), numberCell(i));
    let spillRows = 2;
    const spillInfo = vi.fn(
      (_sheet: number, row: number, col: number): SpillInfo =>
        row >= 10 && row < 10 + spillRows && col >= 5 && col < 7
          ? { anchorRow: 10, anchorCol: 5, rows: spillRows, cols: 2 }
          : null,
    );
    const state = stateWithCells(cells, { rowCount: 3 });
    let currentState = state;
    const canvas = canvasSpy();
    const renderer = makeRenderer(
      () => currentState,
      canvas,
      makeWorkbook(() => [], spillInfo),
    );

    const firstAxes = paintSpills(renderer, state);
    const firstOutline = canvas.strokes.find(({ dash }) => dash.length === 0);
    expect(spillInfo).toHaveBeenCalledTimes(
      firstAxes.rows.visible.length * firstAxes.cols.visible.length,
    );
    expect(firstOutline?.rect[3]).toBe(39);

    spillRows = 3;
    paintSpills(renderer, state);
    const outlines = canvas.strokes.filter(({ dash }) => dash.length === 0);
    expect(outlines[1]?.rect[3]).toBe(59);
    expect(cells.iterations).toBe(1);

    const replacement = new CountingCellMap(cells);
    currentState = stateWithCells(replacement, { rowCount: 3 });
    paintSpills(renderer, currentState);
    expect(replacement.iterations).toBe(1);
  });

  it('indexes error candidates separately per sheet while reusing the same cell map', () => {
    const cells = new CountingCellMap();
    cells.set(addrKey({ sheet: 0, row: 100, col: 100 }), {
      value: { kind: 'error', code: 8, text: '#SPILL!' },
      formula: '=SEQUENCE(2,2)',
    });
    cells.set(addrKey({ sheet: 1, row: 200, col: 200 }), {
      value: { kind: 'error', code: 8, text: '#SPILL!' },
      formula: '=SEQUENCE(2,2)',
    });
    for (let i = 0; i < 4; i += 1)
      cells.set(addrKey({ sheet: 0, row: 300 + i, col: 1 }), numberCell(i));

    const sheetZero = stateWithCells(cells);
    const sheetOne = { ...sheetZero, data: { ...sheetZero.data, sheetIndex: 1 } };
    const spillInfo = vi.fn((): SpillInfo => null);
    const renderer = makeRenderer(
      () => sheetZero,
      canvasSpy(),
      makeWorkbook(() => [], spillInfo),
    );

    paintSpills(renderer, sheetZero);
    expect(spillInfo).toHaveBeenCalledWith(0, 100, 100);
    expect(cells.iterations).toBe(1);

    paintSpills(renderer, sheetOne);
    expect(spillInfo).toHaveBeenCalledWith(1, 200, 200);
    expect(cells.iterations).toBe(2);

    paintSpills(renderer, sheetZero);
    expect(cells.iterations).toBe(2);
  });

  it('uses engine range enumeration for small or spillInfo-unsupported workbooks', () => {
    const cells = new CountingCellMap();
    cells.set(addrKey({ sheet: 0, row: 0, col: 0 }), blankCell());
    const range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    const smallSpillInfo = vi.fn((): SpillInfo => null);
    const smallSpillRanges = vi.fn(() => [range]);
    const smallState = stateWithCells(cells, { rowStart: 0, colStart: 0 });
    const smallCanvas = canvasSpy();
    const smallRenderer = makeRenderer(
      () => smallState,
      smallCanvas,
      makeWorkbook(smallSpillRanges, smallSpillInfo),
    );

    paintSpills(smallRenderer, smallState);
    expect(smallSpillRanges).toHaveBeenCalledOnce();
    expect(smallSpillInfo).not.toHaveBeenCalled();

    const fallbackCells = new CountingCellMap();
    fallbackCells.set(addrKey({ sheet: 0, row: 0, col: 0 }), {
      value: { kind: 'number', value: 1 },
      formula: '=SEQUENCE(2,2)',
    });
    fallbackCells.set(addrKey({ sheet: 0, row: 0, col: 1 }), numberCell(2));
    fallbackCells.set(addrKey({ sheet: 0, row: 1, col: 0 }), numberCell(3));
    fallbackCells.set(addrKey({ sheet: 0, row: 1, col: 1 }), numberCell(4));
    for (let i = 0; i < 4; i += 1)
      fallbackCells.set(addrKey({ sheet: 0, row: 100 + i, col: 100 }), numberCell(i));
    const fallbackState = stateWithCells(fallbackCells, { rowStart: 0, colStart: 0 });

    const unsupportedSpillRanges = vi.fn(() => null);
    const unsupportedSpillInfo = vi.fn((): SpillInfo => null);
    const unsupportedCanvas = canvasSpy();
    const unsupportedRenderer = makeRenderer(
      () => fallbackState,
      unsupportedCanvas,
      makeWorkbook(unsupportedSpillRanges, unsupportedSpillInfo, false),
    );
    paintSpills(unsupportedRenderer, fallbackState);
    expect(unsupportedSpillRanges).toHaveBeenCalledOnce();
    expect(unsupportedSpillInfo).not.toHaveBeenCalled();
    expect(unsupportedCanvas.strokes.filter(({ dash }) => dash.length === 0)).toHaveLength(1);

    const legacySpillRanges = vi.fn(() => null);
    const legacySpillInfo = vi.fn((): SpillInfo => null);
    const legacyCanvas = canvasSpy();
    const legacyRenderer = makeRenderer(
      () => fallbackState,
      legacyCanvas,
      makeWorkbook(legacySpillRanges, legacySpillInfo, null),
    );
    paintSpills(legacyRenderer, fallbackState);
    expect(legacySpillRanges).toHaveBeenCalledOnce();
    expect(legacySpillInfo).not.toHaveBeenCalled();
    expect(legacyCanvas.strokes.filter(({ dash }) => dash.length === 0)).toHaveLength(1);
  });

  it('probes visible freeze/hidden axes and paints only visible blockers', () => {
    const cells = new CountingCellMap();
    for (let i = 0; i < 20; i += 1)
      cells.set(addrKey({ sheet: 0, row: 300 + i, col: 1 }), numberCell(i));
    cells.set(addrKey({ sheet: 0, row: 0, col: 0 }), {
      value: { kind: 'error', code: 8, text: '#SPILL!' },
      formula: '=SEQUENCE(20,20)',
    });
    cells.set(addrKey({ sheet: 0, row: 10, col: 6 }), {
      value: { kind: 'text', value: 'visible blocker' },
      formula: null,
    });
    cells.set(addrKey({ sheet: 0, row: 12, col: 7 }), {
      value: { kind: 'number', value: 5 },
      formula: '=2+3',
    });
    cells.set(addrKey({ sheet: 0, row: 11, col: 6 }), {
      value: { kind: 'text', value: 'hidden row blocker' },
      formula: null,
    });
    cells.set(addrKey({ sheet: 0, row: 10, col: 5 }), {
      value: { kind: 'text', value: 'hidden column blocker' },
      formula: null,
    });
    const spillInfo = vi.fn(
      (_sheet: number, _row: number, _col: number): SpillInfo => ({
        anchorRow: 0,
        anchorCol: 0,
        rows: 20,
        cols: 20,
      }),
    );
    const spillRanges = vi.fn(() => []);
    const state = stateWithCells(cells, {
      rowStart: 10,
      rowCount: 3,
      colStart: 5,
      colCount: 3,
      freezeRows: 2,
      freezeCols: 1,
      hiddenRows: [1, 11],
      hiddenCols: [5],
    });
    const canvas = canvasSpy();
    const renderer = makeRenderer(() => state, canvas, makeWorkbook(spillRanges, spillInfo));

    const { rows, cols } = paintSpills(renderer, state);

    expect(rows.visible).toEqual([0, 10, 12]);
    expect(cols.visible).toEqual([0, 6, 7]);
    expect(spillInfo.mock.calls.map(([, row, col]) => [row, col])).toEqual(
      rows.visible.flatMap((row) => cols.visible.map((col) => [row, col])),
    );
    expect(spillRanges).not.toHaveBeenCalled();
    expect(cells.reads).toBe(rows.visible.length * cols.visible.length - 1);
    expect(canvas.strokes.filter(({ dash }) => dash.join(',') === '3,3')).toHaveLength(2);
  });
});
