import { describe, expect, it, vi } from 'vitest';
import { setFont } from '../../src/commands/format.js';
import {
  associatePaginationEngine,
  pageLayoutGaps,
  paginationFor,
  type SheetPagination,
} from '../../src/commands/pagination.js';
import { setRowsHeight } from '../../src/commands/row-col-layout.js';
import { recordFormatChange } from '../../src/commands/slice-history.js';
import { MAX_COL, MAX_ROW } from '../../src/engine/address.js';
import { fontPointsToCssPx } from '../../src/engine/font-units.js';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import type { ViewState } from '../../src/render/geometry.js';
import {
  frozenRowsHeight,
  gridOriginY,
  layoutForView,
  rowGap,
  rowHeight,
} from '../../src/render/geometry.js';
import { GridRenderer } from '../../src/render/grid.js';
import { frozenRowEnd } from '../../src/store/freeze-band.js';
import { mutators } from '../../src/store/store.js';
import type { ResolvedTheme } from '../../src/theme/resolve.js';
import { mountStubSheet } from '../test-utils/mount.js';

const A1 = { sheet: 0, row: 0, col: 0 } as const;

describe('automatic font row height against the native engine', () => {
  it('projects default rows without native row records and preserves the authored font on reload', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'en-US' });
    const mounted = await mountStubSheet({ workbook });
    let reloaded: WorkbookHandle | undefined;
    try {
      expect(workbook.isStub).toBe(false);
      const store = mounted.instance.store;
      mutators.setActive(store, A1);
      const original = layoutForView(store.getState()).defaultRowHeight;
      expect(workbook.getRowLayouts(0)).toEqual([]);

      mutators.setCellFormat(store, A1, { fontSize: fontPointsToCssPx(24) });
      const enlarged = layoutForView(store.getState());
      expect(enlarged.rowHeights.get(0)).toBeGreaterThan(original);
      expect(enlarged.rowHeights.get(100) ?? enlarged.defaultRowHeight).toBe(original);
      expect(workbook.getRowLayouts(0)).toEqual([]);
      expect(workbook.getEffectiveStyle(0, 0, 0)).not.toBeNull();

      reloaded = await WorkbookHandle.loadBytes(workbook.save());
      expect(reloaded.isStub).toBe(false);
      expect(reloaded.getRowLayouts(0)).toEqual([]);
      const effective = reloaded.getEffectiveStyle(0, 0, 0);
      const xf = effective && reloaded.getCellXf(effective.xfIndex);
      expect(xf && reloaded.getFontRecord(xf.fontIndex)?.size).toBe(24);
    } finally {
      reloaded?.dispose();
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('leaves a native custom-height row under user control', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'en-US' });
    const mounted = await mountStubSheet({ workbook });
    try {
      const store = mounted.instance.store;
      mutators.setActive(store, A1);
      setRowsHeight(store, mounted.instance.history, A1.row, A1.row, 24, workbook);
      const before = layoutForView(store.getState());
      expect(before.rowHeights.get(A1.row)).toBe(24);

      mutators.setCellFormat(store, A1, { fontSize: fontPointsToCssPx(36) });
      const after = layoutForView(store.getState());
      expect(after.rowHeights.get(A1.row)).toBe(24);
      expect(workbook.getRowLayouts(0)).toMatchObject([{ row: 0, height: 24, customHeight: true }]);
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('uses row-axis fonts as the row base across a whole-sheet font and reload', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'en-US' });
    const mounted = await mountStubSheet({ workbook });
    let reloaded: WorkbookHandle | undefined;
    try {
      const store = mounted.instance.store;
      const history = mounted.instance.history;
      const row = 7;
      const original = store.getState().layout.defaultRowHeight;
      expect(original).toBe(20);
      expect(workbook.workbookDisplayFont?.size).toBe(11);
      const fontAt = (source: WorkbookHandle, atRow: number, col: number): number | undefined => {
        const effective = source.getEffectiveStyle(0, atRow, col);
        const xf = effective && source.getCellXf(effective.xfIndex);
        return xf ? source.getFontRecord(xf.fontIndex)?.size : undefined;
      };
      const assertRowAxisCellXf = (source: WorkbookHandle): void => {
        const rowAxis = source.getSheetFontAxisSnapshot(0)?.rows.find((entry) => entry.row === row);
        expect(rowAxis?.styleXf).toBeGreaterThan(0);
        const cellXf = rowAxis && source.getCellXf(rowAxis.styleXf);
        expect(cellXf && source.getFontRecord(cellXf.fontIndex)?.size).toBe(11);
        const namedStyleXf = rowAxis && source.getCellStyleXf(rowAxis.styleXf);
        expect(namedStyleXf && source.getFontRecord(namedStyleXf.fontIndex)?.size).not.toBe(11);
      };

      mutators.selectAll(store);
      recordFormatChange(history, store, () => {
        if (!setFont(store.getState(), store, { fontSize: fontPointsToCssPx(24) })) {
          throw new Error('whole-sheet font was rejected');
        }
      });
      mutators.selectRow(store, row);
      recordFormatChange(history, store, () => {
        if (!setFont(store.getState(), store, { fontSize: fontPointsToCssPx(11) })) {
          throw new Error('row font was rejected');
        }
      });

      const afterRowFormat = layoutForView(store.getState());
      expect(afterRowFormat.defaultRowHeight).toBe(42);
      expect(afterRowFormat.rowHeights.get(row) ?? afterRowFormat.defaultRowHeight).toBe(original);
      expect(afterRowFormat.rowHeights.get(row + 1) ?? afterRowFormat.defaultRowHeight).toBe(
        afterRowFormat.defaultRowHeight,
      );
      expect(fontAt(workbook, row, MAX_COL)).toBe(11);
      expect(fontAt(workbook, row + 1, MAX_COL)).toBe(24);
      assertRowAxisCellXf(workbook);

      mutators.setCellFormat(
        store,
        { sheet: 0, row, col: 0 },
        {
          fontSize: fontPointsToCssPx(30),
        },
      );
      const withCellException = layoutForView(store.getState());
      const normalCss = fontPointsToCssPx(workbook.workbookDisplayFont?.size ?? 11);
      const exceptionHeight =
        original + Math.round(fontPointsToCssPx(30) * 1.28) - Math.round(normalCss * 1.28);
      expect(exceptionHeight).toBe(52);
      expect(withCellException.rowHeights.get(row)).toBe(exceptionHeight);

      reloaded = await WorkbookHandle.loadBytes(workbook.save());
      expect(fontAt(reloaded, row, MAX_COL)).toBe(11);
      assertRowAxisCellXf(reloaded);
      const snapshot = reloaded.getSheetFontAxisSnapshot(0);
      expect(snapshot?.rows.some((entry) => entry.row === row)).toBe(true);
      const state = store.getState();
      const reloadState = {
        ...state,
        format: { ...state.format, sheetAxisOverrides: new Map() },
      };
      associatePaginationEngine(reloadState, reloaded);
      const afterReload = layoutForView(reloadState);
      expect(afterReload.rowHeights.get(row)).toBe(exceptionHeight);
      expect(afterReload.rowHeights.get(row + 1) ?? afterReload.defaultRowHeight).toBe(
        afterReload.defaultRowHeight,
      );
    } finally {
      reloaded?.dispose();
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('invalidates a synchronous undo projection after native format restoration', async () => {
    const frames = new Map<number, FrameRequestCallback>();
    let nextFrameId = 1;
    vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => {
      const id = nextFrameId++;
      frames.set(id, callback);
      return id;
    });
    vi.stubGlobal('cancelAnimationFrame', (id: number) => {
      frames.delete(id);
    });
    const workbook = await WorkbookHandle.createDefault({ locale: 'en-US' });
    const mounted = await mountStubSheet({ workbook });
    let renderer: GridRenderer | undefined;
    let unsubscribe = (): void => {};
    try {
      frames.clear();
      const store = mounted.instance.store;
      const history = mounted.instance.history;
      const original = store.getState().layout.defaultRowHeight;
      const nativeFontAt = (row: number, col: number): number | undefined => {
        const effective = workbook.getEffectiveStyle(0, row, col);
        const xf = effective && workbook.getCellXf(effective.xfIndex);
        return xf ? workbook.getFontRecord(xf.fontIndex)?.size : undefined;
      };
      const originalNativeFontSize = nativeFontAt(0, MAX_COL);
      mutators.selectAll(store);
      recordFormatChange(history, store, () => {
        if (!setFont(store.getState(), store, { fontSize: fontPointsToCssPx(24) })) {
          throw new Error('whole-sheet font was rejected');
        }
      });
      mutators.setActive(store, A1);
      expect(layoutForView(store.getState()).defaultRowHeight).toBeGreaterThan(original);
      expect(nativeFontAt(0, MAX_COL)).toBe(24);

      const grid = mounted.host.querySelector<HTMLElement>('.fc-host__grid');
      if (!grid) throw new Error('mounted spreadsheet grid is missing');
      vi.spyOn(grid, 'getBoundingClientRect').mockReturnValue({
        x: 0,
        y: 0,
        top: 0,
        right: 640,
        bottom: 180,
        left: 0,
        width: 640,
        height: 180,
        toJSON: () => ({}),
      } as DOMRect);
      renderer = new GridRenderer({
        host: grid,
        canvas: document.createElement('canvas'),
        getState: () => store.getState(),
        getTheme: () => ({}) as ResolvedTheme,
        getWb: () => workbook,
        onViewportSize: (rowCount, colCount, widthPx) =>
          mutators.setViewportSize(store, rowCount, colCount, widthPx),
      });
      (renderer as unknown as { paint(): void }).paint = () => {};
      renderer.resize();
      const flushFrames = (): void => {
        let flushed = 0;
        while (frames.size > 0) {
          if (flushed++ > 20) throw new Error('animation frames did not settle');
          const first = frames.entries().next().value as [number, FrameRequestCallback] | undefined;
          if (!first) break;
          const [id, callback] = first;
          frames.delete(id);
          callback(0);
        }
      };
      flushFrames();
      const enlargedRowCount = store.getState().viewport.rowCount;
      expect(enlargedRowCount).toBe(5);

      const synchronousObservations: number[] = [];
      unsubscribe = store.subscribe((state) => {
        synchronousObservations.push(layoutForView(state).defaultRowHeight);
        renderer?.invalidate();
      });
      expect(history.undo()).toBe(true);
      flushFrames();

      expect(synchronousObservations.length).toBeGreaterThan(0);
      expect(synchronousObservations.some((height) => height > original)).toBe(true);
      expect(layoutForView(store.getState()).defaultRowHeight).toBe(original);
      expect(nativeFontAt(0, MAX_COL)).toBe(originalNativeFontSize);
      expect(store.getState().format.sheetFontOverrides?.get(0)?.patch.fontSize).toBeUndefined();
      const settledState = store.getState();
      const settledLayout = layoutForView(settledState);
      const { viewport } = settledState;
      const firstRow = Math.max(
        viewport.rowStart,
        frozenRowEnd(settledLayout),
        viewport.navigationRange?.r0 ?? 0,
      );
      const lastRow = (viewport.navigationRange?.r1 ?? MAX_ROW) + 1;
      const bodyHeight = Math.max(
        0,
        180 - gridOriginY(settledLayout) - frozenRowsHeight(settledLayout, viewport),
      );
      let usedHeight = 0;
      let visibleRows = 0;
      while (visibleRows < Math.min(lastRow - firstRow, 2_000) && usedHeight < bodyHeight) {
        const row = firstRow + visibleRows;
        usedHeight += Math.max(
          0,
          rowHeight(settledLayout, row, viewport) + rowGap(settledLayout, row),
        );
        visibleRows += 1;
      }
      expect(settledState.viewport.rowCount).toBe(Math.max(1, visibleRows + 1));
      expect(settledState.viewport.rowCount).toBe(9);
      expect(settledState.viewport.rowCount).toBeGreaterThan(enlargedRowCount);
      expect(workbook.getRowLayouts(0)).toEqual([]);
    } finally {
      unsubscribe();
      renderer?.dispose();
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('uses projected row geometry in Page Layout and returns to Normal geometry', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'en-US' });
    const mounted = await mountStubSheet({ workbook });
    try {
      const store = mounted.instance.store;
      mutators.setActive(store, A1);
      const normalDefault = layoutForView(store.getState()).defaultRowHeight;
      mutators.setCellFormat(store, A1, { fontSize: fontPointsToCssPx(24) });

      const normal = layoutForView(store.getState());
      const normalHeight = normal.rowHeights.get(0);
      expect(normalHeight).toBeGreaterThan(normalDefault);
      mutators.setWorkbookView(store, 'pageLayout');
      const pageLayout = layoutForView(store.getState());
      expect(pageLayout.rowHeights.get(0)).toBe(normalHeight);
      expect(pageLayout.rulerRowHeight).toBeGreaterThan(0);
      mutators.setWorkbookView(store, 'normal');
      expect(layoutForView(store.getState()).rowHeights.get(0)).toBe(normalHeight);
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('keeps the GridRenderer Page Layout pagination in step with projected rows', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'en-US' });
    const mounted = await mountStubSheet({ workbook });
    const store = mounted.instance.store;
    const renderer = new GridRenderer({
      host: mounted.instance.host,
      canvas: document.createElement('canvas'),
      getState: () => store.getState(),
      getTheme: () => ({}) as ResolvedTheme,
      getWb: () => workbook,
    });
    const productionPagination = (
      renderer as unknown as {
        paginationForView(state: ViewState): SheetPagination | null;
      }
    ).paginationForView.bind(renderer);
    try {
      mutators.setWorkbookView(store, 'pageLayout');
      const plainState = store.getState();
      const plainView = { ...plainState, layout: layoutForView(plainState) };
      const plain = productionPagination(plainView);
      const native = paginationFor(plainState, 0, {
        engine: workbook,
        throughRow: plainState.viewport.rowStart + plainState.viewport.rowCount,
        throughCol: plainState.viewport.colStart + plainState.viewport.colCount,
      });
      expect(plain?.rowBands).toEqual(native.rowBands);

      mutators.selectAll(store);
      expect(setFont(store.getState(), store, { fontSize: fontPointsToCssPx(24) })).toBe(true);
      const formattedState = store.getState();
      const formattedView = { ...formattedState, layout: layoutForView(formattedState) };
      const actual = productionPagination(formattedView);
      const fallback = paginationFor({ ...formattedState, layout: formattedView.layout }, 0, {
        throughRow: formattedState.viewport.rowStart + formattedState.viewport.rowCount,
        throughCol: formattedState.viewport.colStart + formattedState.viewport.colCount,
      });
      const gaps = pageLayoutGaps({ ...formattedState, layout: formattedView.layout }, 0, {
        throughRow: formattedState.viewport.rowStart + formattedState.viewport.rowCount,
        throughCol: formattedState.viewport.colStart + formattedState.viewport.colCount,
      });
      expect(actual?.rowBands).toEqual(fallback.rowBands);
      expect([...gaps.pageGapRows.keys()]).toEqual(fallback.rowBands.map((band) => band.start));
    } finally {
      renderer.dispose();
      mounted.dispose();
      workbook.dispose();
    }
  });
});
