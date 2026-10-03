import type { Addr, Range } from '../engine/types.js';
import { navigationBoundsFor, syncNavigationViewport } from '../interact/navigation-policy.js';
import type { SpreadsheetStore } from './store.js';
import type {
  CopyMode,
  EditorMode,
  EditorRefHighlight,
  State,
  StatusAggKey,
  StatusBarOptionKey,
} from './types.js';

/** UI chrome flags, viewport scroll/zoom, freeze panes and row/column sizes. */
export const viewMutators = {
  setEditor(store: SpreadsheetStore, mode: EditorMode): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, editor: mode } }));
  },

  setPendingFormat(store: SpreadsheetStore, pending: State['ui']['pendingFormat']): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, pendingFormat: pending } }));
  },

  setHover(store: SpreadsheetStore, addr: Addr | null): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, hover: addr } }));
  },

  setTheme(store: SpreadsheetStore, theme: 'paper' | 'ink' | (string & {})): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, theme } }));
  },

  setShowGridLines(store: SpreadsheetStore, on: boolean): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, showGridLines: on } }));
  },

  setShowHeaders(store: SpreadsheetStore, on: boolean): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, showHeaders: on } }));
  },

  setShowZeros(store: SpreadsheetStore, on: boolean): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, showZeros: on } }));
  },

  setRightToLeft(store: SpreadsheetStore, on: boolean): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, rightToLeft: on } }));
  },

  setShowFormulas(store: SpreadsheetStore, on: boolean): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, showFormulas: on } }));
  },

  setEndMode(store: SpreadsheetStore, on: boolean): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, endMode: on } }));
  },

  setSheetBackgroundImage(store: SpreadsheetStore, sheet: number, url: string | undefined): void {
    store.setState((s) => {
      const sheetBackgroundImages = new Map(s.ui.sheetBackgroundImages);
      const normalized = url?.trim();
      if (normalized) sheetBackgroundImages.set(sheet, normalized);
      else sheetBackgroundImages.delete(sheet);
      return { ...s, ui: { ...s.ui, sheetBackgroundImages } };
    });
  },

  setWorkbookView(
    store: SpreadsheetStore,
    mode: 'normal' | 'pageLayout' | 'pageBreakPreview',
  ): void {
    store.setState((s) =>
      s.ui.workbookView === mode ? s : { ...s, ui: { ...s.ui, workbookView: mode } },
    );
  },

  setPageBreakDrag(
    store: SpreadsheetStore,
    drag: { axis: 'row' | 'col'; position: number } | null,
  ): void {
    store.setState((s) =>
      s.ui.pageBreakDrag === drag ? s : { ...s, ui: { ...s.ui, pageBreakDrag: drag } },
    );
  },

  setR1C1(store: SpreadsheetStore, on: boolean): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, r1c1: on } }));
  },

  toggleStatusAgg(store: SpreadsheetStore, key: StatusAggKey): void {
    store.setState((s) => {
      const set = new Set(s.ui.statusAggs);
      if (set.has(key)) set.delete(key);
      else set.add(key);
      return { ...s, ui: { ...s.ui, statusAggs: Array.from(set) } };
    });
  },

  setStatusAggs(store: SpreadsheetStore, keys: StatusAggKey[]): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, statusAggs: [...keys] } }));
  },

  setStatusOption(store: SpreadsheetStore, key: StatusBarOptionKey, on: boolean): void {
    store.setState((s) =>
      s.ui.statusOptions[key] === on
        ? s
        : {
            ...s,
            ui: { ...s.ui, statusOptions: { ...s.ui.statusOptions, [key]: on } },
          },
    );
  },

  toggleStatusOption(store: SpreadsheetStore, key: StatusBarOptionKey): void {
    store.setState((s) => ({
      ...s,
      ui: {
        ...s.ui,
        statusOptions: {
          ...s.ui.statusOptions,
          [key]: !s.ui.statusOptions[key],
        },
      },
    }));
  },

  setFilterRange(store: SpreadsheetStore, range: Range | null): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, filterRange: range } }));
  },

  setEditorRefs(store: SpreadsheetStore, refs: EditorRefHighlight[]): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, editorRefs: refs } }));
  },

  setZoom(store: SpreadsheetStore, zoom: number): void {
    const z = Math.max(0.5, Math.min(4, zoom));
    store.setState((s) => ({ ...s, viewport: { ...s.viewport, zoom: z } }));
  },

  setViewportSize(store: SpreadsheetStore, rowCount: number, colCount: number, widthPx = 0): void {
    const rows = Math.max(1, Math.floor(rowCount));
    const cols = Math.max(1, Math.floor(colCount));
    const width = Math.max(0, widthPx);
    const MAX_ROW = 1_048_575;
    const MAX_COL = 16_383;
    store.setState((s) => {
      if (
        s.viewport.rowCount === rows &&
        s.viewport.colCount === cols &&
        s.viewport.widthPx === width
      ) {
        return s;
      }
      const bounds = navigationBoundsFor(store);
      const minRowStart = Math.max(s.layout.freezeRows, bounds?.r0 ?? 0);
      const minColStart = Math.max(s.layout.freezeCols, bounds?.c0 ?? 0);
      const maxRowStart = bounds
        ? Math.max(minRowStart, bounds.r1 + 1 - rows)
        : Math.max(minRowStart, MAX_ROW + 1 - rows);
      const maxColStart = bounds
        ? Math.max(minColStart, bounds.c1 + 1 - cols)
        : Math.max(minColStart, MAX_COL + 1 - cols);
      return {
        ...s,
        viewport: {
          ...s.viewport,
          rowCount: rows,
          colCount: cols,
          widthPx: width,
          rowStart: Math.min(maxRowStart, Math.max(minRowStart, s.viewport.rowStart)),
          colStart: Math.min(maxColStart, Math.max(minColStart, s.viewport.colStart)),
          ...(bounds ? { navigationRange: { ...bounds } } : { navigationRange: undefined }),
        },
      };
    });
  },

  setFillPreview(store: SpreadsheetStore, range: Range | null): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, fillPreview: range } }));
  },

  setCopyRange(store: SpreadsheetStore, range: Range | null, mode: CopyMode = 'copy'): void {
    store.setState((s) => ({
      ...s,
      ui: {
        ...s.ui,
        copyRevision: (s.ui.copyRevision ?? 0) + 1,
        copyRange: range ? { ...range } : null,
        copyRanges: null,
        copyMode: range ? mode : null,
      },
    }));
  },

  setCopyRanges(store: SpreadsheetStore, ranges: Range[] | null, mode: CopyMode = 'copy'): void {
    store.setState((s) => ({
      ...s,
      ui: {
        ...s.ui,
        copyRevision: (s.ui.copyRevision ?? 0) + 1,
        copyRange: ranges?.[0] ? { ...ranges[0] } : null,
        copyRanges: ranges && ranges.length > 0 ? ranges.map((r) => ({ ...r })) : null,
        copyMode: ranges?.[0] ? mode : null,
      },
    }));
  },

  setColWidth(store: SpreadsheetStore, col: number, px: number): void {
    store.setState((s) => {
      const colWidths = new Map(s.layout.colWidths);
      colWidths.set(col, Math.max(28, Math.min(800, px)));
      return { ...s, layout: { ...s.layout, colWidths } };
    });
  },

  setRowHeight(store: SpreadsheetStore, row: number, px: number): void {
    store.setState((s) => {
      const rowHeights = new Map(s.layout.rowHeights);
      rowHeights.set(row, Math.max(16, Math.min(400, px)));
      return { ...s, layout: { ...s.layout, rowHeights } };
    });
  },

  scrollBy(store: SpreadsheetStore, dRow: number, dCol: number): void {
    // Desktop spreadsheets sheet bounds — keep at least one body row/col visible past the
    // freeze zone, otherwise the viewport disappears off the right/bottom.
    const MAX_ROW = 1_048_575;
    const MAX_COL = 16_383;
    store.setState((s) => {
      const bounds = navigationBoundsFor(store);
      const minRowStart = Math.max(s.layout.freezeRows, bounds?.r0 ?? 0);
      const minColStart = Math.max(s.layout.freezeCols, bounds?.c0 ?? 0);
      const maxRowStart = bounds
        ? Math.max(minRowStart, bounds.r1 + 1 - s.viewport.rowCount)
        : Math.max(minRowStart, MAX_ROW + 1 - s.viewport.rowCount);
      const maxColStart = bounds
        ? Math.max(minColStart, bounds.c1 + 1 - s.viewport.colCount)
        : Math.max(minColStart, MAX_COL + 1 - s.viewport.colCount);
      return {
        ...s,
        viewport: {
          ...s.viewport,
          rowStart: Math.min(maxRowStart, Math.max(minRowStart, s.viewport.rowStart + dRow)),
          colStart: Math.min(maxColStart, Math.max(minColStart, s.viewport.colStart + dCol)),
        },
      };
    });
  },

  /** Pin the first `rows` rows / `cols` columns. Pass 0/0 to unfreeze.
   *  Scrolls past the frozen zone if the body viewport is currently inside it. */
  setFreezePanes(store: SpreadsheetStore, rows: number, cols: number): void {
    const fr = Math.max(0, Math.floor(rows));
    const fc = Math.max(0, Math.floor(cols));
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, freezeRows: fr, freezeCols: fc },
      viewport: {
        ...s.viewport,
        rowStart: Math.max(fr, s.viewport.rowStart),
        colStart: Math.max(fc, s.viewport.colStart),
      },
    }));
    syncNavigationViewport(store);
  },
};
