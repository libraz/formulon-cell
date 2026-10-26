import { MAX_COL, MAX_ROW } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import { navigationBoundsFor, syncNavigationViewport } from '../interact/navigation-policy.js';
import { frozenColEnd, frozenColStart, frozenRowEnd, frozenRowStart } from './freeze-band.js';
import { clampZoom, renderScale } from './render-scale.js';
import { clampSplitPaneOrigin } from './split-panes.js';
import type { SpreadsheetStore } from './store.js';
import type {
  ChartSourceHighlight,
  CopyMode,
  EditorMode,
  EditorRefHighlight,
  SelectedObject,
  State,
  StatusAggKey,
  StatusBarOptionKey,
} from './types.js';

/** Row height ceiling: 409 pt in canonical CSS px. */
const MAX_ROW_HEIGHT_PX = (409 * 96) / 72;

/** Round canonical px to whole screen px under the layout's screen scale. */
const snapToScreenPx = (layout: State['layout'], px: number): number => {
  const s = renderScale(layout);
  return Math.round(px * s) / s;
};

/** One screen pixel in canonical px: the narrowest visible row or column. */
const minScreenPx = (layout: State['layout']): number => 1 / renderScale(layout);

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

  setShowOutlineSymbols(store: SpreadsheetStore, on: boolean): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, showOutlineSymbols: on } }));
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

  setFilterRanges(store: SpreadsheetStore, ranges: readonly Range[]): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, filterRanges: ranges.map((r) => ({ ...r })) } }));
  },

  setEditorRefs(store: SpreadsheetStore, refs: EditorRefHighlight[]): void {
    store.setState((s) => ({ ...s, ui: { ...s.ui, editorRefs: refs } }));
  },

  /** Select a chart or picture, or none; `sourceRefs` outlines a chart's data. */
  setSelectedObject(
    store: SpreadsheetStore,
    object: SelectedObject | null,
    sourceRefs: ChartSourceHighlight[] = [],
  ): void {
    store.setState((s) => ({
      ...s,
      ui: { ...s.ui, selectedObject: object, chartSourceRefs: sourceRefs },
    }));
  },

  setZoom(store: SpreadsheetStore, zoom: number): void {
    const z = clampZoom(zoom);
    store.setState((s) => ({ ...s, viewport: { ...s.viewport, zoom: z } }));
  },

  setViewportSize(store: SpreadsheetStore, rowCount: number, colCount: number, widthPx = 0): void {
    const rows = Math.max(1, Math.floor(rowCount));
    const cols = Math.max(1, Math.floor(colCount));
    const width = Math.max(0, widthPx);
    store.setState((s) => {
      if (
        s.viewport.rowCount === rows &&
        s.viewport.colCount === cols &&
        s.viewport.widthPx === width
      ) {
        return s;
      }
      const bounds = navigationBoundsFor(store);
      const minRowStart = Math.max(frozenRowEnd(s.layout), bounds?.r0 ?? 0);
      const minColStart = Math.max(frozenColEnd(s.layout), bounds?.c0 ?? 0);
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
      colWidths.set(
        col,
        Math.max(minScreenPx(s.layout), Math.min(800, snapToScreenPx(s.layout, px))),
      );
      return { ...s, layout: { ...s.layout, colWidths } };
    });
  },

  setRowHeight(store: SpreadsheetStore, row: number, px: number): void {
    store.setState((s) => {
      const rowHeights = new Map(s.layout.rowHeights);
      rowHeights.set(
        row,
        Math.max(minScreenPx(s.layout), Math.min(MAX_ROW_HEIGHT_PX, snapToScreenPx(s.layout, px))),
      );
      return { ...s, layout: { ...s.layout, rowHeights } };
    });
  },

  /** Scroll the top / left split panes by `dRow` / `dCol`. A pane stops at the
   *  sheet edge and where the pane past the split begins. No-op unless split. */
  scrollSplitPane(store: SpreadsheetStore, dRow: number, dCol: number): void {
    store.setState((s) => {
      const { layout, viewport } = s;
      if (layout.splitPanes !== true) return s;
      const freezeOriginRow = clampSplitPaneOrigin(
        layout.freezeRows,
        Math.max(viewport.rowStart, frozenRowEnd(layout)),
        (layout.freezeOriginRow ?? 0) + dRow,
      );
      const freezeOriginCol = clampSplitPaneOrigin(
        layout.freezeCols,
        Math.max(viewport.colStart, frozenColEnd(layout)),
        (layout.freezeOriginCol ?? 0) + dCol,
      );
      if (
        freezeOriginRow === layout.freezeOriginRow &&
        freezeOriginCol === layout.freezeOriginCol
      ) {
        return s;
      }
      return { ...s, layout: { ...layout, freezeOriginRow, freezeOriginCol } };
    });
  },

  scrollBy(store: SpreadsheetStore, dRow: number, dCol: number): void {
    // Keep at least one body row/col visible past the freeze zone, otherwise the
    // viewport disappears off the right/bottom.
    store.setState((s) => {
      const bounds = navigationBoundsFor(store);
      const minRowStart = Math.max(frozenRowEnd(s.layout), bounds?.r0 ?? 0);
      const minColStart = Math.max(frozenColEnd(s.layout), bounds?.c0 ?? 0);
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

  /** Pin `rows` rows / `cols` columns starting at `origin` (default A1). Pass
   *  0/0 to unfreeze. Freezing scrolls the body past the band; unfreezing an
   *  axis scrolls it to the band's first row/column so nothing jumps. With
   *  `split` the band becomes split panes instead of a freeze. */
  setFreezePanes(
    store: SpreadsheetStore,
    rows: number,
    cols: number,
    origin: { row: number; col: number } = { row: 0, col: 0 },
    split = false,
  ): void {
    const fr = Math.max(0, Math.floor(rows));
    const fc = Math.max(0, Math.floor(cols));
    store.setState((s) => {
      const freezeOriginRow = fr > 0 ? Math.max(0, Math.floor(origin.row)) : 0;
      const freezeOriginCol = fc > 0 ? Math.max(0, Math.floor(origin.col)) : 0;
      const layout = {
        ...s.layout,
        freezeRows: fr,
        freezeCols: fc,
        freezeOriginRow,
        freezeOriginCol,
        splitPanes: split && (fr > 0 || fc > 0),
      };
      const fromRow =
        fr === 0 && s.layout.freezeRows > 0 ? frozenRowStart(s.layout) : s.viewport.rowStart;
      const fromCol =
        fc === 0 && s.layout.freezeCols > 0 ? frozenColStart(s.layout) : s.viewport.colStart;
      return {
        ...s,
        layout,
        viewport: {
          ...s.viewport,
          rowStart: Math.max(frozenRowEnd(layout), fromRow),
          colStart: Math.max(frozenColEnd(layout), fromCol),
        },
      };
    });
    syncNavigationViewport(store);
  },
};
