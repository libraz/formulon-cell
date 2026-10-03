import { createStore } from 'zustand/vanilla';
import type { Addr, Range } from '../engine/types.js';
import { auditingMutators } from './auditing-mutators.js';
import { cellMutators } from './cell-mutators.js';
import { selectionMutators } from './selection-mutators.js';
import { sheetObjectMutators } from './sheet-object-mutators.js';
import { sheetSettingsMutators } from './sheet-settings-mutators.js';
import type { PageSetup, State } from './types.js';
import { defaultPageSetup } from './types.js';
import { viewMutators } from './view-mutators.js';

export type * from './types.js';
export { defaultPageSetup } from './types.js';

const initialAddr = (sheet = 0): Addr => ({ sheet, row: 0, col: 0 });
const initialRange = (sheet = 0): Range => ({ sheet, r0: 0, c0: 0, r1: 0, c1: 0 });

export const createSpreadsheetStore = () =>
  createStore<State>(() => ({
    viewport: { rowStart: 0, rowCount: 40, colStart: 0, colCount: 16, zoom: 1, widthPx: 0 },
    selection: {
      active: initialAddr(),
      range: initialRange(),
      anchor: initialAddr(),
      extraRanges: [],
    },
    layout: {
      colWidths: new Map(),
      rowHeights: new Map(),
      defaultColWidth: 64,
      defaultRowHeight: 20,
      headerColWidth: 32,
      headerRowHeight: 20,
      freezeRows: 0,
      freezeCols: 0,
      hiddenRows: new Set(),
      hiddenCols: new Set(),
      outlineRows: new Map(),
      outlineCols: new Map(),
      outlineRowGutter: 0,
      outlineColGutter: 0,
      hiddenSheets: new Set(),
      veryHiddenSheets: new Set(),
      sheetTabColors: new Map(),
    },
    data: { sheetIndex: 0, cells: new Map() },
    ui: {
      editor: { kind: 'idle' },
      hover: null,
      pendingFormat: null,
      theme: 'paper',
      fillPreview: null,
      copyRange: null,
      copyRanges: null,
      copyMode: null,
      copyRevision: 0,
      showGridLines: true,
      showHeaders: true,
      showZeros: true,
      rightToLeft: false,
      showFormulas: false,
      endMode: false,
      workbookView: 'normal',
      zoomBeforePreview: null,
      pageBreakDrag: null,
      editorRefs: [],
      r1c1: false,
      statusAggs: ['average', 'count', 'sum'],
      statusOptions: {
        capsLock: true,
        numLock: true,
        scrollLock: true,
        uploadStatus: false,
        macroRecording: false,
        viewShortcuts: true,
        zoom: true,
        zoomSlider: true,
      },
      filterRange: null,
      filterCriteria: [],
      watchPanelOpen: false,
      sheetBackgroundImages: new Map(),
    },
    format: { formats: new Map(), customCellStyles: [] },
    merges: { byAnchor: new Map(), byCell: new Map() },
    conditional: { rules: [] },
    sparkline: { sparklines: new Map() },
    charts: { charts: [] },
    illustrations: { illustrations: [] },
    watch: { watches: [] },
    traces: { items: [] },
    errorIndicators: { ignoredErrors: new Set(), validationCircles: new Set() },
    pageSetup: { setupBySheet: new Map() },
    slicers: { slicers: [] },
    tables: { tables: [], customTableStyles: [], customPivotTableStyles: [], pivotTableStyles: [] },
    sheetViews: { views: [], activeViewId: null },
    protection: { protectedSheets: new Map(), allowedEditRanges: [] },
  }));

export type SpreadsheetStore = ReturnType<typeof createSpreadsheetStore>;

// Tiny mutation helpers — single source of truth for state shape changes.
// Each domain group owns its slice of the state; the spread keeps every
// group's function objects as-is.
export const mutators = {
  ...selectionMutators,
  ...viewMutators,
  ...cellMutators,
  ...sheetObjectMutators,
  ...auditingMutators,
  ...sheetSettingsMutators,

  /** Pin every cell in `range` to the Watch Window. Existing entries are
   *  ignored and the append happens in one store update for multi-cell adds. */
  addWatchRange(store: SpreadsheetStore, range: Range): void {
    mutators.addWatchRanges(store, [range]);
  },
};

/** Pure read-helper: return the page-setup for `sheet`, falling back to
 *  `defaultPageSetup()` when no entry exists. Always returns a fully-populated
 *  record so callers can read every field without optional-chaining. */
export function getPageSetup(state: State, sheet: number): PageSetup {
  const entry = state.pageSetup.setupBySheet.get(sheet);
  if (!entry) return defaultPageSetup();
  // Merge missing fields onto the default so callers always see a complete
  // record even if `setPageSetup` was called with a sparse patch.
  const def = defaultPageSetup();
  return {
    ...def,
    ...entry,
    margins: { ...def.margins, ...entry.margins },
    printableBounds: entry.printableBounds ? { ...entry.printableBounds } : undefined,
  };
}
