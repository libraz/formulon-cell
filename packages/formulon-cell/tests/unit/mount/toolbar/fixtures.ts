import { AnchorEditAs, AnchorKind, SparklineType } from '@libraz/formulon';
import { vi } from 'vitest';
import { defaultTableOverlay } from '../../../../src/commands/format-as-table.js';
import { createPivotTableAtPlacement } from '../../../../src/commands/pivot-table.js';
import {
  defaultSparklineGroupSettings,
  setSparklines,
} from '../../../../src/commands/sparkline.js';
import { addrKey } from '../../../../src/engine/address.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { mutators, type SpreadsheetStore } from '../../../../src/store/store.js';
import type { DynamicDropdownsCtx } from '../../../../src/toolbar/ribbon/dynamic-dropdowns.js';
import type { RibbonRenderHelpers } from '../../../../src/toolbar/ribbon/render-ribbon.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { scaledTimeout } from '../../../timeouts.js';

// Minimal helpers stub: enough for the renderer to emit a shell, no real
// dropdown DOM. The toolbar still needs `createSelect/Color/Icon/makeSvg`
// because every command path may reach them.
export const stubHelpers = (): RibbonRenderHelpers => ({
  createSelect: () => document.createElement('div'),
  createColor: () => document.createElement('div'),
  createIcon: () => null,
  makeSvg: (_viewBox, _pathData, className) => {
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.setAttribute('class', className);
    return svg;
  },
  chevronPath: 'M0 0',
});

/** Put a chart in the store and select it, so the ribbon shows the chart tabs. */
export const selectStoreChart = (store: SpreadsheetStore, id = 'chart-native-0-1'): void => {
  mutators.upsertChart(store, {
    id,
    sheet: 0,
    nativeObjectId: 1,
    nativeAnchor: {
      kind: AnchorKind.TwoCell,
      editAs: AnchorEditAs.OneCell,
      fromRow: 1,
      fromCol: 1,
      fromRowOffsetEmu: 0,
      fromColOffsetEmu: 0,
      toRow: 10,
      toCol: 6,
      toRowOffsetEmu: 0,
      toColOffsetEmu: 0,
      widthEmu: 0,
      heightEmu: 0,
    },
    kind: 'column',
    data: {
      domain: 'category',
      categories: [{ label: 'Jan' }],
      series: [{ id: 's0', points: [{ value: 1 }] }],
    },
  });
  mutators.setSelectedObject(store, { id, name: 'Chart 1' });
};

/** Put a rectangle in the store and select it, so the ribbon shows the Shape Format tab. */
export const selectStoreShape = (store: SpreadsheetStore, id = 'shape-test'): void => {
  mutators.upsertIllustration(store, { id, kind: 'shape', shape: 'rectangle', sheet: 0 });
  mutators.setSelectedObject(store, { id, name: 'Rectangle 1' });
};

/** Put a picture in the store and select it, so the ribbon shows the Picture Format tab. */
export const selectStorePicture = (store: SpreadsheetStore, id = 'picture-test'): void => {
  mutators.upsertIllustration(store, { id, kind: 'image', src: 'data:,', sheet: 0 });
  mutators.setSelectedObject(store, { id, name: 'Picture 1' });
};

/** Put a table over the active cell, so the ribbon shows the Table Design tab while no object is selected. */
export const selectStoreTable = (store: SpreadsheetStore): void => {
  const { active } = store.getState().selection;
  mutators.upsertTableOverlay(store, {
    ...defaultTableOverlay('table-test', {
      sheet: active.sheet,
      r0: active.row,
      c0: active.col,
      r1: active.row + 2,
      c1: active.col + 1,
    }),
    source: 'engine',
    name: 'Table1',
  });
};

/** Put a sparkline in the active cell, so the ribbon shows the Sparkline tab while no object is selected. */
export const selectStoreSparkline = (store: SpreadsheetStore): void => {
  const { active } = store.getState().selection;
  setSparklines(
    store,
    [{ addr: active, source: 'A1:A4' }],
    defaultSparklineGroupSettings(SparklineType.Line),
  );
};

/** One selection per set of contextual tabs; together they show every tab the ribbon has. */
export const CONTEXTUAL_SELECTIONS: readonly ((store: SpreadsheetStore) => void)[] = [
  selectStoreSparkline,
  selectStoreTable,
  selectStoreChart,
  selectStoreShape,
  selectStorePicture,
];

/**
 * A sheet on the real engine with the active cell inside a PivotTable report, so the ribbon
 * shows the PivotTable tabs. The stub engine has no pivots.
 */
export const mountPivotSheet = async (locale: 'en' | 'ja' = 'en'): Promise<MountedStubSheet> => {
  const workbook = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Region', 'Sales'],
    ['East', 1],
    ['West', 2],
    ['East', 3],
  ];
  for (const [row, cols] of rows.entries()) {
    for (const [col, value] of cols.entries()) {
      if (typeof value === 'number') workbook.setNumber({ sheet: 0, row, col }, value);
      else workbook.setText({ sheet: 0, row, col }, value);
    }
  }
  const sheet = await mountStubSheet({ locale, workbook });
  mutators.replaceCells(sheet.instance.store, workbook.cells(0));
  createPivotTableAtPlacement(
    {
      origin: 'ribbon' as const,
      commandId: 'test',
      store: sheet.instance.store,
      workbook,
      history: sheet.instance.history,
    },
    { source: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 }, rowField: 'Region', valueField: 'Sales' },
    { row: 8, col: 4 },
  );
  mutators.setActive(sheet.instance.store, { sheet: 0, row: 9, col: 4 });
  return sheet;
};

export const waitFor = async (
  predicate: () => boolean,
  timeoutMs = scaledTimeout(250),
): Promise<void> => {
  const deadline = Date.now() + timeoutMs;
  while (!predicate()) {
    if (Date.now() >= deadline) throw new Error('Timed out waiting for expected state.');
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
};

export const seedNumber = (
  sheet: MountedStubSheet,
  row: number,
  col: number,
  value: number,
): void => {
  sheet.workbook.setNumber({ sheet: 0, row, col }, value);
  sheet.instance.store.setState((state) => {
    const cells = new Map(state.data.cells);
    cells.set(addrKey({ sheet: 0, row, col }), {
      value: { kind: 'number', value },
      formula: null,
    });
    return { ...state, data: { ...state.data, cells } };
  });
};

export const seedText = (
  sheet: MountedStubSheet,
  row: number,
  col: number,
  value: string,
): void => {
  sheet.workbook.setText({ sheet: 0, row, col }, value);
  sheet.instance.store.setState((state) => {
    const cells = new Map(state.data.cells);
    cells.set(addrKey({ sheet: 0, row, col }), {
      value: { kind: 'text', value },
      formula: null,
    });
    return { ...state, data: { ...state.data, cells } };
  });
};

export const dynamicDropdownNoopOverrides = (): Partial<DynamicDropdownsCtx> => ({
  applyRibbonPasteAction: vi.fn(),
  updatePasteMenu: vi.fn(),
  applyPivotTableAction: vi.fn(),
  applyDefinedNameAction: vi.fn(),
  applyLinksAction: vi.fn(),
  applyFillSeries: vi.fn(),
  updateFillMenu: vi.fn(),
  applyFillDirection: vi.fn(),
  applyClearAction: vi.fn(),
  updateClearMenu: vi.fn(),
  applyFreezeAction: vi.fn(),
  updateFreezeMenu: vi.fn(),
  applyTextOrientationAction: vi.fn(),
  updateTextOrientationMenu: vi.fn(),
  applyCellInsertAction: vi.fn(),
  updateCellInsertMenu: vi.fn(),
  applyCellDeleteAction: vi.fn(),
  updateCellDeleteMenu: vi.fn(),
  applyCellFormatAction: vi.fn(),
  applyPageBreakAction: vi.fn(),
  applySheetBackgroundAction: vi.fn(),
  applyPrintAreaAction: vi.fn(),
  applyArrangeAction: vi.fn(),
  applyChartElementAction: vi.fn(),
  updateChartAddElementMenu: vi.fn(),
  applyChartColorsAction: vi.fn(),
  updateChartColorsMenu: vi.fn(),
  applySparklineAction: vi.fn(),
  updateSparklineMenu: vi.fn(),
  updateUseInFormulaMenu: vi.fn(),
  applyShapeFillAction: vi.fn(),
  updateShapeFillMenu: vi.fn(),
  applyShapeOutlineAction: vi.fn(),
  applyShapeOutlineWeightAction: vi.fn(),
  applyShapeOutlineDashAction: vi.fn(),
  updateShapeOutlineMenu: vi.fn(),
  updateShapesInsertMenu: vi.fn(),
  applyUiTheme: vi.fn(),
  applySortMenuAction: vi.fn(),
  applyFindSelectAction: vi.fn(),
  applyAutoSumFormula: vi.fn(),
  applyFormulaAuditAction: vi.fn(),
  applyWatchAction: vi.fn(),
  applyReviewCommentAction: vi.fn(),
  applyProtectAction: vi.fn(),
  applyCalcOptionAction: vi.fn(),
  applyWhatIfAction: vi.fn(),
  updateArrangeMenu: vi.fn(),
  updateCellStylesMenu: vi.fn(),
  updateCurrencyMenu: vi.fn(),
  updatePageBreaksMenu: vi.fn(),
  updatePrintAreaMenu: vi.fn(),
  updateProtectMenu: vi.fn(),
  updatePageThemeMenu: vi.fn(),
  updateReviewCommentsMenu: vi.fn(),
  updateSortMenu: vi.fn(),
  updateTableStylesMenu: vi.fn(),
  updatePivotStylesMenu: vi.fn(),
  updateWatchMenu: vi.fn(),
  createRecommendedChartFromSelection: vi.fn(),
  createChartFromSelection: vi.fn(),
  chartKindFromAction: vi.fn((_action: string): 'column' => 'column'),
  insertPictureFromRibbon: vi.fn(),
  insertShapeFromRibbon: vi.fn(),
  insertScreenshotFromRibbon: vi.fn(),
  applyScriptAction: vi.fn(),
  applyPdfAction: vi.fn(),
  createTableFromSelection: vi.fn(),
  openTableStyleFooterAction: vi.fn(),
  applyCellStyleFromRibbon: vi.fn(),
  openCellStyleFooterAction: vi.fn(),
  applyCurrencyPreset: vi.fn(),
  openCurrencyFooterAction: vi.fn(),
  splitTextToColumns: vi.fn(),
  splitTextToColumnsCustom: vi.fn(),
  applyDataValidationAction: vi.fn(),
  applyAddInAction: vi.fn(),
  applyConditionalMenuAction: vi.fn(),
  applySymbolAction: vi.fn(),
  updateClearArrowsMenu: vi.fn(),
  updateDataValidationMenu: vi.fn(),
  updateLinksMenu: vi.fn(),
  updateErrorCheckingMenu: vi.fn(),
  updateFormatCellsMenu: vi.fn(),
});
