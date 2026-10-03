import { vi } from 'vitest';
import { addrKey } from '../../../../src/engine/address.js';
import type { DynamicDropdownsCtx } from '../../../../src/toolbar/ribbon/dynamic-dropdowns.js';
import type { RibbonRenderHelpers } from '../../../../src/toolbar/ribbon/render-ribbon.js';
import type { MountedStubSheet } from '../../../test-utils/mount.js';

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

export const waitFor = async (predicate: () => boolean, timeoutMs = 250): Promise<void> => {
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
  applyUiTheme: vi.fn(),
  applySortMenuAction: vi.fn(),
  applyFindSelectAction: vi.fn(),
  applyAutoSumFormula: vi.fn(),
  applyFormulaAuditAction: vi.fn(),
  applyWatchAction: vi.fn(),
  applyReviewCommentAction: vi.fn(),
  applyProtectAction: vi.fn(),
  applyCalcOptionAction: vi.fn(),
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
