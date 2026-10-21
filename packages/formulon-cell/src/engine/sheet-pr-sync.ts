import { DEFAULT_OUTLINE_SUMMARY } from '../store/outline-summary.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import type { EngineSheetPrInput, SheetTabColorSpec } from './sheet-pr-xml.js';
import { argbFromCssColor, cssColorFromArgb } from './theme-color.js';
import type { WorkbookHandle } from './workbook-handle.js';

/** Colour kinds of the engine's `ColorSpec`. */
const COLOR_KIND = { rgb: 1, theme: 2, indexed: 3 } as const;

const argbHex = (argb: number): string => (argb >>> 0).toString(16).toUpperCase().padStart(8, '0');

/** CSS colour of a `<tabColor>`, resolving theme and indexed selectors. */
function tabColorToCss(wb: WorkbookHandle, spec: SheetTabColorSpec): string | null {
  if (spec.rgb !== undefined) return cssColorFromArgb(Number.parseInt(spec.rgb, 16));
  const colorSpec =
    spec.theme !== undefined
      ? { kind: COLOR_KIND.theme, rgb: 0, theme: spec.theme, tint: spec.tint ?? 0, indexed: 0 }
      : spec.indexed !== undefined
        ? { kind: COLOR_KIND.indexed, rgb: 0, theme: 0, tint: 0, indexed: spec.indexed }
        : null;
  if (!colorSpec) return null;
  const argb = wb.resolveColorArgb(colorSpec);
  return argb === null ? null : cssColorFromArgb(argb);
}

/**
 * Seed every sheet's tab colour and outline summary position from the
 * workbook's `<sheetPr>`. Sheets that state neither add nothing to the store.
 */
export function hydrateSheetPrFromEngine(wb: WorkbookHandle, store: SpreadsheetStore): void {
  if (!wb.capabilities.sheetPrXml) return;
  for (let sheet = 0; sheet < wb.sheetCount; sheet += 1) {
    const pr = wb.getSheetPr(sheet);
    if (!pr) continue;
    const color = pr.tabColor ? tabColorToCss(wb, pr.tabColor) : null;
    if (color) mutators.setSheetTabColor(store, sheet, color);
    if (!pr.summaryBelow || !pr.summaryRight) {
      mutators.setOutlineSummary(store, sheet, { below: pr.summaryBelow, right: pr.summaryRight });
    }
  }
}

/**
 * Write the store's tab colour and outline summary position for `sheet` into
 * `<sheetPr>`. A tab colour the file already states in another notation (a
 * theme colour) is left alone while it resolves to the colour the store holds.
 */
export function syncSheetPrToEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  if (!wb.capabilities.sheetPrXml) return;
  const { layout } = store.getState();
  const current = wb.getSheetPr(sheet);
  if (!current) return;
  const change: EngineSheetPrInput = {};

  const color = layout.sheetTabColors.get(sheet) ?? null;
  const currentCss = current.tabColor ? tabColorToCss(wb, current.tabColor) : null;
  if (color === null) {
    if (current.tabColor) change.tabColor = null;
  } else if (currentCss?.toLowerCase() !== color.toLowerCase()) {
    const argb = argbFromCssColor(color);
    if (argb !== null) change.tabColor = { rgb: argbHex(argb) };
  }

  const summary = layout.outlineSummary?.get(sheet) ?? DEFAULT_OUTLINE_SUMMARY;
  if (summary.below !== current.summaryBelow) change.summaryBelow = summary.below;
  if (summary.right !== current.summaryRight) change.summaryRight = summary.right;

  if (Object.keys(change).length > 0) wb.setSheetPr(sheet, change);
}
