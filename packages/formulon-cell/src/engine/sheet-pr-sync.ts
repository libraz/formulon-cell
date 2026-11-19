import { ColorKind } from '@libraz/formulon';
import { DEFAULT_OUTLINE_SUMMARY } from '../store/outline-summary.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import type { WriteResult } from './result.js';
import type { EngineSheetPrInput, SheetTabColorSpec } from './sheet-pr-xml.js';
import { argbFromCssColor, cssColorFromArgb } from './theme-color.js';
import type { WorkbookHandle } from './workbook-handle.js';

const argbHex = (argb: number): string => (argb >>> 0).toString(16).toUpperCase().padStart(8, '0');

/** CSS colour of a `<tabColor>`, resolving theme and indexed selectors. */
function tabColorToCss(wb: WorkbookHandle, spec: SheetTabColorSpec): string | null {
  if (spec.rgb !== undefined) return cssColorFromArgb(Number.parseInt(spec.rgb, 16));
  const colorSpec =
    spec.theme !== undefined
      ? { kind: ColorKind.Theme, argb: 0, theme: spec.theme, tint: spec.tint ?? 0, indexed: 0 }
      : spec.indexed !== undefined
        ? { kind: ColorKind.Indexed, argb: 0, theme: 0, tint: 0, indexed: spec.indexed }
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
  for (let sheet = 0; sheet < wb.sheetCount; sheet += 1) hydrateSheetPrForSheet(wb, store, sheet);
}

/** Seed one sheet's tab colour and outline summary position from its `<sheetPr>`. */
export function hydrateSheetPrForSheet(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  if (!wb.capabilities.sheetPrXml) return;
  const pr = wb.getSheetPr(sheet);
  if (!pr) return;
  const color = pr.tabColor ? tabColorToCss(wb, pr.tabColor) : null;
  if (color) mutators.setSheetTabColor(store, sheet, color);
  if (!pr.summaryBelow || !pr.summaryRight) {
    mutators.setOutlineSummary(store, sheet, { below: pr.summaryBelow, right: pr.summaryRight });
  }
}

/**
 * Write the store's tab colour and outline summary position for `sheet` into
 * `<sheetPr>`. A tab colour the file already states in another notation (a
 * theme colour) is left alone while it resolves to the colour the store holds.
 * When the engine refuses the write the store follows the engine: every field
 * written here is re-read from `<sheetPr>` and the refusal is returned.
 */
export function syncSheetPrToEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): WriteResult {
  if (!wb.capabilities.sheetPrXml) return { ok: true };
  const { layout } = store.getState();
  const current = wb.getSheetPr(sheet);
  if (!current) return { ok: true };
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

  if (Object.keys(change).length === 0) return { ok: true };
  const written = wb.setSheetPr(sheet, change);
  if (!written.ok) {
    wb.withEngineSyncMuted(() => refreshSheetPrFromEngine(wb, store, sheet));
  }
  return written;
}

/** Set the store's tab colour and outline summary position for `sheet` to what `<sheetPr>` holds. */
export function refreshSheetPrFromEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  const pr = wb.getSheetPr(sheet);
  if (!pr) return;
  mutators.setSheetTabColor(store, sheet, pr.tabColor ? tabColorToCss(wb, pr.tabColor) : null);
  mutators.setOutlineSummary(store, sheet, { below: pr.summaryBelow, right: pr.summaryRight });
}
