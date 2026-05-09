import {
  defaultPageSetup,
  getPageSetup,
  mutators,
  type PageMargins,
  type PageOrientation,
  type PageSetup,
  type PaperSize,
  type SpreadsheetStore,
  type State,
} from '../store/store.js';
import { parsePrintTitleCols, parsePrintTitleRows } from './print.js';

/** Built-in margin presets. Values match the spreadsheet defaults
 *  (Normal / Wide / Narrow), expressed in inches. */
export type MarginPreset = 'normal' | 'wide' | 'narrow';

const MARGIN_PRESETS: Record<MarginPreset, PageMargins> = {
  normal: { top: 0.75, right: 0.7, bottom: 0.75, left: 0.7 },
  wide: { top: 1, right: 1, bottom: 1, left: 1 },
  narrow: { top: 0.75, right: 0.25, bottom: 0.75, left: 0.25 },
};

export function marginPresetValues(preset: MarginPreset): PageMargins {
  return { ...MARGIN_PRESETS[preset] };
}

/** Match the supplied margins against the named presets and return the
 *  closest one when each side is within `tolerance` inches. Returns `null`
 *  when the margins are bespoke — used by the chrome to render "Custom" in
 *  the margins picker rather than lying about which preset is active. */
export function marginPresetOf(margins: PageMargins, tolerance = 0.01): MarginPreset | null {
  for (const [name, preset] of Object.entries(MARGIN_PRESETS) as [MarginPreset, PageMargins][]) {
    if (
      Math.abs(margins.top - preset.top) <= tolerance &&
      Math.abs(margins.right - preset.right) <= tolerance &&
      Math.abs(margins.bottom - preset.bottom) <= tolerance &&
      Math.abs(margins.left - preset.left) <= tolerance
    ) {
      return name;
    }
  }
  return null;
}

export interface PageSetupEntry {
  sheet: number;
  setup: PageSetup;
}

export type PageSetupPatch = Omit<Partial<PageSetup>, 'margins'> & {
  margins?: Partial<PageSetup['margins']>;
};

export function pageSetupForSheet(state: State, sheet: number): PageSetup {
  return getPageSetup(state, sheet);
}

export function listPageSetups(state: State): readonly PageSetupEntry[] {
  return [...state.pageSetup.setupBySheet.keys()]
    .sort((a, b) => a - b)
    .map((sheet) => ({ sheet, setup: getPageSetup(state, sheet) }));
}

export function setPageSetup(
  store: SpreadsheetStore,
  sheet: number,
  patch: PageSetupPatch,
): PageSetup {
  const current = getPageSetup(store.getState(), sheet);
  const next: Partial<PageSetup> = {
    ...patch,
    margins: patch.margins ? { ...current.margins, ...patch.margins } : undefined,
  };
  if (!patch.margins) delete next.margins;
  mutators.setPageSetup(store, sheet, next);
  return getPageSetup(store.getState(), sheet);
}

export function resetPageSetup(store: SpreadsheetStore, sheet: number): PageSetup {
  mutators.setPageSetup(store, sheet, null);
  return defaultPageSetup();
}

export function clearPrintTitles(store: SpreadsheetStore, sheet: number): PageSetup {
  mutators.setPageSetup(store, sheet, { printTitleRows: undefined, printTitleCols: undefined });
  return getPageSetup(store.getState(), sheet);
}

export function setPrintTitleRows(
  store: SpreadsheetStore,
  sheet: number,
  rows: string | undefined,
): PageSetup | null {
  const normalized = rows?.trim();
  if (normalized && !parsePrintTitleRows(normalized)) return null;
  mutators.setPageSetup(store, sheet, { printTitleRows: normalized || undefined });
  return getPageSetup(store.getState(), sheet);
}

export function setPrintTitleCols(
  store: SpreadsheetStore,
  sheet: number,
  cols: string | undefined,
): PageSetup | null {
  const normalized = cols?.trim();
  if (normalized && !parsePrintTitleCols(normalized)) return null;
  mutators.setPageSetup(store, sheet, { printTitleCols: normalized || undefined });
  return getPageSetup(store.getState(), sheet);
}

export function setPageOrientation(
  store: SpreadsheetStore,
  sheet: number,
  orientation: PageOrientation,
): PageSetup {
  mutators.setPageSetup(store, sheet, { orientation });
  return getPageSetup(store.getState(), sheet);
}

export function togglePageOrientation(store: SpreadsheetStore, sheet: number): PageSetup {
  const current = getPageSetup(store.getState(), sheet);
  const next: PageOrientation = current.orientation === 'portrait' ? 'landscape' : 'portrait';
  return setPageOrientation(store, sheet, next);
}

export function setPaperSize(
  store: SpreadsheetStore,
  sheet: number,
  paperSize: PaperSize,
): PageSetup {
  mutators.setPageSetup(store, sheet, { paperSize });
  return getPageSetup(store.getState(), sheet);
}

export function setMarginPreset(
  store: SpreadsheetStore,
  sheet: number,
  preset: MarginPreset,
): PageSetup {
  mutators.setPageSetup(store, sheet, { margins: marginPresetValues(preset) });
  return getPageSetup(store.getState(), sheet);
}
