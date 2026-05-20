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
import { type History, recordPageSetupChange } from './history.js';
import { parsePrintAreas, parsePrintTitleCols, parsePrintTitleRows } from './print.js';
import {
  normalizePrintableBounds,
  type HostPrinterDevice,
  type HostPrinterPaperOption,
  type PrinterProfile,
  printerProfilesFromHostDevices,
  resolvePrinterProfileBounds,
} from './printer-profile.js';

export type { HostPrinterDevice, HostPrinterPaperOption, PrinterProfile } from './printer-profile.js';
export {
  normalizePrinterProfile,
  normalizePrinterProfileId,
  normalizePrinterProfiles,
  printerProfilesFromHostDevices,
  resolvePrinterProfileBounds,
} from './printer-profile.js';

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

export type PageSetupPatch = Omit<Partial<PageSetup>, 'margins' | 'printableBounds'> & {
  margins?: Partial<PageSetup['margins']>;
  printableBounds?: Partial<PageSetup['printableBounds']>;
};

export type PageBreakAxis = 'row' | 'col';

const normalizeManualBreaks = (breaks: readonly number[] | undefined): number[] | undefined => {
  const normalized = [
    ...new Set((breaks ?? []).filter((value) => Number.isInteger(value) && value > 0)),
  ].sort((a, b) => a - b);
  return normalized.length ? normalized : undefined;
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
  history: History | null = null,
): PageSetup {
  const current = getPageSetup(store.getState(), sheet);
  const next: Partial<PageSetup> = {
    ...patch,
    margins: patch.margins ? { ...current.margins, ...patch.margins } : undefined,
    printableBounds: patch.printableBounds
      ? { ...(current.printableBounds ?? current.margins), ...patch.printableBounds }
      : undefined,
  };
  if (!patch.margins) delete next.margins;
  if (!patch.printableBounds) delete next.printableBounds;
  recordPageSetupChange(history, store, () => {
    mutators.setPageSetup(store, sheet, next);
  });
  return getPageSetup(store.getState(), sheet);
}

export function setPrintableBounds(
  store: SpreadsheetStore,
  sheet: number,
  bounds: Partial<PageMargins> | undefined,
  history: History | null = null,
): PageSetup {
  recordPageSetupChange(history, store, () => {
    mutators.setPageSetup(store, sheet, { printableBounds: normalizePrintableBounds(bounds) });
  });
  return getPageSetup(store.getState(), sheet);
}

export function clearPrintableBounds(
  store: SpreadsheetStore,
  sheet: number,
  history: History | null = null,
): PageSetup {
  recordPageSetupChange(history, store, () => {
    mutators.setPageSetup(store, sheet, { printableBounds: undefined });
  });
  return getPageSetup(store.getState(), sheet);
}

export function applyPrinterProfileBounds(
  store: SpreadsheetStore,
  sheet: number,
  profiles: readonly PrinterProfile[],
  history: History | null = null,
  printerProfileId?: string,
): PageSetup {
  const setup = getPageSetup(store.getState(), sheet);
  const bounds = resolvePrinterProfileBounds(setup, profiles, printerProfileId);
  return bounds
    ? setPrintableBounds(store, sheet, bounds, history)
    : clearPrintableBounds(store, sheet, history);
}

export function resetPageSetup(
  store: SpreadsheetStore,
  sheet: number,
  history: History | null = null,
): PageSetup {
  recordPageSetupChange(history, store, () => {
    mutators.setPageSetup(store, sheet, null);
  });
  return defaultPageSetup();
}

export function clearPrintTitles(
  store: SpreadsheetStore,
  sheet: number,
  history: History | null = null,
): PageSetup {
  recordPageSetupChange(history, store, () => {
    mutators.setPageSetup(store, sheet, { printTitleRows: undefined, printTitleCols: undefined });
  });
  return getPageSetup(store.getState(), sheet);
}

export function clearPrintArea(
  store: SpreadsheetStore,
  sheet: number,
  history: History | null = null,
): PageSetup {
  recordPageSetupChange(history, store, () => {
    mutators.setPageSetup(store, sheet, { printArea: undefined });
  });
  return getPageSetup(store.getState(), sheet);
}

export function setPrintArea(
  store: SpreadsheetStore,
  sheet: number,
  area: string | undefined,
  history: History | null = null,
): PageSetup | null {
  const normalized = area?.trim();
  if (normalized && !parsePrintAreas(normalized)) return null;
  recordPageSetupChange(history, store, () => {
    mutators.setPageSetup(store, sheet, { printArea: normalized || undefined });
  });
  return getPageSetup(store.getState(), sheet);
}

export function addPrintArea(
  store: SpreadsheetStore,
  sheet: number,
  area: string,
  history: History | null = null,
): PageSetup | null {
  const normalized = area.trim();
  if (!normalized || !parsePrintAreas(normalized)) return null;
  const current = getPageSetup(store.getState(), sheet).printArea?.trim();
  const next = current ? `${current},${normalized}` : normalized;
  return setPrintArea(store, sheet, next, history);
}

export function setPrintTitleRows(
  store: SpreadsheetStore,
  sheet: number,
  rows: string | undefined,
  history: History | null = null,
): PageSetup | null {
  const normalized = rows?.trim();
  if (normalized && !parsePrintTitleRows(normalized)) return null;
  recordPageSetupChange(history, store, () => {
    mutators.setPageSetup(store, sheet, { printTitleRows: normalized || undefined });
  });
  return getPageSetup(store.getState(), sheet);
}

export function setPrintTitleCols(
  store: SpreadsheetStore,
  sheet: number,
  cols: string | undefined,
  history: History | null = null,
): PageSetup | null {
  const normalized = cols?.trim();
  if (normalized && !parsePrintTitleCols(normalized)) return null;
  recordPageSetupChange(history, store, () => {
    mutators.setPageSetup(store, sheet, { printTitleCols: normalized || undefined });
  });
  return getPageSetup(store.getState(), sheet);
}

export function setPageOrientation(
  store: SpreadsheetStore,
  sheet: number,
  orientation: PageOrientation,
  history: History | null = null,
): PageSetup {
  recordPageSetupChange(history, store, () => {
    mutators.setPageSetup(store, sheet, { orientation });
  });
  return getPageSetup(store.getState(), sheet);
}

export function togglePageOrientation(
  store: SpreadsheetStore,
  sheet: number,
  history: History | null = null,
): PageSetup {
  const current = getPageSetup(store.getState(), sheet);
  const next: PageOrientation = current.orientation === 'portrait' ? 'landscape' : 'portrait';
  return setPageOrientation(store, sheet, next, history);
}

export function setPaperSize(
  store: SpreadsheetStore,
  sheet: number,
  paperSize: PaperSize,
  history: History | null = null,
): PageSetup {
  recordPageSetupChange(history, store, () => {
    mutators.setPageSetup(store, sheet, { paperSize });
  });
  return getPageSetup(store.getState(), sheet);
}

export function setMarginPreset(
  store: SpreadsheetStore,
  sheet: number,
  preset: MarginPreset,
  history: History | null = null,
): PageSetup {
  recordPageSetupChange(history, store, () => {
    mutators.setPageSetup(store, sheet, { margins: marginPresetValues(preset) });
  });
  return getPageSetup(store.getState(), sheet);
}

export function setPageScale(
  store: SpreadsheetStore,
  sheet: number,
  scale: number,
  history: History | null = null,
): PageSetup {
  const clamped = Math.max(0.1, Math.min(4, scale));
  recordPageSetupChange(history, store, () => {
    mutators.setPageSetup(store, sheet, {
      scale: clamped,
      fitWidth: undefined,
      fitHeight: undefined,
    });
  });
  return getPageSetup(store.getState(), sheet);
}

export function setFitToPages(
  store: SpreadsheetStore,
  sheet: number,
  axis: 'width' | 'height',
  pages: number | undefined,
  history: History | null = null,
): PageSetup {
  const normalized =
    pages == null || pages <= 0 ? undefined : Math.max(1, Math.min(999, Math.trunc(pages)));
  recordPageSetupChange(history, store, () => {
    mutators.setPageSetup(
      store,
      sheet,
      axis === 'width' ? { fitWidth: normalized } : { fitHeight: normalized },
    );
  });
  return getPageSetup(store.getState(), sheet);
}

export function insertManualPageBreak(
  store: SpreadsheetStore,
  sheet: number,
  axis: PageBreakAxis,
  index: number,
  history: History | null = null,
): PageSetup {
  const normalized = Math.trunc(index);
  if (normalized <= 0) return getPageSetup(store.getState(), sheet);
  const current = getPageSetup(store.getState(), sheet);
  const key = axis === 'row' ? 'manualPageBreakRows' : 'manualPageBreakCols';
  recordPageSetupChange(history, store, () => {
    mutators.setPageSetup(store, sheet, {
      [key]: normalizeManualBreaks([...(current[key] ?? []), normalized]),
    });
  });
  return getPageSetup(store.getState(), sheet);
}

export function removeManualPageBreak(
  store: SpreadsheetStore,
  sheet: number,
  axis: PageBreakAxis,
  index: number,
  history: History | null = null,
): PageSetup {
  const normalized = Math.trunc(index);
  const current = getPageSetup(store.getState(), sheet);
  const key = axis === 'row' ? 'manualPageBreakRows' : 'manualPageBreakCols';
  recordPageSetupChange(history, store, () => {
    mutators.setPageSetup(store, sheet, {
      [key]: normalizeManualBreaks((current[key] ?? []).filter((value) => value !== normalized)),
    });
  });
  return getPageSetup(store.getState(), sheet);
}

export function resetManualPageBreaks(
  store: SpreadsheetStore,
  sheet: number,
  history: History | null = null,
): PageSetup {
  recordPageSetupChange(history, store, () => {
    mutators.setPageSetup(store, sheet, {
      manualPageBreakRows: undefined,
      manualPageBreakCols: undefined,
    });
  });
  return getPageSetup(store.getState(), sheet);
}

export function setPrintGridlines(
  store: SpreadsheetStore,
  sheet: number,
  visible: boolean,
  history: History | null = null,
): PageSetup {
  recordPageSetupChange(history, store, () => {
    mutators.setPageSetup(store, sheet, { showGridlines: visible });
  });
  return getPageSetup(store.getState(), sheet);
}

export function setPrintHeadings(
  store: SpreadsheetStore,
  sheet: number,
  visible: boolean,
  history: History | null = null,
): PageSetup {
  recordPageSetupChange(history, store, () => {
    mutators.setPageSetup(store, sheet, { showHeadings: visible });
  });
  return getPageSetup(store.getState(), sheet);
}
