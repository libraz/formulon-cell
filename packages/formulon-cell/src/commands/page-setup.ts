import {
  defaultPageSetup,
  getPageSetup,
  mutators,
  type PageSetup,
  type SpreadsheetStore,
  type State,
} from '../store/store.js';
import { parsePrintTitleCols, parsePrintTitleRows } from './print.js';

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
