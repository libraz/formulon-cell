import { hydrateAutoFilterFromEngine } from '../engine/auto-filter-sync.js';
import { hydrateCellFormatsFromEngine } from '../engine/cell-format-sync.js';
import { hydrateConditionalRulesFromEngine } from '../engine/cf-sync.js';
import { hydrateCommentsAndHyperlinksFromEngine } from '../engine/format-sync.js';
import { hydrateLayoutFromEngine } from '../engine/layout-sync.js';
import { hydrateMergesFromEngine } from '../engine/merges-sync.js';
import { summarizePassthroughs, summarizeTables } from '../engine/passthrough-sync.js';
import { hydratePageSetupFromEngine } from '../engine/print-sync.js';
import { hydrateProtectionFromEngine } from '../engine/protection-sync.js';
import { hydrateTableOverlaysFromEngine } from '../engine/table-sync.js';
import { hydrateValidationsFromEngine } from '../engine/validation-sync.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { type ViewportOptions, validateViewportOptions } from '../interact/navigation-policy.js';
import { createSpreadsheetStore, mutators, type SpreadsheetStore } from '../store/store.js';

export function hydrateActiveSheetFromEngine(wb: WorkbookHandle, store: SpreadsheetStore): void {
  const hydrate = (): void => {
    const sheet = store.getState().data.sheetIndex;
    mutators.replaceCells(store, wb.cells(sheet));
    hydrateLayoutFromEngine(wb, store, sheet);
    hydrateCommentsAndHyperlinksFromEngine(wb, store, sheet);
    hydrateMergesFromEngine(wb, store, sheet);
    hydrateValidationsFromEngine(wb, store, sheet);
    hydrateCellFormatsFromEngine(wb, store, sheet);
    hydrateConditionalRulesFromEngine(wb, store, sheet);
    hydrateAutoFilterFromEngine(wb, store, sheet);
    hydratePageSetupFromEngine(wb, store, sheet);
  };
  if (typeof wb.withEngineSyncMuted === 'function') wb.withEngineSyncMuted(hydrate);
  else hydrate();
}

/** Validate viewport options against `wb`, including its merges on the target sheet. */
export function validateViewportAgainstWorkbook(
  options: ViewportOptions | undefined,
  wb: WorkbookHandle,
  sheet?: number,
): void {
  validateViewportOptions(options, wb);
  const preview = createSpreadsheetStore();
  if (sheet !== undefined) mutators.setSheetIndex(preview, sheet);
  hydrateActiveSheetFromEngine(wb, preview);
  validateViewportOptions(options, wb, preview.getState());
}

export function hydrateWorkbookMetadataFromEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
): void {
  hydrateProtectionFromEngine(wb, store);
  hydrateTableOverlaysFromEngine(wb, store);
}

export function dispatchWorkbookObjectSummaries(host: HTMLElement, wb: WorkbookHandle): void {
  host.dispatchEvent(new CustomEvent('fc:passthroughs', { detail: summarizePassthroughs(wb) }));
  host.dispatchEvent(new CustomEvent('fc:tables', { detail: summarizeTables(wb) }));
}
