import type { TableOverlay } from '../commands/format-as-table.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { parseRangeRef } from './range-resolver.js';
import type { Range } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type TableReader = Pick<WorkbookHandle, 'getTables'> &
  Partial<Pick<WorkbookHandle, 'withEngineSyncMuted'>>;

function tableId(name: string, sheet: number, ref: string): string {
  const safe = `${name || 'Table'}-${sheet}-${ref}`.replace(/[^A-Za-z0-9_-]+/g, '-');
  return `engine-table-${safe}`;
}

function tableRange(sheet: number, ref: string): Range | null {
  const parsed = parseRangeRef(ref);
  if (!parsed) return null;
  return {
    sheet,
    r0: parsed.r0,
    c0: parsed.c0,
    r1: parsed.r1,
    c1: parsed.c1,
  };
}

/** Convert the engine's worksheet tables into engine-backed overlays. The
 *  engine reports no table style, so the look is the default one. */
export function tableOverlaysFromEngine(wb: TableReader): TableOverlay[] {
  const out: TableOverlay[] = [];
  for (const table of wb.getTables()) {
    const range = tableRange(table.sheetIndex, table.ref);
    if (!range) continue;
    out.push({
      id: tableId(table.displayName || table.name, table.sheetIndex, table.ref),
      source: 'engine',
      name: table.name,
      range,
      style: 'medium',
      showHeader: true,
      showTotal: false,
      banded: true,
    });
  }
  return out;
}

/** Load the engine's tables into the store without echoing them back; known
 *  tables keep the look the user gave them. */
export function hydrateTableOverlaysFromEngine(wb: TableReader, store: SpreadsheetStore): void {
  const hydrate = (): void =>
    mutators.replaceEngineTableOverlays(store, tableOverlaysFromEngine(wb));
  if (typeof wb.withEngineSyncMuted === 'function') wb.withEngineSyncMuted(hydrate);
  else hydrate();
}
