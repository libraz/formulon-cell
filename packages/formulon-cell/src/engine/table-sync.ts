import type { TableOverlay } from '../commands/format-as-table.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { parseRangeRef } from './range-resolver.js';
import type { Range } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type TableReader = Pick<WorkbookHandle, 'getTables'>;

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

/** Convert loaded spreadsheet ListObjects into renderer table overlays. The engine
 *  currently exposes table metadata as read-only, so these overlays are also
 *  read-only: they give users the full spreadsheet visual affordance for loaded
 *  files while session-created Format-as-Table overlays stay separate. */
export function tableOverlaysFromEngine(wb: TableReader): TableOverlay[] {
  const out: TableOverlay[] = [];
  for (const table of wb.getTables()) {
    const range = tableRange(table.sheetIndex, table.ref);
    if (!range) continue;
    out.push({
      id: tableId(table.displayName || table.name, table.sheetIndex, table.ref),
      source: 'engine',
      range,
      style: 'medium',
      showHeader: true,
      showTotal: false,
      banded: true,
    });
  }
  return out;
}

export function hydrateTableOverlaysFromEngine(wb: TableReader, store: SpreadsheetStore): void {
  mutators.replaceEngineTableOverlays(store, tableOverlaysFromEngine(wb));
}
