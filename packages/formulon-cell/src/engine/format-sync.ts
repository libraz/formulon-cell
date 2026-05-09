import type { CellFormat, SpreadsheetStore } from '../store/store.js';
import type { WorkbookHandle } from './workbook-handle.js';
import { addrKey } from './workbook-handle.js';

/**
 * Seed cell-level comment and hyperlink fields from engine state for `sheet`.
 * Called after a workbook loads (and inside `setWorkbook`) so notes / links
 * stored in an .xlsx survive the round-trip.
 *
 * No-op for unsupported engines (the stub doesn't implement either method).
 *
 * Comments: the engine has no sheet-wide enumerator, so we probe `getComment`
 * for every populated cell. Comments on otherwise-empty cells are not picked
 * up — acceptable for now.
 */
export function hydrateCommentsAndHyperlinksFromEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  if (!wb.capabilities.comments && !wb.capabilities.hyperlinks) return;

  const updates: Array<{ key: string; patch: Partial<CellFormat> }> = [];

  if (wb.capabilities.comments) {
    const cells =
      typeof (wb as WorkbookHandle & { physicalCells?: WorkbookHandle['cells'] }).physicalCells ===
      'function'
        ? wb.physicalCells(sheet)
        : wb.cells(sheet);
    for (const c of cells) {
      const e = wb.getComment(sheet, c.addr.row, c.addr.col);
      if (e && e.text.length > 0) {
        updates.push({ key: addrKey(c.addr), patch: { comment: e.text } });
      }
    }
  }

  if (wb.capabilities.hyperlinks) {
    for (const h of wb.getHyperlinks(sheet)) {
      if (h.target.length === 0) continue;
      updates.push({
        key: addrKey({ sheet, row: h.row, col: h.col }),
        patch: { hyperlink: h.target },
      });
    }
  }

  if (updates.length === 0) return;

  store.setState((s) => {
    const formats = new Map(s.format.formats);
    for (const u of updates) {
      const prev = formats.get(u.key) ?? {};
      formats.set(u.key, { ...prev, ...u.patch });
    }
    return { ...s, format: { formats } };
  });
}

/**
 * Replace the engine's hyperlink set on `sheet` with whatever FormatSlice
 * currently asserts. Each cell with a non-empty `.hyperlink` becomes a single
 * engine hyperlink entry whose `target` carries the URL; `display` and
 * `tooltip` stay default since the UI does not surface them yet. No-op when
 * `capabilities.hyperlinks` is off.
 */
export function syncHyperlinksToEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  if (!wb.capabilities.hyperlinks) return;
  wb.clearHyperlinks(sheet);
  const formats = store.getState().format.formats;
  for (const [key, fmt] of formats) {
    if (!fmt.hyperlink) continue;
    const [sStr, rStr, cStr] = key.split(':');
    if (sStr === undefined || rStr === undefined || cStr === undefined) continue;
    if (Number.parseInt(sStr, 10) !== sheet) continue;
    const row = Number.parseInt(rStr, 10);
    const col = Number.parseInt(cStr, 10);
    wb.addHyperlink(sheet, row, col, fmt.hyperlink);
  }
}
