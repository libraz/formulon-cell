import { mutators, type SpreadsheetStore } from '../store/store.js';
import type { WorkbookHandle } from './workbook-handle.js';

/**
 * Read `<sheetProtection>` flags off the engine for every sheet in the
 * workbook and mirror them into the protection store. Called once after
 * `loadBytes` so xlsx files arrive with the right toggle state. The
 * legacy password is preserved as the slice's `password` field so a
 * round-trip back through `setSheetProtection` keeps it intact.
 */
export function hydrateProtectionFromEngine(wb: WorkbookHandle, store: SpreadsheetStore): void {
  if (!wb.capabilities.sheetProtectionRoundtrip) return;
  for (let i = 0; i < wb.sheetCount; i += 1) {
    const p = wb.getSheetProtection(i);
    if (!p || !p.enabled) continue;
    mutators.setSheetProtected(
      store,
      i,
      true,
      p.legacyPassword ? { password: p.legacyPassword } : undefined,
    );
  }
}

/**
 * Push the JS-side protection toggle to the engine. Mirrors the
 * `<sheetProtection enabled="1" sheet="1" password="…">` shape Excel
 * emits: when `on`, the sheet flag flips on and a legacy password is
 * stored if the caller provided one; when off, the protection block is
 * cleared (`enabled=0`). Future flag-rich UIs can call
 * `wb.setSheetProtection` directly with a richer payload — this helper
 * covers the round-trip path the v1 toggle exposes today.
 */
export function flushProtectionToEngine(
  wb: WorkbookHandle,
  sheet: number,
  on: boolean,
  password?: string,
): void {
  if (!wb.capabilities.sheetProtectionRoundtrip) return;
  if (on) {
    wb.setSheetProtection(sheet, {
      enabled: true,
      sheet: true,
      legacyPassword: password ?? '',
    });
  } else {
    wb.setSheetProtection(sheet, { enabled: false });
  }
}
