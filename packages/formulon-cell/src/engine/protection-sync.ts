import {
  mutators,
  type SheetProtectionPasswordHash,
  type SheetProtectionPermissions,
  type SpreadsheetStore,
} from '../store/store.js';
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
    if (!p?.enabled) continue;
    mutators.setSheetProtected(store, i, true, {
      ...(p.legacyPassword ? { password: p.legacyPassword } : {}),
      ...(sheetProtectionPasswordHashFromEngine(p)
        ? { passwordHash: sheetProtectionPasswordHashFromEngine(p) }
        : {}),
      permissions: sheetProtectionPermissionsFromEngine(p),
    });
  }
}

const sheetProtectionPasswordHashFromEngine = (p: {
  algorithmName: string;
  hashValue: string;
  saltValue: string;
  spinCount: number;
}): SheetProtectionPasswordHash | undefined => {
  if (!p.algorithmName && !p.hashValue && !p.saltValue && p.spinCount === 0) return undefined;
  return {
    algorithmName: p.algorithmName,
    hashValue: p.hashValue,
    saltValue: p.saltValue,
    spinCount: p.spinCount,
  };
};

const sheetProtectionPermissionsFromEngine = (p: {
  objects: boolean;
  scenarios: boolean;
  selectLockedCells: boolean;
  selectUnlockedCells: boolean;
  formatCells: boolean;
  formatColumns: boolean;
  formatRows: boolean;
  insertColumns: boolean;
  insertRows: boolean;
  insertHyperlinks: boolean;
  deleteColumns: boolean;
  deleteRows: boolean;
  sort: boolean;
  autoFilter: boolean;
  pivotTables: boolean;
}): SheetProtectionPermissions => ({
  objects: p.objects,
  scenarios: p.scenarios,
  selectLockedCells: p.selectLockedCells,
  selectUnlockedCells: p.selectUnlockedCells,
  formatCells: p.formatCells,
  formatColumns: p.formatColumns,
  formatRows: p.formatRows,
  insertColumns: p.insertColumns,
  insertRows: p.insertRows,
  insertHyperlinks: p.insertHyperlinks,
  deleteColumns: p.deleteColumns,
  deleteRows: p.deleteRows,
  sort: p.sort,
  autoFilter: p.autoFilter,
  pivotTables: p.pivotTables,
});

/**
 * Push the JS-side protection toggle to the engine. Mirrors the
 * `<sheetProtection enabled="1" sheet="1" password="…">` shape desktop spreadsheets
 * emit: when `on`, the sheet flag flips on and the caller-supplied legacy
 * password plus permission flags are stored; when off, the protection block
 * is cleared (`enabled=0`).
 */
export function flushProtectionToEngine(
  wb: WorkbookHandle,
  sheet: number,
  on: boolean,
  password?: string,
  permissions?: SheetProtectionPermissions,
  passwordHash?: SheetProtectionPasswordHash,
): void {
  if (!wb.capabilities.sheetProtectionRoundtrip) return;
  if (on) {
    wb.setSheetProtection(sheet, {
      enabled: true,
      sheet: true,
      legacyPassword: password ?? '',
      ...(passwordHash ?? {}),
      ...(permissions ?? {}),
    });
  } else {
    wb.setSheetProtection(sheet, { enabled: false });
  }
}
