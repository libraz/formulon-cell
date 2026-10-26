import { newAllowedEditRangeId } from '../store/sheet-settings-mutators.js';
import {
  type AllowedEditRange,
  mutators,
  type SheetProtectionPasswordHash,
  type SheetProtectionPermissions,
  type SpreadsheetStore,
  type WorkbookStructureProtectionState,
} from '../store/store.js';
import { colFromLetters, formatA1Range, MAX_COL, MAX_ROW } from './address.js';
import { legacySheetPasswordHash } from './protection-password.js';
import { parseRangeRef } from './range-resolver.js';
import type { ProtectedRange, Range, WorkbookProtection } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

/**
 * Read `<sheetProtection>` flags off the engine for every sheet in the
 * workbook and mirror them into the protection store. Called once after
 * `loadBytes` so xlsx files arrive with the right toggle state. The
 * legacy password is preserved as the slice's `password` field so a
 * round-trip back through `setSheetProtection` keeps it intact.
 */
export function hydrateProtectionFromEngine(wb: WorkbookHandle, store: SpreadsheetStore): void {
  hydrateWorkbookProtectionFromEngine(wb, store);
  for (let i = 0; i < wb.sheetCount; i += 1) {
    hydrateProtectedRangesFromEngine(wb, store, i);
    if (!wb.capabilities.sheetProtectionRoundtrip) continue;
    const p = wb.getSheetProtection(i);
    if (!p?.enabled) continue;
    mutators.setSheetProtected(store, i, true, {
      ...(p.legacyPassword ? { legacyPasswordHash: p.legacyPassword } : {}),
      ...(sheetProtectionPasswordHashFromEngine(p)
        ? { passwordHash: sheetProtectionPasswordHashFromEngine(p) }
        : {}),
      permissions: sheetProtectionPermissionsFromEngine(p),
    });
  }
}

/** Mirror `<workbookProtection lockStructure>` and its password into the store. */
export function hydrateWorkbookProtectionFromEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
): void {
  if (!wb.capabilities.workbookProtectionRoundtrip) return;
  const p = wb.getWorkbookProtection();
  if (!p?.lockStructure) {
    // The file carries no structure lock, so none may linger from another workbook.
    mutators.setWorkbookStructureProtected(store, false);
    return;
  }
  const passwordHash = sheetProtectionPasswordHashFromEngine({
    algorithmName: p.workbookAlgorithmName,
    hashValue: p.workbookHashValue,
    saltValue: p.workbookSaltValue,
    spinCount: p.workbookSpinCount,
  });
  mutators.setWorkbookStructureProtected(store, true, {
    ...(p.workbookPassword ? { legacyPasswordHash: p.workbookPassword } : {}),
    ...(passwordHash ? { passwordHash } : {}),
  });
}

/**
 * Write the workbook structure lock to the engine. Starts from the engine's
 * current record so `lockWindows`, `lockRevision` and the revisions password
 * pass through untouched. Unlocking keeps the password while windows are
 * still locked, since the one password guards both.
 */
export function flushWorkbookProtectionToEngine(
  wb: WorkbookHandle,
  state: WorkbookStructureProtectionState | null,
): void {
  if (!wb.capabilities.workbookProtectionRoundtrip) return;
  const current = wb.getWorkbookProtection() ?? emptyWorkbookProtection();
  const next: WorkbookProtection = { ...current, lockStructure: state ? 1 : 0 };
  if (state) {
    const hash = state.passwordHash;
    next.workbookPassword = hash
      ? ''
      : (state.legacyPasswordHash ??
        (state.password ? legacySheetPasswordHash(state.password) : ''));
    next.workbookAlgorithmName = hash?.algorithmName ?? '';
    next.workbookHashValue = hash?.hashValue ?? '';
    next.workbookSaltValue = hash?.saltValue ?? '';
    next.workbookSpinCount = hash?.spinCount ?? 0;
  } else if (!current.lockWindows) {
    next.workbookPassword = '';
    next.workbookAlgorithmName = '';
    next.workbookHashValue = '';
    next.workbookSaltValue = '';
    next.workbookSpinCount = 0;
  }
  wb.setWorkbookProtection(next);
}

const emptyWorkbookProtection = (): WorkbookProtection => ({
  lockStructure: 0,
  lockWindows: 0,
  lockRevision: 0,
  workbookPassword: '',
  workbookAlgorithmName: '',
  workbookHashValue: '',
  workbookSaltValue: '',
  workbookSpinCount: 0,
  revisionsPassword: '',
  revisionsAlgorithmName: '',
  revisionsHashValue: '',
  revisionsSaltValue: '',
  revisionsSpinCount: 0,
});

const parseSqrefArea = (token: string, sheet: number): Range | null => {
  const col = /^\$?([A-Za-z]+)(?::\$?([A-Za-z]+))?$/.exec(token);
  if (col) {
    const c0 = colFromLetters(col[1] ?? '');
    const c1 = colFromLetters(col[2] ?? col[1] ?? '');
    if (c0 < 0 || c1 < 0 || c0 > MAX_COL || c1 > MAX_COL) return null;
    return { sheet, r0: 0, c0: Math.min(c0, c1), r1: MAX_ROW, c1: Math.max(c0, c1) };
  }
  const row = /^\$?([1-9][0-9]*)(?::\$?([1-9][0-9]*))?$/.exec(token);
  if (row) {
    const r0 = Number.parseInt(row[1] ?? '', 10) - 1;
    const r1 = Number.parseInt(row[2] ?? row[1] ?? '', 10) - 1;
    if (r0 > MAX_ROW || r1 > MAX_ROW) return null;
    return { sheet, r0: Math.min(r0, r1), c0: 0, r1: Math.max(r0, r1), c1: MAX_COL };
  }
  const ref = parseRangeRef(token);
  return ref ? { sheet, r0: ref.r0, c0: ref.c0, r1: ref.r1, c1: ref.c1 } : null;
};

/** The areas of an A1 range list (`"B2:C3 E5"`); null when any area is unreadable. */
export function areasFromSqref(sqref: string, sheet: number): Range[] | null {
  const tokens = sqref.split(/\s+/).filter(Boolean);
  const areas: Range[] = [];
  for (const token of tokens) {
    const area = parseSqrefArea(token, sheet);
    if (!area) return null;
    areas.push(area);
  }
  return areas.length > 0 ? areas : null;
}

/** An A1 range list for `areas`, in order. */
export const sqrefFromAreas = (areas: readonly Range[]): string =>
  areas.map((area) => formatA1Range(area)).join(' ');

/** Every area of a protected range entry, first area first. */
export const allowedEditRangeAreas = (entry: AllowedEditRange): readonly Range[] => [
  entry.range,
  ...(entry.extraAreas ?? []),
];

const protectedRangeFromEntry = (entry: AllowedEditRange): ProtectedRange => {
  const hash = entry.passwordHash;
  return {
    name: entry.title,
    sqref: sqrefFromAreas(allowedEditRangeAreas(entry)),
    password: hash
      ? ''
      : (entry.legacyPasswordHash ??
        (entry.password ? legacySheetPasswordHash(entry.password) : '')),
    algorithmName: hash?.algorithmName ?? '',
    hashValue: hash?.hashValue ?? '',
    saltValue: hash?.saltValue ?? '',
    spinCount: hash?.spinCount ?? 0,
    securityDescriptor: entry.securityDescriptor ?? '',
  };
};

/** Read the protected ranges of `sheet` into the store. Entries that already
 *  exist under the same title keep their id and typed password. */
export function hydrateProtectedRangesFromEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  if (!wb.capabilities.protectedRanges) return;
  const existing = store.getState().protection.allowedEditRanges;
  const entries: AllowedEditRange[] = [];
  for (const record of wb.getProtectedRanges(sheet)) {
    const areas = areasFromSqref(record.sqref, sheet);
    const [range, ...extraAreas] = areas ?? [];
    if (!range) continue;
    const previous = existing.find((e) => e.range.sheet === sheet && e.title === record.name);
    const passwordHash = sheetProtectionPasswordHashFromEngine(record);
    entries.push({
      id: previous?.id ?? newAllowedEditRangeId(),
      title: record.name,
      range,
      ...(extraAreas.length > 0 ? { extraAreas } : {}),
      ...(previous?.password !== undefined ? { password: previous.password } : {}),
      ...(record.password ? { legacyPasswordHash: record.password } : {}),
      ...(passwordHash ? { passwordHash } : {}),
      ...(record.securityDescriptor ? { securityDescriptor: record.securityDescriptor } : {}),
    });
  }
  if (entries.length === 0 && !existing.some((e) => e.range.sheet === sheet)) return;
  mutators.setAllowedEditRangesForSheet(store, sheet, entries);
}

const sameProtectedRange = (a: AllowedEditRange, b: AllowedEditRange): boolean =>
  JSON.stringify(protectedRangeFromEntry(a)) === JSON.stringify(protectedRangeFromEntry(b));

/**
 * Apply the difference between two lists of allowed edit ranges to the engine:
 * entries gone from `after` are removed, changed ones updated, new ones added.
 * Engine records the store never saw are not touched.
 */
export function syncAllowedEditRangesToEngine(
  wb: WorkbookHandle,
  before: readonly AllowedEditRange[],
  after: readonly AllowedEditRange[],
): void {
  if (!wb.capabilities.protectedRanges) return;
  const indexOf = (entry: AllowedEditRange): number =>
    wb.getProtectedRanges(entry.range.sheet).findIndex((r) => r.name === entry.title);
  const afterById = new Map(after.map((e) => [e.id, e]));
  for (const old of before) {
    if (afterById.has(old.id)) continue;
    const at = indexOf(old);
    if (at >= 0) wb.removeProtectedRange(old.range.sheet, at);
  }
  const beforeById = new Map(before.map((e) => [e.id, e]));
  for (const entry of after) {
    const old = beforeById.get(entry.id);
    if (!old) continue;
    if (sameProtectedRange(old, entry)) continue;
    const at = indexOf(old);
    if (at >= 0) wb.updateProtectedRange(old.range.sheet, at, protectedRangeFromEntry(entry));
  }
  for (const entry of after) {
    if (!beforeById.has(entry.id)) {
      wb.addProtectedRange(entry.range.sheet, protectedRangeFromEntry(entry));
    }
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
  legacyPasswordHash?: string,
): void {
  if (!wb.capabilities.sheetProtectionRoundtrip) return;
  if (on) {
    // A modern hash replaces the 16-bit verifier; typed text becomes a verifier.
    const legacyPassword = passwordHash
      ? ''
      : (legacyPasswordHash ?? (password ? legacySheetPasswordHash(password) : ''));
    wb.setSheetProtection(sheet, {
      enabled: true,
      sheet: true,
      legacyPassword,
      ...(passwordHash ?? {}),
      ...(permissions ?? {}),
    });
  } else {
    wb.setSheetProtection(sheet, { enabled: false });
  }
}
