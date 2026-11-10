import type {
  ProtectedRangeInput,
  Sheet,
  SheetProtectionInput,
  WorkbookProtectionInput,
} from '@libraz/formulon';
import { isFormulonError, orFallback } from './result.js';
import type { EngineCapabilities, ProtectedRange, Workbook, WorkbookProtection } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  assertAlive(): void;
};

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandleProtectionMethods {}
}

function internals(handle: unknown): WorkbookHandleInternals {
  return handle as WorkbookHandleInternals;
}

function assertAlive(handle: unknown): void {
  internals(handle).assertAlive();
}

function wb(handle: unknown): Workbook {
  return internals(handle).wb;
}

function ws(handle: unknown, sheet: number): Sheet {
  return wb(handle).sheet(sheet);
}

/** Run an engine call; an engine refusal is a soft `false`, anything else rethrows. */
function attempt(fn: () => void): boolean {
  try {
    fn();
    return true;
  } catch (e) {
    if (!isFormulonError(e)) throw e;
    return false;
  }
}

export abstract class WorkbookHandleProtectionMethods {
  declare readonly capabilities: EngineCapabilities;

  /** Reads the round-trip `<sheetProtection>` flags. Returns `null` under the
   *  stub. Each permission flag is true when users may perform that action;
   *  `enabled` is whether the sheet is protected, and an absent password
   *  string reads empty. */
  getSheetProtection(sheet: number): {
    enabled: boolean;
    algorithmName: string;
    hashValue: string;
    saltValue: string;
    spinCount: number;
    legacyPassword: string;
    sheet: boolean;
    objects: boolean;
    scenarios: boolean;
    formatCells: boolean;
    formatColumns: boolean;
    formatRows: boolean;
    insertColumns: boolean;
    insertRows: boolean;
    insertHyperlinks: boolean;
    deleteColumns: boolean;
    deleteRows: boolean;
    selectLockedCells: boolean;
    selectUnlockedCells: boolean;
    sort: boolean;
    autoFilter: boolean;
    pivotTables: boolean;
  } | null {
    assertAlive(this);
    if (!this.capabilities.sheetProtectionRoundtrip) return null;
    const read = orFallback(null, () => ({ p: ws(this, sheet).getSheetProtection() }));
    if (!read) return null;
    const { p } = read;
    return {
      enabled: p !== null,
      algorithmName: p?.algorithmName ?? '',
      hashValue: p?.hashValue ?? '',
      saltValue: p?.saltValue ?? '',
      spinCount: p?.spinCount ?? 0,
      legacyPassword: p?.legacyPasswordHash ?? '',
      sheet: p !== null,
      objects: p?.allowEditObjects ?? false,
      scenarios: p?.allowEditScenarios ?? false,
      formatCells: p?.allowFormatCells ?? false,
      formatColumns: p?.allowFormatColumns ?? false,
      formatRows: p?.allowFormatRows ?? false,
      insertColumns: p?.allowInsertColumns ?? false,
      insertRows: p?.allowInsertRows ?? false,
      insertHyperlinks: p?.allowInsertHyperlinks ?? false,
      deleteColumns: p?.allowDeleteColumns ?? false,
      deleteRows: p?.allowDeleteRows ?? false,
      selectLockedCells: p?.allowSelectLockedCells ?? true,
      selectUnlockedCells: p?.allowSelectUnlockedCells ?? true,
      sort: p?.allowSort ?? false,
      autoFilter: p?.allowAutoFilter ?? false,
      pivotTables: p?.allowPivotTables ?? false,
    };
  }

  /** Replaces `<sheetProtection>` flags wholesale. Setting `enabled` to
   *  `false` unprotects the sheet. A permission left out takes Excel's
   *  dialog default. Returns `false` (no-op) under stub mode. */
  setSheetProtection(
    sheet: number,
    protection: {
      enabled: boolean;
      legacyPassword?: string;
      algorithmName?: string;
      hashValue?: string;
      saltValue?: string;
      spinCount?: number;
      sheet?: boolean;
      objects?: boolean;
      scenarios?: boolean;
      formatCells?: boolean;
      formatColumns?: boolean;
      formatRows?: boolean;
      insertColumns?: boolean;
      insertRows?: boolean;
      insertHyperlinks?: boolean;
      deleteColumns?: boolean;
      deleteRows?: boolean;
      selectLockedCells?: boolean;
      selectUnlockedCells?: boolean;
      sort?: boolean;
      autoFilter?: boolean;
      pivotTables?: boolean;
    },
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetProtectionRoundtrip) return false;
    if (!protection.enabled) return attempt(() => ws(this, sheet).setSheetProtection(null));
    const input: SheetProtectionInput = {
      algorithmName: protection.algorithmName || null,
      hashValue: protection.hashValue || null,
      saltValue: protection.saltValue || null,
      spinCount: protection.spinCount ?? 0,
      legacyPasswordHash: protection.legacyPassword || null,
      allowEditObjects: protection.objects,
      allowEditScenarios: protection.scenarios,
      allowFormatCells: protection.formatCells,
      allowFormatColumns: protection.formatColumns,
      allowFormatRows: protection.formatRows,
      allowInsertColumns: protection.insertColumns,
      allowInsertRows: protection.insertRows,
      allowInsertHyperlinks: protection.insertHyperlinks,
      allowDeleteColumns: protection.deleteColumns,
      allowDeleteRows: protection.deleteRows,
      allowSelectLockedCells: protection.selectLockedCells,
      allowSelectUnlockedCells: protection.selectUnlockedCells,
      allowSort: protection.sort,
      allowAutoFilter: protection.autoFilter,
      allowPivotTables: protection.pivotTables,
    };
    // An omitted permission takes the engine's dialog default; undefined is not omission.
    const defined = Object.fromEntries(
      Object.entries(input).filter(([, value]) => value !== undefined),
    ) as SheetProtectionInput;
    return attempt(() => ws(this, sheet).setSheetProtection(defined));
  }

  /** Reads `<workbookProtection>`. `null` when the workbook carries none or
   *  under the stub. */
  getWorkbookProtection(): WorkbookProtection | null {
    assertAlive(this);
    if (!this.capabilities.workbookProtectionRoundtrip) return null;
    return orFallback(null, () => wb(this).getWorkbookProtection());
  }

  /** Replaces `<workbookProtection>`; an all-empty record removes the element. */
  setWorkbookProtection(protection: WorkbookProtectionInput | null): boolean {
    assertAlive(this);
    if (!this.capabilities.workbookProtectionRoundtrip) return false;
    return attempt(() => wb(this).setWorkbookProtection(protection));
  }

  /** The `<protectedRange>` records of `sheet` in file order; empty when the
   *  engine lacks the API. */
  getProtectedRanges(sheet: number): ProtectedRange[] {
    assertAlive(this);
    if (!this.capabilities.protectedRanges) return [];
    return [...orFallback([], () => ws(this, sheet).listProtectedRanges())];
  }

  addProtectedRange(sheet: number, range: ProtectedRangeInput): boolean {
    assertAlive(this);
    if (!this.capabilities.protectedRanges) return false;
    return attempt(() => {
      ws(this, sheet).addProtectedRange(range);
    });
  }

  updateProtectedRange(sheet: number, index: number, range: ProtectedRangeInput): boolean {
    assertAlive(this);
    if (!this.capabilities.protectedRanges) return false;
    return attempt(() => ws(this, sheet).updateProtectedRange(index, range));
  }

  removeProtectedRange(sheet: number, index: number): boolean {
    assertAlive(this);
    if (!this.capabilities.protectedRanges) return false;
    return attempt(() => ws(this, sheet).removeProtectedRange(index));
  }
}

export function installProtectionMethods(target: WorkbookHandleCtor): void {
  for (const key of Object.getOwnPropertyNames(WorkbookHandleProtectionMethods.prototype)) {
    if (key === 'constructor') continue;
    const descriptor = Object.getOwnPropertyDescriptor(
      WorkbookHandleProtectionMethods.prototype,
      key,
    );
    if (!descriptor) continue;
    Object.defineProperty(target.prototype, key, descriptor);
  }
}
