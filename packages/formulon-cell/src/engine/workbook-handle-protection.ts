import type { EngineCapabilities, Workbook } from './types.js';
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

export abstract class WorkbookHandleProtectionMethods {
  declare readonly capabilities: EngineCapabilities;

  /** Reads the round-trip `<sheetProtection>` flags. Returns `null` when
   *  the engine doesn't expose `getSheetProtection`. The booleans are
   *  reported as JS booleans (the engine wires them as 0/1 numbers); the
   *  `enabled` flag denotes whether the protection block is emitted on
   *  save. */
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
    const r = wb(this).getSheetProtection(sheet);
    if (!r.status.ok) return null;
    const p = r.protection;
    return {
      enabled: p.enabled !== 0,
      algorithmName: p.algorithmName,
      hashValue: p.hashValue,
      saltValue: p.saltValue,
      spinCount: p.spinCount,
      legacyPassword: p.legacyPassword,
      sheet: p.sheet !== 0,
      objects: p.objects !== 0,
      scenarios: p.scenarios !== 0,
      formatCells: p.formatCells !== 0,
      formatColumns: p.formatColumns !== 0,
      formatRows: p.formatRows !== 0,
      insertColumns: p.insertColumns !== 0,
      insertRows: p.insertRows !== 0,
      insertHyperlinks: p.insertHyperlinks !== 0,
      deleteColumns: p.deleteColumns !== 0,
      deleteRows: p.deleteRows !== 0,
      selectLockedCells: p.selectLockedCells !== 0,
      selectUnlockedCells: p.selectUnlockedCells !== 0,
      sort: p.sort !== 0,
      autoFilter: p.autoFilter !== 0,
      pivotTables: p.pivotTables !== 0,
    };
  }

  /** Replaces `<sheetProtection>` flags wholesale. Setting `enabled` to
   *  `false` clears the protection block on save. Returns `false` (no-op)
   *  under stub mode and older engine package builds. */
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
    const b = (v: boolean | undefined): number => (v ? 1 : 0);
    const s = wb(this).setSheetProtection(sheet, {
      enabled: b(protection.enabled),
      algorithmName: protection.algorithmName ?? '',
      hashValue: protection.hashValue ?? '',
      saltValue: protection.saltValue ?? '',
      spinCount: protection.spinCount ?? 0,
      legacyPassword: protection.legacyPassword ?? '',
      sheet: b(protection.sheet ?? true),
      objects: b(protection.objects),
      scenarios: b(protection.scenarios),
      formatCells: b(protection.formatCells),
      formatColumns: b(protection.formatColumns),
      formatRows: b(protection.formatRows),
      insertColumns: b(protection.insertColumns),
      insertRows: b(protection.insertRows),
      insertHyperlinks: b(protection.insertHyperlinks),
      deleteColumns: b(protection.deleteColumns),
      deleteRows: b(protection.deleteRows),
      selectLockedCells: b(protection.selectLockedCells),
      selectUnlockedCells: b(protection.selectUnlockedCells),
      sort: b(protection.sort),
      autoFilter: b(protection.autoFilter),
      pivotTables: b(protection.pivotTables),
    });
    return s.ok;
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
