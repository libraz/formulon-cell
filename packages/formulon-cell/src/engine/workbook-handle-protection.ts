import type {
  ProtectedRangeInput,
  SheetProtectionInput,
  WorkbookProtectionInput,
} from '@libraz/formulon';
import { attemptWrite, orFallback } from './result.js';
import type { EngineCapabilities, ProtectedRange, WorkbookProtection } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';
import { assertAlive, wb, ws } from './workbook-handle-internals.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandleProtectionMethods {}
}

/** The engine refuses an empty hash, salt, algorithm or descriptor string; an empty one is the absent one. */
const blankToNull = <T extends object>(input: T, keys: readonly (keyof T)[]): T => {
  const out = { ...input };
  for (const key of keys) if ((out[key] as unknown) === '') out[key] = null as T[keyof T];
  return out;
};

const WORKBOOK_PROTECTION_TEXT_KEYS = [
  'workbookLegacyPasswordHash',
  'workbookAlgorithmName',
  'workbookHashValue',
  'workbookSaltValue',
  'revisionsLegacyPasswordHash',
  'revisionsAlgorithmName',
  'revisionsHashValue',
  'revisionsSaltValue',
] as const satisfies readonly (keyof WorkbookProtectionInput)[];

const PROTECTED_RANGE_TEXT_KEYS = [
  'legacyPasswordHash',
  'algorithmName',
  'hashValue',
  'saltValue',
  'securityDescriptor',
] as const satisfies readonly (keyof ProtectedRangeInput)[];

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
    if (!protection.enabled) return attemptWrite(() => ws(this, sheet).setSheetProtection(null));
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
    return attemptWrite(() => ws(this, sheet).setSheetProtection(defined));
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
    const input = protection && blankToNull(protection, WORKBOOK_PROTECTION_TEXT_KEYS);
    return attemptWrite(() => wb(this).setWorkbookProtection(input));
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
    return attemptWrite(() => {
      ws(this, sheet).addProtectedRange(blankToNull(range, PROTECTED_RANGE_TEXT_KEYS));
    });
  }

  updateProtectedRange(sheet: number, index: number, range: ProtectedRangeInput): boolean {
    assertAlive(this);
    if (!this.capabilities.protectedRanges) return false;
    const input = blankToNull(range, PROTECTED_RANGE_TEXT_KEYS);
    return attemptWrite(() => ws(this, sheet).updateProtectedRange(index, input));
  }

  removeProtectedRange(sheet: number, index: number): boolean {
    assertAlive(this);
    if (!this.capabilities.protectedRanges) return false;
    return attemptWrite(() => ws(this, sheet).removeProtectedRange(index));
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
