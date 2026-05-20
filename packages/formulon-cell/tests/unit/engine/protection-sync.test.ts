import { describe, expect, it } from 'vitest';
import {
  flushProtectionToEngine,
  hydrateProtectionFromEngine,
} from '../../../src/engine/protection-sync.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

type EngineProtection = ReturnType<WorkbookHandle['getSheetProtection']>;

const makeWb = (opts: {
  sheetCount: number;
  initial: Map<number, NonNullable<EngineProtection>>;
  capabilityOff?: boolean;
}): { wb: WorkbookHandle; storage: Map<number, NonNullable<EngineProtection>> } => {
  const storage = new Map(opts.initial);
  const wb = {
    capabilities: { sheetProtectionRoundtrip: !opts.capabilityOff },
    sheetCount: opts.sheetCount,
    getSheetProtection: (sheet: number): EngineProtection => {
      if (opts.capabilityOff) return null;
      return storage.get(sheet) ?? null;
    },
    setSheetProtection: (
      sheet: number,
      patch: {
        enabled: boolean;
        algorithmName?: string;
        hashValue?: string;
        saltValue?: string;
        spinCount?: number;
        legacyPassword?: string;
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
    ): boolean => {
      if (opts.capabilityOff) return false;
      if (patch.enabled) {
        storage.set(sheet, {
          enabled: true,
          legacyPassword: patch.legacyPassword ?? '',
          algorithmName: patch.algorithmName ?? '',
          hashValue: patch.hashValue ?? '',
          saltValue: patch.saltValue ?? '',
          spinCount: patch.spinCount ?? 0,
          sheet: true,
          objects: patch.objects ?? false,
          scenarios: patch.scenarios ?? false,
          formatCells: patch.formatCells ?? false,
          formatColumns: patch.formatColumns ?? false,
          formatRows: patch.formatRows ?? false,
          insertColumns: patch.insertColumns ?? false,
          insertRows: patch.insertRows ?? false,
          insertHyperlinks: patch.insertHyperlinks ?? false,
          deleteColumns: patch.deleteColumns ?? false,
          deleteRows: patch.deleteRows ?? false,
          selectLockedCells: patch.selectLockedCells ?? false,
          selectUnlockedCells: patch.selectUnlockedCells ?? false,
          sort: patch.sort ?? false,
          autoFilter: patch.autoFilter ?? false,
          pivotTables: patch.pivotTables ?? false,
        });
      } else {
        storage.delete(sheet);
      }
      return true;
    },
  } as unknown as WorkbookHandle;
  return { wb, storage };
};

const enabledRecord = (password = ''): NonNullable<EngineProtection> => ({
  enabled: true,
  legacyPassword: password,
  algorithmName: '',
  hashValue: '',
  saltValue: '',
  spinCount: 0,
  sheet: true,
  objects: false,
  scenarios: false,
  formatCells: false,
  formatColumns: false,
  formatRows: false,
  insertColumns: false,
  insertRows: false,
  insertHyperlinks: false,
  deleteColumns: false,
  deleteRows: false,
  selectLockedCells: false,
  selectUnlockedCells: false,
  sort: false,
  autoFilter: false,
  pivotTables: false,
});

describe('hydrateProtectionFromEngine', () => {
  it('mirrors enabled sheets into the store', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeWb({
      sheetCount: 3,
      initial: new Map([
        [
          1,
          {
            ...enabledRecord('secret'),
            algorithmName: 'SHA-512',
            hashValue: 'hash',
            saltValue: 'salt',
            spinCount: 100000,
          },
        ],
      ]),
    });
    hydrateProtectionFromEngine(wb, store);
    const map = store.getState().protection.protectedSheets;
    expect(map.has(0)).toBe(false);
    expect(map.has(1)).toBe(true);
    expect(map.get(1)?.password).toBe('secret');
    expect(map.get(1)?.passwordHash).toEqual({
      algorithmName: 'SHA-512',
      hashValue: 'hash',
      saltValue: 'salt',
      spinCount: 100000,
    });
    expect(map.get(1)?.permissions).toMatchObject({
      objects: false,
      scenarios: false,
      formatCells: false,
      formatColumns: false,
      formatRows: false,
      insertColumns: false,
      insertRows: false,
      insertHyperlinks: false,
      deleteColumns: false,
      deleteRows: false,
      selectLockedCells: false,
      selectUnlockedCells: false,
      sort: false,
      autoFilter: false,
      pivotTables: false,
    });
    expect(map.has(2)).toBe(false);
  });

  it('skips sheets with enabled=false', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeWb({
      sheetCount: 1,
      initial: new Map([[0, { ...enabledRecord(), enabled: false }]]),
    });
    hydrateProtectionFromEngine(wb, store);
    expect(store.getState().protection.protectedSheets.size).toBe(0);
  });

  it('is a no-op when the engine lacks the capability', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeWb({ sheetCount: 1, initial: new Map(), capabilityOff: true });
    hydrateProtectionFromEngine(wb, store);
    expect(store.getState().protection.protectedSheets.size).toBe(0);
  });
});

describe('flushProtectionToEngine', () => {
  it('writes a protection block when toggling on', () => {
    const { wb, storage } = makeWb({ sheetCount: 1, initial: new Map() });
    flushProtectionToEngine(
      wb,
      0,
      true,
      'pw',
      {
        formatCells: true,
        formatColumns: true,
        insertColumns: true,
        insertRows: true,
        insertHyperlinks: true,
        deleteRows: true,
        selectLockedCells: true,
        selectUnlockedCells: true,
        sort: true,
        autoFilter: true,
        pivotTables: true,
        objects: true,
      },
      {
        algorithmName: 'SHA-512',
        hashValue: 'hash',
        saltValue: 'salt',
        spinCount: 100000,
      },
    );
    const stored = storage.get(0);
    expect(stored?.enabled).toBe(true);
    expect(stored?.legacyPassword).toBe('pw');
    expect(stored?.algorithmName).toBe('SHA-512');
    expect(stored?.hashValue).toBe('hash');
    expect(stored?.saltValue).toBe('salt');
    expect(stored?.spinCount).toBe(100000);
    expect(stored?.formatCells).toBe(true);
    expect(stored?.formatColumns).toBe(true);
    expect(stored?.insertColumns).toBe(true);
    expect(stored?.insertRows).toBe(true);
    expect(stored?.insertHyperlinks).toBe(true);
    expect(stored?.deleteRows).toBe(true);
    expect(stored?.selectLockedCells).toBe(true);
    expect(stored?.selectUnlockedCells).toBe(true);
    expect(stored?.sort).toBe(true);
    expect(stored?.autoFilter).toBe(true);
    expect(stored?.pivotTables).toBe(true);
    expect(stored?.objects).toBe(true);
  });

  it('clears the protection block when toggling off', () => {
    const { wb, storage } = makeWb({
      sheetCount: 1,
      initial: new Map([[0, enabledRecord('pw')]]),
    });
    flushProtectionToEngine(wb, 0, false);
    expect(storage.has(0)).toBe(false);
  });

  it('is a no-op when the engine lacks the capability', () => {
    const { wb, storage } = makeWb({
      sheetCount: 1,
      initial: new Map([[0, enabledRecord()]]),
      capabilityOff: true,
    });
    flushProtectionToEngine(wb, 0, false);
    // capabilityOff masked the read, but the underlying storage is untouched.
    expect(storage.has(0)).toBe(true);
  });
});
