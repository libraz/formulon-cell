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
      patch: { enabled: boolean; legacyPassword?: string },
    ): boolean => {
      if (opts.capabilityOff) return false;
      if (patch.enabled) {
        storage.set(sheet, {
          enabled: true,
          legacyPassword: patch.legacyPassword ?? '',
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
      initial: new Map([[1, enabledRecord('secret')]]),
    });
    hydrateProtectionFromEngine(wb, store);
    const map = store.getState().protection.protectedSheets;
    expect(map.has(0)).toBe(false);
    expect(map.has(1)).toBe(true);
    expect(map.get(1)?.password).toBe('secret');
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
    flushProtectionToEngine(wb, 0, true, 'pw');
    const stored = storage.get(0);
    expect(stored?.enabled).toBe(true);
    expect(stored?.legacyPassword).toBe('pw');
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
