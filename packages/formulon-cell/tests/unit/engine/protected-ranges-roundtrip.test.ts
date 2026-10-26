import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  addAllowedEditRange,
  allowedEditRangeHasPassword,
  allowedEditRangesForSheet,
  clearAllowedEditRanges,
  isAddrInAllowedEditRange,
  removeAllowedEditRange,
  updateAllowedEditRange,
  verifyAllowedEditRangePassword,
  whenSheetPasswordsHashed,
} from '../../../src/commands/protection.js';
import { deleteRows, insertRows } from '../../../src/commands/structure.js';
import { hydrateProtectionFromEngine } from '../../../src/engine/protection-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const addRange = (...args: Parameters<typeof addAllowedEditRange>): string =>
  addAllowedEditRange(...args) as string;

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';

const reload = async (wb: WorkbookHandle) => {
  const restored = await WorkbookHandle.loadBytes(wb.save());
  const store = createSpreadsheetStore();
  hydrateProtectionFromEngine(restored, store);
  return { restored, store };
};

const rect = (r0: number, c0: number, r1: number, c1: number, sheet = 0) => ({
  sheet,
  r0,
  c0,
  r1,
  c1,
});

describe.skipIf(!canLoadWasm())('protected ranges round-trip', () => {
  it('saves a range with a hashed password and reloads it', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      const id = addRange(store, rect(1, 1, 2, 2), {
        title: 'Inputs',
        password: 'open',
        workbook: wb,
      });
      expect(wb.getProtectedRanges(0)).toMatchObject([{ name: 'Inputs', sqref: 'B2:C3' }]);
      await whenSheetPasswordsHashed();
      const stored = store.getState().protection.allowedEditRanges[0];
      expect(stored?.passwordHash?.algorithmName).toBe('SHA-512');
      expect(stored?.password).toBeUndefined();
      expect(wb.getProtectedRanges(0)[0]).toMatchObject({
        password: '',
        algorithmName: 'SHA-512',
      });

      const { restored, store: next } = await reload(wb);
      try {
        const [entry] = allowedEditRangesForSheet(next.getState(), 0);
        expect(entry).toMatchObject({ title: 'Inputs', range: rect(1, 1, 2, 2) });
        if (!entry) throw new Error('expected a hydrated range');
        expect(allowedEditRangeHasPassword(next.getState(), entry.id)).toBe(true);
        await expect(
          verifyAllowedEditRangePassword(next.getState(), entry.id, 'open'),
        ).resolves.toBe(true);
        await expect(
          verifyAllowedEditRangePassword(next.getState(), entry.id, 'nope'),
        ).resolves.toBe(false);
        // The password gates the range until it is entered this session.
        expect(isAddrInAllowedEditRange(next.getState(), { sheet: 0, row: 2, col: 2 })).toBe(false);
        mutators.setAllowedEditRangeUnlocked(next, entry.id, true);
        expect(isAddrInAllowedEditRange(next.getState(), { sheet: 0, row: 2, col: 2 })).toBe(true);
        expect(isAddrInAllowedEditRange(next.getState(), { sheet: 0, row: 3, col: 3 })).toBe(false);
      } finally {
        restored.dispose();
      }
      expect(id).toBe(store.getState().protection.allowedEditRanges[0]?.id);
    } finally {
      wb.dispose();
    }
  });

  it('adds, modifies and removes ranges in step with the engine', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      const first = addRange(store, rect(0, 0, 0, 0), { title: 'A', workbook: wb });
      const second = addRange(store, rect(4, 4, 5, 5), { title: 'B', workbook: wb });
      expect(wb.getProtectedRanges(0).map((r) => [r.name, r.sqref])).toEqual([
        ['A', 'A1'],
        ['B', 'E5:F6'],
      ]);

      updateAllowedEditRange(
        store,
        first,
        { title: 'Renamed', range: rect(1, 0, 2, 0), extraAreas: [rect(9, 9, 9, 9)] },
        { workbook: wb },
      );
      expect(wb.getProtectedRanges(0).map((r) => [r.name, r.sqref])).toEqual([
        ['Renamed', 'A2:A3 J10'],
        ['B', 'E5:F6'],
      ]);

      const { restored, store: next } = await reload(wb);
      try {
        const [one] = allowedEditRangesForSheet(next.getState(), 0);
        expect(one).toMatchObject({
          title: 'Renamed',
          range: rect(1, 0, 2, 0),
          extraAreas: [rect(9, 9, 9, 9)],
        });
        expect(isAddrInAllowedEditRange(next.getState(), { sheet: 0, row: 9, col: 9 })).toBe(true);
      } finally {
        restored.dispose();
      }

      removeAllowedEditRange(store, second, { workbook: wb });
      expect(wb.getProtectedRanges(0).map((r) => r.name)).toEqual(['Renamed']);
      clearAllowedEditRanges(store, 0, { workbook: wb });
      expect(wb.getProtectedRanges(0)).toEqual([]);
    } finally {
      wb.dispose();
    }
  });

  it('keeps the security descriptor and an unrelated password verbatim through edits', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(
        wb.addProtectedRange(0, {
          name: 'Locked',
          sqref: 'C3',
          password: 'CC3B',
          algorithmName: '',
          hashValue: '',
          saltValue: '',
          spinCount: 0,
          securityDescriptor: 'O:BAG:BAD:(A;;GA;;;WD)',
        }),
      ).toBe(true);
      const store = createSpreadsheetStore();
      hydrateProtectionFromEngine(wb, store);
      const [entry] = store.getState().protection.allowedEditRanges;
      if (!entry) throw new Error('expected a hydrated range');
      expect(entry.legacyPasswordHash).toBe('CC3B');

      updateAllowedEditRange(store, entry.id, { range: rect(2, 2, 3, 3) }, { workbook: wb });
      expect(wb.getProtectedRanges(0)[0]).toMatchObject({
        name: 'Locked',
        sqref: 'C3:D4',
        password: 'CC3B',
        securityDescriptor: 'O:BAG:BAD:(A;;GA;;;WD)',
      });
    } finally {
      wb.dispose();
    }
  });

  it('removes the password with password: null and keeps it when unspecified', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      const id = addRange(store, rect(0, 0, 0, 0), {
        title: 'Guarded',
        password: 'open',
        workbook: wb,
      });
      await whenSheetPasswordsHashed();
      updateAllowedEditRange(store, id, { title: 'Renamed' }, { workbook: wb });
      expect(allowedEditRangeHasPassword(store.getState(), id)).toBe(true);
      expect(wb.getProtectedRanges(0)[0]?.algorithmName).toBe('SHA-512');

      updateAllowedEditRange(store, id, { password: null }, { workbook: wb });
      expect(allowedEditRangeHasPassword(store.getState(), id)).toBe(false);
      expect(wb.getProtectedRanges(0)[0]).toMatchObject({
        name: 'Renamed',
        password: '',
        algorithmName: '',
        hashValue: '',
        spinCount: 0,
      });
    } finally {
      wb.dispose();
    }
  });

  it('follows a row insert and its undo', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      const id = addRange(store, rect(4, 1, 5, 2), { title: 'Moves', workbook: wb });
      const history = new History();
      wb.attachHistory(history);

      expect(insertRows(store, wb, history, 0, 2)).toBe(true);
      const [moved] = allowedEditRangesForSheet(store.getState(), 0);
      expect(moved).toMatchObject({ id, title: 'Moves', range: rect(6, 1, 7, 2) });
      expect(wb.getProtectedRanges(0)[0]?.sqref).toBe('B7:C8');

      history.undo();
      expect(allowedEditRangesForSheet(store.getState(), 0)[0]?.range).toEqual(rect(4, 1, 5, 2));
      expect(wb.getProtectedRanges(0)[0]?.sqref).toBe('B5:C6');
    } finally {
      wb.dispose();
    }
  });

  it('keeps a range through deleting the rows that held it, and its undo', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      addRange(store, rect(2, 1, 3, 2), { title: 'Gone', password: 'pw', workbook: wb });
      await whenSheetPasswordsHashed();
      const before = wb.getProtectedRanges(0);
      const history = new History();
      wb.attachHistory(history);

      expect(deleteRows(store, wb, history, 1, 4)).toBe(true);
      expect(wb.getProtectedRanges(0)).toEqual([]);
      expect(allowedEditRangesForSheet(store.getState(), 0)).toHaveLength(0);

      history.undo();
      expect(wb.getProtectedRanges(0)).toEqual(before);
      expect(allowedEditRangesForSheet(store.getState(), 0)[0]).toMatchObject({
        title: 'Gone',
        range: rect(2, 1, 3, 2),
      });
    } finally {
      wb.dispose();
    }
  });

  it('gives a duplicate title a distinct engine name', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      addAllowedEditRange(store, rect(0, 0, 0, 0), { title: 'Same', workbook: wb });
      addAllowedEditRange(store, rect(1, 1, 1, 1), { title: 'Same', workbook: wb });
      expect(wb.getProtectedRanges(0).map((r) => r.name)).toEqual(['Same', 'Same (2)']);
    } finally {
      wb.dispose();
    }
  });
});
