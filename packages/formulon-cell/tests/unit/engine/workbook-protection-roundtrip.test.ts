import { describe, expect, it } from 'vitest';
import {
  isWorkbookStructureProtected,
  setWorkbookStructureProtected,
  verifyWorkbookStructurePassword,
  whenSheetPasswordsHashed,
  workbookStructureHasPassword,
} from '../../../src/commands/protection.js';
import { legacySheetPasswordHash } from '../../../src/engine/protection-password.js';
import { hydrateProtectionFromEngine } from '../../../src/engine/protection-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';

const reload = async (wb: WorkbookHandle) => {
  const restored = await WorkbookHandle.loadBytes(wb.save());
  const store = createSpreadsheetStore();
  hydrateProtectionFromEngine(restored, store);
  return { restored, store };
};

describe.skipIf(!canLoadWasm())('workbook structure protection round-trip', () => {
  it('saves a hashed password and verifies it after reload', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      setWorkbookStructureProtected(store, true, { password: 'secret', workbook: wb });
      await whenSheetPasswordsHashed();

      const entry = store.getState().protection.workbookStructure;
      expect(entry?.passwordHash?.algorithmName).toBe('SHA-512');
      expect(entry?.password).toBeUndefined();
      const record = wb.getWorkbookProtection();
      expect(record).toMatchObject({ lockStructure: 1, workbookAlgorithmName: 'SHA-512' });
      expect(record?.workbookPassword).toBe('');

      const { restored, store: next } = await reload(wb);
      try {
        expect(isWorkbookStructureProtected(next.getState())).toBe(true);
        expect(workbookStructureHasPassword(next.getState())).toBe(true);
        await expect(verifyWorkbookStructurePassword(next.getState(), 'secret')).resolves.toBe(
          true,
        );
        await expect(verifyWorkbookStructurePassword(next.getState(), 'wrong')).resolves.toBe(
          false,
        );
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('writes the legacy verifier first, then replaces it with the hash', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      setWorkbookStructureProtected(store, true, { password: 'abc', workbook: wb });
      expect(wb.getWorkbookProtection()?.workbookPassword).toBe(legacySheetPasswordHash('abc'));
      await whenSheetPasswordsHashed();
      expect(wb.getWorkbookProtection()?.workbookPassword).toBe('');
    } finally {
      wb.dispose();
    }
  });

  it('verifies a legacy verifier loaded from a file and an unguarded lock', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const base = wb.getWorkbookProtection();
      if (!base) throw new Error('expected a workbook protection record');
      wb.setWorkbookProtection({
        ...base,
        lockStructure: 1,
        workbookPassword: legacySheetPasswordHash('abc'),
      });
      const { restored, store } = await reload(wb);
      try {
        expect(workbookStructureHasPassword(store.getState())).toBe(true);
        await expect(verifyWorkbookStructurePassword(store.getState(), 'abc')).resolves.toBe(true);
        await expect(verifyWorkbookStructurePassword(store.getState(), 'abd')).resolves.toBe(false);
      } finally {
        restored.dispose();
      }

      const open = createSpreadsheetStore();
      setWorkbookStructureProtected(open, true, { workbook: wb });
      expect(wb.getWorkbookProtection()).toMatchObject({ lockStructure: 1, workbookPassword: '' });
      const again = await reload(wb);
      try {
        expect(isWorkbookStructureProtected(again.store.getState())).toBe(true);
        expect(workbookStructureHasPassword(again.store.getState())).toBe(false);
      } finally {
        again.restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('keeps window and revision locks untouched across protect and unprotect', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const base = wb.getWorkbookProtection();
      if (!base) throw new Error('expected a workbook protection record');
      wb.setWorkbookProtection({
        ...base,
        lockWindows: 1,
        lockRevision: 1,
        revisionsPassword: 'CC3B',
        revisionsAlgorithmName: 'SHA-512',
        revisionsHashValue: 'aGFzaA==',
        revisionsSaltValue: 'c2FsdA==',
        revisionsSpinCount: 7,
      });
      const store = createSpreadsheetStore();
      setWorkbookStructureProtected(store, true, { password: 'pw', workbook: wb });
      await whenSheetPasswordsHashed();
      const locked = { ...wb.getWorkbookProtection() };
      expect(locked).toMatchObject({
        lockStructure: 1,
        lockWindows: 1,
        lockRevision: 1,
        revisionsPassword: 'CC3B',
        revisionsAlgorithmName: 'SHA-512',
        revisionsHashValue: 'aGFzaA==',
        revisionsSaltValue: 'c2FsdA==',
        revisionsSpinCount: 7,
      });

      setWorkbookStructureProtected(store, false, { workbook: wb });
      const unlocked = wb.getWorkbookProtection();
      expect(unlocked).toMatchObject({
        lockStructure: 0,
        lockWindows: 1,
        lockRevision: 1,
        revisionsPassword: 'CC3B',
        revisionsSpinCount: 7,
      });
      // The password still guards the windows lock.
      expect(unlocked?.workbookHashValue).toBe(locked.workbookHashValue);
    } finally {
      wb.dispose();
    }
  });

  it('hydrating a workbook without a structure lock clears a lock left in the store', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      setWorkbookStructureProtected(store, true, { password: 'old' });
      hydrateProtectionFromEngine(wb, store);
      expect(isWorkbookStructureProtected(store.getState())).toBe(false);
    } finally {
      wb.dispose();
    }
  });

  it('drops the password when the last lock is removed', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      setWorkbookStructureProtected(store, true, { password: 'pw', workbook: wb });
      await whenSheetPasswordsHashed();
      setWorkbookStructureProtected(store, false, { workbook: wb });
      expect(wb.getWorkbookProtection()).toMatchObject({
        lockStructure: 0,
        workbookHashValue: '',
        workbookPassword: '',
      });
      const { restored, store: next } = await reload(wb);
      try {
        expect(isWorkbookStructureProtected(next.getState())).toBe(false);
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });
});
