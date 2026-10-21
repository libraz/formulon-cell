// @vitest-environment node

import { createHash } from 'node:crypto';
import { describe, expect, it } from 'vitest';
import {
  setProtectedSheet,
  verifySheetProtectionPassword,
  whenSheetPasswordsHashed,
} from '../../../src/commands/protection.js';
import {
  hashSheetProtectionPassword,
  legacySheetPasswordHash,
  verifySheetProtectionPasswordHash,
} from '../../../src/engine/protection-password.js';
import { hydrateProtectionFromEngine } from '../../../src/engine/protection-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

/** [MS-OFFCRYPTO] 2.3.7.1 transcribed step by step, as an independent reference. */
const referenceVerifier = (password: string): string => {
  const passwordArray = [password.length, ...Array.from(password, (c) => c.charCodeAt(0))];
  let verifier = 0;
  for (let i = passwordArray.length - 1; i >= 0; i -= 1) {
    const intermediate = (verifier & 0x4000) === 0 ? 0 : 1;
    verifier = (verifier * 2) & 0x7fff;
    verifier |= intermediate;
    verifier ^= passwordArray[i] as number;
  }
  verifier ^= 0xce4b;
  return verifier.toString(16).toUpperCase().padStart(4, '0');
};

describe('legacy sheet-protection verifier', () => {
  it('matches hand-derived values', () => {
    // "" folds only the length 0: 0 ^ 0xCE4B.
    expect(legacySheetPasswordHash('')).toBe('CE4B');
    // "a": 0x61, then rotate(0x61) = 0xC2 ^ length 1 = 0xC3, then ^ 0xCE4B.
    expect(legacySheetPasswordHash('a')).toBe('CE88');
  });

  it('agrees with the step-by-step reference', () => {
    for (const password of ['password', 'test', 'abc', 'secret', 'Password123', 'x'.repeat(40)]) {
      expect(legacySheetPasswordHash(password)).toBe(referenceVerifier(password));
    }
    expect(legacySheetPasswordHash('password')).toBe('83AF');
  });
});

describe('modern sheet-protection hash', () => {
  it('is SHA-512 over salt + UTF-16LE, spun 100000 times', async () => {
    const hash = await hashSheetProtectionPassword('hunter2');
    expect(hash).toMatchObject({ algorithmName: 'SHA-512', spinCount: 100000 });
    if (!hash) throw new Error('no hash');
    const salt = Buffer.from(hash.saltValue, 'base64');
    expect(salt.length).toBe(16);

    let h = createHash('sha512')
      .update(Buffer.concat([salt, Buffer.from('hunter2', 'utf16le')]))
      .digest();
    for (let i = 0; i < 100000; i += 1) {
      const counter = Buffer.alloc(4);
      counter.writeUInt32LE(i);
      h = createHash('sha512')
        .update(Buffer.concat([h, counter]))
        .digest();
    }
    expect(hash.hashValue).toBe(h.toString('base64'));
    expect(await verifySheetProtectionPasswordHash('hunter2', hash)).toBe(true);
    expect(await verifySheetProtectionPasswordHash('hunter3', hash)).toBe(false);
  });

  it('uses a fresh salt each time', async () => {
    const a = await hashSheetProtectionPassword('same');
    const b = await hashSheetProtectionPassword('same');
    expect(a?.saltValue).not.toBe(b?.saltValue);
  });
});

describe('protecting a sheet through the engine', () => {
  it('writes a verifier at once and the modern hash once it is computed', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      setProtectedSheet(store, 0, true, { password: 'password', workbook: wb });
      expect(wb.getSheetProtection(0)).toMatchObject({
        enabled: true,
        legacyPassword: '83AF',
        algorithmName: '',
      });

      await whenSheetPasswordsHashed();
      expect(wb.getSheetProtection(0)).toMatchObject({
        legacyPassword: '',
        algorithmName: 'SHA-512',
        spinCount: 100000,
      });

      const restored = await WorkbookHandle.loadBytes(wb.save());
      try {
        const next = createSpreadsheetStore();
        hydrateProtectionFromEngine(restored, next);
        expect(next.getState().protection.protectedSheets.get(0)?.password).toBeUndefined();
        expect(await verifySheetProtectionPassword(next.getState(), 0, 'password')).toBe(true);
        expect(await verifySheetProtectionPassword(next.getState(), 0, 'nope')).toBe(false);
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('unprotects a loaded file that carries only the 16-bit verifier', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.setSheetProtection(0, { enabled: true, legacyPassword: '83AF' })).toBe(true);
      const restored = await WorkbookHandle.loadBytes(wb.save());
      try {
        const store = createSpreadsheetStore();
        hydrateProtectionFromEngine(restored, store);
        const entry = store.getState().protection.protectedSheets.get(0);
        expect(entry?.legacyPasswordHash).toBe('83AF');
        expect(await verifySheetProtectionPassword(store.getState(), 0, 'password')).toBe(true);
        expect(await verifySheetProtectionPassword(store.getState(), 0, 'Password')).toBe(false);
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });
});
