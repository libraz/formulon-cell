import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  type AllowedRangeUnlockUi,
  allowedEditRangesForSheet,
  isWorkbookStructureProtected,
  registerAllowedRangeUnlockUi,
  requestAllowedRangeUnlock,
  setProtectedSheet,
  setWorkbookStructureProtected,
} from '../../../src/commands/protection.js';
import { legacySheetPasswordHash } from '../../../src/engine/protection-password.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

const cellA1 = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };

describe('setWorkbook starts from the new workbook protection', () => {
  let sheet: MountedStubSheet;
  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
  });
  afterEach(() => sheet.dispose());

  it('drops the old unlocks, ranges and structure lock; the new password is the one that counts', async () => {
    const { store } = sheet.instance;
    mutators.addAllowedEditRange(store, {
      title: 'R',
      range: cellA1,
      legacyPasswordHash: legacySheetPasswordHash('passA'),
    });
    setProtectedSheet(store, 0, true);
    setWorkbookStructureProtected(store, true);
    const answersA = ['passA'];
    registerAllowedRangeUnlockUi(store, {
      askPassword: async () => answersA.shift() ?? null,
      reportIncorrect: async () => {},
    });
    await requestAllowedRangeUnlock(store, [cellA1]);
    const oldId = allowedEditRangesForSheet(store.getState(), 0)[0]?.id ?? '';
    expect(store.getState().protection.unlockedAllowedRanges?.has(oldId)).toBe(true);

    const next = await WorkbookHandle.createDefault();
    next.setSheetProtection(0, { enabled: true });
    next.addProtectedRange(0, {
      name: 'R',
      sqref: 'A1',
      password: legacySheetPasswordHash('passB'),
      algorithmName: '',
      hashValue: '',
      saltValue: '',
      spinCount: 0,
      securityDescriptor: '',
    });
    await sheet.instance.setWorkbook(next);

    expect(isWorkbookStructureProtected(store.getState())).toBe(false);
    expect(store.getState().protection.unlockedAllowedRanges?.size ?? 0).toBe(0);
    const [entry] = allowedEditRangesForSheet(store.getState(), 0);
    expect(entry?.id).not.toBe(oldId);
    expect(entry?.password).toBeUndefined();

    const ui: AllowedRangeUnlockUi & {
      askPassword: ReturnType<typeof vi.fn>;
      reportIncorrect: ReturnType<typeof vi.fn>;
    } = {
      askPassword: vi.fn(async () => (ui.askPassword.mock.calls.length === 1 ? 'passA' : 'passB')),
      reportIncorrect: vi.fn(async () => {}),
    };
    registerAllowedRangeUnlockUi(store, ui);
    await expect(requestAllowedRangeUnlock(store, [cellA1])).resolves.toBe(false);
    expect(ui.reportIncorrect).toHaveBeenCalledTimes(1);
    await expect(requestAllowedRangeUnlock(store, [cellA1])).resolves.toBe(true);
  });
});
