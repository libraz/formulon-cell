import { describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  type AllowedRangeUnlockUi,
  addAllowedEditRange,
  allowedEditRangesForSheet,
  clearAllowedEditRanges,
  deferForAllowedRangeUnlock,
  recordProtectionChange,
  registerAllowedRangeUnlockUi,
  removeAllowedEditRange,
  requestAllowedRangeUnlock,
  setProtectedSheet,
  updateAllowedEditRange,
  whenSheetPasswordsHashed,
} from '../../../src/commands/protection.js';
import { executeRibbonProtectionAction } from '../../../src/commands/ribbon-protection.js';
import { legacySheetPasswordHash } from '../../../src/engine/protection-password.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings/en.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const area = (r0: number, c0: number, r1: number, c1: number): Range => ({
  sheet: 0,
  r0,
  c0,
  r1,
  c1,
});
const settle = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 0));
const fakeUi = (answers: (string | null)[]) => {
  const ui: AllowedRangeUnlockUi & {
    askPassword: ReturnType<typeof vi.fn>;
    reportIncorrect: ReturnType<typeof vi.fn>;
  } = {
    askPassword: vi.fn(async () => {
      await settle();
      return answers.shift() ?? null;
    }),
    reportIncorrect: vi.fn(async () => {}),
  };
  return ui;
};

describe('Allow Edit Ranges are unavailable on a protected sheet', () => {
  const protectedStore = () => {
    const store = createSpreadsheetStore();
    const id = addAllowedEditRange(store, area(0, 0, 0, 0), { title: 'Keep' });
    setProtectedSheet(store, 0, true);
    return { store, id: id as string };
  };

  it('refuses add, modify, remove and clear in the command layer', () => {
    const { store, id } = protectedStore();
    expect(addAllowedEditRange(store, area(2, 2, 2, 2), { title: 'New' })).toBeNull();
    expect(updateAllowedEditRange(store, id, { title: 'Renamed' })).toBe(false);
    removeAllowedEditRange(store, id);
    clearAllowedEditRanges(store, 0);
    expect(allowedEditRangesForSheet(store.getState(), 0).map((e) => e.title)).toEqual(['Keep']);
  });

  it('refuses the ribbon add action and says why', () => {
    const { store } = protectedStore();
    mutators.setRange(store, area(3, 3, 3, 3));
    const report = executeRibbonProtectionAction({
      store,
      action: 'allow-edit-range',
      strings: en.ribbonMenu,
    });
    expect(allowedEditRangesForSheet(store.getState(), 0)).toHaveLength(1);
    expect(report.items[0]?.severity).toBe('warning');
  });
});

describe('one unlock prompt at a time', () => {
  it('a second request waits for the first instead of asking again', async () => {
    const store = createSpreadsheetStore();
    mutators.addAllowedEditRange(store, {
      title: 'G',
      range: area(0, 0, 0, 0),
      legacyPasswordHash: legacySheetPasswordHash('pw'),
    });
    setProtectedSheet(store, 0, true);
    const ui = fakeUi(['pw']);
    registerAllowedRangeUnlockUi(store, ui);
    const [a, b] = await Promise.all([
      requestAllowedRangeUnlock(store, [area(0, 0, 0, 0)]),
      requestAllowedRangeUnlock(store, [area(0, 0, 0, 0)]),
    ]);
    expect([a, b]).toEqual([true, true]);
    expect(ui.askPassword).toHaveBeenCalledTimes(1);
  });
});

describe('locked cells outside every allowed range', () => {
  const setup = (answers: (string | null)[]) => {
    const store = createSpreadsheetStore();
    mutators.addAllowedEditRange(store, {
      title: 'G',
      range: area(0, 0, 0, 0),
      legacyPasswordHash: legacySheetPasswordHash('pw'),
    });
    setProtectedSheet(store, 0, true);
    const ui = fakeUi(answers);
    registerAllowedRangeUnlockUi(store, ui);
    return { store, ui };
  };

  it('refuses the whole edit before prompting', async () => {
    const { store, ui } = setup(['pw']);
    const run = vi.fn();
    expect(deferForAllowedRangeUnlock(store, [area(0, 0, 0, 1)], run)).toBe(true);
    await settle();
    await settle();
    expect(ui.askPassword).not.toHaveBeenCalled();
    expect(run).not.toHaveBeenCalled();
  });

  it('does not prompt when every target cell is unlocked', () => {
    const { store, ui } = setup(['pw']);
    mutators.setRangeFormat(store, area(0, 0, 0, 0), { locked: false });
    const run = vi.fn();
    expect(deferForAllowedRangeUnlock(store, [area(0, 0, 0, 0)], run)).toBe(false);
    expect(ui.askPassword).not.toHaveBeenCalled();
  });

  it('still prompts when the target is fully inside the guarded range', async () => {
    const { store, ui } = setup(['pw']);
    const run = vi.fn();
    expect(deferForAllowedRangeUnlock(store, [area(0, 0, 0, 0)], run)).toBe(true);
    await settle();
    await settle();
    expect(ui.askPassword).toHaveBeenCalledTimes(1);
    expect(run).toHaveBeenCalledTimes(1);
  });
});

describe('undo and redo of protection changes', () => {
  it('redo of a range typed before its hash existed still ends up hashed', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      const history = new History();
      recordProtectionChange(history, store, wb, () =>
        addAllowedEditRange(store, area(0, 0, 0, 0), { title: 'R', password: 'pw', workbook: wb }),
      );
      await whenSheetPasswordsHashed();
      history.undo();
      history.redo();
      await whenSheetPasswordsHashed();
      expect(wb.getProtectedRanges(0)[0]).toMatchObject({ password: '', algorithmName: 'SHA-512' });
      expect(store.getState().protection.allowedEditRanges[0]?.password).toBeUndefined();
    } finally {
      wb.dispose();
    }
  });

  it('undoing a sheet protection clears the unlocks of that sheet', async () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const id = mutators.addAllowedEditRange(store, {
      title: 'G',
      range: area(0, 0, 0, 0),
      legacyPasswordHash: legacySheetPasswordHash('pw'),
    });
    recordProtectionChange(history, store, undefined, () => setProtectedSheet(store, 0, true));
    registerAllowedRangeUnlockUi(store, fakeUi(['pw']));
    await requestAllowedRangeUnlock(store, [area(0, 0, 0, 0)]);
    expect(store.getState().protection.unlockedAllowedRanges?.has(id)).toBe(true);
    history.undo();
    expect(store.getState().protection.unlockedAllowedRanges?.has(id) ?? false).toBe(false);
  });
});
