import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  type AllowedRangeUnlockUi,
  addAllowedEditRange,
  allowedRangesPendingUnlock,
  isAddrInAllowedEditRange,
  isCellWritable,
  recordProtectionChange,
  registerAllowedRangeUnlockUi,
  requestAllowedRangeUnlock,
  setProtectedSheet,
  whenSheetPasswordsHashed,
} from '../../../src/commands/protection.js';
import { legacySheetPasswordHash } from '../../../src/engine/protection-password.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachClipboard } from '../../../src/interact/clipboard.js';
import { attachKeyboard } from '../../../src/interact/keyboard.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const addRange = (...args: Parameters<typeof addAllowedEditRange>): string =>
  addAllowedEditRange(...args) as string;

const cell = (row: number, col: number, sheet = 0) => ({ sheet, row, col });
const area = (r0: number, c0: number, r1: number, c1: number, sheet = 0) => ({
  sheet,
  r0,
  c0,
  r1,
  c1,
});

const fakeUi = (answers: Array<string | null>) => {
  const ui: AllowedRangeUnlockUi & {
    askPassword: ReturnType<typeof vi.fn>;
    reportIncorrect: ReturnType<typeof vi.fn>;
  } = {
    askPassword: vi.fn(async () => answers.shift() ?? null),
    reportIncorrect: vi.fn(async () => {}),
  };
  return ui;
};

const settle = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 0));

describe('password-guarded allowed edit ranges', () => {
  it('leaves a range without a password freely editable', () => {
    const store = createSpreadsheetStore();
    addAllowedEditRange(store, area(0, 0, 1, 1));
    setProtectedSheet(store, 0, true);
    expect(isCellWritable(store.getState(), cell(0, 0))).toBe(true);
    expect(allowedRangesPendingUnlock(store.getState(), [area(0, 0, 0, 0)])).toEqual([]);
  });

  it('holds a guarded range back until its password is entered, once per session', async () => {
    const store = createSpreadsheetStore();
    addAllowedEditRange(store, area(0, 0, 1, 1), { password: 'open' });
    setProtectedSheet(store, 0, true);
    const ui = fakeUi(['open']);
    registerAllowedRangeUnlockUi(store, ui);

    expect(isCellWritable(store.getState(), cell(0, 0))).toBe(false);
    expect(isAddrInAllowedEditRange(store.getState(), cell(0, 0))).toBe(false);
    await expect(requestAllowedRangeUnlock(store, [area(0, 0, 0, 0)])).resolves.toBe(true);
    expect(isCellWritable(store.getState(), cell(1, 1))).toBe(true);
    await expect(requestAllowedRangeUnlock(store, [area(1, 1, 1, 1)])).resolves.toBe(true);
    expect(ui.askPassword).toHaveBeenCalledTimes(1);
    // Cells outside the range stay locked.
    expect(isCellWritable(store.getState(), cell(5, 5))).toBe(false);
  });

  it('refuses a wrong or cancelled password', async () => {
    const store = createSpreadsheetStore();
    addAllowedEditRange(store, area(0, 0, 0, 0), { password: 'open' });
    setProtectedSheet(store, 0, true);
    const ui = fakeUi(['nope', null]);
    registerAllowedRangeUnlockUi(store, ui);

    await expect(requestAllowedRangeUnlock(store, [area(0, 0, 0, 0)])).resolves.toBe(false);
    expect(ui.reportIncorrect).toHaveBeenCalledTimes(1);
    await expect(requestAllowedRangeUnlock(store, [area(0, 0, 0, 0)])).resolves.toBe(false);
    expect(ui.reportIncorrect).toHaveBeenCalledTimes(1);
    expect(isCellWritable(store.getState(), cell(0, 0))).toBe(false);
  });

  it('refuses without a registered prompt', async () => {
    const store = createSpreadsheetStore();
    addAllowedEditRange(store, area(0, 0, 0, 0), { password: 'open' });
    setProtectedSheet(store, 0, true);
    await expect(requestAllowedRangeUnlock(store, [area(0, 0, 0, 0)])).resolves.toBe(false);
  });

  it('verifies the hashed password and the legacy verifier of a loaded file', async () => {
    const store = createSpreadsheetStore();
    addAllowedEditRange(store, area(0, 0, 0, 0), { password: 'hashed' });
    await whenSheetPasswordsHashed();
    expect(store.getState().protection.allowedEditRanges[0]?.passwordHash).toBeTruthy();
    mutators.addAllowedEditRange(store, {
      title: 'legacy',
      range: area(3, 3, 3, 3),
      legacyPasswordHash: legacySheetPasswordHash('abc'),
    });
    setProtectedSheet(store, 0, true);
    registerAllowedRangeUnlockUi(store, fakeUi(['hashed', 'abd', 'abc']));

    await expect(requestAllowedRangeUnlock(store, [area(0, 0, 0, 0)])).resolves.toBe(true);
    await expect(requestAllowedRangeUnlock(store, [area(3, 3, 3, 3)])).resolves.toBe(false);
    await expect(requestAllowedRangeUnlock(store, [area(3, 3, 3, 3)])).resolves.toBe(true);
  });

  it('keeps the unlock out of undo and clears it on re-protect or a password change', async () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const id = addRange(store, area(0, 0, 0, 0), { password: 'open' });
    setProtectedSheet(store, 0, true);
    registerAllowedRangeUnlockUi(store, fakeUi(['open', 'open', 'new']));
    await requestAllowedRangeUnlock(store, [area(0, 0, 0, 0)]);
    expect(store.getState().protection.unlockedAllowedRanges?.has(id)).toBe(true);

    recordProtectionChange(history, store, undefined, () =>
      addAllowedEditRange(store, area(7, 7, 7, 7), { title: 'other' }),
    );
    history.undo();
    history.redo();
    expect(isCellWritable(store.getState(), cell(0, 0))).toBe(true);

    setProtectedSheet(store, 0, false);
    setProtectedSheet(store, 0, true);
    expect(isCellWritable(store.getState(), cell(0, 0))).toBe(false);

    await requestAllowedRangeUnlock(store, [area(0, 0, 0, 0)]);
    expect(isCellWritable(store.getState(), cell(0, 0))).toBe(true);
  });

  it('does not write the unlock state to the file', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      addAllowedEditRange(store, area(0, 0, 0, 0), { password: 'open', workbook: wb });
      await whenSheetPasswordsHashed();
      setProtectedSheet(store, 0, true, { workbook: wb });
      registerAllowedRangeUnlockUi(store, fakeUi(['open']));
      const before = JSON.stringify(wb.getProtectedRanges(0));
      await requestAllowedRangeUnlock(store, [area(0, 0, 0, 0)]);
      expect(JSON.stringify(wb.getProtectedRanges(0))).toBe(before);
    } finally {
      wb.dispose();
    }
  });
});

describe('edit entry points prompt for a guarded range', () => {
  let host: HTMLElement;
  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
  });
  afterEach(() => {
    document.body.innerHTML = '';
  });

  const setup = async (answers: Array<string | null>) => {
    const store = createSpreadsheetStore();
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    wb.setNumber(cell(0, 0), 5);
    wb.recalc();
    addAllowedEditRange(store, area(0, 0, 0, 0), { password: 'open' });
    setProtectedSheet(store, 0, true);
    const ui = fakeUi(answers);
    registerAllowedRangeUnlockUi(store, ui);
    return { store, wb, ui };
  };

  it('Delete clears after the right password and keeps the cell after a wrong one', async () => {
    const { store, wb, ui } = await setup(['bad', 'open']);
    const detach = attachKeyboard({
      host,
      store,
      wb,
      history: null,
      onBeginEdit: vi.fn(),
      onClearActive: vi.fn(),
      onGoTo: vi.fn(),
      onSwitchSheet: vi.fn(),
      onPasteSpecial: vi.fn(),
    });
    const press = (): void => {
      host.dispatchEvent(
        new KeyboardEvent('keydown', { key: 'Delete', bubbles: true, cancelable: true }),
      );
    };
    press();
    await settle();
    expect(ui.reportIncorrect).toHaveBeenCalledTimes(1);
    expect(wb.getValue(cell(0, 0))).toEqual({ kind: 'number', value: 5 });

    press();
    await settle();
    expect(wb.getValue(cell(0, 0)).kind).toBe('blank');
    detach();
  });

  it('paste lands only once the range is unlocked', async () => {
    const { store, wb, ui } = await setup(['open']);
    const handle = attachClipboard({ host, store, wb, onAfterCommit: vi.fn() });
    const transfer = new DataTransfer();
    transfer.setData('text/plain', '9');
    host.dispatchEvent(
      new ClipboardEvent('paste', { clipboardData: transfer, bubbles: true, cancelable: true }),
    );
    await settle();
    expect(ui.askPassword).toHaveBeenCalledTimes(1);
    expect(wb.getValue(cell(0, 0))).toEqual({ kind: 'number', value: 9 });
    handle.detach();
  });
});
