import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as cellShift from '../../../src/commands/cell-shift.js';
import * as structure from '../../../src/commands/structure.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings/en.js';
import { ja } from '../../../src/i18n/strings/ja.js';
import {
  disposeCellShiftDialog,
  openCellShiftDialog,
  runCellShiftChoice,
} from '../../../src/interact/cell-shift-dialog.js';
import { createDialogShell } from '../../../src/interact/dialog-shell.js';
import { disposeOverlayPortal } from '../../../src/interact/overlay-portal.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('openCellShiftDialog', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
  });

  afterEach(() => {
    disposeCellShiftDialog(host);
    disposeOverlayPortal(host);
    document.body.replaceChildren();
  });

  // Excel: '挿入' / '削除' with four choices — two shifts, then Entire row and
  // Entire column — and the downward / upward shift selected.
  it.each([
    ['insert', ['right', 'down', 'rows', 'cols'], 'down', en.cellShiftDialog.insertTitle],
    ['delete', ['left', 'up', 'rows', 'cols'], 'up', en.cellShiftDialog.deleteTitle],
  ] as const)(
    '%s keeps its direction choices, default selection, and submits after closing',
    (kind, values, defaultValue, title) => {
      const onSubmit = vi.fn(() => {
        expect(document.querySelector('.fc-cellshift')).toBeNull();
      });
      host.focus();

      openCellShiftDialog({ host, strings: en, kind, onSubmit });

      const dialog = document.querySelector<HTMLElement>('.fc-cellshift');
      const radios = Array.from(
        dialog?.querySelectorAll<HTMLInputElement>('input[type="radio"]') ?? [],
      );
      const ok = dialog?.querySelector<HTMLButtonElement>('.fc-cellshift__button--primary');
      expect(dialog?.getAttribute('role')).toBe('dialog');
      expect(dialog?.getAttribute('aria-label')).toBe(title);
      expect(Array.from(radios, (radio) => radio.value)).toEqual(values);
      expect(radios.filter((radio) => radio.checked).map((radio) => radio.value)).toEqual([
        defaultValue,
      ]);
      expect(document.activeElement).toBe(ok);

      ok?.click();

      expect(onSubmit).toHaveBeenCalledExactlyOnceWith(defaultValue);
      expect(document.querySelector('.fc-cellshift')).toBeNull();
      expect(document.activeElement).toBe(host);
    },
  );

  it('labels the choices with the dialog wording, not the menu wording', () => {
    openCellShiftDialog({ host, strings: ja, kind: 'insert', onSubmit: vi.fn() });
    const dialog = document.querySelector<HTMLElement>('.fc-cellshift');
    expect(dialog?.querySelector('.fc-cellshift__title')?.textContent).toBe('挿入');
    expect(
      Array.from(dialog?.querySelectorAll('.fc-cellshift__choice') ?? [], (c) => c.textContent),
    ).toEqual(['右方向にシフト', '下方向にシフト', '行全体', '列全体']);
    document.querySelector<HTMLButtonElement>('.fc-cellshift__button--secondary')?.click();
  });

  it('defaults the Mac delete dialog to the left shift', () => {
    // Excel for Mac: '削除' opens with 左方向にシフト selected (insert keeps 下方向).
    host.dataset.fcPlatform = 'mac';
    openCellShiftDialog({ host, strings: ja, kind: 'delete', onSubmit: vi.fn() });
    const checked = document.querySelector<HTMLInputElement>('.fc-cellshift input:checked');
    expect(checked?.value).toBe('left');
    document.querySelector<HTMLButtonElement>('.fc-cellshift__button--secondary')?.click();
    openCellShiftDialog({ host, strings: ja, kind: 'insert', onSubmit: vi.fn() });
    expect(document.querySelector<HTMLInputElement>('.fc-cellshift input:checked')?.value).toBe(
      'down',
    );
    document.querySelector<HTMLButtonElement>('.fc-cellshift__button--secondary')?.click();
  });

  it('submits an entire-row choice', () => {
    const onSubmit = vi.fn();
    openCellShiftDialog({ host, strings: en, kind: 'delete', onSubmit });
    const dialog = document.querySelector<HTMLElement>('.fc-cellshift');
    const rows = dialog?.querySelector<HTMLInputElement>('input[value="rows"]');
    if (rows) rows.checked = true;
    dialog?.querySelector<HTMLButtonElement>('.fc-cellshift__button--primary')?.click();
    expect(onSubmit).toHaveBeenCalledExactlyOnceWith('rows');
  });

  it('wraps Tab from OK to the checked radio, then Shift+Tab back to OK', () => {
    openCellShiftDialog({ host, strings: en, kind: 'insert', onSubmit: vi.fn() });
    const dialog = document.querySelector<HTMLElement>('.fc-cellshift');
    const ok = dialog?.querySelector<HTMLButtonElement>('.fc-cellshift__button--primary');
    const selected = dialog?.querySelector<HTMLInputElement>('input[type="radio"]:checked');

    ok?.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true }),
    );
    expect(document.activeElement).toBe(selected);

    selected?.dispatchEvent(
      new KeyboardEvent('keydown', {
        key: 'Tab',
        shiftKey: true,
        bubbles: true,
        cancelable: true,
      }),
    );
    expect(document.activeElement).toBe(ok);
    document.querySelector<HTMLButtonElement>('.fc-cellshift__button--secondary')?.click();
  });

  it('Escape closes and restores the host that opened the dialog', () => {
    host.focus();
    openCellShiftDialog({ host, strings: en, kind: 'delete', onSubmit: vi.fn() });
    const ok = document.querySelector<HTMLButtonElement>('.fc-cellshift__button--primary');

    ok?.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }),
    );

    expect(document.querySelector('.fc-cellshift')).toBeNull();
    expect(document.activeElement).toBe(host);
  });

  it('Escape closes this dialog while leaving an older dialog open', () => {
    let parent: ReturnType<typeof createDialogShell> | null = null;
    parent = createDialogShell({
      host,
      className: 'fc-cellshift-parent-test',
      ariaLabel: 'Parent dialog',
      onDismiss: () => parent?.close(),
    });
    parent.panel.appendChild(document.createElement('button'));
    parent.open();
    openCellShiftDialog({ host, strings: en, kind: 'insert', onSubmit: vi.fn() });
    const ok = document.querySelector<HTMLButtonElement>('.fc-cellshift__button--primary');

    try {
      ok?.dispatchEvent(
        new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }),
      );

      expect(document.querySelector('.fc-cellshift')).toBeNull();
      expect(parent.isOpen()).toBe(true);
      expect(document.activeElement).toBe(parent.panel.querySelector('button'));
    } finally {
      parent.dispose();
    }
  });

  it('Cancel closes without submitting and restores host focus', () => {
    const onSubmit = vi.fn();
    host.focus();
    openCellShiftDialog({ host, strings: en, kind: 'insert', onSubmit });

    document.querySelector<HTMLButtonElement>('.fc-cellshift__button--secondary')?.click();

    expect(onSubmit).not.toHaveBeenCalled();
    expect(document.querySelector('.fc-cellshift')).toBeNull();
    expect(document.activeElement).toBe(host);
  });

  it('Enter on the selected radio submits once after removing the dialog', () => {
    const onSubmit = vi.fn(() => {
      expect(document.querySelector('.fc-cellshift')).toBeNull();
    });
    openCellShiftDialog({ host, strings: en, kind: 'insert', onSubmit });
    const selected = document.querySelector<HTMLInputElement>(
      '.fc-cellshift input[type="radio"]:checked',
    );
    selected?.focus();

    selected?.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }),
    );

    expect(onSubmit).toHaveBeenCalledExactlyOnceWith('down');
    expect(document.querySelector('.fc-cellshift')).toBeNull();
  });

  it('does not submit when Enter is part of IME composition', () => {
    const onSubmit = vi.fn();
    openCellShiftDialog({ host, strings: en, kind: 'insert', onSubmit });
    const selected = document.querySelector<HTMLInputElement>(
      '.fc-cellshift input[type="radio"]:checked',
    );
    selected?.focus();
    const event = new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true });
    Object.defineProperty(event, 'isComposing', { value: true });

    selected?.dispatchEvent(event);

    expect(onSubmit).not.toHaveBeenCalled();
    expect(document.querySelector('.fc-cellshift')).not.toBeNull();
    document.querySelector<HTMLButtonElement>('.fc-cellshift__button--secondary')?.click();
  });

  it('replacing a dialog prevents the old OK button from submitting', () => {
    const oldSubmit = vi.fn();
    const newSubmit = vi.fn();
    openCellShiftDialog({ host, strings: en, kind: 'insert', onSubmit: oldSubmit });
    const oldDialog = document.querySelector('.fc-cellshift');
    const oldOk = oldDialog?.querySelector<HTMLButtonElement>('.fc-cellshift__button--primary');

    openCellShiftDialog({ host, strings: en, kind: 'delete', onSubmit: newSubmit });
    const currentDialog = document.querySelector('.fc-cellshift');
    oldOk?.click();

    expect(oldDialog?.isConnected).toBe(false);
    expect(oldSubmit).not.toHaveBeenCalled();
    expect(currentDialog?.isConnected).toBe(true);
    expect(newSubmit).not.toHaveBeenCalled();
    document.querySelector<HTMLButtonElement>('.fc-cellshift__button--secondary')?.click();
  });

  it('ignores stale instance cleanup and disposes the current host dialog without stealing focus', () => {
    host.dataset.fcInstId = 'older-instance';
    openCellShiftDialog({ host, strings: en, kind: 'insert', onSubmit: vi.fn() });

    host.dataset.fcInstId = 'current-instance';
    openCellShiftDialog({ host, strings: en, kind: 'delete', onSubmit: vi.fn() });
    const currentDialog = document.querySelector('.fc-cellshift');
    const outside = document.createElement('button');
    document.body.appendChild(outside);
    outside.focus();

    disposeCellShiftDialog(host, 'older-instance');
    expect(currentDialog?.isConnected).toBe(true);
    expect(document.activeElement).toBe(outside);

    disposeCellShiftDialog(host, 'current-instance');
    expect(currentDialog?.isConnected).toBe(false);
    expect(document.activeElement).toBe(outside);
  });

  it('removes the current dialog before a nested submit callback opens its replacement', () => {
    let oldDialog: Element | null = null;
    let dialogsSeenByCallback = -1;
    const nestedSubmit = vi.fn();
    const onSubmit = vi.fn(() => {
      dialogsSeenByCallback = document.querySelectorAll('.fc-cellshift').length;
      openCellShiftDialog({ host, strings: en, kind: 'delete', onSubmit: nestedSubmit });
    });
    openCellShiftDialog({ host, strings: en, kind: 'insert', onSubmit });
    oldDialog = document.querySelector('.fc-cellshift');

    document.querySelector<HTMLButtonElement>('.fc-cellshift__button--primary')?.click();

    const replacement = document.querySelector('.fc-cellshift');
    expect(onSubmit).toHaveBeenCalledExactlyOnceWith('down');
    expect(dialogsSeenByCallback).toBe(0);
    expect(oldDialog?.isConnected).toBe(false);
    expect(replacement).not.toBeNull();
    document.querySelector<HTMLButtonElement>('.fc-cellshift__button--secondary')?.click();
  });
});

describe('runCellShiftChoice', () => {
  const setup = async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const store = createSpreadsheetStore();
    return { wb, store };
  };

  it('routes each dialog choice to its command over the normalized selection', async () => {
    const { wb, store } = await setup();
    const range = { sheet: 0, r0: 4, c0: 3, r1: 2, c1: 1 };
    const spies = {
      insertCells: vi.spyOn(cellShift, 'insertCells').mockReturnValue(true),
      deleteCells: vi.spyOn(cellShift, 'deleteCells').mockReturnValue(true),
      insertRows: vi.spyOn(structure, 'insertRows').mockReturnValue(true),
      deleteRows: vi.spyOn(structure, 'deleteRows').mockReturnValue(true),
      insertCols: vi.spyOn(structure, 'insertCols').mockReturnValue(true),
      deleteCols: vi.spyOn(structure, 'deleteCols').mockReturnValue(true),
    };
    const norm = { sheet: 0, r0: 2, c0: 1, r1: 4, c1: 3 };
    expect(runCellShiftChoice({ store, wb, history: null }, 'insert', range, 'down')).toBe(true);
    expect(spies.insertCells).toHaveBeenCalledWith(store, wb, null, norm, 'down');
    runCellShiftChoice({ store, wb, history: null }, 'delete', range, 'left');
    expect(spies.deleteCells).toHaveBeenCalledWith(store, wb, null, norm, 'left');
    runCellShiftChoice({ store, wb, history: null }, 'insert', range, 'rows');
    expect(spies.insertRows).toHaveBeenCalledWith(store, wb, null, 2, 3);
    runCellShiftChoice({ store, wb, history: null }, 'delete', range, 'cols');
    expect(spies.deleteCols).toHaveBeenCalledWith(store, wb, null, 1, 3);
    // A direction that belongs to the other dialog does nothing.
    expect(runCellShiftChoice({ store, wb, history: null }, 'insert', range, 'up')).toBe(false);
    vi.restoreAllMocks();
    wb.dispose();
  });
});
