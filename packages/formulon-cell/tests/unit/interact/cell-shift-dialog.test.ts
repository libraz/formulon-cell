import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { en } from '../../../src/i18n/strings/en.js';
import {
  disposeCellShiftDialog,
  openCellShiftDialog,
} from '../../../src/interact/cell-shift-dialog.js';
import { createDialogShell } from '../../../src/interact/dialog-shell.js';
import { disposeOverlayPortal } from '../../../src/interact/overlay-portal.js';

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

  it.each([
    ['insert', ['right', 'down'], 'down', en.ribbonMenu.insertCells],
    ['delete', ['left', 'up'], 'up', en.ribbonMenu.deleteCells],
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
