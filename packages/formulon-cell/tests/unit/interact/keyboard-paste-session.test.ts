import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachClipboard } from '../../../src/interact/clipboard.js';
import { attachKeyboard } from '../../../src/interact/keyboard.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const fire = (
  host: HTMLElement,
  key: string,
  init: Partial<KeyboardEventInit> = {},
): KeyboardEvent => {
  const event = new KeyboardEvent('keydown', {
    key,
    bubbles: true,
    cancelable: true,
    ...init,
  });
  host.dispatchEvent(event);
  return event;
};

const settleMicrotasks = async (): Promise<void> => {
  await Promise.resolve();
  await Promise.resolve();
  await Promise.resolve();
};

describe('keyboard Enter paste sessions', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let clipboard: ReturnType<typeof attachClipboard>;
  let detachKeyboard: () => void;
  let readText: ReturnType<typeof vi.fn>;

  const select = (row: number, col = 0): void => {
    const active = { sheet: 0, row, col };
    mutators.setActive(store, active);
    mutators.setRange(store, { sheet: 0, r0: row, c0: col, r1: row, c1: col });
  };

  beforeEach(async () => {
    host = document.createElement('div');
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await WorkbookHandle.createDefault({ preferStub: true });
    readText = vi.fn().mockRejectedValue(new DOMException('Denied', 'NotAllowedError'));
    vi.stubGlobal('navigator', {
      clipboard: { readText, writeText: vi.fn().mockResolvedValue(undefined) },
    });

    wb.setText({ sheet: 0, row: 0, col: 0 }, 'Copied');
    mutators.replaceCells(store, wb.cells(0));
    clipboard = attachClipboard({
      host,
      store,
      wb,
      onAfterCommit: () => mutators.replaceCells(store, wb.cells(0)),
    });
    detachKeyboard = attachKeyboard({
      host,
      store,
      wb,
      onBeginEdit: vi.fn(),
      onClearActive: vi.fn(),
      onClipboardShortcut: (kind) => clipboard.runShortcut(kind),
    });
  });

  afterEach(() => {
    detachKeyboard?.();
    clipboard.detach();
    wb.dispose();
    document.body.innerHTML = '';
    vi.unstubAllGlobals();
  });

  it('pastes through the real clipboard adapter before clearing the marquee', async () => {
    select(0);
    await clipboard.runShortcut('copy');
    select(1);

    const event = fire(host, 'Enter');
    expect(event.defaultPrevented).toBe(true);
    expect(store.getState().ui.copyRange).not.toBeNull();

    await settleMicrotasks();

    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'text',
      value: 'Copied',
    });
    expect(store.getState().ui.copyRange).toBeNull();
  });

  it('does not let a pending Enter cleanup clear a replacement marquee', async () => {
    select(0);
    await clipboard.runShortcut('copy');
    let rejectRead: ((reason: unknown) => void) | undefined;
    readText.mockImplementation(
      () =>
        new Promise<string>((_resolve, reject) => {
          rejectRead = reject;
        }),
    );
    select(1);
    fire(host, 'Enter');

    const replacement = { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 };
    mutators.setCopyRange(store, replacement);
    rejectRead?.(new DOMException('Denied', 'NotAllowedError'));
    await settleMicrotasks();

    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'blank' });
    expect(store.getState().ui.copyRange).toEqual(replacement);
  });

  it('allows a newer Enter session while the older one is pending', async () => {
    select(0);
    await clipboard.runShortcut('copy');
    wb.setText({ sheet: 0, row: 2, col: 0 }, 'Replacement');
    mutators.replaceCells(store, wb.cells(0));

    const resolveRead: Array<(text: string) => void> = [];
    const rejectRead: Array<(reason: unknown) => void> = [];
    readText.mockImplementation(
      () =>
        new Promise<string>((resolve, reject) => {
          resolveRead.push(resolve);
          rejectRead.push(reject);
        }),
    );

    select(1);
    fire(host, 'Enter');
    select(2);
    await clipboard.runShortcut('copy');
    select(3);
    fire(host, 'Enter');
    expect(readText).toHaveBeenCalledTimes(2);

    rejectRead[0]?.(new DOMException('Denied', 'NotAllowedError'));
    await settleMicrotasks();
    expect(store.getState().ui.copyRange).toEqual({
      sheet: 0,
      r0: 2,
      c0: 0,
      r1: 2,
      c1: 0,
    });
    expect(wb.getValue({ sheet: 0, row: 3, col: 0 })).toEqual({ kind: 'blank' });

    resolveRead[1]?.('Replacement');
    await settleMicrotasks();
    expect(wb.getValue({ sheet: 0, row: 3, col: 0 })).toEqual({
      kind: 'text',
      value: 'Replacement',
    });
    expect(store.getState().ui.copyRange).toBeNull();
  });

  it('ignores a repeated Enter while the first paste is still pending', async () => {
    select(0);
    await clipboard.runShortcut('copy');
    let finishRead: ((text: string) => void) | undefined;
    readText.mockImplementation(
      () =>
        new Promise<string>((resolve) => {
          finishRead = resolve;
        }),
    );
    select(1);

    fire(host, 'Enter');
    const second = fire(host, 'Enter');
    expect(second.defaultPrevented).toBe(true);
    expect(readText).toHaveBeenCalledTimes(1);

    finishRead?.('Copied');
    await settleMicrotasks();

    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'text',
      value: 'Copied',
    });
    expect(store.getState().ui.copyRange).toBeNull();
  });

  it('clears the marquee immediately for a synchronous shortcut callback', () => {
    detachKeyboard();
    const onClipboardShortcut = vi.fn();
    detachKeyboard = attachKeyboard({
      host,
      store,
      wb,
      onBeginEdit: vi.fn(),
      onClearActive: vi.fn(),
      onClipboardShortcut,
    });
    mutators.setCopyRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });

    fire(host, 'Enter');

    expect(onClipboardShortcut).toHaveBeenCalledWith('paste');
    expect(store.getState().ui.copyRange).toBeNull();
  });
});
