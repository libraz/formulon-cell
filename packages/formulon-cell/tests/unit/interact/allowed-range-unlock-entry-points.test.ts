import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  type AllowedRangeUnlockUi,
  registerAllowedRangeUnlockUi,
  setProtectedSheet,
} from '../../../src/commands/protection.js';
import { legacySheetPasswordHash } from '../../../src/engine/protection-password.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings/en.js';
import { attachAutoFillOptions } from '../../../src/interact/auto-fill-options.js';
import { attachFindReplace } from '../../../src/interact/find-replace.js';
import { createDataDropdownDefaults } from '../../../src/mount/dynamic-dropdown-defaults/data.js';
import { createEditingDropdownDefaults } from '../../../src/mount/dynamic-dropdown-defaults/editing.js';
import { createHostShortcutHandler } from '../../../src/mount/host-shortcuts.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';
import { handleAutoSum, handlePasteAction } from '../../../src/wrappers/toolbar-actions.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

const PASSWORD = 'pw';
const settle = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 0));
const area = (r0: number, c0: number, r1: number, c1: number): Range => ({
  sheet: 0,
  r0,
  c0,
  r1,
  c1,
});

/** Protect sheet 0 and guard `range` with PASSWORD; the fake prompt answers in order. */
const guard = (store: SpreadsheetStore, range: Range, answers: Array<string | null>) => {
  mutators.addAllowedEditRange(store, {
    title: 'Guarded',
    range,
    legacyPasswordHash: legacySheetPasswordHash(PASSWORD),
  });
  setProtectedSheet(store, 0, true);
  const ui: AllowedRangeUnlockUi & {
    askPassword: ReturnType<typeof vi.fn>;
    reportIncorrect: ReturnType<typeof vi.fn>;
  } = {
    askPassword: vi.fn(async () => answers.shift() ?? null),
    reportIncorrect: vi.fn(async () => {}),
  };
  registerAllowedRangeUnlockUi(store, ui);
  return ui;
};

const num = (wb: WorkbookHandle, row: number, col: number): number | string => {
  const value = wb.getValue({ sheet: 0, row, col });
  return value.kind === 'number' ? value.value : value.kind;
};

const select = (store: SpreadsheetStore, range: Range): void => {
  mutators.setRange(store, range);
  mutators.setActive(store, { sheet: 0, row: range.r0, col: range.c0 });
  mutators.setRange(store, range);
};

describe('fill shortcuts (Ctrl+D / Ctrl+R / Ctrl+E)', () => {
  const setup = async (answers: Array<string | null>) => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const store = createSpreadsheetStore();
    const history = new History();
    wb.attachHistory(history);
    const invalidate = vi.fn();
    const handler = createHostShortcutHandler({
      addSheet: vi.fn(),
      findReplace: () => null,
      formatDialog: () => null,
      formatPainter: () => null,
      goToDialog: () => null,
      history,
      host: document.createElement('div'),
      hostTag: document.createElement('input'),
      hyperlinkDialog: () => null,
      invalidate,
      getLocale: () => 'en-US',
      namedRangeDialog: () => null,
      pasteSpecialDialog: () => null,
      quickAnalysis: () => null,
      store,
      strings: () => en,
      wb: () => wb,
    });
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 5);
    wb.setNumber({ sheet: 0, row: 0, col: 1 }, 7);
    wb.recalc();
    mutators.replaceCells(store, wb.cells(0));
    select(store, area(0, 0, 2, 0));
    const ui = guard(store, area(0, 0, 2, 1), answers);
    return { wb, store, handler, invalidate, ui };
  };
  const key = (init: KeyboardEventInit) =>
    new KeyboardEvent('keydown', { ...init, cancelable: true });

  it('Ctrl+D fills down only after the range is unlocked', async () => {
    const { wb, handler, ui } = await setup(['bad', PASSWORD]);
    handler(key({ key: 'd', ctrlKey: true }));
    await settle();
    expect(ui.reportIncorrect).toHaveBeenCalledTimes(1);
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 }).kind).toBe('blank');

    handler(key({ key: 'd', ctrlKey: true }));
    await settle();
    expect(num(wb, 1, 0)).toBe(5);
    expect(num(wb, 2, 0)).toBe(5);
  });

  it('Ctrl+R fills right after the unlock', async () => {
    const { wb, store, handler } = await setup([PASSWORD]);
    select(store, area(0, 0, 0, 1));
    wb.setNumber({ sheet: 0, row: 0, col: 1 }, 0);
    mutators.replaceCells(store, wb.cells(0));
    handler(key({ key: 'r', ctrlKey: true }));
    await settle();
    expect(num(wb, 0, 1)).toBe(5);
  });

  it('Ctrl+E (flash fill) asks first and runs once unlocked', async () => {
    const { handler, invalidate, ui } = await setup([PASSWORD]);
    handler(key({ key: 'e', ctrlKey: true }));
    expect(invalidate).not.toHaveBeenCalled();
    await settle();
    expect(ui.askPassword).toHaveBeenCalledTimes(1);
    expect(invalidate).toHaveBeenCalled();
  });
});

describe('AutoFill options', () => {
  let host: HTMLDivElement;
  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
  });
  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('re-applying a fill mode into a guarded range asks first', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const store = createSpreadsheetStore();
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      cells.set('0:0:0', { value: { kind: 'number', value: 1 }, formula: null });
      cells.set('0:1:0', { value: { kind: 'number', value: 2 }, formula: null });
      return { ...s, data: { ...s.data, cells } };
    });
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    wb.setNumber({ sheet: 0, row: 1, col: 0 }, 2);
    wb.recalc();
    const ui = guard(store, area(2, 0, 3, 0), ['bad', PASSWORD]);
    const handle = attachAutoFillOptions({ host, store, wb, strings: en });
    host.dispatchEvent(
      new CustomEvent('fc:autofilloptions', {
        detail: {
          src: area(0, 0, 1, 0),
          dest: area(0, 0, 3, 0),
          mode: 'series',
          clientX: 1,
          clientY: 1,
        },
      }),
    );
    const copy = (): void => {
      document.querySelector<HTMLButtonElement>('.fc-autofill-options__button')?.click();
      Array.from(document.querySelectorAll<HTMLButtonElement>('.fc-autofill-options__item'))
        .find((item) => item.textContent === 'Copy Cells')
        ?.click();
    };
    copy();
    await settle();
    expect(ui.reportIncorrect).toHaveBeenCalledTimes(1);
    expect(wb.getValue({ sheet: 0, row: 2, col: 0 }).kind).toBe('blank');
    copy();
    await settle();
    expect(num(wb, 2, 0)).toBe(1);
    handle.detach();
  });
});

describe('ribbon commands on a mounted sheet', () => {
  let sheet: MountedStubSheet;
  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
  });
  afterEach(() => sheet.dispose());

  const seed = (cells: [number, number, number | string][]): void => {
    for (const [row, col, value] of cells) {
      if (typeof value === 'number') sheet.workbook.setNumber({ sheet: 0, row, col }, value);
      else sheet.workbook.setText({ sheet: 0, row, col }, value);
    }
    sheet.workbook.recalc();
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
  };

  it('paste commands (Paste Values) wait for the unlock', async () => {
    seed([[0, 0, 5]]);
    const { store, host } = sheet.instance;
    select(store, area(0, 0, 0, 0));
    const transfer = new DataTransfer();
    host.dispatchEvent(
      new ClipboardEvent('copy', { clipboardData: transfer, bubbles: true, cancelable: true }),
    );
    select(store, area(0, 2, 0, 2));
    const ui = guard(store, area(0, 2, 0, 2), [PASSWORD]);
    handlePasteAction(sheet.instance, 'pasteValues');
    await settle();
    expect(ui.askPassword).toHaveBeenCalledTimes(1);
    expect(num(sheet.workbook, 0, 2)).toBe(5);
  });

  it('Sort A to Z waits for the unlock', async () => {
    seed([
      [0, 0, 3],
      [1, 0, 1],
      [2, 0, 2],
    ]);
    const { store } = sheet.instance;
    select(store, area(0, 0, 0, 0));
    const ui = guard(store, area(0, 0, 2, 0), [PASSWORD]);
    createEditingDropdownDefaults(sheet.instance).applySortMenuAction('asc');
    expect(num(sheet.workbook, 0, 0)).toBe(3);
    await settle();
    expect(ui.askPassword).toHaveBeenCalledTimes(1);
    expect([0, 1, 2].map((row) => num(sheet.workbook, row, 0))).toEqual([1, 2, 3]);
  });

  it('AutoSum writes its total only once the range is unlocked', async () => {
    seed([
      [0, 0, 1],
      [1, 0, 2],
    ]);
    const { store } = sheet.instance;
    select(store, area(0, 0, 1, 0));
    const ui = guard(store, area(2, 0, 2, 0), ['bad']);
    expect(handleAutoSum(sheet.instance, 'SUM')).toBe(false);
    await settle();
    expect(ui.reportIncorrect).toHaveBeenCalledTimes(1);
    expect(sheet.workbook.getValue({ sheet: 0, row: 2, col: 0 }).kind).toBe('blank');
  });

  it('Text to Columns waits for the unlock', async () => {
    seed([[0, 0, 'a,b']]);
    const { store } = sheet.instance;
    select(store, area(0, 0, 0, 0));
    const ui = guard(store, area(0, 0, 0, 1), [PASSWORD]);
    createDataDropdownDefaults(sheet.instance).splitTextToColumns(',');
    await settle();
    expect(ui.askPassword).toHaveBeenCalledTimes(1);
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
      kind: 'text',
      value: 'b',
    });
  });
});

describe('Replace All', () => {
  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('replaces inside a guarded range only after the unlock', async () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const store = createSpreadsheetStore();
    wb.setText({ sheet: 0, row: 0, col: 0 }, 'cat');
    wb.recalc();
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      cells.set('0:0:0', { value: { kind: 'text', value: 'cat' }, formula: null });
      return { ...s, data: { ...s.data, cells } };
    });
    const ui = guard(store, area(0, 0, 0, 0), [PASSWORD]);
    const handle = attachFindReplace({ host, store, wb });
    handle.open('replace');
    const input = (sel: string): HTMLInputElement => host.querySelector(sel) as HTMLInputElement;
    for (const [sel, value] of [
      ['#fc-find-find-input', 'cat'],
      ['#fc-find-replace-input', 'dog'],
    ] as const) {
      input(sel).value = value;
      input(sel).dispatchEvent(new Event('input', { bubbles: true }));
    }
    (host.querySelector('.fc-find__btn--replace-all') as HTMLButtonElement).click();
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'cat' });
    await settle();
    expect(ui.askPassword).toHaveBeenCalledTimes(1);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'dog' });
    handle.detach();
  });
});
