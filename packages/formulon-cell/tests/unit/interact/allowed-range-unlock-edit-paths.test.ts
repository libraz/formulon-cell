import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { commentAt, setComment } from '../../../src/commands/comment.js';
import { toggleBold } from '../../../src/commands/format.js';
import { History } from '../../../src/commands/history.js';
import { hyperlinkAt, setHyperlink } from '../../../src/commands/hyperlinks.js';
import {
  type AllowedRangeUnlockUi,
  registerAllowedRangeUnlockUi,
  setProtectedSheet,
} from '../../../src/commands/protection.js';
import { executeRibbonCommentAction } from '../../../src/commands/ribbon-comment.js';
import { recordRepeatableFormatChange } from '../../../src/commands/slice-history.js';
import { legacySheetPasswordHash } from '../../../src/engine/protection-password.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings/en.js';
import { executeRibbonHyperlinkAction } from '../../../src/interact/ribbon-hyperlink.js';
import { createHostShortcutHandler } from '../../../src/mount/host-shortcuts.js';
import { formatWithPending } from '../../../src/store/pending-format.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';
import { applyMergeAction } from '../../../src/toolbar/merge-action.js';

const PASSWORD = 'pw';
const settle = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 0));
const area = (r0: number, c0: number, r1: number, c1: number): Range => ({
  sheet: 0,
  r0,
  c0,
  r1,
  c1,
});
const A1 = { sheet: 0, row: 0, col: 0 };

const guard = (store: SpreadsheetStore, range: Range, answers: (string | null)[]) => {
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

const select = (store: SpreadsheetStore, range: Range): void => {
  mutators.setActive(store, { sheet: 0, row: range.r0, col: range.c0 });
  mutators.setRange(store, range);
};

const bold = (store: SpreadsheetStore): boolean | undefined =>
  formatWithPending(store.getState(), A1)?.bold;

describe('format paths', () => {
  let store: SpreadsheetStore;
  let history: History;
  beforeEach(() => {
    store = createSpreadsheetStore();
    history = new History();
    select(store, area(0, 0, 0, 0));
  });

  it('recordRepeatableFormatChange (ribbon, dropdowns, context menu) asks, then applies', async () => {
    const ui = guard(store, area(0, 0, 1, 1), ['bad', PASSWORD]);
    const run = () =>
      recordRepeatableFormatChange(history, store, () => toggleBold(store.getState(), store));
    run();
    await settle();
    expect(ui.reportIncorrect).toHaveBeenCalledTimes(1);
    expect(bold(store)).toBeUndefined();
    run();
    await settle();
    expect(bold(store)).toBe(true);
  });

  it('keyboard format shortcuts (Ctrl+B) ask, then apply', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    wb.attachHistory(history);
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
      invalidate: vi.fn(),
      getLocale: () => 'en-US',
      namedRangeDialog: () => null,
      pasteSpecialDialog: () => null,
      quickAnalysis: () => null,
      store,
      strings: () => en,
      wb: () => wb,
    });
    const ui = guard(store, area(0, 0, 1, 1), [PASSWORD]);
    handler(new KeyboardEvent('keydown', { key: 'b', ctrlKey: true, cancelable: true }));
    expect(bold(store)).toBeUndefined();
    await settle();
    expect(ui.askPassword).toHaveBeenCalledTimes(1);
    expect(bold(store)).toBe(true);
  });
});

describe('merge', () => {
  it('Merge Cells waits for the unlock and is dropped on a wrong password', async () => {
    const store = createSpreadsheetStore();
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    select(store, area(0, 0, 0, 1));
    const ui = guard(store, area(0, 0, 0, 1), ['bad', PASSWORD]);
    const context = { store, workbook: wb, history: null, strings: en };
    const merges = () => store.getState().merges.byAnchor.size;

    await expect(applyMergeAction(context, 'mergeCells')).resolves.toBe(false);
    await settle();
    expect(ui.reportIncorrect).toHaveBeenCalledTimes(1);
    expect(merges()).toBe(0);

    await applyMergeAction(context, 'mergeCells');
    await settle();
    expect(merges()).toBe(1);
  });
});

describe('hyperlinks', () => {
  it('Remove Hyperlink asks first, then clears', async () => {
    const store = createSpreadsheetStore();
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    setHyperlink(store, A1, 'https://example.com/', wb);
    select(store, area(0, 0, 0, 0));
    const ui = guard(store, area(0, 0, 0, 0), [PASSWORD]);
    const result = executeRibbonHyperlinkAction({
      store,
      workbook: wb,
      history: new History(),
      action: 'clear',
      strings: en.ribbonMenu,
    });
    expect(result.kind).toBe('deferred');
    expect(hyperlinkAt(store.getState(), A1)).not.toBeNull();
    await settle();
    expect(ui.askPassword).toHaveBeenCalledTimes(1);
    expect(hyperlinkAt(store.getState(), A1)).toBeNull();
  });
});

describe('hyperlink outcome', () => {
  it('reports cleared when the link is removed at once and refused on a locked cell', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const deps = (store: SpreadsheetStore) => ({
      store,
      workbook: wb,
      history: new History(),
      action: 'clear' as const,
      strings: en.ribbonMenu,
    });
    const open = createSpreadsheetStore();
    setHyperlink(open, A1, 'https://example.com/', wb);
    select(open, area(0, 0, 0, 0));
    expect(executeRibbonHyperlinkAction(deps(open)).kind).toBe('cleared');

    const locked = createSpreadsheetStore();
    setHyperlink(locked, A1, 'https://example.com/', wb);
    select(locked, area(0, 0, 0, 0));
    setProtectedSheet(locked, 0, true);
    expect(executeRibbonHyperlinkAction(deps(locked)).kind).toBe('refused');
    expect(hyperlinkAt(locked.getState(), A1)).not.toBeNull();
  });
});

describe('comments', () => {
  it('Delete Comment asks first, then deletes; a wrong password keeps it', async () => {
    const store = createSpreadsheetStore();
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    setComment(store, A1, 'note', wb);
    select(store, area(0, 0, 0, 0));
    const ui = guard(store, area(0, 0, 0, 0), ['bad', PASSWORD]);
    const deps = {
      store,
      workbook: wb,
      history: new History(),
      action: 'delete-active' as const,
    };
    executeRibbonCommentAction(deps);
    await settle();
    expect(ui.reportIncorrect).toHaveBeenCalledTimes(1);
    expect(commentAt(store.getState(), A1)).toBe('note');
    executeRibbonCommentAction(deps);
    await settle();
    expect(commentAt(store.getState(), A1)).toBeNull();
  });
});

describe('mounted ribbon commands', () => {
  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('Clear Validation Rules asks first', async () => {
    const { mountStubSheet } = await import('../../test-utils/mount.js');
    const { createDataDropdownDefaults } = await import(
      '../../../src/mount/dynamic-dropdown-defaults/data.js'
    );
    const sheet = await mountStubSheet({ locale: 'en' });
    try {
      const { store } = sheet.instance;
      select(store, area(0, 0, 0, 0));
      const ui = guard(store, area(0, 0, 0, 0), [PASSWORD]);
      createDataDropdownDefaults(sheet.instance).applyDataValidationAction('clear-rules');
      await settle();
      expect(ui.askPassword).toHaveBeenCalledTimes(1);
    } finally {
      sheet.dispose();
    }
  });
});

describe('date and time shortcuts (Ctrl+; and Ctrl+Shift+:)', () => {
  const setup = async () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    wb.attachHistory(history);
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
      invalidate: vi.fn(),
      getLocale: () => 'en-US',
      namedRangeDialog: () => null,
      pasteSpecialDialog: () => null,
      quickAnalysis: () => null,
      store,
      strings: () => en,
      wb: () => wb,
    });
    select(store, area(0, 0, 0, 0));
    const press = (init: KeyboardEventInit) =>
      handler(new KeyboardEvent('keydown', { ctrlKey: true, cancelable: true, ...init }));
    return { store, wb, press };
  };
  const kind = (wb: WorkbookHandle) => wb.getValue(A1).kind;

  it.each([
    ['date', { key: ';' }],
    ['time', { key: ':', shiftKey: true }],
  ] as const)('leaves a locked cell on a protected sheet unchanged (%s)', async (_name, init) => {
    const { store, wb, press } = await setup();
    setProtectedSheet(store, 0, true);
    press(init);
    await settle();
    expect(kind(wb)).toBe('blank');
  });

  it.each([
    ['date', { key: ';' }],
    ['time', { key: ':', shiftKey: true }],
  ] as const)('still writes on an unprotected sheet (%s)', async (_name, init) => {
    const { wb, press } = await setup();
    press(init);
    expect(kind(wb)).toBe('number');
  });

  it('asks for the password of a guarded range, then writes', async () => {
    const { store, wb, press } = await setup();
    const ui = guard(store, area(0, 0, 0, 0), ['bad', PASSWORD]);
    press({ key: ';' });
    await settle();
    expect(ui.reportIncorrect).toHaveBeenCalledTimes(1);
    expect(kind(wb)).toBe('blank');
    press({ key: ';' });
    await settle();
    expect(kind(wb)).toBe('number');
  });
});
