import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { commentAt, setComment } from '../../../src/commands/comment.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { viewerPolicy } from '../../../src/commands/interaction-policy.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachCommentDialog } from '../../../src/interact/comment-dialog.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const wrap = (): HTMLElement | null => document.querySelector<HTMLElement>('.fc-cmtnote');
const textarea = (): HTMLTextAreaElement | null =>
  document.querySelector<HTMLTextAreaElement>('.fc-cmtnote__textarea');
const okBtn = (): HTMLButtonElement | null =>
  document.querySelector<HTMLButtonElement>('.fc-cmtnote__btn--primary');
const cancelBtn = (): HTMLButtonElement | null =>
  Array.from(document.querySelectorAll<HTMLButtonElement>('.fc-cmtnote__btn')).find(
    (b) => !b.classList.contains('fc-cmtnote__btn--primary'),
  ) ?? null;
const removeBtn = (): HTMLButtonElement | null =>
  document.querySelector<HTMLButtonElement>('.fc-cmtnote__icon');
const threadMeta = (): HTMLElement | null =>
  document.querySelector<HTMLElement>('.fc-cmtnote__thread-meta');
const dialogError = (): HTMLElement | null =>
  document.querySelector<HTMLElement>('.fc-cmtnote__error');

describe('attachCommentDialog', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 1, col: 1 });
  });

  afterEach(() => {
    while (document.body.firstChild) document.body.removeChild(document.body.firstChild);
  });

  it('open with no existing comment shows a blank textarea and hides the remove button', () => {
    const handle = attachCommentDialog({ host, store });
    handle.open();
    expect(wrap()?.hidden).toBe(false);
    expect(textarea()?.value).toBe('');
    expect(textarea()?.getAttribute('aria-label')).toBeTruthy();
    expect(removeBtn()?.hidden).toBe(true);
    handle.detach();
  });

  it('open populates the textarea from the existing comment and shows the remove button', () => {
    setComment(store, { sheet: 0, row: 1, col: 1 }, 'preexisting note');
    const handle = attachCommentDialog({ host, store });
    handle.open();
    expect(textarea()?.value).toBe('preexisting note');
    expect(removeBtn()?.hidden).toBe(false);
    handle.detach();
  });

  it('OK with non-empty text persists via setComment to the store format slice', () => {
    const handle = attachCommentDialog({ host, store });
    handle.open();
    const ta = textarea();
    if (!ta) throw new Error('textarea missing');
    ta.value = 'fresh note';
    okBtn()?.click();
    expect(commentAt(store.getState(), { sheet: 0, row: 1, col: 1 })).toBe('fresh note');
    expect(wrap()?.hidden).toBe(true);
    handle.detach();
  });

  it('remove icon clears the existing comment via clearComment', () => {
    setComment(store, { sheet: 0, row: 1, col: 1 }, 'will be wiped');
    const handle = attachCommentDialog({ host, store });
    handle.open();
    expect(commentAt(store.getState(), { sheet: 0, row: 1, col: 1 })).toBe('will be wiped');
    removeBtn()?.click();
    expect(commentAt(store.getState(), { sheet: 0, row: 1, col: 1 })).toBeNull();
    expect(wrap()?.hidden).toBe(true);
    handle.detach();
  });

  it('opens native threads with localized labels and person display names', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.isStub).toBe(false);
      expect(wb.capabilities.threadedComments).toBe(true);
      const personId = '{AAAAAAAA-AAAA-4AAA-8AAA-AAAAAAAAAAAA}';
      const threadId = '{BBBBBBBB-BBBB-4BBB-8BBB-BBBBBBBBBBBB}';
      expect(
        wb.addPerson({ id: personId, displayName: 'Alice', userId: null, providerId: 'None' }),
      ).toBe(true);
      expect(
        wb.addThreadedComment(0, {
          id: threadId,
          cell: { row: 1, col: 1 },
          personId,
          created: { year: 2026, month: 1, day: 1, hour: 10, minute: 0, second: 0 },
          text: 'Please review',
        }),
      ).toBe(true);
      const handle = attachCommentDialog({ host, store, getWb: () => wb });
      handle.open({ mode: 'thread' });
      expect(document.querySelector('.fc-cmtnote__title')?.textContent).toBe('コメント');
      expect(textarea()?.placeholder).toBe('コメントを入力');
      expect(threadMeta()?.textContent).toBe('Alice');
      expect(document.querySelector('.fc-cmtnote__reply')?.textContent).toBe('返信');
      expect(document.querySelector('.fc-cmtnote__resolve')?.textContent).toBe('解決');
      handle.detach();
    } finally {
      wb.dispose();
    }
  });

  it('starts a Mac conversation on a floating card with author, address and send', async () => {
    host.classList.add('fc-host');
    host.dataset.fcPlatform = 'mac';
    mutators.setActive(store, { sheet: 0, row: 4, col: 7 });
    const wb = await WorkbookHandle.createDefault();
    try {
      const handle = attachCommentDialog({ host, store, getWb: () => wb });
      handle.open({ mode: 'thread' });
      const card = wrap();
      expect(card?.classList.contains('fc-cmtnote--start')).toBe(true);
      expect(card?.querySelector('.fc-cmtnote__start-name')?.textContent).toBe('Formulon');
      expect(card?.querySelector('.fc-cmtnote__start-addr')?.textContent).toBe('H5');
      expect(card?.querySelector('.fc-cmtnote__start-avatar svg')).toBeTruthy();
      expect(card?.querySelector('.fc-cmtnote__send svg')).toBeTruthy();
      const close = card?.querySelector<HTMLButtonElement>('.fc-cmtnote__start-close');
      expect(close?.getAttribute('aria-label')).toBe('閉じる');
      expect(textarea()?.placeholder).toBe('会話を始める');
      expect(card?.querySelector<HTMLElement>('.fc-cmtnote__head')?.hidden).toBe(true);
      expect(card?.querySelector<HTMLElement>('.fc-cmtnote__footer')?.hidden).toBe(true);
      const send = card?.querySelector<HTMLButtonElement>('.fc-cmtnote__send');
      expect(send?.getAttribute('aria-label')).toBe('投稿');
      expect(send?.disabled).toBe(true);

      const ta = textarea();
      if (!ta || !send) throw new Error('composer missing');
      ta.value = 'Please check';
      ta.dispatchEvent(new Event('input', { bubbles: true }));
      expect(send.disabled).toBe(false);
      send.click();
      expect(wrap()?.hidden).toBe(true);
      const thread = wb.getThreadedComments(0);
      expect(thread.map((entry) => entry.text)).toEqual(['Please check']);
      expect(wb.getPersons().find((p) => p.id === thread[0]?.personId)?.displayName).toBe(
        'Formulon',
      );
      handle.detach();
    } finally {
      wb.dispose();
    }
  });

  it('closes the Mac conversation card from its X without posting', async () => {
    host.classList.add('fc-host');
    host.dataset.fcPlatform = 'mac';
    const wb = await WorkbookHandle.createDefault();
    try {
      const handle = attachCommentDialog({ host, store, getWb: () => wb });
      handle.open({ mode: 'thread' });
      const ta = textarea();
      if (!ta) throw new Error('textarea missing');
      ta.value = 'draft';
      document.querySelector<HTMLButtonElement>('.fc-cmtnote__start-close')?.click();
      expect(wrap()?.hidden).toBe(true);
      expect(wb.getThreadedComments(0)).toEqual([]);
      handle.detach();
    } finally {
      wb.dispose();
    }
  });

  it('keeps the titled thread editor on the default profile', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const handle = attachCommentDialog({ host, store, getWb: () => wb });
      handle.open({ mode: 'thread' });
      expect(wrap()?.classList.contains('fc-cmtnote--start')).toBe(false);
      expect(document.querySelector<HTMLElement>('.fc-cmtnote__head')?.hidden).toBe(false);
      expect(textarea()?.placeholder).toBe('コメントを入力');
      handle.detach();
    } finally {
      wb.dispose();
    }
  });

  it('keeps the native thread editor open when a note blocks creation', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.isStub).toBe(false);
      expect(wb.capabilities.threadedComments).toBe(true);
      setComment(store, { sheet: 0, row: 1, col: 1 }, 'existing note');
      const handle = attachCommentDialog({ host, store, getWb: () => wb });
      handle.open({ mode: 'thread' });
      const ta = textarea();
      if (!ta) throw new Error('textarea missing');
      ta.value = 'new thread';
      document.querySelector<HTMLButtonElement>('.fc-cmtnote__btn--primary')?.click();
      expect(wrap()?.hidden).toBe(false);
      expect(dialogError()?.hidden).toBe(false);
      expect(dialogError()?.textContent).toContain('メモ');
      expect(commentAt(store.getState(), { sheet: 0, row: 1, col: 1 })).toBe('existing note');
      expect(wb.getThreadedComments(0)).toHaveLength(0);
      handle.detach();
    } finally {
      wb.dispose();
    }
  });

  it('rechecks note authorization when the policy changes after opening', async () => {
    const wb = await WorkbookHandle.createDefault();
    const history = new History();
    const controller = new InteractionController({
      store,
      getWb: () => wb,
      history,
    });
    const unregister = registerInteractionController(store, controller);
    try {
      expect(wb.isStub).toBe(false);
      expect(setComment(store, { sheet: 0, row: 1, col: 1 }, 'before', wb)).toBe(true);
      const handle = attachCommentDialog({ host, store, history, getWb: () => wb });
      handle.open({ mode: 'note' });
      controller.setPolicy(viewerPolicy());
      const ta = textarea();
      if (!ta) throw new Error('textarea missing');
      ta.value = 'must stay before';
      okBtn()?.click();
      expect(wrap()?.hidden).toBe(false);
      expect(dialogError()?.hidden).toBe(false);
      expect(commentAt(store.getState(), { sheet: 0, row: 1, col: 1 })).toBe('before');
      expect(wb.getComment(0, 1, 1)).toEqual({ author: 'Formulon', text: 'before' });
      expect(history.canUndo()).toBe(false);
      handle.detach();
    } finally {
      unregister();
      controller.dispose();
      wb.dispose();
    }
  });

  it('Escape closes without writing the typed text', () => {
    const handle = attachCommentDialog({ host, store });
    handle.open();
    const ta = textarea();
    if (!ta) throw new Error('textarea missing');
    ta.value = 'never persisted';
    wrap()?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(wrap()?.hidden).toBe(true);
    expect(commentAt(store.getState(), { sheet: 0, row: 1, col: 1 })).toBeNull();
    handle.detach();
  });

  it('Cancel button closes without writing the typed text', () => {
    const handle = attachCommentDialog({ host, store });
    handle.open();
    const ta = textarea();
    if (!ta) throw new Error('textarea missing');
    ta.value = 'discarded';
    cancelBtn()?.click();
    expect(wrap()?.hidden).toBe(true);
    expect(commentAt(store.getState(), { sheet: 0, row: 1, col: 1 })).toBeNull();
    handle.detach();
  });

  it('close returns focus to the spreadsheet host', () => {
    host.tabIndex = 0;
    const handle = attachCommentDialog({ host, store });
    handle.open();
    cancelBtn()?.click();
    expect(document.activeElement).toBe(host);
    handle.detach();
  });

  it('detach removes the note element from the host', () => {
    const handle = attachCommentDialog({ host, store });
    expect(wrap()).not.toBeNull();
    handle.detach();
    expect(wrap()).toBeNull();
  });

  it('keeps the note title in normal desktop note casing', () => {
    const css = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/comment-note.css'),
      'utf8',
    );

    expect(css).toMatch(/\.fc-cmtnote__title\s*\{[\s\S]*?letter-spacing: 0;/);
    expect(css).not.toContain('text-transform: uppercase');
  });
});
