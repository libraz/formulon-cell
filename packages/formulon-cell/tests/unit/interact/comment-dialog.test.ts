import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { commentAt, setComment } from '../../../src/commands/comment.js';
import { attachCommentDialog } from '../../../src/interact/comment-dialog.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

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
});
