import { describe, expect, it } from 'vitest';
import {
  addThreadedComment,
  editThreadedComment,
  hydrateThreadedComments,
  listThreadedComments,
  removeThreadedComment,
  resolveThread,
} from '../../../src/commands/comment.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';
const bytesFromBase64 = (value: string): Uint8Array =>
  Uint8Array.from(atob(value), (char) => char.charCodeAt(0));

// 1×1 transparent PNG.
const PNG = bytesFromBase64(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
);

describe.skipIf(!canLoadWasm())('formulon 0.13 annotation and image round-trip', () => {
  it('does not reuse hydrated threads after the store slice is reset', () => {
    const store = createSpreadsheetStore();
    hydrateThreadedComments(store, 0, [
      {
        id: '{88888888-8888-4888-8888-888888888888}',
        row: 0,
        col: 0,
        personId: '{99999999-9999-4999-8999-999999999999}',
        created: '2026-01-01T10:00:00.00',
        text: 'stale',
        parentId: '',
        done: false,
        mentions: [],
      },
    ]);
    expect(listThreadedComments(store, 0)).toHaveLength(1);
    store.setState((state) => ({
      ...state,
      threadedComments: { bySheet: new Map(), persons: [] },
    }));
    expect(listThreadedComments(store, 0)).toEqual([]);
  });

  it('keeps threaded comment ids, replies, and resolution through save/load', async () => {
    const first = await WorkbookHandle.createDefault();
    try {
      expect(first.isStub).toBe(false);
      expect(first.capabilities.threadedComments).toBe(true);
      const personId = '{11111111-1111-4111-8111-111111111111}';
      const threadId = '{22222222-2222-4222-8222-222222222222}';
      const replyId = '{33333333-3333-4333-8333-333333333333}';
      expect(
        first.addPerson({ id: personId, displayName: 'Alice', userId: '', providerId: 'None' }),
      ).toBe(true);
      expect(
        first.addThreadedComment(0, {
          id: threadId,
          row: 2,
          col: 3,
          personId,
          created: '2026-01-01T10:00:00.00',
          text: 'Review this',
        }),
      ).toBe(true);
      expect(
        first.addThreadedComment(0, {
          id: replyId,
          personId,
          created: '2026-01-01T10:01:00.00',
          text: 'Done',
          parentId: threadId,
        }),
      ).toBe(true);
      expect(first.setThreadResolved(0, threadId, true)).toBe(true);

      const reloaded = await WorkbookHandle.loadBytes(first.save());
      try {
        const comments = reloaded.getThreadedComments(0);
        expect(comments.map((comment) => comment.id)).toEqual([threadId, replyId]);
        expect(comments[0]).toMatchObject({
          row: 2,
          col: 3,
          text: 'Review this',
          done: true,
        });
        expect(comments[1]).toMatchObject({ parentId: threadId, text: 'Done' });
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('keeps inserted image bytes and anchor metadata through save/load', async () => {
    const first = await WorkbookHandle.createDefault();
    try {
      expect(first.isStub).toBe(false);
      expect(first.capabilities.images).toBe(true);
      expect(first.probeImage(PNG)).toMatchObject({ pxWidth: 1, pxHeight: 1 });
      const objectId = first.insertImage(0, PNG, {
        row: 4,
        col: 2,
        name: 'logo',
        descr: 'Logo',
      });
      expect(objectId).not.toBeNull();
      const reloaded = await WorkbookHandle.loadBytes(first.save());
      try {
        const object = reloaded.listDrawingObjects(0).find((entry) => entry.objectId > 0);
        expect(object).toMatchObject({ fromRow: 4, fromCol: 2, name: 'logo', descr: 'Logo' });
        expect(object).toBeDefined();
        const image = reloaded.getImage(0, object?.objectId ?? 0);
        expect(image?.bytes).toEqual(PNG);
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('preserves the original annotation when notes and threads share a cell', async () => {
    const first = await WorkbookHandle.createDefault();
    try {
      expect(first.isStub).toBe(false);
      expect(first.capabilities.threadedComments).toBe(true);
      const personId = '{44444444-4444-4444-8444-444444444444}';
      const threadId = '{55555555-5555-4555-8555-555555555555}';
      expect(
        first.addPerson({ id: personId, displayName: 'Bob', userId: '', providerId: 'None' }),
      ).toBe(true);
      expect(first.setCommentEntry(0, 1, 1, 'Bob', 'legacy note')).toBe(true);
      expect(
        first.addThreadedComment(0, {
          id: threadId,
          row: 1,
          col: 1,
          personId,
          created: '2026-01-01T10:00:00.00',
          text: 'thread attempt',
        }),
      ).toBe(false);
      expect(first.getComment(0, 1, 1)).toEqual({ author: 'Bob', text: 'legacy note' });

      expect(first.setCommentEntry(0, 1, 1, 'Bob', '')).toBe(true);
      expect(
        first.addThreadedComment(0, {
          id: threadId,
          row: 1,
          col: 1,
          personId,
          created: '2026-01-01T10:00:00.00',
          text: 'thread',
        }),
      ).toBe(true);
      expect(first.setCommentEntry(0, 1, 1, 'Bob', 'note attempt')).toBe(false);
      const reloaded = await WorkbookHandle.loadBytes(first.save());
      try {
        expect(reloaded.getThreadedComments(0).map((comment) => comment.id)).toEqual([threadId]);
        expect(reloaded.getComment(0, 1, 1)).toBeNull();
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('does not mutate a protected threaded comment after the dialog opened', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.isStub).toBe(false);
      expect(wb.capabilities.threadedComments).toBe(true);
      const store = createSpreadsheetStore();
      const personId = '{66666666-6666-4666-8666-666666666666}';
      const threadId = '{77777777-7777-4777-8777-777777777777}';
      expect(
        wb.addPerson({ id: personId, displayName: 'Carol', userId: '', providerId: 'None' }),
      ).toBe(true);
      expect(
        wb.addThreadedComment(0, {
          id: threadId,
          row: 3,
          col: 4,
          personId,
          created: '2026-01-01T10:00:00.00',
          text: 'Protected thread',
        }),
      ).toBe(true);
      hydrateThreadedComments(store, 0, wb.getThreadedComments(0));
      setProtectedSheet(store, 0, true);
      const addr = { sheet: 0, row: 3, col: 4 };
      expect(addThreadedComment(store, addr, 'blocked', wb)).toBeNull();
      expect(editThreadedComment(store, 0, threadId, 'blocked', wb)).toBe(false);
      expect(resolveThread(store, 0, threadId, true, wb)).toBe(false);
      expect(removeThreadedComment(store, 0, threadId, wb)).toBe(false);
      expect(wb.getThreadedComments(0)[0]?.text).toBe('Protected thread');
      expect(wb.getThreadedComments(0)[0]?.done).toBe(false);
    } finally {
      wb.dispose();
    }
  });
});
