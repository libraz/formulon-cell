import { afterEach, describe, expect, it } from 'vitest';
import { commentAt } from '../../../src/commands/comment.js';
import { MAX_COL, MAX_ROW } from '../../../src/engine/address.js';
import { attachCommentDialog } from '../../../src/interact/comment-dialog.js';
import { cellRect, layoutForView } from '../../../src/render/geometry.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const NOTE_W = 104;
const NOTE_H = 76;

describe('comment dialog new-note shape', () => {
  afterEach(() => {
    document.body.innerHTML = '';
  });

  const open = (
    platform: string | null,
    row: number,
    col: number,
    zoom = 1,
    screenScale = 0.75,
  ) => {
    const host = document.createElement('div');
    if (platform) host.dataset.fcPlatform = platform;
    Object.defineProperty(host, 'clientWidth', { value: 4000 });
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, screenScale },
      viewport: {
        ...s.viewport,
        zoom,
        rowStart: Math.max(0, row - 10),
        colStart: Math.max(0, col - 3),
      },
    }));
    mutators.setActive(store, { sheet: 0, row, col });
    const handle = attachCommentDialog({ host, store });
    handle.open({ mode: 'note' });
    const s = store.getState();
    const r = cellRect(layoutForView(s), s.viewport, row, col);
    const wrap = document.querySelector('.fc-cmtnote') as HTMLElement;
    return { handle, wrap, r, store };
  };

  it('sizes the Mac note 104 x 76 and puts it right of the cell, 10 up', () => {
    const { handle, wrap, r } = open('mac', 20, 5);
    expect(wrap.style.width).toBe(`${NOTE_W}px`);
    expect(wrap.style.height).toBe(`${NOTE_H}px`);
    expect(wrap.style.left).toBe(`${Math.round(r.x + r.w + 15)}px`);
    expect(wrap.style.top).toBe(`${Math.round(r.y - 10)}px`);
    expect(wrap.style.fontSize).toBe('10px');
    handle.detach();
  });

  it('scales the Mac note with the zoom', () => {
    const { handle, wrap, r } = open('mac', 20, 5, 2);
    expect(wrap.style.width).toBe(`${NOTE_W * 2}px`);
    expect(wrap.style.height).toBe(`${NOTE_H * 2}px`);
    expect(wrap.style.left).toBe(`${Math.round(r.x + r.w + 30)}px`);
    expect(wrap.style.top).toBe(`${Math.round(r.y - 20)}px`);
    expect(wrap.style.fontSize).toBe('20px');
    handle.detach();
  });

  it('flips up at the last row so the note stays on the sheet', () => {
    const { handle, wrap, r } = open('mac', MAX_ROW, 5);
    expect(wrap.style.top).toBe(`${Math.round(r.y + r.h - 88)}px`);
    handle.detach();
  });

  it('flips left at the last column', () => {
    const { handle, wrap, r } = open('mac', 20, MAX_COL);
    expect(wrap.style.left).toBe(`${Math.round(r.x - NOTE_W - 15)}px`);
    handle.detach();
  });

  it('keeps 1 pt = 1 browser px at the Mac screen scale', () => {
    const { handle, wrap } = open('mac', 20, 5, 1, 0.75);
    expect(wrap.style.width).toBe('104px');
    expect(wrap.style.height).toBe('76px');
    expect(wrap.style.fontSize).toBe('10px');
    handle.detach();
  });

  it('commits a Mac note to its own cell when a grid click has moved the selection', async () => {
    const { handle, store } = open('mac', 2, 2);
    const ta = document.querySelector<HTMLTextAreaElement>('.fc-cmtnote__textarea');
    if (!ta) throw new Error('no textarea');
    ta.value = 'grid click';
    ta.dispatchEvent(new Event('input'));
    await new Promise((resolve) => setTimeout(resolve, 0));
    mutators.setActive(store, { sheet: 0, row: 15, col: 2 });
    document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    expect(commentAt(store.getState(), { sheet: 0, row: 2, col: 2 })).toBe('grid click');
    expect(commentAt(store.getState(), { sheet: 0, row: 15, col: 2 })).toBeNull();
    handle.detach();
  });

  it('leaves the default profile on the fixed chrome', () => {
    const { handle, wrap } = open(null, 20, 5);
    expect(wrap.style.width).toBe('240px');
    expect(wrap.style.height).toBe('');
    expect(wrap.style.fontSize).toBe('');
    handle.detach();
  });
});
