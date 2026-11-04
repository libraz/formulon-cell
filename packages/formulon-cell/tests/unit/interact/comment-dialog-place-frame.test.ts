import { afterEach, describe, expect, it, vi } from 'vitest';
import { attachCommentDialog } from '../../../src/interact/comment-dialog.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('comment dialog repositioning', () => {
  afterEach(() => {
    vi.restoreAllMocks();
    document.body.innerHTML = '';
  });

  const setup = () => {
    const host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    const frames: FrameRequestCallback[] = [];
    vi.spyOn(window, 'requestAnimationFrame').mockImplementation((cb) => frames.push(cb));
    const cancel = vi.spyOn(window, 'cancelAnimationFrame').mockImplementation(() => {});
    const handle = attachCommentDialog({ host, store: createSpreadsheetStore() });
    handle.open();
    const wrap = document.querySelector('.fc-cmtnote') as HTMLElement;
    frames.length = 0;
    const rect = vi.spyOn(wrap, 'getBoundingClientRect');
    return { handle, frames, cancel, wrap, rect };
  };

  it('coalesces a burst of scroll events into one frame', () => {
    const { handle, frames, rect } = setup();
    rect.mockClear();
    for (let i = 0; i < 5; i += 1) window.dispatchEvent(new Event('scroll'));
    expect(frames).toHaveLength(1);
    expect(rect).not.toHaveBeenCalled();
    handle.detach();
  });

  it('ignores scroll that originates inside the panel', () => {
    const { handle, frames, wrap } = setup();
    wrap.dispatchEvent(new Event('scroll'));
    expect(frames).toHaveLength(0);
    handle.detach();
  });

  it('cancels a pending frame on close', () => {
    const { handle, cancel } = setup();
    window.dispatchEvent(new Event('resize'));
    handle.close();
    expect(cancel).toHaveBeenCalled();
    handle.detach();
  });
});
