import { afterEach, describe, expect, it } from 'vitest';
import { addThreadedComment } from '../../../src/commands/comment.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

describe('Mac Comments pane navigation policy', () => {
  let sheet: MountedStubSheet | undefined;
  afterEach(() => {
    sheet?.dispose();
    sheet = undefined;
    document.body.innerHTML = '';
  });

  const pane = () =>
    sheet?.instance.host.querySelector<HTMLElement>('.fc-host__taskpane-dock .fc-commentspane');

  const mount = async (selectable: (addr: { row: number }) => boolean) => {
    const workbook = await WorkbookHandle.createDefault();
    sheet = await mountStubSheet({
      locale: 'en',
      ui: { platform: 'mac' },
      workbook,
      viewport: { selectable },
    });
    return { store: sheet.instance.store, workbook };
  };

  it('lists only threads on cells the navigation policy allows', async () => {
    const { store, workbook } = await mount((addr) => addr.row !== 1);
    addThreadedComment(store, { sheet: 0, row: 1, col: 1 }, 'blocked', workbook);
    addThreadedComment(store, { sheet: 0, row: 4, col: 2 }, 'allowed', workbook);
    mutators.setCommentsPaneOpen(store, true);
    const items = [...(pane()?.querySelectorAll('.fc-commentspane__item') ?? [])];
    expect(items.map((item) => item.textContent)).toEqual([expect.stringContaining('C5')]);
    expect(pane()?.textContent).not.toContain('blocked');
  });

  it('shows the empty state when every thread is filtered out', async () => {
    const { store, workbook } = await mount(() => false);
    addThreadedComment(store, { sheet: 0, row: 1, col: 1 }, 'blocked', workbook);
    mutators.setCommentsPaneOpen(store, true);
    expect(pane()?.querySelectorAll('.fc-commentspane__item').length).toBe(0);
    expect(pane()?.querySelector('.fc-commentspane__empty')).not.toBeNull();
  });

  it('re-filters the open pane when the host narrows navigation', async () => {
    const { store, workbook } = await mount(() => true);
    addThreadedComment(store, { sheet: 0, row: 3, col: 2 }, 'kept', workbook);
    addThreadedComment(store, { sheet: 0, row: 7, col: 2 }, 'dropped', workbook);
    mutators.setCommentsPaneOpen(store, true);
    expect(pane()?.querySelectorAll('.fc-commentspane__item').length).toBe(2);
    sheet?.instance.setViewportOptions({ selectable: (addr: { row: number }) => addr.row !== 7 });
    const items = [...(pane()?.querySelectorAll('.fc-commentspane__item') ?? [])];
    expect(items.map((item) => item.textContent)).toEqual([expect.stringContaining('C4')]);
  });
});
