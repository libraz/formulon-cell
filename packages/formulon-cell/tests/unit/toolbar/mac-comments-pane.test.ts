import { afterEach, describe, expect, it } from 'vitest';
import { addThreadedComment } from '../../../src/commands/comment.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { Spreadsheet } from '../../../src/mount.js';
import { mutators } from '../../../src/store/store.js';
import { MAC_BUNDLE_LABELS } from '../../../src/toolbar/ribbon/mac/labels-bundle.js';
import { MAC_LOCALES } from '../../../src/toolbar/ribbon/mac/labels-data.js';
import { projectMacRibbonState } from '../../../src/toolbar/ribbon/mac/state-projection.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';
import { stubHelpers } from '../mount/toolbar/fixtures.js';

const ID = 'mac.review.showComments';

describe('Mac Show Comments toggle', () => {
  let sheet: MountedStubSheet | undefined;
  afterEach(() => {
    sheet?.dispose();
    sheet = undefined;
    document.body.innerHTML = '';
  });

  const pane = () =>
    sheet?.instance.host.querySelector<HTMLElement>('.fc-host__taskpane-dock .fc-commentspane');

  it('opens the Comments pane on press and closes it on the next press', async () => {
    sheet = await mountStubSheet({
      locale: 'en',
      ui: { platform: 'mac' },
      workbook: await WorkbookHandle.createDefault(),
    });
    const { instance } = sheet;
    const host = document.createElement('div');
    document.body.appendChild(host);
    const tb = Spreadsheet.mountToolbar(host, instance, { helpers: stubHelpers() });
    expect(pane()?.hidden).toBe(true);
    expect(tb.applyCommand(ID)).toBe(true);
    expect(instance.store.getState().ui.commentsPaneOpen).toBe(true);
    expect(pane()?.hidden).toBe(false);
    expect(document.querySelector('.fc-tb__dlg')).toBeNull();
    expect(tb.applyCommand(ID)).toBe(true);
    expect(instance.store.getState().ui.commentsPaneOpen).toBe(false);
    expect(pane()?.hidden).toBe(true);
    tb.dispose();
  });

  it('draws the control pressed while the pane is open, however it was opened or closed', async () => {
    sheet = await mountStubSheet({ locale: 'en', ui: { platform: 'mac' } });
    const { instance } = sheet;
    const host = document.createElement('div');
    const button = document.createElement('button');
    button.dataset.ribbonCommand = ID;
    host.appendChild(button);
    const pressed = () => {
      projectMacRibbonState(host, instance);
      return button.getAttribute('aria-pressed');
    };
    expect(pressed()).toBe('false');
    mutators.setCommentsPaneOpen(instance.store, true);
    expect(pressed()).toBe('true');
    pane()?.querySelector<HTMLButtonElement>('.fc-commentspane__close')?.click();
    expect(instance.store.getState().ui.commentsPaneOpen).toBe(false);
    expect(pressed()).toBe('false');
  });

  it.each(MAC_LOCALES)('titles the pane with the bundle wording in %s', async (locale) => {
    const title = MAC_BUNDLE_LABELS['pane:comments']?.[locale];
    expect(title).toBeTruthy();
    sheet = await mountStubSheet({ locale, ui: { platform: 'mac' } });
    mutators.setCommentsPaneOpen(sheet.instance.store, true);
    expect(pane()?.querySelector('.fc-commentspane__title')?.textContent).toBe(title);
    expect(pane()?.getAttribute('aria-label')).toBe(title);
  });

  it('reads the bundle empty notice when the file has no comments', async () => {
    sheet = await mountStubSheet({ locale: 'ja', ui: { platform: 'mac' } });
    mutators.setCommentsPaneOpen(sheet.instance.store, true);
    expect(pane()?.querySelector('.fc-commentspane__empty')?.textContent).toBe(
      'このファイル内にコメントはありません。',
    );
    expect(pane()?.querySelectorAll('.fc-commentspane__item').length).toBe(0);
  });

  it('lists each thread under its cell', async () => {
    const workbook = await WorkbookHandle.createDefault();
    sheet = await mountStubSheet({ locale: 'en', ui: { platform: 'mac' }, workbook });
    const { instance } = sheet;
    addThreadedComment(instance.store, { sheet: 0, row: 1, col: 1 }, 'hello', workbook);
    mutators.setCommentsPaneOpen(instance.store, true);
    const items = [...(pane()?.querySelectorAll('.fc-commentspane__item') ?? [])];
    expect(items.map((item) => item.textContent)).toEqual([expect.stringContaining('B2')]);
    expect(items[0]?.textContent).toContain('hello');
    expect(pane()?.querySelector('.fc-commentspane__empty')).toBeNull();
  });
});
