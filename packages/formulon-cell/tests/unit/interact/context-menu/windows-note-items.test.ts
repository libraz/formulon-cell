import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { en } from '../../../../src/i18n/strings/en.js';
import {
  attachContextMenu,
  type ContextMenuHandle,
} from '../../../../src/interact/context-menu.js';
import { disposeOverlayPortal } from '../../../../src/interact/overlay-portal.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../../src/store/store.js';
import { fireContextMenu, item, setFormat } from './fixtures.js';

const label = (id: string): string | null | undefined =>
  item(id)?.querySelector('.fc-ctxmenu__label')?.textContent;

// Windows-unmeasured: the cell menu follows the Mac's note rows.
describe('cell menu note rows on the default profile', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let detach: ContextMenuHandle | null = null;

  beforeEach(async () => {
    host = document.createElement('div');
    host.tabIndex = -1;
    host.classList.add('fc-host');
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await WorkbookHandle.createDefault();
  });

  afterEach(() => {
    detach?.();
    detach = null;
    disposeOverlayPortal(host);
    document.body.replaceChildren();
    wb.dispose();
  });

  const attach = (): void => {
    detach = attachContextMenu({
      getLocale: () => 'en-US',
      host,
      store,
      wb,
      strings: en,
      onEditComment: () => {},
    });
  };

  it('offers New Note alone on a cell without a note', () => {
    attach();
    fireContextMenu(host, 200, 70);
    expect(label('insertComment')).toBe('New Note');
    expect(item('deleteComment')).toBeNull();
    expect(item('showHideNote')).toBeNull();
  });

  it('offers Edit, Delete and Show/Hide Note on a cell with a note', () => {
    setFormat(store, 0, 0, { comment: 'memo' });
    attach();
    fireContextMenu(host, 200, 70);
    expect([label('insertComment'), label('deleteComment'), label('showHideNote')]).toEqual([
      'Edit Note',
      'Delete Note',
      'Show/Hide Note',
    ]);
  });
});
