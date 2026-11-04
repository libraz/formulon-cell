import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import type { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { ja } from '../../../../src/i18n/strings/ja.js';
import {
  attachContextMenu,
  type ContextMenuHandle,
} from '../../../../src/interact/context-menu.js';
import {
  currentDialogAnchor,
  type DialogAnchor,
} from '../../../../src/interact/dialog-placement.js';
import { disposeOverlayPortal } from '../../../../src/interact/overlay-portal.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../../src/store/store.js';
import { fireContextMenu, item, newWb } from './fixtures.js';

// Excel for Mac: a dialog opened from a cell context menu centres on the clicked row.
describe('context menu dialog anchor', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let detach: ContextMenuHandle | null;
  let seen: (DialogAnchor | null)[];

  beforeEach(async () => {
    host = document.createElement('div');
    host.tabIndex = -1;
    host.classList.add('fc-host');
    host.dataset.fcPlatform = 'mac';
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await newWb();
    seen = [];
    detach = attachContextMenu({
      host,
      store,
      wb,
      strings: ja,
      onInsertHyperlink: () => {
        seen.push(currentDialogAnchor());
      },
    });
  });

  afterEach(() => {
    detach?.();
    disposeOverlayPortal(host);
    document.body.replaceChildren();
  });

  const clickRow = (id: string, rect: Partial<DOMRect>): void => {
    fireContextMenu(host, 100, 100);
    const row = item(id);
    if (!row) throw new Error(`missing ${id}`);
    row.getBoundingClientRect = () =>
      ({ left: 0, top: 0, width: 0, height: 0, ...rect }) as DOMRect;
    row.click();
  };

  it('hands the clicked row centre to the dialog the item opens', () => {
    clickRow('insertHyperlink', { left: 900, top: 600, width: 280, height: 30 });
    expect(seen).toEqual([{ x: 1040, y: 615 }]);
  });

  it('leaves no anchor behind once the dispatch returns', () => {
    clickRow('insertHyperlink', { left: 900, top: 600, width: 280, height: 30 });
    expect(currentDialogAnchor()).toBeNull();
  });

  it('leaves no anchor for a later dialog when the item opens none', () => {
    clickRow('copy', { left: 900, top: 600, width: 280, height: 30 });
    expect(currentDialogAnchor()).toBeNull();
    // A dialog opened later, from the ribbon, sees no anchor and centres on the window.
    seen.push(currentDialogAnchor());
    expect(seen).toEqual([null]);
  });
});
