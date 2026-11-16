import { afterEach, describe, expect, it } from 'vitest';
import { contextMenu } from '../../../src/extensions/built-ins.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/index.js';

const DIALOG_ITEMS = ['formatCells', 'insertHyperlink', 'insertComment', 'defineName'];

/** The cell menu's dialog-opening items, the Paste Special dialog item inside its submenu included. */
const openCellMenu = (sheet: MountedStubSheet): string[] => {
  mutators.setActive(sheet.instance.store, { sheet: 0, row: 2, col: 2 });
  sheet.host.dispatchEvent(
    new MouseEvent('contextmenu', { clientX: 200, clientY: 90, bubbles: true, cancelable: true }),
  );
  document.querySelector<HTMLButtonElement>('[data-fc-submenu="pasteSpecialMenu"]')?.click();
  return [...DIALOG_ITEMS, 'pasteSpecial'].filter(
    (id) =>
      document.querySelector(`.fc-ctxmenu__item[data-fc-action="${id}"]:not([data-fc-submenu])`) !==
      null,
  );
};

describe('context-menu items that open a dialog', () => {
  let sheet: MountedStubSheet | null = null;
  afterEach(() => {
    sheet?.dispose();
    sheet = null;
    document.body.replaceChildren();
  });

  it('show while their dialogs are attached', async () => {
    sheet = await mountStubSheet();
    expect(openCellMenu(sheet)).toEqual([...DIALOG_ITEMS, 'pasteSpecial']);
  });

  it('are left out while their dialogs are not attached', async () => {
    sheet = await mountStubSheet({
      features: {
        formatDialog: false,
        hyperlink: false,
        commentDialog: false,
        namedRanges: false,
        pasteSpecial: false,
      },
    });
    expect(openCellMenu(sheet)).toEqual([]);
  });

  it('are left out by the context-menu extension while their extensions are not registered', async () => {
    sheet = await mountStubSheet({
      features: {
        contextMenu: false,
        formatDialog: false,
        hyperlink: false,
        commentDialog: false,
        namedRanges: false,
        pasteSpecial: false,
      },
      extensions: [contextMenu()],
    });
    expect(openCellMenu(sheet)).toEqual([]);
  });
});
