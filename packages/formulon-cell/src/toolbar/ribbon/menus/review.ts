// Review tab menus: notes, watch window, sheet/workbook protection. Each
// factory builds shared icon menu rows.

import type { ToolbarMenuText } from '../../menu-text.js';
import type { ToolbarText } from '../../ribbon-model.js';

import { createMenu, menuIconButton, menuIdForCommand, menuSeparator } from './general.js';

export interface ReviewMenuFactories {
  createWatchMenu: (id: string) => HTMLDivElement;
  createReviewNotesMenu: () => HTMLDivElement;
  createProtectMenu: (id: string) => HTMLDivElement;
}

export const createReviewMenuFactories = (
  ribbonMenuText: ToolbarMenuText,
  ribbonText: ToolbarText,
): ReviewMenuFactories => {
  const t = ribbonMenuText;
  const r = ribbonText;

  const createWatchMenu = (id: string): HTMLDivElement => {
    const menu = createMenu(menuIdForCommand(id));
    menu.append(
      menuIconButton(t.watchWindow, 'watchAction', 'open', 'watch-open'),
      menuIconButton(t.watchAdd, 'watchAction', 'add', 'watch-add'),
      menuIconButton(t.watchDelete, 'watchAction', 'delete', 'watch-delete'),
      menuSeparator(),
      menuIconButton(t.watchDeleteAll, 'watchAction', 'delete-all', 'watch-delete-all'),
    );
    return menu;
  };

  /** Review > Notes: New Note (Edit Note on a note cell), note navigation, Show/Hide Note, Show All
   *  Notes and Convert to Comments. */
  const createReviewNotesMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-review-notes');
    menu.append(
      menuIconButton(r.newNote, 'commentAction', 'new-note', 'comment-add'),
      menuIconButton(r.previousNote, 'commentAction', 'previous-note', 'comment-previous'),
      menuIconButton(r.nextNote, 'commentAction', 'next-note', 'comment-next'),
      menuIconButton(r.showHideNote, 'commentAction', 'toggle-note', 'note-show-hide'),
      menuIconButton(r.showAllNotes, 'commentAction', 'toggle-all-notes', 'note-show-all'),
      menuIconButton(r.convertNotes, 'commentAction', 'convert-notes', 'comment-convert'),
    );
    return menu;
  };

  const createProtectMenu = (id: string): HTMLDivElement => {
    const menu = createMenu(menuIdForCommand(id));
    menu.append(
      menuIconButton(t.protectSheetCommand, 'protectAction', 'protect-sheet', 'protect-sheet'),
      menuIconButton(
        t.unprotectSheetCommand,
        'protectAction',
        'unprotect-sheet',
        'protect-unprotect-sheet',
      ),
    );
    if (id === 'protect') return menu;
    menu.append(
      menuSeparator(),
      menuIconButton(t.lockCell, 'protectAction', 'lock-cell', 'protect-lock-cell'),
      menuIconButton(t.unlockCell, 'protectAction', 'unlock-cell', 'protect-unlock-cell'),
      menuSeparator(),
      menuIconButton(
        t.protectWorkbookCommand,
        'protectAction',
        'protect-workbook',
        'protect-workbook',
      ),
      menuIconButton(
        t.unprotectWorkbookCommand,
        'protectAction',
        'unprotect-workbook',
        'protect-unprotect-workbook',
      ),
      menuIconButton(
        t.allowEditRangesCommand,
        'protectAction',
        'allow-edit-ranges',
        'protect-allow-ranges',
      ),
      menuIconButton(
        t.allowEditRangesClearCommand,
        'protectAction',
        'clear-allowed-edit-ranges',
        'protect-clear-ranges',
      ),
    );
    return menu;
  };

  return { createWatchMenu, createReviewNotesMenu, createProtectMenu };
};
