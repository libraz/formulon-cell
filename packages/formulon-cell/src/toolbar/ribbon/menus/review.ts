// Review tab menus: comments, watch window, sheet/workbook protection. Each
// factory builds shared icon menu rows extracted from main.ts.

import type { ToolbarMenuText } from '../../menu-text.js';

import { createMenu, menuIconButton, menuIdForCommand, menuSeparator } from './general.js';

export interface ReviewMenuFactories {
  createWatchMenu: (id: string) => HTMLDivElement;
  createReviewCommentsMenu: (id?: string) => HTMLDivElement;
  createProtectMenu: (id: string) => HTMLDivElement;
}

export const createReviewMenuFactories = (ribbonMenuText: ToolbarMenuText): ReviewMenuFactories => {
  const t = ribbonMenuText;

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

  const createReviewCommentsMenu = (id = 'menu-review-comments'): HTMLDivElement => {
    const menu = createMenu(id);
    menu.append(menuIconButton(t.commentThread, 'commentAction', 'thread', 'comment-add'));
    if (id === 'menu-review-new-comment') return menu;
    menu.append(
      menuIconButton(t.commentDelete, 'commentAction', 'delete-active', 'comment-delete'),
      menuIconButton(t.commentDeleteAll, 'commentAction', 'delete-all', 'comment-delete-all'),
    );
    return menu;
  };

  const createProtectMenu = (id: string): HTMLDivElement => {
    const menu = createMenu(menuIdForCommand(id));
    if (id === 'protect') {
      menu.append(
        menuIconButton(t.protectSheetCommand, 'protectAction', 'protect-sheet', 'protect-sheet'),
        menuIconButton(
          t.unprotectSheetCommand,
          'protectAction',
          'unprotect-sheet',
          'protect-unprotect-sheet',
        ),
      );
      return menu;
    }
    menu.append(
      menuIconButton(t.protectSheetCommand, 'protectAction', 'protect-sheet', 'protect-sheet'),
      menuIconButton(
        t.unprotectSheetCommand,
        'protectAction',
        'unprotect-sheet',
        'protect-unprotect-sheet',
      ),
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

  return { createWatchMenu, createReviewCommentsMenu, createProtectMenu };
};
