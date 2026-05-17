// Review tab menus: comments, watch window, sheet/workbook protection. Each
// factory is a static label list extracted from main.ts.

import type { ToolbarMenuText } from '@libraz/formulon-cell';

import { createMenu, menuButton, menuIdForCommand, menuSeparator } from './general.js';

export interface ReviewMenuFactories {
  createWatchMenu: (id: string) => HTMLDivElement;
  createReviewCommentsMenu: () => HTMLDivElement;
  createProtectMenu: (id: string) => HTMLDivElement;
}

export const createReviewMenuFactories = (ribbonMenuText: ToolbarMenuText): ReviewMenuFactories => {
  const t = ribbonMenuText;

  const createWatchMenu = (id: string): HTMLDivElement => {
    const menu = createMenu(menuIdForCommand(id));
    menu.append(
      menuButton(t.watchWindow, 'watchAction', 'open'),
      menuButton(t.watchAdd, 'watchAction', 'add'),
      menuButton(t.watchDelete, 'watchAction', 'delete'),
      menuSeparator(),
      menuButton(t.watchDeleteAll, 'watchAction', 'delete-all'),
    );
    return menu;
  };

  const createReviewCommentsMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-review-comments');
    menu.append(
      menuButton(t.commentDelete, 'commentAction', 'delete-active'),
      menuButton(t.commentDeleteAll, 'commentAction', 'delete-all'),
    );
    return menu;
  };

  const createProtectMenu = (id: string): HTMLDivElement => {
    const menu = createMenu(menuIdForCommand(id));
    menu.append(
      menuButton(t.protectSheetCommand, 'protectAction', 'protect-sheet'),
      menuButton(t.unprotectSheetCommand, 'protectAction', 'unprotect-sheet'),
      menuSeparator(),
      menuButton(t.lockCell, 'protectAction', 'lock-cell'),
      menuButton(t.unlockCell, 'protectAction', 'unlock-cell'),
      menuSeparator(),
      menuButton(t.protectWorkbookCommand, 'protectAction', 'protect-workbook'),
      menuButton(t.unprotectWorkbookCommand, 'protectAction', 'unprotect-workbook'),
      menuButton(t.allowEditRangesCommand, 'protectAction', 'allow-edit-ranges'),
      menuButton(t.allowEditRangesClearCommand, 'protectAction', 'clear-allowed-edit-ranges'),
    );
    return menu;
  };

  return { createWatchMenu, createReviewCommentsMenu, createProtectMenu };
};
