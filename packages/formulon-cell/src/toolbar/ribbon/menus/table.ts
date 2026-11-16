// Table Design tab menus. The Refresh split button lists the external-data commands; only
// Refresh All has something to run on a table without a connection.

import type { ToolbarLang } from '../../ribbon-model.js';
import { t as macLabel, text as macText } from '../mac/labels.js';
import { createMenu, menuIconButton, menuIdForCommand, menuSeparator } from './general.js';

export interface TableMenuFactories {
  createTableRefreshMenu: (commandId: string) => HTMLDivElement;
}

export const createTableMenuFactories = (
  lang: ToolbarLang,
  refreshLabel: string,
): TableMenuFactories => {
  const label = (key: string): string => macText(lang, macLabel(key));
  const createTableRefreshMenu = (commandId: string): HTMLDivElement => {
    const menu = createMenu(menuIdForCommand(commandId));
    menu.append(
      menuIconButton(refreshLabel, 'tableRefresh', 'refresh', 'table-refresh'),
      menuIconButton(label('refreshAll'), 'tableRefresh', 'refresh-all', 'table-refresh'),
      menuSeparator(),
      menuIconButton(label('refreshStatus'), 'tableRefresh', 'refresh-status', 'table-refresh'),
      menuIconButton(label('cancelRefresh'), 'tableRefresh', 'cancel-refresh', 'table-refresh'),
      menuSeparator(),
      menuIconButton(
        label('connectionProperties'),
        'tableRefresh',
        'connection-properties',
        'table-refresh',
      ),
    );
    return menu;
  };
  return { createTableRefreshMenu };
};
