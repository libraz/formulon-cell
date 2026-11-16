// Table Design tab menus. The Refresh split button lists the external-data commands; only
// Refresh All has something to run on a table without a connection.

import type { ToolbarLang } from '../../ribbon-model.js';
import { t as macLabel, text as macText } from '../mac/labels.js';
import { createMenu, menuIconButton, menuIdForCommand, menuSeparator } from './general.js';

/** Wording of the external-data menu, from the desktop app in every language. */
const LABELS = {
  refreshStatus: {
    en: 'Refresh Status',
    ja: 'データ更新の状況',
    de: 'Aktualisierungsstatus',
    fr: 'État de l’actualisation',
    zh: '刷新状态',
    ko: '새로 고침 상태',
    th: 'สถานะการรีเฟรช',
  },
  cancelRefresh: {
    en: 'Cancel Refresh',
    ja: '更新のキャンセル',
    de: 'Aktualisierung abbrechen',
    fr: 'Annuler l’actualisation',
    zh: '取消刷新',
    ko: '새로 고침 취소',
    th: 'ยกเลิกการรีเฟรช',
  },
  connectionProperties: {
    en: 'Connection Properties…',
    ja: '接続のプロパティ…',
    de: 'Verbindungseigenschaften…',
    fr: 'Propriétés de connexion…',
    zh: '连接属性…',
    ko: '연결 속성…',
    th: 'คุณสมบัติการเชื่อมต่อ…',
  },
} as const satisfies Record<string, Readonly<Record<ToolbarLang, string>>>;

export interface TableMenuFactories {
  createTableRefreshMenu: (commandId: string) => HTMLDivElement;
}

export const createTableMenuFactories = (
  lang: ToolbarLang,
  refreshLabel: string,
): TableMenuFactories => {
  const label = (key: keyof typeof LABELS): string => LABELS[key][lang];
  const createTableRefreshMenu = (commandId: string): HTMLDivElement => {
    const menu = createMenu(menuIdForCommand(commandId));
    menu.append(
      menuIconButton(refreshLabel, 'tableRefresh', 'refresh', 'table-refresh'),
      menuIconButton(
        macText(lang, macLabel('refreshAll')),
        'tableRefresh',
        'refresh-all',
        'table-refresh',
      ),
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
