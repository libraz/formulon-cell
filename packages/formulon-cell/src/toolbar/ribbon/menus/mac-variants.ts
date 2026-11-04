import { projectDisabledState } from '../../menu-a11y.js';
import type { ToolbarLang } from '../../ribbon-model.js';
import { menuSectionHeader } from './general.js';

/** Which command surface a menu is built for. */
export interface MenuVariant {
  /** Build the Mac app's row set. */
  mac?: boolean;
}

/** Rows the Mac menu does not carry, by menu id then `dataset:value` row key. */
const MAC_OMITTED_ROWS: Readonly<Record<string, readonly string[]>> = {
  'menu-clear': ['clear:conditional'],
  'menu-sort-home': [
    'sort:filter-by-value',
    'sort:filter-advanced',
    'sort:dedupe',
    'sort:conditional',
    'sort:named',
  ],
};

/** Menus whose section captions are greyed rows on the Mac. */
const MAC_DISABLED_HEADING_MENUS: ReadonlySet<string> = new Set(['menu-format-cells']);

/** Greyed heading the Mac menu opens with, by menu id; Japanese is the measured wording. */
const MAC_MENU_HEADINGS: Readonly<Record<string, { ja: string; en: string }>> = {
  'menu-paste': { ja: '貼り付けのオプション:', en: 'Paste Options:' },
};

/** Japanese wording of the Mac menus, by `dataset:value` row key. */
const MAC_JA_ROW_LABELS: Readonly<Record<string, string>> = {
  'wrapAction:wrapText': '折り返して全体を表示する',
  'cellInsert:cells': 'セルの挿入...',
  'cellDelete:cells': 'セルの削除...',
  'cellFormat:col-autofit': '列の幅に合わせて自動調整',
  'clear:comments': 'コメントとメモをクリア',
  'cellFormat:protect-sheet': 'シートの保護...',
  'sort:asc': '昇順',
  'sort:desc': '降順',
  'pasteAction:all': 'ペースト',
  'pasteAction:dialog': '形式を選択してペースト...',
  'findSelect:go-to-special': '選択オプション...',
  'findSelect:comments': 'メモ',
  'findSelect:selection-pane': '選択ウインドウ...',
};

const rowKeys = (row: HTMLElement): string[] =>
  Object.entries(row.dataset).map(([attr, value]) => `${attr}:${value}`);

/**
 * Give a built menu the Mac row set and wording. The row set is a property of the
 * platform and applies in every language; only the wording is per language, and
 * only Japanese is measured so far.
 */
export const applyMenuVariant = (
  menu: HTMLDivElement,
  lang: ToolbarLang,
  variant: MenuVariant | undefined,
): HTMLDivElement => {
  if (!variant?.mac) return menu;
  const omitted = MAC_OMITTED_ROWS[menu.id] ?? [];
  for (const row of menu.querySelectorAll<HTMLElement>('.fc-tb__menu-item')) {
    const keys = rowKeys(row);
    if (keys.some((key) => omitted.includes(key))) {
      row.remove();
      continue;
    }
    const label =
      lang === 'ja' ? keys.map((key) => MAC_JA_ROW_LABELS[key]).find(Boolean) : undefined;
    if (!label) continue;
    const text = row.querySelector<HTMLElement>('.fc-tb__menu-item__text') ?? row;
    const previous = (text.textContent ?? '').trim();
    text.textContent = label;
    if (row.title === previous) row.title = label;
    if (row.getAttribute('aria-label') === previous) row.setAttribute('aria-label', label);
  }
  if (MAC_DISABLED_HEADING_MENUS.has(menu.id)) {
    for (const caption of menu.querySelectorAll<HTMLElement>(':scope > .fc-tb__menu-heading')) {
      projectDisabledState(caption, true, null);
    }
  }
  const heading = MAC_MENU_HEADINGS[menu.id];
  if (heading) {
    const row = menuSectionHeader(lang === 'ja' ? heading.ja : heading.en);
    projectDisabledState(row, true, null);
    menu.prepend(row);
  }
  // A separator left at an edge, or doubled up, divides nothing.
  let previousWasSeparator = true;
  for (const child of [...menu.children]) {
    const isSeparator = child.getAttribute('role') === 'separator';
    if (isSeparator && previousWasSeparator) child.remove();
    else previousWasSeparator = isSeparator;
  }
  if (menu.lastElementChild?.getAttribute('role') === 'separator') menu.lastElementChild.remove();
  return menu;
};
