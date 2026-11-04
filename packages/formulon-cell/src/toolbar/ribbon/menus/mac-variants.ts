import type { ToolbarLang } from '../../ribbon-model.js';

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

/** Japanese wording of the Mac menus, by `dataset:value` row key. */
const MAC_JA_ROW_LABELS: Readonly<Record<string, string>> = {
  'wrapAction:wrapText': '折り返して全体を表示する',
  'cellInsert:cells': 'セルの挿入...',
  'cellDelete:cells': 'セルの削除...',
  'cellFormat:col-autofit': '列の幅に合わせて自動調整',
  'clear:comments': 'コメントとメモをクリア',
  'sort:asc': '昇順',
  'sort:desc': '降順',
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
