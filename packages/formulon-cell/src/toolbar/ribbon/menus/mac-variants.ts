import { projectDisabledState } from '../../menu-a11y.js';
import type { ToolbarLang } from '../../ribbon-model.js';
import { findMacBundleLabel, macBundleText } from '../mac/bundle-text.js';
import { menuSectionHeader } from './general.js';

/** Which command surface a menu is built for. */
export interface MenuVariant {
  /** Build the Mac app's row set. */
  mac?: boolean;
}

/** Rows the Mac menu does not carry, by menu id then `dataset:value` row key. */
const MAC_OMITTED_ROWS: Readonly<Record<string, readonly string[]>> = {
  'menu-format-cells': ['cellFormat:tab-color-high-contrast'],
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

/** Menus that open with a greyed heading on the Mac; its wording is generated as `heading:<menu id>`. */
const MAC_HEADING_MENUS: ReadonlySet<string> = new Set(['menu-paste']);

const rowKeys = (row: HTMLElement): string[] =>
  Object.entries(row.dataset).map(([attr, value]) => `${attr}:${value}`);

/** The Mac wording of a menu row, generated as `row:<dataset:value>`, in `lang`. */
const macMenuRowText = (rowKey: string, lang: ToolbarLang): string | undefined =>
  findMacBundleLabel(`row:${rowKey}`)?.[lang];

/**
 * Give a built menu the Mac row set and wording. Both are properties of the platform and apply
 * in every language; the wording is the Mac app's own string for each row that has one.
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
    const label = keys.map((key) => macMenuRowText(key, lang)).find(Boolean);
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
  if (MAC_HEADING_MENUS.has(menu.id)) {
    const row = menuSectionHeader(macBundleText(`heading:${menu.id}`, lang));
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
