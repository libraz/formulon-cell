// Menu spec for the right-click context menu. The DOM-wiring half lives in
// `context-menu.ts`; this module exposes only the data shape (item ids, menu
// kinds, entry tree) plus the small set of label builders that turn a
// localized `Strings` dictionary into the cell/row/col entry lists.

import type { Strings } from '../i18n/strings.js';

export type ItemId =
  | 'copy'
  | 'cut'
  | 'paste'
  | 'pasteSpecial'
  | 'pasteAll'
  | 'pasteFormulas'
  | 'pasteFormulasNumFmt'
  | 'pasteValues'
  | 'pasteValuesNumFmt'
  | 'pasteFormatsOnly'
  | 'pasteTranspose'
  | 'insertCopiedCells'
  | 'clear'
  | 'bold'
  | 'italic'
  | 'underline'
  | 'alignLeft'
  | 'alignCenter'
  | 'alignRight'
  | 'borders'
  | 'formatCells'
  | 'defineName'
  | 'filterClear'
  | 'filterReapply'
  | 'filterByValue'
  | 'sortAsc'
  | 'sortDesc'
  | 'selectAll'
  | 'rowHeight'
  | 'colWidth'
  | 'rowInsertAbove'
  | 'rowInsertBelow'
  | 'rowDelete'
  | 'rowHide'
  | 'rowUnhide'
  | 'colInsertLeft'
  | 'colInsertRight'
  | 'colDelete'
  | 'colHide'
  | 'colUnhide'
  | 'rowGroup'
  | 'rowUngroup'
  | 'colGroup'
  | 'colUngroup'
  | 'insertComment'
  | 'deleteComment'
  | 'insertHyperlink'
  | 'toggleWatch';

export type MenuKind = 'cell' | 'row' | 'col';

export type MenuEntry =
  | { kind: 'item'; id: ItemId; label: string; hint?: string }
  | { kind: 'submenu'; id: string; label: string; children: MenuEntry[] }
  | { kind: 'sep'; id: string };

/** Quick-paste entries inside the Paste Special submenu — disabled when no
 *  structured clipboard snapshot is available. */
export const PASTE_QUICK_IDS: readonly ItemId[] = [
  'pasteAll',
  'pasteFormulas',
  'pasteFormulasNumFmt',
  'pasteValues',
  'pasteValuesNumFmt',
  'pasteFormatsOnly',
  'pasteTranspose',
];

/** Maps a submenu id to the `data-fc-action` value used for its CSS icon. */
export const SUBMENU_ICON_ACTION: Record<string, string> = {
  pasteSpecialMenu: 'pasteSpecial',
  filterMenu: 'filter',
  sortMenu: 'sort',
};

function pasteSpecialSubmenu(s: Strings): MenuEntry {
  const t = s.contextMenu;
  return {
    kind: 'submenu',
    id: 'pasteSpecialMenu',
    label: t.pasteSpecial,
    children: [
      { kind: 'item', id: 'pasteAll', label: t.paste },
      { kind: 'item', id: 'pasteFormulas', label: t.pasteFormulas },
      { kind: 'item', id: 'pasteFormulasNumFmt', label: t.pasteFormulasNumFmt },
      { kind: 'sep', id: 'psSep1' },
      { kind: 'item', id: 'pasteValues', label: t.pasteValues },
      { kind: 'item', id: 'pasteValuesNumFmt', label: t.pasteValuesNumFmt },
      { kind: 'item', id: 'pasteFormatsOnly', label: t.pasteFormatsOnly },
      { kind: 'item', id: 'pasteTranspose', label: t.pasteTranspose },
      { kind: 'sep', id: 'psSep2' },
      { kind: 'item', id: 'pasteSpecial', label: t.pasteSpecialDialog },
    ],
  };
}

function filterSubmenu(s: Strings): MenuEntry {
  const t = s.contextMenu;
  return {
    kind: 'submenu',
    id: 'filterMenu',
    label: t.filter,
    children: [
      { kind: 'item', id: 'filterClear', label: t.filterClear },
      { kind: 'item', id: 'filterReapply', label: t.filterReapply },
      { kind: 'sep', id: 'flSep1' },
      { kind: 'item', id: 'filterByValue', label: t.filterByValue },
    ],
  };
}

function sortSubmenu(s: Strings): MenuEntry {
  const t = s.contextMenu;
  return {
    kind: 'submenu',
    id: 'sortMenu',
    label: t.sort,
    children: [
      { kind: 'item', id: 'sortAsc', label: t.sortAsc },
      { kind: 'item', id: 'sortDesc', label: t.sortDesc },
    ],
  };
}

export function buildCellEntries(s: Strings): MenuEntry[] {
  const t = s.contextMenu;
  return [
    { kind: 'item', id: 'cut', label: t.cut, hint: '⌘X' },
    { kind: 'item', id: 'copy', label: t.copy, hint: '⌘C' },
    { kind: 'item', id: 'paste', label: t.paste, hint: '⌘V' },
    pasteSpecialSubmenu(s),
    { kind: 'sep', id: 'sep1' },
    { kind: 'item', id: 'insertCopiedCells', label: t.insertCopiedCells },
    { kind: 'item', id: 'clear', label: t.clear, hint: 'Del' },
    { kind: 'sep', id: 'sep2' },
    filterSubmenu(s),
    sortSubmenu(s),
    { kind: 'sep', id: 'sep3' },
    { kind: 'item', id: 'insertComment', label: t.insertComment, hint: '⇧F2' },
    { kind: 'item', id: 'deleteComment', label: t.deleteComment },
    { kind: 'sep', id: 'sep4' },
    { kind: 'item', id: 'formatCells', label: t.formatCells, hint: '⌘1' },
    { kind: 'item', id: 'defineName', label: t.defineName },
    { kind: 'sep', id: 'sep5' },
    { kind: 'item', id: 'insertHyperlink', label: t.insertHyperlink, hint: '⌘K' },
    { kind: 'sep', id: 'sep6' },
    { kind: 'item', id: 'toggleWatch', label: t.addWatch },
    { kind: 'item', id: 'selectAll', label: t.selectAll, hint: '⌘A' },
  ];
}

export function buildRowEntries(s: Strings): MenuEntry[] {
  const t = s.contextMenu;
  return [
    { kind: 'item', id: 'cut', label: t.cut, hint: '⌘X' },
    { kind: 'item', id: 'copy', label: t.copy, hint: '⌘C' },
    { kind: 'item', id: 'paste', label: t.paste, hint: '⌘V' },
    pasteSpecialSubmenu(s),
    { kind: 'sep', id: 'sepR1' },
    { kind: 'item', id: 'rowInsertAbove', label: t.insert },
    { kind: 'item', id: 'rowInsertBelow', label: t.rowInsertBelow },
    { kind: 'item', id: 'rowDelete', label: t.delete },
    { kind: 'item', id: 'clear', label: t.clear, hint: 'Del' },
    { kind: 'sep', id: 'sepR2' },
    { kind: 'item', id: 'formatCells', label: t.formatCells, hint: '⌘1' },
    { kind: 'item', id: 'rowHeight', label: t.rowHeight },
    { kind: 'item', id: 'rowHide', label: t.rowHide },
    { kind: 'item', id: 'rowUnhide', label: t.rowUnhide },
    { kind: 'sep', id: 'sepR3' },
    { kind: 'item', id: 'selectAll', label: t.selectAll, hint: '⌘A' },
  ];
}

export function buildColEntries(s: Strings): MenuEntry[] {
  const t = s.contextMenu;
  return [
    { kind: 'item', id: 'cut', label: t.cut, hint: '⌘X' },
    { kind: 'item', id: 'copy', label: t.copy, hint: '⌘C' },
    { kind: 'item', id: 'paste', label: t.paste, hint: '⌘V' },
    pasteSpecialSubmenu(s),
    { kind: 'sep', id: 'sepC1' },
    { kind: 'item', id: 'colInsertLeft', label: t.insert },
    { kind: 'item', id: 'colInsertRight', label: t.colInsertRight },
    { kind: 'item', id: 'colDelete', label: t.delete },
    { kind: 'item', id: 'clear', label: t.clear, hint: 'Del' },
    { kind: 'sep', id: 'sepC2' },
    filterSubmenu(s),
    sortSubmenu(s),
    { kind: 'sep', id: 'sepC3' },
    { kind: 'item', id: 'formatCells', label: t.formatCells, hint: '⌘1' },
    { kind: 'item', id: 'colWidth', label: t.colWidth },
    { kind: 'item', id: 'colHide', label: t.colHide },
    { kind: 'item', id: 'colUnhide', label: t.colUnhide },
    { kind: 'sep', id: 'sepC4' },
    { kind: 'item', id: 'selectAll', label: t.selectAll, hint: '⌘A' },
  ];
}

export function compactMenuEntries(entries: MenuEntry[]): MenuEntry[] {
  const out: MenuEntry[] = [];
  for (const entry of entries) {
    if (entry.kind === 'sep') {
      const prev = out[out.length - 1];
      if (!prev || prev.kind === 'sep') continue;
      out.push(entry);
      continue;
    }
    out.push(entry);
  }
  while (out[out.length - 1]?.kind === 'sep') out.pop();
  return out;
}
