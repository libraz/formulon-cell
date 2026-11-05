// Menu spec for the right-click context menu. The DOM-wiring half lives in
// `context-menu.ts`; this module exposes only the data shape (item ids, menu
// kinds, entry tree) plus the small set of label builders that turn a
// localized `Strings` dictionary into the cell/row/col entry lists.

import { PivotAggregation, PivotShowValuesAs } from '../engine/types.js';
import type { Strings } from '../i18n/strings.js';
import type { ContextMenuItem } from './context-menu-options.js';

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
  | 'pasteKeepSource'
  | 'pasteNoBorders'
  | 'pasteColumnWidths'
  | 'pasteLink'
  | 'insertCopiedCells'
  | 'insertCells'
  | 'insertTableRows'
  | 'insertTableColumns'
  | 'deleteCells'
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
  | 'showHideNote'
  | 'insertHyperlink'
  | 'openHyperlink'
  | 'editPhonetic'
  | 'showPhonetic'
  | 'toggleWatch'
  | 'pivotRefresh'
  | 'pivotSortAsc'
  | 'pivotSortDesc'
  | 'pivotFieldSettings'
  | 'pivotShowFieldList'
  | 'pivotNumberFormat'
  | 'pivotDelete'
  | 'pivotSortMore'
  | 'pivotFilterClear'
  | 'pivotFilterKeepSelected'
  | 'pivotFilterHideSelected'
  | 'pivotFilterTop10'
  | 'pivotFilterLabel'
  | 'pivotFilterValue'
  | 'pivotSelectLabels'
  | 'pivotSelectData'
  | 'pivotSelectLabelsAndData'
  | 'pivotSelectTable'
  | 'pivotSelectEnable'
  | 'pivotCalcField'
  | 'pivotCalcItem'
  | 'pivotSolveOrder'
  | 'pivotListFormulas'
  | 'pivotRemoveField'
  | 'pivotRemoveValues'
  | 'pivotRemoveGrandTotal'
  | 'pivotSubtotal'
  | 'pivotExpand'
  | 'pivotCollapse'
  | 'pivotExpandField'
  | 'pivotCollapseField'
  | 'pivotCollapseTo'
  | 'pivotGroup'
  | 'pivotUngroup'
  | 'pivotMoveItemFirst'
  | 'pivotMoveItemUp'
  | 'pivotMoveItemDown'
  | 'pivotMoveItemLast'
  | 'pivotMoveFieldFirst'
  | 'pivotMoveFieldLeft'
  | 'pivotMoveFieldRight'
  | 'pivotMoveFieldLast'
  | 'pivotMoveFieldToColumns'
  | 'pivotMoveFieldToRows'
  | 'pivotSummarizeSum'
  | 'pivotSummarizeCount'
  | 'pivotSummarizeAverage'
  | 'pivotSummarizeMax'
  | 'pivotSummarizeMin'
  | 'pivotSummarizeProduct'
  | 'pivotSummarizeDistinctCount'
  | 'pivotSummarizeMore'
  | 'pivotShowAsNormal'
  | 'pivotShowAsPercentOfTotal'
  | 'pivotShowAsPercentOfColumn'
  | 'pivotShowAsPercentOfRow'
  | 'pivotShowAsPercentOfParentRow'
  | 'pivotShowAsPercentOfParentColumn'
  | 'pivotShowAsMore'
  | 'pivotShowDetails'
  | 'pivotOptions';

/** Menu-level IDs are also stable composition targets. They are not commands
 * themselves; selecting one selects the corresponding nested item tree. */
export type MenuContainerId =
  | 'pasteSpecialMenu'
  | 'filterMenu'
  | 'sortMenu'
  | 'insertTableMenu'
  | 'pivotSortMenu'
  | 'pivotFilterMenu'
  | 'pivotSelectMenu'
  | 'pivotFormulasMenu'
  | 'pivotSummarizeMenu'
  | 'pivotShowAsMenu'
  | 'pivotExpandMenu'
  | 'pivotMoveMenu';

export type MenuKind = 'cell' | 'row' | 'col';

/** Plain insert entries. While a copy is pending the desktop menus replace
 *  these with "Insert Copied Cells" rather than listing both. */
export const PLAIN_INSERT_IDS: readonly ItemId[] = [
  'insertCells',
  'rowInsertAbove',
  'rowInsertBelow',
  'colInsertLeft',
  'colInsertRight',
];

export type MenuEntry =
  | {
      kind: 'item';
      id: ItemId;
      label: string;
      hint?: string;
      /** Drawn disabled, with `disabledReason` as its explanation. */
      disabled?: boolean;
      disabledReason?: string;
      /** A check-mark row: true draws the mark. */
      checked?: boolean;
    }
  | { kind: 'submenu'; id: string; label: string; children: MenuEntry[] }
  | { kind: 'sep'; id: string };

/** Every registered executable built-in ID. Custom IDs must stay outside this
 * set so a host cannot accidentally turn a callback item into a built-in. */
export const BUILTIN_ITEM_IDS: readonly ItemId[] = [
  'copy',
  'cut',
  'paste',
  'pasteSpecial',
  'pasteAll',
  'pasteFormulas',
  'pasteFormulasNumFmt',
  'pasteValues',
  'pasteValuesNumFmt',
  'pasteFormatsOnly',
  'pasteTranspose',
  'pasteKeepSource',
  'pasteNoBorders',
  'pasteColumnWidths',
  'pasteLink',
  'insertCopiedCells',
  'insertCells',
  'insertTableRows',
  'insertTableColumns',
  'deleteCells',
  'clear',
  'bold',
  'italic',
  'underline',
  'alignLeft',
  'alignCenter',
  'alignRight',
  'borders',
  'formatCells',
  'defineName',
  'filterClear',
  'filterReapply',
  'filterByValue',
  'sortAsc',
  'sortDesc',
  'selectAll',
  'rowHeight',
  'colWidth',
  'rowInsertAbove',
  'rowInsertBelow',
  'rowDelete',
  'rowHide',
  'rowUnhide',
  'colInsertLeft',
  'colInsertRight',
  'colDelete',
  'colHide',
  'colUnhide',
  'rowGroup',
  'rowUngroup',
  'colGroup',
  'colUngroup',
  'insertComment',
  'deleteComment',
  'showHideNote',
  'insertHyperlink',
  'openHyperlink',
  'editPhonetic',
  'showPhonetic',
  'toggleWatch',
  'pivotRefresh',
  'pivotSortAsc',
  'pivotSortDesc',
  'pivotFieldSettings',
  'pivotShowFieldList',
  'pivotNumberFormat',
  'pivotDelete',
  'pivotSortMore',
  'pivotFilterClear',
  'pivotFilterKeepSelected',
  'pivotFilterHideSelected',
  'pivotFilterTop10',
  'pivotFilterLabel',
  'pivotFilterValue',
  'pivotSelectLabels',
  'pivotSelectData',
  'pivotSelectLabelsAndData',
  'pivotSelectTable',
  'pivotSelectEnable',
  'pivotCalcField',
  'pivotCalcItem',
  'pivotSolveOrder',
  'pivotListFormulas',
  'pivotRemoveField',
  'pivotRemoveValues',
  'pivotRemoveGrandTotal',
  'pivotSubtotal',
  'pivotExpand',
  'pivotCollapse',
  'pivotExpandField',
  'pivotCollapseField',
  'pivotCollapseTo',
  'pivotGroup',
  'pivotUngroup',
  'pivotMoveItemFirst',
  'pivotMoveItemUp',
  'pivotMoveItemDown',
  'pivotMoveItemLast',
  'pivotMoveFieldFirst',
  'pivotMoveFieldLeft',
  'pivotMoveFieldRight',
  'pivotMoveFieldLast',
  'pivotMoveFieldToColumns',
  'pivotMoveFieldToRows',
  'pivotSummarizeSum',
  'pivotSummarizeCount',
  'pivotSummarizeAverage',
  'pivotSummarizeMax',
  'pivotSummarizeMin',
  'pivotSummarizeProduct',
  'pivotSummarizeDistinctCount',
  'pivotSummarizeMore',
  'pivotShowAsNormal',
  'pivotShowAsPercentOfTotal',
  'pivotShowAsPercentOfColumn',
  'pivotShowAsPercentOfRow',
  'pivotShowAsPercentOfParentRow',
  'pivotShowAsPercentOfParentColumn',
  'pivotShowAsMore',
  'pivotShowDetails',
  'pivotOptions',
];

const BUILTIN_ITEM_ID_SET: ReadonlySet<string> = new Set(BUILTIN_ITEM_IDS);

export const isBuiltinItemId = (id: string): id is ItemId => BUILTIN_ITEM_ID_SET.has(id);

/** Convert the legacy rendering tree to the public stable-ID shape. */
export function menuEntryToContextItem(entry: MenuEntry): ContextMenuItem {
  if (entry.kind === 'sep') return { id: entry.id, separator: true };
  if (entry.kind === 'submenu') {
    return {
      id: entry.id,
      label: entry.label,
      children: entry.children.map(menuEntryToContextItem),
    };
  }
  return {
    id: entry.id,
    label: entry.label,
    ...(entry.hint === undefined ? {} : { hint: entry.hint }),
    ...(entry.disabled ? { disabled: true } : {}),
    ...(entry.disabledReason === undefined ? {} : { disabledReason: entry.disabledReason }),
    ...(entry.checked === undefined ? {} : { checked: entry.checked }),
    builtIn: entry.id,
  };
}

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
  'pasteKeepSource',
  'pasteNoBorders',
  'pasteColumnWidths',
  'pasteLink',
];

/** Maps a submenu id to the `data-fc-action` value used for its CSS icon. */
export const SUBMENU_ICON_ACTION: Record<string, string> = {
  pasteSpecialMenu: 'pasteSpecial',
  filterMenu: 'filter',
  sortMenu: 'sort',
  pivotSortMenu: 'sort',
  pivotFilterMenu: 'filter',
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

/**
 * Inside a table the flat Insert entry becomes a submenu of the table
 * insertions. `rows` is false where Table Rows Above has no row to go above
 * (the header row), which draws it disabled.
 */
export function withTableInsertMenu(
  entries: readonly MenuEntry[],
  s: Strings,
  { rows }: { rows: boolean },
): MenuEntry[] {
  const t = s.contextMenu;
  return entries.map((e) =>
    e.kind === 'item' && e.id === 'insertCells'
      ? {
          kind: 'submenu',
          id: 'insertTableMenu',
          label: e.label,
          children: [
            { kind: 'item', id: 'insertTableColumns', label: t.tableColumnsLeft },
            {
              kind: 'item',
              id: 'insertTableRows',
              label: t.tableRowsAbove,
              ...(rows ? {} : { disabled: true }),
            },
          ],
        }
      : e,
  );
}

export function buildCellEntries(s: Strings): MenuEntry[] {
  const t = s.contextMenu;
  return [
    { kind: 'item', id: 'cut', label: t.cut, hint: '⌘X' },
    { kind: 'item', id: 'copy', label: t.copy, hint: '⌘C' },
    { kind: 'item', id: 'paste', label: t.paste, hint: '⌘V' },
    pasteSpecialSubmenu(s),
    { kind: 'sep', id: 'sep1' },
    { kind: 'item', id: 'insertCells', label: s.ribbonMenu.insertCells },
    { kind: 'item', id: 'insertCopiedCells', label: t.insertCopiedCells },
    { kind: 'item', id: 'deleteCells', label: s.ribbonMenu.deleteCells },
    { kind: 'item', id: 'clear', label: t.clear, hint: 'Del' },
    { kind: 'sep', id: 'sep2' },
    filterSubmenu(s),
    sortSubmenu(s),
    { kind: 'sep', id: 'sep3' },
    { kind: 'item', id: 'insertComment', label: t.insertComment, hint: '⇧F2' },
    { kind: 'item', id: 'deleteComment', label: t.deleteComment },
    { kind: 'item', id: 'showHideNote', label: t.showHideNote },
    { kind: 'sep', id: 'sep4' },
    { kind: 'item', id: 'formatCells', label: t.formatCells, hint: '⌘1' },
    { kind: 'item', id: 'defineName', label: t.defineName },
    { kind: 'sep', id: 'sep5' },
    { kind: 'item', id: 'insertHyperlink', label: t.insertHyperlink, hint: '⌘K' },
    { kind: 'item', id: 'openHyperlink', label: s.ribbonMenu.linkOpen },
    { kind: 'item', id: 'editPhonetic', label: t.editPhonetic },
    { kind: 'sep', id: 'sep6' },
    { kind: 'item', id: 'toggleWatch', label: t.addWatch },
    { kind: 'item', id: 'selectAll', label: t.selectAll, hint: '⌘A' },
  ];
}

/** The kind of report cell a pivot menu is opened on, as the desktop tells them apart. */
export type PivotMenuCell = 'value' | 'item' | 'header' | 'total' | 'dataHeader';

/** What the pivot menu needs to know about the clicked report cell. */
export interface PivotMenuTarget {
  cell: PivotMenuCell;
  /** The row or column field of an item cell. */
  fieldName: string;
  /** The item an item cell shows. */
  itemName: string;
  /** The data field of a value, total or data-header cell. */
  dataFieldName: string;
  /** Whether the item's field sits on the columns rather than the rows. */
  itemOnColumns: boolean;
  /** The data field's current function and calculation, for the check marks. */
  aggregation?: PivotAggregation;
  showValuesAs?: number;
  /** The item's field shows its automatic or custom subtotal. */
  subtotalOn: boolean;
  /** The item's field hides an item or carries a label / value filter. */
  filtered: boolean;
  /** Where the item sits among its field's items, and how many there are. */
  itemPosition: number;
  itemCount: number;
  /** Where the item's field sits on its axis, and how many fields share it. */
  fieldPosition: number;
  fieldCount: number;
  /** The item's field is grouped by date. */
  grouped: boolean;
  /** Whether the field list is hidden, when the host can tell; draws the Hide
   *  Field List check mark. */
  fieldListHidden?: boolean;
  /** The report can be refreshed from its source. */
  canRefresh: boolean;
  /** The field list panel can be opened. */
  canOpenFieldList: boolean;
}

const SUMMARIZE_ITEMS: readonly [ItemId, PivotAggregation | null, keyof Strings['contextMenu']][] =
  [
    ['pivotSummarizeSum', PivotAggregation.Sum, 'pivotSummarizeSum'],
    ['pivotSummarizeCount', PivotAggregation.Count, 'pivotSummarizeCount'],
    ['pivotSummarizeAverage', PivotAggregation.Average, 'pivotSummarizeAverage'],
    ['pivotSummarizeMax', PivotAggregation.Max, 'pivotSummarizeMax'],
    ['pivotSummarizeMin', PivotAggregation.Min, 'pivotSummarizeMin'],
    ['pivotSummarizeProduct', PivotAggregation.Product, 'pivotSummarizeProduct'],
    ['pivotSummarizeDistinctCount', null, 'pivotSummarizeDistinctCount'],
  ];

const SHOW_AS_ITEMS: readonly [ItemId, number, keyof Strings['contextMenu']][] = [
  ['pivotShowAsNormal', PivotShowValuesAs.Normal, 'pivotShowAsNormal'],
  ['pivotShowAsPercentOfTotal', PivotShowValuesAs.PercentOfTotal, 'pivotShowAsPercentOfTotal'],
  ['pivotShowAsPercentOfColumn', PivotShowValuesAs.PercentOfCol, 'pivotShowAsPercentOfColumn'],
  ['pivotShowAsPercentOfRow', PivotShowValuesAs.PercentOfRow, 'pivotShowAsPercentOfRow'],
  [
    'pivotShowAsPercentOfParentRow',
    PivotShowValuesAs.PercentOfParentRow,
    'pivotShowAsPercentOfParentRow',
  ],
  [
    'pivotShowAsPercentOfParentColumn',
    PivotShowValuesAs.PercentOfParentCol,
    'pivotShowAsPercentOfParentColumn',
  ],
];

/** The data field's function and calculation, by the engine kind they write. */
export const PIVOT_MENU_AGGREGATION: Partial<Record<ItemId, PivotAggregation>> = Object.fromEntries(
  SUMMARIZE_ITEMS.filter(([, value]) => value !== null).map(([id, value]) => [id, value]),
);
export const PIVOT_MENU_SHOW_AS: Partial<Record<ItemId, number>> = Object.fromEntries(
  SHOW_AS_ITEMS.map(([id, value]) => [id, value]),
);

/** Cell menu for a PivotTable report cell: the entries, separators and
 *  submenus of the desktop menu for that kind of cell, in its order. An entry
 *  the host cannot carry out is listed disabled with the reason. */
export function buildPivotEntries(s: Strings, target: PivotMenuTarget): MenuEntry[] {
  const t = s.contextMenu;
  const { cell } = target;
  const reason = t.pivotUnavailable;
  const off = (id: ItemId, label: string, checked?: boolean): MenuEntry => ({
    kind: 'item',
    id,
    label,
    disabled: true,
    disabledReason: reason,
    ...(checked === undefined ? {} : { checked }),
  });
  const on = (id: ItemId, label: string, checked?: boolean): MenuEntry => ({
    kind: 'item',
    id,
    label,
    ...(checked === undefined ? {} : { checked }),
  });
  // A move that has nowhere to go is drawn disabled, as on the desktop.
  const onIf = (id: ItemId, label: string, enabled: boolean): MenuEntry =>
    enabled ? on(id, label) : { kind: 'item', id, label, disabled: true };
  const field = (template: string): string => template.replace('{field}', target.fieldName);
  const item = (template: string): string => template.replace('{item}', target.itemName);
  const dataField = (template: string): string => template.replace('{field}', target.dataFieldName);
  let sep = 0;
  const separator = (): MenuEntry => {
    sep += 1;
    return { kind: 'sep', id: `pvSep${sep}` };
  };
  const valueLike = cell === 'value' || cell === 'total' || cell === 'dataHeader';
  const sortMenu: MenuEntry = {
    kind: 'submenu',
    id: 'pivotSortMenu',
    label: t.pivotSort,
    children: [
      on('pivotSortAsc', t.pivotSortAsc),
      on('pivotSortDesc', t.pivotSortDesc),
      separator(),
      off('pivotSortMore', t.pivotSortMore),
    ],
  };
  const filterMenu: MenuEntry = {
    kind: 'submenu',
    id: 'pivotFilterMenu',
    label: t.pivotFilter,
    children: [
      target.filtered
        ? on('pivotFilterClear', t.pivotFilterClear)
        : off('pivotFilterClear', t.pivotFilterClear),
      separator(),
      on('pivotFilterKeepSelected', t.pivotFilterKeepSelected),
      on('pivotFilterHideSelected', t.pivotFilterHideSelected),
      separator(),
      on('pivotFilterTop10', t.pivotFilterTop10),
      on('pivotFilterLabel', t.pivotFilterLabel),
      on('pivotFilterValue', t.pivotFilterValue),
    ],
  };
  const selectMenu: MenuEntry = {
    kind: 'submenu',
    id: 'pivotSelectMenu',
    label: t.pivotSelect,
    children: [
      off('pivotSelectLabels', t.pivotSelectLabels),
      off('pivotSelectData', t.pivotSelectData),
      off('pivotSelectLabelsAndData', t.pivotSelectLabelsAndData),
      on('pivotSelectTable', t.pivotSelectTable),
      separator(),
      off('pivotSelectEnable', t.pivotSelectEnable, true),
    ],
  };
  const formulasMenu: MenuEntry = {
    kind: 'submenu',
    id: 'pivotFormulasMenu',
    label: t.pivotFormulas,
    children: [
      off('pivotCalcField', t.pivotCalcField),
      off('pivotCalcItem', t.pivotCalcItem),
      separator(),
      off('pivotSolveOrder', t.pivotSolveOrder),
      off('pivotListFormulas', t.pivotListFormulas),
    ],
  };
  const summarizeMenu: MenuEntry = {
    kind: 'submenu',
    id: 'pivotSummarizeMenu',
    label: t.pivotSummarizeBy,
    children: [
      ...SUMMARIZE_ITEMS.map(([id, value, key]) =>
        value === null ? off(id, t[key]) : on(id, t[key], value === target.aggregation),
      ),
      separator(),
      on('pivotSummarizeMore', t.pivotMoreOptions),
    ],
  };
  // On an item the desktop draws the whole list, but none of it applies there.
  const showAsMenu: MenuEntry = {
    kind: 'submenu',
    id: 'pivotShowAsMenu',
    label: t.pivotShowValuesAs,
    children: [
      ...SHOW_AS_ITEMS.map(([id, value, key]) =>
        valueLike
          ? on(id, t[key], value === (target.showValuesAs ?? PivotShowValuesAs.Normal))
          : off(id, t[key], value === PivotShowValuesAs.Normal),
      ),
      separator(),
      valueLike
        ? on('pivotShowAsMore', t.pivotMoreOptions)
        : off('pivotShowAsMore', t.pivotMoreOptions),
    ],
  };
  const expandMenu: MenuEntry = {
    kind: 'submenu',
    id: 'pivotExpandMenu',
    label: t.pivotExpandCollapse,
    children: [
      off('pivotExpand', t.pivotExpand),
      off('pivotCollapse', t.pivotCollapse),
      off('pivotExpandField', t.pivotExpandField),
      off('pivotCollapseField', t.pivotCollapseField),
      separator(),
      off('pivotCollapseTo', field(t.pivotCollapseTo)),
    ],
  };
  const moveMenu: MenuEntry = {
    kind: 'submenu',
    id: 'pivotMoveMenu',
    label: t.pivotMove,
    children: [
      onIf('pivotMoveItemFirst', item(t.pivotMoveItemFirst), target.itemPosition > 0),
      onIf('pivotMoveItemUp', item(t.pivotMoveItemUp), target.itemPosition > 0),
      onIf(
        'pivotMoveItemDown',
        item(t.pivotMoveItemDown),
        target.itemPosition < target.itemCount - 1,
      ),
      onIf(
        'pivotMoveItemLast',
        item(t.pivotMoveItemLast),
        target.itemPosition < target.itemCount - 1,
      ),
      onIf('pivotMoveFieldFirst', field(t.pivotMoveFieldFirst), target.fieldPosition > 0),
      onIf('pivotMoveFieldLeft', field(t.pivotMoveFieldLeft), target.fieldPosition > 0),
      onIf(
        'pivotMoveFieldRight',
        field(t.pivotMoveFieldRight),
        target.fieldPosition < target.fieldCount - 1,
      ),
      onIf(
        'pivotMoveFieldLast',
        field(t.pivotMoveFieldLast),
        target.fieldPosition < target.fieldCount - 1,
      ),
      separator(),
      target.itemOnColumns
        ? on('pivotMoveFieldToRows', field(t.pivotMoveFieldToRows))
        : on('pivotMoveFieldToColumns', field(t.pivotMoveFieldToColumns)),
    ],
  };

  const entries: MenuEntry[] = [
    { kind: 'item', id: 'copy', label: t.copy, hint: '⌘C' },
    { kind: 'item', id: 'formatCells', label: t.formatCells, hint: '⌘1' },
  ];
  if (valueLike) entries.push(on('pivotNumberFormat', t.pivotNumberFormat));
  entries.push(separator());
  if (target.canRefresh) entries.push(on('pivotRefresh', t.pivotRefresh));
  entries.push(separator(), on('pivotDelete', t.pivotDelete), separator());
  if (cell === 'value' || cell === 'dataHeader') entries.push(sortMenu, separator());
  if (cell === 'item') entries.push(sortMenu, filterMenu, separator());
  entries.push(selectMenu, formulasMenu, separator());
  if (cell === 'item') {
    entries.push(
      off('pivotSubtotal', field(t.pivotSubtotalOf), target.subtotalOn),
      separator(),
      expandMenu,
      separator(),
      off('pivotGroup', t.pivotGroup),
      onIf('pivotUngroup', t.pivotUngroup, target.grouped),
      separator(),
      moveMenu,
      on('pivotRemoveField', field(t.pivotRemoveField)),
      separator(),
      showAsMenu,
    );
  } else if (cell === 'header') {
    entries.push(showAsMenu);
  } else {
    if (cell === 'total') entries.push(on('pivotRemoveGrandTotal', t.pivotRemoveGrandTotal));
    else entries.push(on('pivotRemoveField', dataField(t.pivotRemoveField)));
    if (cell === 'dataHeader') entries.push(on('pivotRemoveValues', t.pivotRemoveValues));
    entries.push(separator(), summarizeMenu, showAsMenu);
    if (cell === 'value') entries.push(separator(), off('pivotShowDetails', t.pivotShowDetails));
  }
  entries.push(separator());
  if (cell !== 'header') {
    entries.push(
      on('pivotFieldSettings', valueLike ? t.pivotValueFieldSettings : t.pivotFieldSettings),
    );
  }
  entries.push(on('pivotOptions', t.pivotOptions));
  if (target.canOpenFieldList) {
    entries.push(on('pivotShowFieldList', t.pivotHideFieldList, target.fieldListHidden));
  }
  return entries;
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
    { kind: 'item', id: 'insertCopiedCells', label: t.insertCopiedBand },
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
    { kind: 'item', id: 'insertCopiedCells', label: t.insertCopiedBand },
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

function macPasteSpecialSubmenu(s: Strings): MenuEntry {
  const t = s.contextMenu;
  const m = s.macContextMenu;
  return {
    kind: 'submenu',
    id: 'pasteSpecialMenu',
    label: m.pasteSpecial,
    children: [
      { kind: 'item', id: 'pasteAll', label: m.paste, hint: '⌘V' },
      { kind: 'sep', id: 'psSep1' },
      { kind: 'item', id: 'pasteFormulas', label: t.pasteFormulas },
      { kind: 'item', id: 'pasteFormulasNumFmt', label: t.pasteFormulasNumFmt },
      { kind: 'item', id: 'pasteKeepSource', label: m.pasteKeepSource },
      { kind: 'item', id: 'pasteNoBorders', label: m.pasteNoBorders },
      { kind: 'item', id: 'pasteColumnWidths', label: m.pasteColumnWidths },
      { kind: 'item', id: 'pasteTranspose', label: t.pasteTranspose },
      { kind: 'sep', id: 'psSep2' },
      { kind: 'item', id: 'pasteValues', label: t.pasteValues },
      { kind: 'item', id: 'pasteValuesNumFmt', label: t.pasteValuesNumFmt },
      { kind: 'sep', id: 'psSep3' },
      { kind: 'item', id: 'pasteFormatsOnly', label: m.pasteFormatsOnly },
      { kind: 'item', id: 'pasteLink', label: m.pasteLink },
      { kind: 'sep', id: 'psSep4' },
      { kind: 'item', id: 'pasteSpecial', label: m.pasteSpecialDialog, hint: '^⌘V' },
    ],
  };
}

/** Cell menu in the Mac desktop order. A note is edited or created through
 *  the same command, so its label (and the delete row) follow `hasNote`. */
export function buildMacCellEntries(s: Strings, hasNote: boolean): MenuEntry[] {
  const t = s.contextMenu;
  const m = s.macContextMenu;
  return [
    { kind: 'item', id: 'cut', label: m.cut, hint: '⌘X' },
    { kind: 'item', id: 'copy', label: m.copy, hint: '⌘C' },
    { kind: 'item', id: 'paste', label: m.paste, hint: '⌘V' },
    macPasteSpecialSubmenu(s),
    { kind: 'sep', id: 'sep1' },
    { kind: 'item', id: 'insertCells', label: m.insert },
    { kind: 'item', id: 'insertCopiedCells', label: m.insertCopiedCells },
    { kind: 'item', id: 'deleteCells', label: m.delete },
    { kind: 'item', id: 'clear', label: m.clear },
    { kind: 'sep', id: 'sep2' },
    filterSubmenu(s),
    sortSubmenu(s),
    { kind: 'sep', id: 'sep3' },
    { kind: 'item', id: 'insertComment', label: hasNote ? m.editNote : m.newNote },
    ...(hasNote
      ? ([
          { kind: 'item', id: 'deleteComment', label: m.deleteNote },
          { kind: 'item', id: 'showHideNote', label: t.showHideNote },
        ] as const)
      : []),
    { kind: 'sep', id: 'sep4' },
    { kind: 'item', id: 'formatCells', label: m.formatCells, hint: '⌘1' },
    { kind: 'item', id: 'showPhonetic', label: m.showPhonetic },
    { kind: 'item', id: 'defineName', label: m.defineName },
    { kind: 'sep', id: 'sep5' },
    { kind: 'item', id: 'insertHyperlink', label: m.hyperlink, hint: '⌘K' },
    { kind: 'item', id: 'openHyperlink', label: s.ribbonMenu.linkOpen },
  ];
}

/** Select-all corner menu: the Mac cell list with the band wording for Insert and Delete. */
export function buildMacCornerEntries(s: Strings, hasNote: boolean): MenuEntry[] {
  const m = s.macContextMenu;
  return buildMacCellEntries(s, hasNote).map((entry) => {
    if (entry.kind !== 'item') return entry;
    if (entry.id === 'insertCells') return { ...entry, label: m.insertBand };
    if (entry.id === 'deleteCells') return { ...entry, label: m.deleteBand };
    return entry;
  });
}

/** Row header menu in the Mac desktop order. */
export function buildMacRowEntries(s: Strings): MenuEntry[] {
  const m = s.macContextMenu;
  return [
    ...macBandHead(s),
    { kind: 'item', id: 'rowInsertAbove', label: m.insertBand },
    { kind: 'item', id: 'insertCopiedCells', label: m.insertCopiedBand },
    { kind: 'item', id: 'rowDelete', label: m.deleteBand },
    { kind: 'item', id: 'clear', label: m.clear },
    { kind: 'sep', id: 'sepR2' },
    { kind: 'item', id: 'formatCells', label: m.formatCells, hint: '⌘1' },
    { kind: 'item', id: 'rowHeight', label: m.rowHeight },
    { kind: 'item', id: 'rowHide', label: m.hide, hint: '^9' },
    { kind: 'item', id: 'rowUnhide', label: m.unhide, hint: '^⇧9' },
  ];
}

/** Column header menu in the Mac desktop order. */
export function buildMacColEntries(s: Strings): MenuEntry[] {
  const m = s.macContextMenu;
  return [
    ...macBandHead(s),
    { kind: 'item', id: 'colInsertLeft', label: m.insertBand },
    { kind: 'item', id: 'insertCopiedCells', label: m.insertCopiedBand },
    { kind: 'item', id: 'colDelete', label: m.deleteBand },
    { kind: 'item', id: 'clear', label: m.clear },
    { kind: 'sep', id: 'sepC2' },
    { kind: 'item', id: 'formatCells', label: m.formatCells, hint: '⌘1' },
    { kind: 'item', id: 'colWidth', label: m.colWidth },
    { kind: 'item', id: 'colHide', label: m.hide, hint: '^0' },
    { kind: 'item', id: 'colUnhide', label: m.unhide, hint: '^⇧0' },
  ];
}

function macBandHead(s: Strings): MenuEntry[] {
  const m = s.macContextMenu;
  return [
    { kind: 'item', id: 'cut', label: m.cut, hint: '⌘X' },
    { kind: 'item', id: 'copy', label: m.copy, hint: '⌘C' },
    { kind: 'item', id: 'paste', label: m.paste, hint: '⌘V' },
    macPasteSpecialSubmenu(s),
    { kind: 'sep', id: 'sepB1' },
  ];
}

/** Row/column header menus on Mac: the default entries with the Mac clipboard
 *  wording, no Delete-key hint and no Select All row. */
export function toMacBandEntries(entries: readonly MenuEntry[], s: Strings): MenuEntry[] {
  const m = s.macContextMenu;
  const labels: Partial<Record<ItemId, string>> = {
    cut: m.cut,
    copy: m.copy,
    paste: m.paste,
    clear: m.clear,
    formatCells: m.formatCells,
  };
  return entries
    .filter((e) => !(e.kind === 'item' && e.id === 'selectAll'))
    .map((e) => {
      if (e.kind === 'submenu' && e.id === 'pasteSpecialMenu') return macPasteSpecialSubmenu(s);
      if (e.kind !== 'item') return e;
      const label = labels[e.id];
      if (label === undefined) return e;
      return e.id === 'clear' ? { kind: 'item', id: e.id, label } : { ...e, label };
    });
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
