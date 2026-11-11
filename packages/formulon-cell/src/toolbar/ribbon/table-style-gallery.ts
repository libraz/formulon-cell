import {
  currentTableStyleId,
  listCustomTableStyles,
  overlayStyleName,
  tableOverlayAt,
  tableStyleGallerySections,
} from '../../commands/format-as-table.js';
import { applyTableStyle } from '../../commands/table-style-actions.js';
import { builtinTableStyleElements } from '../../engine/builtin-table-styles.js';
import { dictionaries } from '../../i18n/strings.js';
import type { SpreadsheetInstance } from '../../mount/types.js';
import type { ToolbarLang } from '../ribbon-model.js';
import { tableStyleSwatchGrid } from './menus/styles.js';
import { createQuickStyleStrip, type QuickStyleEntry } from './quick-style-strip.js';

// One gallery row of the seven palette columns; the desktop strip's width is unmeasured.
const PAGE_SIZE = 7;

export interface TableStyleGalleryOptions {
  lang: ToolbarLang;
  getInst: () => SpreadsheetInstance | null;
  /** The ribbon's Quick Styles button; it opens the full gallery. */
  arrow: HTMLButtonElement;
  menu: HTMLElement | null;
}

/** Table Design > Table Styles: the custom and built-in styles a row at a time, with the Quick
 *  Styles button as the more arrow. A pick restyles the active table. */
export const createTableStyleGallery = (opts: TableStyleGalleryOptions): HTMLDivElement => {
  const { lang, getInst } = opts;
  const activeTable = (inst: SpreadsheetInstance) => {
    const state = inst.store.getState();
    const { sheet, row, col } = state.selection.active;
    return tableOverlayAt(state, sheet, row, col);
  };
  const wrap = createQuickStyleStrip({
    lang,
    classPrefix: 'fc-tb__tablestyle-quick',
    chipKey: 'tableQuickStyle',
    pageSize: PAGE_SIZE,
    entries: () => {
      const inst = getInst();
      if (!inst) return [];
      const names = dictionaries[lang].ribbonMenu.tableStyleSwatches;
      const theme = inst.workbook.getTheme()?.colors;
      return tableStyleGallerySections(listCustomTableStyles(inst.store.getState())).flatMap(
        (section): QuickStyleEntry[] =>
          section.kind === 'custom'
            ? section.styles.map((style) => ({
                id: style.id,
                label: style.label,
                thumbnail: () => tableStyleSwatchGrid(style.elements),
              }))
            : section.styles.map((entry) => ({
                id: entry.name,
                label: names[entry.name as keyof typeof names] ?? entry.name,
                thumbnail: () =>
                  tableStyleSwatchGrid(builtinTableStyleElements(entry.name, theme) ?? {}),
              })),
      );
    },
    current: () => {
      const inst = getInst();
      const table = inst ? activeTable(inst) : null;
      if (!inst || !table) return null;
      return currentTableStyleId(inst.store.getState(), overlayStyleName(table));
    },
    apply: (id) => {
      const inst = getInst();
      const table = inst ? activeTable(inst) : null;
      if (!inst || !table) return;
      applyTableStyle(
        {
          store: inst.store,
          history: inst.history,
          origin: 'ribbon',
          commandId: 'tableQuickStyles',
        },
        table,
        id,
      );
      inst.host.focus();
    },
    more: { arrow: opts.arrow, menu: opts.menu },
  });
  wrap.dataset.ribbonGallery = 'tableQuickStyles';
  return wrap;
};
