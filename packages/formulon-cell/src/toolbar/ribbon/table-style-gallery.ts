import { canExecuteBuiltInRequest } from '../../commands/built-in-command-policy.js';
import {
  listCustomTableStyles,
  tableStyleGallerySections,
} from '../../commands/format-as-table.js';
import { localizedDecisionReason } from '../../commands/permission-reasons.js';
import { activeStyleTarget } from '../../commands/table-style-actions.js';
import { builtinTableStyleElements } from '../../engine/builtin-table-styles.js';
import { dictionaries } from '../../i18n/strings.js';
import type { SpreadsheetInstance } from '../../mount/types.js';
import type { ToolbarLang } from '../ribbon-model.js';
import { tableStyleSwatchGrid } from './menus/styles.js';
import { createQuickStyleStrip, type QuickStyleEntry } from './quick-style-strip.js';

export interface TableStyleGalleryOptions {
  lang: ToolbarLang;
  getInst: () => SpreadsheetInstance | null;
  /** The ribbon's Quick Styles button; it opens the full gallery. */
  arrow: HTMLButtonElement;
  menu: HTMLElement | null;
  /** The Mac profile's strip: no page arrows, scrolled to the current style. */
  mac: boolean;
}

/** Table Design > Table Styles: the custom and built-in styles a row at a time, with the Quick
 *  Styles button as the more arrow. A pick restyles the active table. */
export const createTableStyleGallery = (opts: TableStyleGalleryOptions): HTMLDivElement => {
  const { lang, getInst } = opts;
  const target = () => {
    const inst = getInst();
    return inst
      ? activeStyleTarget(
          {
            store: inst.store,
            history: inst.history,
            workbook: inst.workbook,
            origin: 'ribbon',
            commandId: 'tableQuickStyles',
          },
          'table',
        )
      : null;
  };
  const wrap = createQuickStyleStrip({
    lang,
    classPrefix: 'fc-tb__tablestyle-quick',
    kind: 'table',
    mac: opts.mac,
    chipKey: 'tableQuickStyle',
    entries: () => {
      const inst = getInst();
      if (!inst) return [];
      const { tableStyleSwatches: names, tableStyleNone } = dictionaries[lang].ribbonMenu;
      const theme = inst.workbook.getTheme()?.colors;
      return tableStyleGallerySections(listCustomTableStyles(inst.store.getState()), {
        none: true,
      }).flatMap((section): QuickStyleEntry[] =>
        section.kind === 'custom'
          ? section.styles.map((style) => ({
              id: style.id,
              label: style.label,
              thumbnail: () => tableStyleSwatchGrid(style.elements),
            }))
          : section.styles.map((entry) => ({
              id: entry.name,
              label:
                entry.name === ''
                  ? tableStyleNone
                  : (names[entry.name as keyof typeof names] ?? entry.name),
              thumbnail: () =>
                tableStyleSwatchGrid(builtinTableStyleElements(entry.name, theme) ?? {}),
            })),
      );
    },
    current: () => target()?.currentId ?? null,
    // A refused restyle (a protected sheet) greys the whole strip, with the reason.
    blocked: () => {
      const inst = getInst();
      if (!inst) return { reason: null };
      const decision = canExecuteBuiltInRequest(inst.store, 'tableQuickStyles', 'ribbon');
      return decision.allowed
        ? null
        : { reason: localizedDecisionReason(decision, inst.i18n.strings) };
    },
    apply: (id) => {
      const picked = target();
      if (!picked) return;
      picked.apply(id);
      getInst()?.host.focus();
    },
    more: { arrow: opts.arrow, menu: opts.menu },
  });
  wrap.dataset.ribbonGallery = 'tableQuickStyles';
  return wrap;
};
