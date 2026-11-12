import {
  listCustomTableStyles,
  tableStyleGallerySections,
} from '../../commands/format-as-table.js';
import { ribbonStyleTarget } from '../../commands/table-style-actions.js';
import { builtinTableStyleElements } from '../../engine/builtin-table-styles.js';
import { dictionaries } from '../../i18n/strings.js';
import { tableStyleSwatchGrid } from './menus/styles.js';
import {
  blockedUnlessPermitted,
  createQuickStyleStrip,
  type QuickStyleEntry,
  type RibbonStyleGalleryOptions,
} from './quick-style-strip.js';

/** Table Design > Table Styles: the custom and built-in styles a row at a time, with the Quick
 *  Styles button as the more arrow. A pick restyles the active table. */
export const createTableStyleGallery = (opts: RibbonStyleGalleryOptions): HTMLDivElement => {
  const { lang, getInst } = opts;
  const target = () => {
    const inst = getInst();
    return inst ? ribbonStyleTarget(inst, 'table', 'tableQuickStyles') : null;
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
    blocked: blockedUnlessPermitted(getInst, 'tableQuickStyles'),
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
