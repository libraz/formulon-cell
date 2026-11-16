import {
  listCustomPivotTableStyles,
  pivotStyleGallerySections,
} from '../../commands/format-as-table.js';
import { ribbonStyleTarget } from '../../commands/table-style-actions.js';
import { BUILTIN_PIVOT_STYLE, builtinPivotStyleElements } from '../../engine/pivot-style.js';
import type { StyleElementFormat, StyleElementFormats } from '../../engine/table-style-elements.js';
import { focusSheet } from '../../interact/dialog-focus.js';
import type { ToolbarLang } from '../ribbon-model.js';
import { findMacBundleLabel, macBundleText } from './mac/bundle-text.js';
import {
  blockedUnlessPermitted,
  createQuickStyleStrip,
  type QuickStyleEntry,
  type RibbonStyleGalleryOptions,
} from './quick-style-strip.js';

type PivotStyleFamily = 'Light' | 'Medium' | 'Dark';

/** The colour each built-in style's accessible name opens with, by family and style number. */
const COLOURS: Readonly<Record<PivotStyleFamily, readonly string[]>> = {
  Light: [
    ...Array<string>(8).fill('white'),
    ...['iceBlue', 'rose', 'lightGreen', 'lavender', 'skyBlue', 'lightOrange', 'lightGray'],
    ...['iceBlue', 'rose', 'lightGreen', 'lavender', 'skyBlue', 'lightOrange'],
    ...Array<string>(7).fill('white'),
  ],
  Medium: [
    ...['white', 'lightBlue', 'rose', 'lightGreen', 'lavender', 'skyBlue', 'lightOrange'],
    ...['white', 'iceBlue', 'rose', 'lightGreen', 'lavender', 'skyBlue', 'lightOrange'],
    ...['white', 'iceBlue', 'rose', 'lightGreen', 'lavender', 'skyBlue', 'lightOrange'],
    ...['lightGray', 'iceBlue', 'rose', 'lightGreen', 'lavender', 'skyBlue', 'lightOrange'],
  ],
  Dark: [
    ...['lightGray', 'darkBlue', 'darkRed', 'darkGreen', 'darkPurple', 'darkTeal', 'brown'],
    ...Array<string>(7).fill('darkGray'),
    ...['gray', 'blue', 'red', 'oliveGreen', 'darkPurple', 'aqua', 'orange'],
    ...['gray', 'blue', 'red', 'oliveGreen', 'darkPurple', 'aqua', 'orange'],
  ],
};

/** The accessible name of a built-in PivotTable style, "<colour>, <name>", or None for ''. */
export const pivotStyleLabel = (name: string, lang: ToolbarLang): string => {
  if (name === '') return macBundleText('pivot-style:none', lang);
  const match = BUILTIN_PIVOT_STYLE.exec(name);
  const style = findMacBundleLabel(`pivot-style:${name}`);
  const colour = match && COLOURS[match[1] as PivotStyleFamily][Number(match[2]) - 1];
  if (!style || !colour) return name;
  return `${macBundleText(`colour:${colour}`, lang)}, ${style[lang]}`;
};

const edgeColor = (edge: unknown): string =>
  edge && typeof edge === 'object' && 'color' in edge && typeof edge.color === 'string'
    ? edge.color
    : '#999';

const paint = (cell: HTMLElement, format: StyleElementFormat | undefined): void => {
  if (!format) return;
  if (format.fill) cell.style.background = format.fill;
  if (format.color) cell.style.color = format.color;
  if (format.bold) cell.style.fontWeight = '700';
  if (format.borders?.top) cell.style.borderTop = `1px solid ${edgeColor(format.borders.top)}`;
  if (format.borders?.bottom) {
    cell.style.borderBottom = `1px solid ${edgeColor(format.borders.bottom)}`;
  }
};

/** A small report drawn in `elements`: header band, plain rows, total band. */
export const pivotStyleThumbnail = (elements: StyleElementFormats | null): HTMLDivElement => {
  const thumb = document.createElement('div');
  thumb.className = 'fc-tb__pivotstyle-thumb';
  const rows = 5;
  for (let r = 0; r < rows; r += 1) {
    const kind = r === 0 ? 'header' : r === rows - 1 ? 'total' : 'body';
    for (let c = 0; c < 3; c += 1) {
      const cell = document.createElement('span');
      cell.className = `fc-tb__pivotstyle-cell fc-tb__pivotstyle-cell--${kind}`;
      paint(
        cell,
        kind === 'header' ? elements?.headerRow : kind === 'total' ? elements?.totalRow : undefined,
      );
      thumb.appendChild(cell);
    }
  }
  return thumb;
};

/** The PivotTable Design Quick Styles: custom styles, None and the built-in styles, a page at a
 *  time, with the Quick Styles button as the more arrow. */
export const createPivotStyleGallery = (opts: RibbonStyleGalleryOptions): HTMLDivElement => {
  const { lang, getInst } = opts;
  const target = () => {
    const inst = getInst();
    return inst ? ribbonStyleTarget(inst, 'pivot', 'pivotQuickStyles') : null;
  };
  const wrap = createQuickStyleStrip({
    lang,
    classPrefix: 'fc-tb__pivotstyle',
    kind: 'pivot',
    mac: opts.mac,
    chipKey: 'pivotStyle',
    entries: () => {
      const inst = getInst();
      if (!inst) return [];
      const theme = inst.workbook.getTheme()?.colors;
      return pivotStyleGallerySections(listCustomPivotTableStyles(inst.store.getState()), {
        none: true,
      }).flatMap((section): QuickStyleEntry[] =>
        section.kind === 'custom'
          ? section.styles.map((style) => ({
              id: style.id,
              label: style.label,
              thumbnail: () => pivotStyleThumbnail(style.elements),
            }))
          : section.names.map((name) => ({
              id: name,
              label: pivotStyleLabel(name, lang),
              thumbnail: () => pivotStyleThumbnail(builtinPivotStyleElements(name, theme)),
            })),
      );
    },
    // No report, or one without a style element, shows None as current.
    current: () => target()?.currentId ?? '',
    blocked: blockedUnlessPermitted(getInst, 'pivotQuickStyles'),
    apply: (id) => {
      const picked = target();
      if (!picked) return;
      picked.apply(id);
      const sheetHost = getInst()?.host;
      if (sheetHost) focusSheet(sheetHost);
    },
    more: { arrow: opts.arrow, menu: opts.menu },
  });
  wrap.dataset.ribbonGallery = 'pivotQuickStyles';
  return wrap;
};
