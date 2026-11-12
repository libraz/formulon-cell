// Home tab style-gallery dropdowns: Table Styles, Cell Styles, Currency. Each
// builds a swatch / chip / preset list that emits a `data-table-*`,
// `data-cell-style*`, or `data-currency*` attribute the playground dispatcher
// reads. Kept in one file because they share the same factory pattern (deps
// bundle of ribbonLang + dictionaries + ribbonText) and they all sit in the
// Home tab's Styles group.

import {
  CELL_STYLE_GROUPS,
  CELL_STYLES,
  type CellStyleFallbackProfile,
  type CellStyleGroupId,
  type CellStyleId,
  resolveCellStyleFormat,
} from '../../../commands/cell-styles.js';
import {
  type BuiltinTableStyle,
  type CustomTableStyle,
  pivotStyleGallerySections,
  type TableStyle,
  tableStyleGallerySections,
} from '../../../commands/format-as-table.js';
import { builtinTableStyleElements } from '../../../engine/builtin-table-styles.js';
import {
  composeStyleElements,
  type StyleElementFormats,
  type StyleElementRegion,
  TABLE_ELEMENT_ORDER,
} from '../../../engine/table-style-elements.js';
import type { WorkbookHandle } from '../../../engine/workbook-handle.js';
import { dictionaries } from '../../../i18n/strings.js';
import type { CellBorderSide, CellFormat } from '../../../store/types.js';
import type { ToolbarMenuText } from '../../menu-text.js';
import type { ToolbarLang, ToolbarText } from '../../ribbon-model.js';
import {
  builtinPivotThumbnailElements,
  pivotStyleLabel,
  pivotStyleThumbnail,
} from '../pivot-style-gallery.js';

import {
  createMenu,
  createMenuButton,
  menuIconButton,
  menuIconSpacer,
  menuIdForCommand,
  menuLabeledGrid,
  menuPresetButton,
  menuScrollBody,
  menuSeparator,
  menuTextChip,
} from './general.js';

/** A custom table or PivotTable style as the Format as Table menu lists it. */
interface CustomStyleSwatchItem {
  id: string;
  label: string;
  elements: StyleElementFormats;
  stripeSizes?: CustomTableStyle['stripeSizes'];
}

export interface StylesMenuDeps {
  ribbonLang: ToolbarLang;
  ribbonMenuText: ToolbarMenuText;
  ribbonText: ToolbarText;
  getWorkbook?: () => WorkbookHandle | null;
  getFallbackProfile?: () => CellStyleFallbackProfile;
  customCellStyles?: () => readonly {
    id: string;
    label: string;
    format: {
      bold?: boolean;
      italic?: boolean;
      underline?: CellFormat['underline'];
      color?: string;
      fill?: string;
      fontSize?: number;
    };
  }[];
  customTableStyles?: () => readonly CustomStyleSwatchItem[];
  customPivotTableStyles?: () => readonly CustomStyleSwatchItem[];
}

export interface StylesMenuFactories {
  createTableStyleMenu: (id: string) => HTMLDivElement;
  createPivotStyleMenu: (id: string) => HTMLDivElement;
  createCellStylesMenu: () => HTMLDivElement;
  createCurrencyMenu: () => HTMLDivElement;
}

type StylesMenuText = ToolbarMenuText & {
  tableStyleCustom: string;
  cellStyleCustom: string;
};

const tableStyleSwatchPart = (className: string, background?: string): HTMLDivElement => {
  const part = document.createElement('div');
  part.className = className;
  if (background) part.style.background = background;
  return part;
};

// A built-in swatch is a 5 x 5 table: a header row over four body rows that
// alternate the two row stripes, as the desktop gallery draws it.
const SWATCH_SIZE = 5;
const SWATCH_REGIONS: readonly StyleElementRegion[] = [
  { type: 'headerRow', r0: 0, r1: 0, c0: 0, c1: SWATCH_SIZE - 1 },
  ...[1, 2, 3, 4].map(
    (row): StyleElementRegion => ({
      type: row % 2 === 1 ? 'firstRowStripe' : 'secondRowStripe',
      r0: row,
      r1: row,
      c0: 0,
      c1: SWATCH_SIZE - 1,
    }),
  ),
];
const SWATCH_TEXT = '#7f7f7f';

const swatchBorder = (side: CellBorderSide | undefined): string => {
  if (!side || side === true) return '';
  const width =
    side.style === 'medium' ? 2 : side.style === 'thick' || side.style === 'double' ? 3 : 1;
  return `${width}px ${side.style === 'double' ? 'double' : 'solid'} ${side.color ?? SWATCH_TEXT}`;
};

/** A table style drawn as the desktop gallery draws it: a header row over striped body rows. */
export const tableStyleSwatchGrid = (elements: StyleElementFormats): HTMLDivElement => {
  const grid = tableStyleSwatchPart('fc-tb__tablestyle-swatch__grid');
  for (let row = 0; row < SWATCH_SIZE; row += 1) {
    for (let col = 0; col < SWATCH_SIZE; col += 1) {
      const format = composeStyleElements(elements, TABLE_ELEMENT_ORDER, SWATCH_REGIONS, row, col);
      const cell = document.createElement('span');
      cell.className = 'fc-tb__tablestyle-swatch__cell';
      if (format?.fill) cell.style.background = format.fill;
      cell.style.color = format?.color ?? SWATCH_TEXT;
      if (format?.bold) cell.classList.add('fc-tb__tablestyle-swatch__cell--bold');
      const borders = format?.borders;
      cell.style.borderTop = swatchBorder(borders?.top);
      cell.style.borderBottom = swatchBorder(borders?.bottom);
      cell.style.borderLeft = swatchBorder(borders?.left);
      cell.style.borderRight = swatchBorder(borders?.right);
      grid.appendChild(cell);
    }
  }
  return grid;
};

const createStyleSwatch = (
  name: string,
  elements: StyleElementFormats,
  label: string,
  dataset: 'tableStyle' | 'pivotTableStyle' = 'tableStyle',
): HTMLButtonElement => {
  // A style is applied by name and leaves the table's options alone.
  const btn = createMenuButton({
    className: 'fc-tb__tablestyle-swatch fc-tb__tablestyle-swatch--builtin',
    attr: dataset,
    value: name,
    title: label,
    ariaLabel: label,
  });
  btn.appendChild(tableStyleSwatchGrid(elements));
  return btn;
};

const tableStyleFooterButton = (label: string, action: string): HTMLButtonElement => {
  const btn = menuIconButton(
    label,
    'tableStyleFooter',
    action,
    action === 'new-pivot-style'
      ? 'pivot-style-new'
      : action === 'clear-table-style' || action === 'clear-pivot-style'
        ? 'clear'
        : 'table-style-new',
  );
  btn.classList.add('fc-tb__tablestyle-footer');
  return btn;
};

const cellStyleFooterButton = (label: string, action: string): HTMLButtonElement => {
  const btn = menuIconButton(
    label,
    'cellStyleFooter',
    action,
    action === 'merge-cell-style' ? 'cell-style-merge' : 'cell-style-new',
  );
  btn.classList.add('fc-tb__cellstyle-footer');
  return btn;
};

const currencyFooterButton = (label: string, action: string): HTMLButtonElement => {
  const btn = menuPresetButton(label, 'currencyFooter', action, menuIconSpacer());
  btn.classList.add('fc-tb__currency-footer');
  return btn;
};

const currencyPresetItem = (label: string, symbol: string): HTMLButtonElement => {
  return menuPresetButton(label, 'currencyPreset', symbol, menuIconSpacer());
};

/** Paints a cell-style swatch with the style's own fill, ink and type treatment. */
export const paintCellStyleChip = (
  btn: HTMLElement,
  fmt: {
    bold?: boolean;
    italic?: boolean;
    underline?: CellFormat['underline'];
    color?: string;
    fill?: string;
    fontSize?: number;
  },
): void => {
  btn.style.background = fmt.fill ?? '#ffffff';
  btn.style.color = fmt.color ?? '#1f1f1f';
  if (fmt.bold) btn.style.fontWeight = '700';
  if (fmt.italic) btn.style.fontStyle = 'italic';
  if (fmt.underline) btn.style.textDecoration = 'underline';
  if (fmt.fontSize) btn.style.fontSize = `${Math.min(fmt.fontSize, 13)}px`;
};

export const createStylesMenuFactories = (deps: StylesMenuDeps): StylesMenuFactories => {
  const { ribbonLang, ribbonMenuText, ribbonText } = deps;
  const t = ribbonMenuText as StylesMenuText;

  const cellStyleGalleryLabel = (id: CellStyleId): string => {
    const strings = dictionaries[ribbonLang].cellStylesGallery.styles;
    return strings[id] ?? CELL_STYLES.find((s) => s.id === id)?.label ?? id;
  };

  const cellStyleGroupLabel = (id: CellStyleGroupId): string =>
    dictionaries[ribbonLang].cellStylesGallery.groups[id];

  const createCellStyleChipFromDef = (def: {
    id: string;
    label: string;
    format: {
      bold?: boolean;
      italic?: boolean;
      underline?: CellFormat['underline'];
      color?: string;
      fill?: string;
      fontSize?: number;
    };
  }): HTMLButtonElement => {
    const label = def.label;
    const btn = menuTextChip({
      label,
      className: 'fc-tb__menu-item fc-tb__cellstyle-chip',
      attr: 'cellStyle',
      value: def.id,
    });
    paintCellStyleChip(btn, def.format ?? {});
    return btn;
  };
  const createCellStyleChip = (id: CellStyleId): HTMLButtonElement => {
    const def = CELL_STYLES.find((s) => s.id === id);
    return createCellStyleChipFromDef({
      id,
      label: cellStyleGalleryLabel(id),
      format: def
        ? resolveCellStyleFormat(
            id,
            deps.getWorkbook?.() ?? null,
            deps.getFallbackProfile?.() ?? 'default',
          )
        : {},
    });
  };

  /** Home > Format as Table, or the Table Design gallery (`tableQuickStyles`): only the latter
   *  offers None and Clear, and only the former New PivotTable Style. */
  const createTableStyleMenu = (id: string): HTMLDivElement => {
    const tableDesign = id === 'tableQuickStyles';
    const menu = createMenu(menuIdForCommand(id));
    menu.classList.add('fc-tb__tablestyle-menu');
    const scrollBody = menuScrollBody('fc-tb__tablestyle-scroll', ribbonText.formatTable);
    const familyLabels: Record<TableStyle, string> = {
      light: t.tableStyleLight,
      medium: t.tableStyleMedium,
      dark: t.tableStyleDark,
    };
    const themeColors = deps.getWorkbook?.()?.getTheme?.()?.colors;
    const builtinSwatch = (entry: BuiltinTableStyle): HTMLButtonElement =>
      createStyleSwatch(
        entry.name,
        builtinTableStyleElements(entry.name, themeColors) ?? {},
        entry.name === ''
          ? t.tableStyleNone
          : (t.tableStyleSwatches[entry.name as keyof typeof t.tableStyleSwatches] ?? entry.name),
      );
    for (const section of tableStyleGallerySections(deps.customTableStyles?.() ?? [], {
      none: tableDesign,
    })) {
      scrollBody.append(
        ...menuLabeledGrid({
          label: section.kind === 'custom' ? t.tableStyleCustom : familyLabels[section.kind],
          headingClassName: 'fc-tb__tablestyle-heading',
          gridClassName: 'fc-tb__tablestyle-grid',
          children:
            section.kind === 'custom'
              ? section.styles.map((style) =>
                  createStyleSwatch(style.id, style.elements, style.label),
                )
              : section.styles.map(builtinSwatch),
        }),
      );
    }
    menu.appendChild(scrollBody);
    menu.appendChild(menuSeparator());
    menu.appendChild(tableStyleFooterButton(t.tableStyleNew, 'new-table-style'));
    if (!tableDesign) {
      menu.appendChild(tableStyleFooterButton(t.tableStyleNewPivot, 'new-pivot-style'));
      return menu;
    }
    // Clear shows only while a table is active; the menu's update toggles it.
    const clear = tableStyleFooterButton(t.tableStyleClear, 'clear-table-style');
    clear.hidden = true;
    menu.appendChild(clear);
    return menu;
  };

  /** PivotTable Design > Quick Styles: the full PivotTable style gallery in the shared section
   *  order, New PivotTable Style, and Clear while a PivotTable is active. */
  const createPivotStyleMenu = (id: string): HTMLDivElement => {
    const menu = createMenu(menuIdForCommand(id));
    menu.classList.add('fc-tb__tablestyle-menu', 'fc-tb__pivotstyle-menu');
    const scrollBody = menuScrollBody('fc-tb__tablestyle-scroll', ribbonText.quickStyles);
    const familyLabels: Record<TableStyle, string> = {
      light: t.tableStyleLight,
      medium: t.tableStyleMedium,
      dark: t.tableStyleDark,
    };
    const themeColors = deps.getWorkbook?.()?.getTheme?.()?.colors;
    const swatch = (value: string, label: string, thumb: HTMLElement): HTMLButtonElement => {
      const btn = createMenuButton({
        className: 'fc-tb__tablestyle-swatch fc-tb__pivotstyle-swatch',
        attr: 'pivotTableStyle',
        value,
        title: label,
        ariaLabel: label,
      });
      btn.appendChild(thumb);
      return btn;
    };
    for (const section of pivotStyleGallerySections(deps.customPivotTableStyles?.() ?? [], {
      none: true,
    })) {
      scrollBody.append(
        ...menuLabeledGrid({
          label: section.kind === 'custom' ? t.tableStyleCustom : familyLabels[section.kind],
          headingClassName: 'fc-tb__tablestyle-heading',
          gridClassName: 'fc-tb__tablestyle-grid',
          children:
            section.kind === 'custom'
              ? section.styles.map((style) =>
                  swatch(style.id, style.label, pivotStyleThumbnail(style.elements)),
                )
              : section.names.map((name) =>
                  swatch(
                    name,
                    pivotStyleLabel(name, ribbonLang),
                    pivotStyleThumbnail(builtinPivotThumbnailElements(name, themeColors)),
                  ),
                ),
        }),
      );
    }
    menu.appendChild(scrollBody);
    menu.appendChild(menuSeparator());
    menu.appendChild(tableStyleFooterButton(t.tableStyleNewPivot, 'new-pivot-style'));
    // Clear shows only while a PivotTable is active; the menu's update toggles it.
    const clear = tableStyleFooterButton(t.tableStyleClear, 'clear-pivot-style');
    clear.hidden = true;
    menu.appendChild(clear);
    return menu;
  };

  const createCellStylesMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-cell-styles-home');
    menu.classList.add('fc-tb__cellstyle-menu');
    const scrollBody = menuScrollBody('fc-tb__cellstyle-scroll', ribbonText.cellStyles);
    for (const group of CELL_STYLE_GROUPS) {
      const label = cellStyleGroupLabel(group.id);
      scrollBody.append(
        ...menuLabeledGrid({
          label,
          headingClassName: 'fc-tb__cellstyle-heading',
          gridClassName: 'fc-tb__cellstyle-grid',
          children: group.styleIds.map((id) => createCellStyleChip(id)),
        }),
      );
    }
    const customStyles = deps.customCellStyles?.() ?? [];
    if (customStyles.length > 0) {
      const label = t.cellStyleCustom;
      scrollBody.append(
        ...menuLabeledGrid({
          label,
          headingClassName: 'fc-tb__cellstyle-heading',
          gridClassName: 'fc-tb__cellstyle-grid',
          children: customStyles.map((style) => createCellStyleChipFromDef(style)),
        }),
      );
    }
    menu.appendChild(scrollBody);
    menu.appendChild(menuSeparator());
    menu.appendChild(cellStyleFooterButton(t.cellStyleNew, 'new-cell-style'));
    menu.appendChild(cellStyleFooterButton(t.cellStyleMerge, 'merge-cell-style'));
    return menu;
  };

  const createCurrencyMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-currency-home');
    menu.classList.add('fc-tb__currency-menu');
    menu.append(
      currencyPresetItem(ribbonText.currencyPresetJpy, '¥'),
      currencyPresetItem(ribbonText.currencyPresetUsd, '$'),
      currencyPresetItem(ribbonText.currencyPresetEur, '€'),
      currencyPresetItem(ribbonText.currencyPresetGbp, '£'),
      currencyPresetItem(ribbonText.currencyPresetChf, 'CHF'),
      menuSeparator(),
      currencyFooterButton(ribbonText.moreCurrencyFormats, 'more'),
    );
    return menu;
  };

  return { createTableStyleMenu, createPivotStyleMenu, createCellStylesMenu, createCurrencyMenu };
};
