// Home tab style-gallery dropdowns: Table Styles, Cell Styles, Currency. Each
// builds a swatch / chip / preset list that emits a `data-table-*`,
// `data-cell-style*`, or `data-currency*` attribute the playground dispatcher
// reads. Kept in one file because they share the same factory pattern (deps
// bundle of ribbonLang + dictionaries + ribbonText) and they all sit in the
// Home tab's Styles group.

import {
  CELL_STYLE_GROUPS,
  CELL_STYLES,
  type CellStyleGroupId,
  type CellStyleId,
  dictionaries,
  TABLE_STYLE_COLORS,
  type TableStyle,
  type ToolbarLang,
  type ToolbarMenuText,
  type ToolbarText,
  tableStyleSwatch,
} from '@libraz/formulon-cell';

import {
  createMenu,
  createMenuButton,
  menuIconButton,
  menuIdForCommand,
  menuLabeledGrid,
  menuScrollBody,
  menuSeparator,
  menuTextChip,
} from './general.js';

export interface StylesMenuDeps {
  ribbonLang: ToolbarLang;
  ribbonMenuText: ToolbarMenuText;
  ribbonText: ToolbarText;
  customCellStyles?: () => readonly {
    id: string;
    label: string;
    format: {
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      color?: string;
      fill?: string;
      fontSize?: number;
    };
  }[];
  customTableStyles?: () => readonly {
    id: string;
    label: string;
    style: TableStyle;
    color?: string;
    variant: TableVariantId;
  }[];
  customPivotTableStyles?: () => readonly {
    id: string;
    label: string;
    style: TableStyle;
    color?: string;
    variant: TableVariantId;
  }[];
}

export interface StylesMenuFactories {
  createTableStyleMenu: (id: string) => HTMLDivElement;
  createCellStylesMenu: () => HTMLDivElement;
  createCurrencyMenu: () => HTMLDivElement;
}

type StylesMenuText = ToolbarMenuText & {
  tableStyleCustom: string;
  pivotTableStyleCustom: string;
  cellStyleCustom: string;
};

export type TableVariantId = 'plain' | 'banded' | 'firstCol' | 'bandedFirstCol';

const TABLE_VARIANTS_LIGHT_MEDIUM: TableVariantId[] = [
  'plain',
  'banded',
  'firstCol',
  'bandedFirstCol',
];
const TABLE_VARIANTS_DARK: TableVariantId[] = ['banded'];

export const tableVariantOptions = (
  variant: TableVariantId,
): { banded: boolean; firstCol: boolean } => {
  switch (variant) {
    case 'plain':
      return { banded: false, firstCol: false };
    case 'banded':
      return { banded: true, firstCol: false };
    case 'firstCol':
      return { banded: false, firstCol: true };
    case 'bandedFirstCol':
      return { banded: true, firstCol: true };
  }
};

const tableStyleSwatchPart = (className: string, background?: string): HTMLDivElement => {
  const part = document.createElement('div');
  part.className = className;
  if (background) part.style.background = background;
  return part;
};

const createTableStyleSwatch = (
  style: TableStyle,
  color: string,
  variant: TableVariantId,
  label: string,
  actionStyleId: string = style,
  dataset: 'tableStyle' | 'pivotTableStyle' = 'tableStyle',
): HTMLButtonElement => {
  const swatch = tableStyleSwatch(style, color);
  const { banded, firstCol } = tableVariantOptions(variant);
  const btn = createMenuButton({
    className: 'app__tablestyle-swatch',
    attr: dataset,
    value: actionStyleId,
    title: label,
    ariaLabel: label,
  });
  btn.dataset.tableColor = color;
  btn.dataset.tableVariant = variant;
  btn.appendChild(tableStyleSwatchPart('app__tablestyle-swatch__head', swatch.header));
  for (let i = 0; i < 3; i += 1) {
    const rowFill = banded && i % 2 === 1 ? swatch.band : '#ffffff';
    const row = tableStyleSwatchPart('app__tablestyle-swatch__row', rowFill);
    if (firstCol) {
      row.append(
        tableStyleSwatchPart('app__tablestyle-swatch__first-col', swatch.header),
        tableStyleSwatchPart('app__tablestyle-swatch__rest'),
      );
    }
    btn.appendChild(row);
  }
  return btn;
};

const tableStyleFooterButton = (label: string, action: string): HTMLButtonElement => {
  const btn = menuIconButton(
    label,
    'tableStyleFooter',
    action,
    action === 'new-pivot-style' ? 'pivot-style-new' : 'table-style-new',
  );
  btn.classList.add('app__tablestyle-footer');
  return btn;
};

const cellStyleFooterButton = (label: string, action: string): HTMLButtonElement => {
  const btn = menuIconButton(
    label,
    'cellStyleFooter',
    action,
    action === 'merge-cell-style' ? 'cell-style-merge' : 'cell-style-new',
  );
  btn.classList.add('app__cellstyle-footer');
  return btn;
};

const currencyFooterButton = (label: string, action: string): HTMLButtonElement => {
  const btn = menuIconButton(label, 'currencyFooter', action, 'currency-more');
  btn.classList.add('app__currency-footer');
  return btn;
};

const currencyPresetItem = (label: string, symbol: string): HTMLButtonElement => {
  const icon =
    symbol === '¥'
      ? 'currency-yen'
      : symbol === '$'
        ? 'currency-dollar'
        : symbol === '€'
          ? 'currency-euro'
          : symbol === '£'
            ? 'currency-pound'
            : 'currency-chf';
  return menuIconButton(label, 'currencyPreset', symbol, icon);
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
      underline?: boolean;
      color?: string;
      fill?: string;
      fontSize?: number;
    };
  }): HTMLButtonElement => {
    const label = def.label;
    const btn = menuTextChip({
      label,
      className: 'app__menu-item app__cellstyle-chip',
      attr: 'cellStyle',
      value: def.id,
    });
    const fmt = def.format ?? {};
    btn.style.background = fmt.fill ?? '#ffffff';
    btn.style.color = fmt.color ?? '#1f1f1f';
    if (fmt.bold) btn.style.fontWeight = '700';
    if (fmt.italic) btn.style.fontStyle = 'italic';
    if (fmt.underline) btn.style.textDecoration = 'underline';
    if (fmt.fontSize) btn.style.fontSize = `${Math.min(fmt.fontSize, 13)}px`;
    return btn;
  };
  const createCellStyleChip = (id: CellStyleId): HTMLButtonElement => {
    const def = CELL_STYLES.find((s) => s.id === id);
    return createCellStyleChipFromDef({
      id,
      label: cellStyleGalleryLabel(id),
      format: def?.format ?? {},
    });
  };

  const createTableStyleMenu = (id: string): HTMLDivElement => {
    const menu = createMenu(menuIdForCommand(id));
    menu.classList.add('app__tablestyle-menu');
    const scrollBody = menuScrollBody('app__tablestyle-scroll', ribbonText.formatTable);
    const intensities: {
      id: TableStyle;
      label: string;
      variants: readonly TableVariantId[];
    }[] = [
      { id: 'light', label: t.tableStyleLight, variants: TABLE_VARIANTS_LIGHT_MEDIUM },
      { id: 'medium', label: t.tableStyleMedium, variants: TABLE_VARIANTS_LIGHT_MEDIUM },
      { id: 'dark', label: t.tableStyleDark, variants: TABLE_VARIANTS_DARK },
    ];
    for (const intensity of intensities) {
      scrollBody.append(
        ...menuLabeledGrid({
          label: intensity.label,
          headingClassName: 'app__tablestyle-heading',
          gridClassName: 'app__tablestyle-grid',
          children: intensity.variants.flatMap((variant) =>
            TABLE_STYLE_COLORS.map((color) =>
              createTableStyleSwatch(intensity.id, color, variant, intensity.label),
            ),
          ),
        }),
      );
    }
    const customTableStyles = deps.customTableStyles?.() ?? [];
    if (customTableStyles.length > 0) {
      const label = t.tableStyleCustom;
      scrollBody.append(
        ...menuLabeledGrid({
          label,
          headingClassName: 'app__tablestyle-heading',
          gridClassName: 'app__tablestyle-grid',
          children: customTableStyles.map((style) =>
            createTableStyleSwatch(
              style.style,
              style.color ?? '#5b9bd5',
              style.variant,
              style.label,
              style.id,
            ),
          ),
        }),
      );
    }
    const customPivotTableStyles = deps.customPivotTableStyles?.() ?? [];
    if (customPivotTableStyles.length > 0) {
      const label = t.pivotTableStyleCustom;
      scrollBody.append(
        ...menuLabeledGrid({
          label,
          headingClassName: 'app__tablestyle-heading',
          gridClassName: 'app__tablestyle-grid',
          children: customPivotTableStyles.map((style) =>
            createTableStyleSwatch(
              style.style,
              style.color ?? '#5b9bd5',
              style.variant,
              style.label,
              style.id,
              'pivotTableStyle',
            ),
          ),
        }),
      );
    }
    menu.appendChild(scrollBody);
    menu.appendChild(menuSeparator());
    menu.appendChild(tableStyleFooterButton(t.tableStyleNew, 'new-table-style'));
    menu.appendChild(tableStyleFooterButton(t.tableStyleNewPivot, 'new-pivot-style'));
    return menu;
  };

  const createCellStylesMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-cell-styles-home');
    menu.classList.add('app__cellstyle-menu');
    const scrollBody = menuScrollBody('app__cellstyle-scroll', ribbonText.cellStyles);
    for (const group of CELL_STYLE_GROUPS) {
      const label = cellStyleGroupLabel(group.id);
      scrollBody.append(
        ...menuLabeledGrid({
          label,
          headingClassName: 'app__cellstyle-heading',
          gridClassName: 'app__cellstyle-grid',
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
          headingClassName: 'app__cellstyle-heading',
          gridClassName: 'app__cellstyle-grid',
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
    menu.classList.add('app__currency-menu');
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

  return { createTableStyleMenu, createCellStylesMenu, createCurrencyMenu };
};
