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

import { createMenu, menuIdForCommand, menuSeparator } from './general.js';

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
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'app__tablestyle-swatch';
  btn.setAttribute('role', 'menuitem');
  btn.dataset[dataset] = actionStyleId;
  btn.dataset.tableColor = color;
  btn.dataset.tableVariant = variant;
  btn.title = label;
  btn.setAttribute('aria-label', label);
  btn.style.cssText =
    'display:flex;flex-direction:column;width:46px;height:34px;padding:0;' +
    'border:1px solid #c8c6c4;border-radius:2px;overflow:hidden;cursor:pointer;background:#fff;';
  const head = document.createElement('div');
  head.style.cssText = `flex:0 0 9px;background:${swatch.header};`;
  btn.appendChild(head);
  for (let i = 0; i < 3; i += 1) {
    const row = document.createElement('div');
    const rowFill = banded && i % 2 === 1 ? swatch.band : '#ffffff';
    row.style.cssText = `flex:1;display:flex;background:${rowFill};`;
    if (firstCol) {
      const emphasis = document.createElement('div');
      emphasis.style.cssText = `flex:0 0 10px;background:${swatch.header};`;
      row.appendChild(emphasis);
      const rest = document.createElement('div');
      rest.style.cssText = 'flex:1;';
      row.appendChild(rest);
    }
    btn.appendChild(row);
  }
  return btn;
};

const tableStyleFooterButton = (label: string, action: string): HTMLButtonElement => {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'app__menu-item app__tablestyle-footer';
  btn.setAttribute('role', 'menuitem');
  btn.dataset.tableStyleFooter = action;
  btn.textContent = label;
  btn.style.cssText =
    'display:flex;width:100%;padding:6px 12px;border:0;background:transparent;cursor:pointer;text-align:left;font-size:12px;color:#1f1f1f;';
  return btn;
};

const cellStyleFooterButton = (label: string, action: string): HTMLButtonElement => {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'app__menu-item app__cellstyle-footer';
  btn.setAttribute('role', 'menuitem');
  btn.dataset.cellStyleFooter = action;
  btn.textContent = label;
  btn.style.cssText =
    'display:flex;width:100%;padding:6px 12px;border:0;background:transparent;cursor:pointer;text-align:left;font-size:12px;color:#1f1f1f;';
  return btn;
};

const currencyFooterButton = (label: string, action: string): HTMLButtonElement => {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'app__menu-item app__currency-footer';
  btn.setAttribute('role', 'menuitem');
  btn.dataset.currencyFooter = action;
  btn.textContent = label;
  btn.style.cssText =
    'display:flex;width:100%;padding:6px 12px;border:0;background:transparent;cursor:pointer;text-align:left;font-size:12px;color:#1f1f1f;';
  return btn;
};

const currencyPresetItem = (label: string, symbol: string): HTMLButtonElement => {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'app__menu-item';
  btn.setAttribute('role', 'menuitem');
  btn.dataset.currencyPreset = symbol;
  btn.textContent = label;
  return btn;
};

export const createStylesMenuFactories = (deps: StylesMenuDeps): StylesMenuFactories => {
  const { ribbonLang, ribbonMenuText: t, ribbonText } = deps;

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
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'app__menu-item app__cellstyle-chip';
    btn.setAttribute('role', 'menuitem');
    btn.dataset.cellStyle = def.id;
    const label = def.label;
    btn.title = label;
    btn.setAttribute('aria-label', label);
    btn.textContent = label;
    const fmt = def.format ?? {};
    const css: string[] = [
      'display:flex;align-items:center;justify-content:center;',
      'min-width:88px;height:28px;padding:2px 8px;',
      'border:1px solid #d0cfcd;border-radius:2px;cursor:pointer;',
      'font-size:11px;line-height:1.1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;',
      `background:${fmt.fill ?? '#ffffff'};`,
      `color:${fmt.color ?? '#1f1f1f'};`,
    ];
    if (fmt.bold) css.push('font-weight:700;');
    if (fmt.italic) css.push('font-style:italic;');
    if (fmt.underline) css.push('text-decoration:underline;');
    if (fmt.fontSize) css.push(`font-size:${Math.min(fmt.fontSize, 13)}px;`);
    btn.style.cssText = css.join('');
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
    menu.style.width = 'auto';
    menu.style.maxWidth = '420px';
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
      const heading = document.createElement('div');
      heading.textContent = intensity.label;
      heading.style.cssText = 'padding:6px 10px 2px;font-size:11px;font-weight:600;color:#605e5c;';
      menu.appendChild(heading);
      const grid = document.createElement('div');
      grid.setAttribute('role', 'group');
      grid.setAttribute('aria-label', intensity.label);
      grid.style.cssText =
        'display:grid;grid-template-columns:repeat(7,46px);gap:4px;padding:2px 8px 6px;';
      for (const variant of intensity.variants) {
        for (const color of TABLE_STYLE_COLORS) {
          grid.appendChild(createTableStyleSwatch(intensity.id, color, variant, intensity.label));
        }
      }
      menu.appendChild(grid);
    }
    const customTableStyles = deps.customTableStyles?.() ?? [];
    if (customTableStyles.length > 0) {
      const heading = document.createElement('div');
      heading.textContent = ribbonLang === 'ja' ? 'ユーザー設定' : 'Custom';
      heading.style.cssText = 'padding:6px 10px 2px;font-size:11px;font-weight:600;color:#605e5c;';
      menu.appendChild(heading);
      const grid = document.createElement('div');
      grid.setAttribute('role', 'group');
      grid.setAttribute('aria-label', heading.textContent);
      grid.style.cssText =
        'display:grid;grid-template-columns:repeat(7,46px);gap:4px;padding:2px 8px 6px;';
      for (const style of customTableStyles) {
        grid.appendChild(
          createTableStyleSwatch(
            style.style,
            style.color ?? '#5b9bd5',
            style.variant,
            style.label,
            style.id,
          ),
        );
      }
      menu.appendChild(grid);
    }
    const customPivotTableStyles = deps.customPivotTableStyles?.() ?? [];
    if (customPivotTableStyles.length > 0) {
      const heading = document.createElement('div');
      heading.textContent =
        ribbonLang === 'ja' ? 'ピボットテーブル ユーザー設定' : 'Custom PivotTable';
      heading.style.cssText = 'padding:6px 10px 2px;font-size:11px;font-weight:600;color:#605e5c;';
      menu.appendChild(heading);
      const grid = document.createElement('div');
      grid.setAttribute('role', 'group');
      grid.setAttribute('aria-label', heading.textContent);
      grid.style.cssText =
        'display:grid;grid-template-columns:repeat(7,46px);gap:4px;padding:2px 8px 6px;';
      for (const style of customPivotTableStyles) {
        grid.appendChild(
          createTableStyleSwatch(
            style.style,
            style.color ?? '#5b9bd5',
            style.variant,
            style.label,
            style.id,
            'pivotTableStyle',
          ),
        );
      }
      menu.appendChild(grid);
    }
    menu.appendChild(menuSeparator());
    menu.appendChild(tableStyleFooterButton(t.tableStyleNew, 'new-table-style'));
    menu.appendChild(tableStyleFooterButton(t.tableStyleNewPivot, 'new-pivot-style'));
    return menu;
  };

  const createCellStylesMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-cell-styles-home');
    menu.style.width = 'auto';
    menu.style.maxWidth = '620px';
    for (const group of CELL_STYLE_GROUPS) {
      const heading = document.createElement('div');
      heading.textContent = cellStyleGroupLabel(group.id);
      heading.style.cssText = 'padding:6px 10px 2px;font-size:11px;font-weight:600;color:#605e5c;';
      menu.appendChild(heading);
      const grid = document.createElement('div');
      grid.setAttribute('role', 'group');
      grid.setAttribute('aria-label', cellStyleGroupLabel(group.id));
      grid.style.cssText =
        'display:grid;grid-template-columns:repeat(6,minmax(88px,1fr));gap:4px;padding:2px 8px 6px;';
      for (const id of group.styleIds) grid.appendChild(createCellStyleChip(id));
      menu.appendChild(grid);
    }
    const customStyles = deps.customCellStyles?.() ?? [];
    if (customStyles.length > 0) {
      const heading = document.createElement('div');
      heading.textContent = ribbonLang === 'ja' ? 'ユーザー設定' : 'Custom';
      heading.style.cssText = 'padding:6px 10px 2px;font-size:11px;font-weight:600;color:#605e5c;';
      menu.appendChild(heading);
      const grid = document.createElement('div');
      grid.setAttribute('role', 'group');
      grid.setAttribute('aria-label', heading.textContent);
      grid.style.cssText =
        'display:grid;grid-template-columns:repeat(6,minmax(88px,1fr));gap:4px;padding:2px 8px 6px;';
      for (const style of customStyles) grid.appendChild(createCellStyleChipFromDef(style));
      menu.appendChild(grid);
    }
    menu.appendChild(menuSeparator());
    menu.appendChild(cellStyleFooterButton(t.cellStyleNew, 'new-cell-style'));
    menu.appendChild(cellStyleFooterButton(t.cellStyleMerge, 'merge-cell-style'));
    return menu;
  };

  const createCurrencyMenu = (): HTMLDivElement => {
    const menu = createMenu('menu-currency-home');
    menu.style.width = 'auto';
    menu.style.maxWidth = '320px';
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
