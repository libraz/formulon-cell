// Generic menu primitives shared across all ribbon dropdown factories.
// Extracted from main.ts so the per-tab menu modules can build off the same
// building blocks without dragging the whole playground entry along with them.

import { createExcelRibbonSvg } from '../../excel-ribbon-icons.js';
import { prepareMenu } from '../../menu-a11y.js';
import { createRibbonButton } from '../button.js';
import { ribbonDropdownMenuIdForCommand } from '../dynamic-dropdowns.js';

const MENU_EXCEL_ICON_OVERRIDES: Readonly<Record<string, string>> = {
  'addin-get': 'addIn',
  'addin-manage': 'addIn',
  'addin-my': 'addIn',
  'audit-clear-all': 'clearArrows',
  'audit-clear-dependents': 'clearArrows',
  'audit-clear-precedents': 'clearArrows',
  'autosum-average': 'autosum',
  'autosum-count': 'autosum',
  'autosum-max': 'autosum',
  'autosum-min': 'autosum',
  'autosum-more': 'function',
  'autosum-sum': 'autosum',
  'break-page': 'pageBreaks',
  'break-remove': 'pageBreaks',
  'bring-forward': 'shapes',
  'bring-front': 'shapes',
  clear: 'clear',
  'cell-style-merge': 'cellStyles',
  'cell-style-new': 'cellStyles',
  'clear-all': 'clearAll',
  'clear-conditional': 'clearConditional',
  'clear-contents': 'clearContents',
  'clear-formats': 'clearFormats',
  'clear-hyperlinks': 'clearHyperlinks',
  'clear-comments': 'clearComments',
  'comment-delete': 'commentAdd',
  'comment-delete-all': 'commentAdd',
  conditional: 'conditional',
  copy: 'copy',
  'currency-chf': 'currency',
  'currency-dollar': 'currency',
  'currency-euro': 'currency',
  'currency-more': 'currency',
  'currency-pound': 'currency',
  'currency-yen': 'currency',
  'defined-name-define': 'names',
  'defined-name-create-bottom': 'namesCreateBottom',
  'defined-name-create-left': 'namesCreateLeft',
  'defined-name-create-right': 'namesCreateRight',
  'defined-name-create-top': 'namesCreateTop',
  'defined-name-manager': 'names',
  'defined-name-use': 'names',
  'delete-cells': 'deleteCells',
  'delete-col': 'deleteCols',
  'delete-row': 'deleteRows',
  'delete-sheet': 'deleteRows',
  'delete-shift-left': 'deleteCols',
  'delete-shift-up': 'deleteRows',
  'error-checking': 'errorChecking',
  'fill-days': 'fill',
  'fill-down': 'fillDown',
  'fill-group': 'fillGroup',
  'fill-justify': 'fillJustify',
  'fill-left': 'fillLeft',
  'fill-months': 'fill',
  'fill-right': 'fillRight',
  'fill-series': 'fillSeries',
  'fill-up': 'fillUp',
  'fill-weekdays': 'fill',
  'fill-years': 'fill',
  'flash-fill': 'flashFill',
  'filter-advanced': 'filterAdvanced',
  'filter-by-value': 'filterByValue',
  'filter-clear': 'filterClear',
  'filter-reapply': 'filterReapply',
  'filter-toggle': 'filterToggle',
  'format-col-autofit': 'formatCells',
  'format-col-width': 'formatCells',
  'format-dialog': 'formatCells',
  'format-hide-cols': 'formatCells',
  'format-hide-rows': 'formatCells',
  'format-hide-sheet': 'formatCells',
  'format-lock': 'protect',
  'format-move-copy': 'formatCells',
  'format-move-left': 'formatCells',
  'format-move-right': 'formatCells',
  'format-protect': 'protect',
  'format-rename-sheet': 'formatCells',
  'format-row-autofit': 'formatCells',
  'format-row-height': 'formatCells',
  'format-show-cols': 'formatCells',
  'format-show-rows': 'formatCells',
  'format-unhide-sheet': 'formatCells',
  'format-unlock': 'protect',
  find: 'find',
  'find-comments': 'findComments',
  'find-conditional': 'findConditional',
  'find-constants': 'findConstants',
  'find-formulas': 'findFormulas',
  'find-validation': 'findValidation',
  'freeze-col': 'freeze',
  'freeze-panes': 'freeze',
  'freeze-row': 'freeze',
  'go-to': 'goTo',
  'go-to-special': 'goToSpecial',
  'ignore-error': 'errorChecking',
  'insert-cells': 'insertCells',
  'insert-col': 'insertCols',
  'insert-row': 'insertRows',
  'insert-sheet': 'insertRows',
  'insert-shift-down': 'insertRows',
  'insert-shift-right': 'insertCols',
  'link-clear': 'link',
  'link-edit': 'link',
  'link-external': 'link',
  'link-open': 'link',
  merge: 'merge',
  'name-manager': 'names',
  'new-table-style': 'tableStyle',
  'object-select': 'objectSelect',
  'paste-all': 'paste',
  'paste-formats': 'paint',
  'paste-formulas': 'pasteFormulas',
  'paste-formulas-numfmt': 'pasteFormulas',
  'paste-special': 'pasteSpecial',
  'paste-transpose': 'pasteTranspose',
  'paste-values': 'pasteValues',
  'paste-values-numfmt': 'pasteValues',
  pane: 'page',
  picture: 'picture',
  'pdf-create': 'pdf',
  'pdf-preferences': 'pdf',
  'pdf-share': 'pdf',
  'pivot-new-sheet': 'pivotTable',
  'pivot-range': 'pivotTable',
  'pivot-existing-sheet': 'pivotExistingSheet',
  'pivot-refresh': 'pivotTable',
  'pivot-recommended': 'pivotRecommended',
  'pivot-style-new': 'tableStyle',
  'print-area-add': 'printArea',
  'print-area-set': 'printArea',
  'protect-lock-cell': 'protect',
  'protect-allow-ranges': 'protect',
  'protect-clear-ranges': 'protect',
  'protect-sheet': 'protect',
  'protect-unlock-cell': 'protect',
  'protect-unprotect-sheet': 'protect',
  'protect-unprotect-workbook': 'protect',
  'protect-workbook': 'protect',
  replace: 'replaceFind',
  reset: 'clear',
  'remove-duplicates': 'removeDuplicates',
  wrap: 'wrap',
  'script-clear': 'script',
  'script-custom': 'script',
  'script-lowercase': 'script',
  'script-trim': 'script',
  'script-uppercase': 'script',
  'send-back': 'shapes',
  'send-backward': 'shapes',
  'selection-pane': 'selectionPane',
  'sort-asc': 'sortAsc',
  'sort-custom': 'sortCustom',
  'sort-desc': 'sortDesc',
  'sort-filter': 'sortFilter',
  'symbol-more': 'function',
  'table-style-new': 'tableStyle',
  'text-column-comma': 'textToColumns',
  'text-column-custom': 'textToColumns',
  'text-column-semicolon': 'textToColumns',
  'text-column-space': 'textToColumns',
  'text-column-tab': 'textToColumns',
  'trace-error': 'trace',
  'title-autosave': 'autosave',
  'title-comments': 'commentAdd',
  'title-save': 'save',
  'title-save-as': 'saveAs',
  'title-share': 'share',
  'underline-double': 'underlineDouble',
  'underline-single': 'underlineSingle',
  'validation-circle': 'dataValidationCircle',
  'validation-clear-circles': 'dataValidationClearCircles',
  'validation-clear-rules': 'dataValidationClearRules',
  'validation-settings': 'dataValidation',
  'watch-add': 'watch',
  'watch-delete': 'watch',
  'watch-delete-all': 'watch',
  'watch-open': 'watch',
};

const VISUAL_TILE_EXCEL_ICON_OVERRIDES: Readonly<Record<string, string>> = {
  'chart-area': 'chartArea',
  'chart-bar': 'chartBar',
  'chart-column': 'chartColumn',
  'chart-line': 'chartLine',
  'chart-pie': 'chartPie',
  'chart-recommended': 'chartRecommended',
  'chart-scatter': 'chartScatter',
  'device-picture': 'devicePicture',
  'online-picture': 'onlinePicture',
  'screen-clipping': 'screenClipping',
  'shape-arrow': 'shapeArrow',
  'shape-diamond': 'shapeDiamond',
  'shape-line': 'shapeLine',
  'shape-oval': 'shapeOval',
  'shape-rectangle': 'shapeRectangle',
  'shape-rounded-rectangle': 'shapeRoundedRectangle',
  'shape-triangle': 'shapeTriangle',
  'screenshot-window': 'screenshotWindow',
  'stock-picture': 'stockPicture',
  'theme-contrast': 'themeContrast',
  'theme-dark': 'themeDark',
  'theme-light': 'themeLight',
};

const menuDiv = (
  className: string,
  opts: { id?: string; text?: string; role?: string; ariaLabel?: string; hidden?: boolean } = {},
): HTMLDivElement => {
  const div = document.createElement('div');
  div.className = className;
  if (opts.id) div.id = opts.id;
  if (opts.text !== undefined) div.textContent = opts.text;
  if (opts.role) div.setAttribute('role', opts.role);
  if (opts.ariaLabel) div.setAttribute('aria-label', opts.ariaLabel);
  if (opts.hidden) div.hidden = true;
  return div;
};

/** Maps a ribbon command id to the DOM id its dropdown menu should use.
 *  Falls back to the command id when no entry exists (callers can pass any
 *  string and get a stable id back). */
export const menuIdForCommand = (commandId: string): string =>
  ribbonDropdownMenuIdForCommand(commandId) ?? commandId;

export const createMenu = (id: string): HTMLDivElement => {
  const menu = menuDiv('app__menu', { id, hidden: true });
  prepareMenu(menu);
  return menu;
};

export type MenuButtonOptions = {
  className: string;
  attr: string;
  value: string;
  title?: string;
  ariaLabel?: string;
};

export const createMenuButton = (opts: MenuButtonOptions): HTMLButtonElement => {
  return createRibbonButton({
    className: opts.className,
    role: 'menuitem',
    dataset: { [opts.attr]: opts.value },
    title: opts.title,
    ariaLabel: opts.ariaLabel,
  });
};

const menuSpan = (
  className: string,
  opts: { text?: string; ariaHidden?: boolean } = {},
): HTMLSpanElement => {
  const span = document.createElement('span');
  span.className = className;
  if (opts.text !== undefined) span.textContent = opts.text;
  if (opts.ariaHidden) span.setAttribute('aria-hidden', 'true');
  return span;
};

export const menuIconButton = (
  label: string,
  attr: string,
  value: string,
  icon: string,
): HTMLButtonElement => {
  const button = createMenuButton({
    className: 'app__menu-item app__menu-item--iconic',
    attr,
    value,
  });

  const iconSpan = menuSpan(`app__menu-icon app__menu-icon--${icon}`, { ariaHidden: true });
  const svg = createExcelRibbonSvg(MENU_EXCEL_ICON_OVERRIDES[icon] ?? '', 'app__menu-icon-svg');
  if (svg) {
    iconSpan.classList.add('app__menu-icon--svg');
    iconSpan.append(svg);
  }

  button.append(iconSpan, menuSpan('app__menu-item__text', { text: label }));
  return button;
};

export const menuPresetButton = (
  label: string,
  attr: string,
  value: string,
  leading: Node,
): HTMLButtonElement => {
  const button = createMenuButton({
    className: 'app__menu-item app__menu-item--preset',
    attr,
    value,
  });

  button.append(leading, menuSpan('app__menu-item__text', { text: label }));
  return button;
};

export type MenuTextChipOptions = {
  label: string;
  attr: string;
  value: string;
  className: string;
  labelClassName?: string;
};

export const menuTextChip = (opts: MenuTextChipOptions): HTMLButtonElement => {
  const button = createMenuButton({
    className: opts.className,
    attr: opts.attr,
    value: opts.value,
    title: opts.label,
    ariaLabel: opts.label,
  });
  button.append(
    menuSpan(opts.labelClassName ?? 'app__menu-text-chip__label', { text: opts.label }),
  );
  return button;
};

export const menuIconSpacer = (): HTMLSpanElement => {
  return menuSpan('app__menu-item__icon-spacer');
};

export const menuSubmenuTrigger = (
  button: HTMLButtonElement,
  dataset?: Record<string, string>,
  opts: { controlsId?: string } = {},
): HTMLButtonElement => {
  button.classList.add('app__menu-item--submenu');
  button.setAttribute('aria-haspopup', 'menu');
  button.setAttribute('aria-expanded', 'false');
  if (opts.controlsId) button.setAttribute('aria-controls', opts.controlsId);
  for (const [key, value] of Object.entries(dataset ?? {})) button.dataset[key] = value;

  button.appendChild(menuSpan('app__menu-item__caret', { ariaHidden: true }));
  return button;
};

export type ColorSwatchButtonOptions = {
  label: string;
  attr: string;
  value: string;
  color: string | null;
};

export const colorSwatchButton = (opts: ColorSwatchButtonOptions): HTMLButtonElement => {
  const button = createMenuButton({
    className: ['app__color-swatch', opts.color === null ? 'app__color-swatch--none' : '']
      .filter(Boolean)
      .join(' '),
    attr: opts.attr,
    value: opts.value,
    title: opts.label,
    ariaLabel: opts.label,
  });
  if (opts.color !== null) button.style.setProperty('--app-menu-swatch-color', opts.color);

  button.append(menuSpan('app__color-swatch__chip', { ariaHidden: true }));
  return button;
};

export const colorSwatchGrid = (className?: string): HTMLDivElement => {
  return menuDiv(['app__color-swatch-grid', className].filter(Boolean).join(' '), {
    role: 'presentation',
  });
};

export const symbolMenuTile = (symbol: string): HTMLButtonElement => {
  const button = createMenuButton({
    className: 'app__symbol-tile',
    attr: 'symbol',
    value: symbol,
    title: symbol,
    ariaLabel: symbol,
  });

  button.append(menuSpan('app__symbol-tile__glyph', { text: symbol, ariaHidden: true }));
  return button;
};

export const symbolMenuGrid = (groupLabel: string, symbols: readonly string[]): HTMLDivElement => {
  const grid = menuDiv('app__symbol-grid', { role: 'presentation', ariaLabel: groupLabel });
  grid.append(...symbols.map((symbol) => symbolMenuTile(symbol)));
  return grid;
};

export type VisualMenuTileOptions = {
  label: string;
  attr: string;
  value: string;
  icon: string;
  className?: string;
};

export const visualMenuTile = (opts: VisualMenuTileOptions): HTMLButtonElement => {
  const button = createMenuButton({
    className: ['app__visual-tile', opts.className].filter(Boolean).join(' '),
    attr: opts.attr,
    value: opts.value,
    title: opts.label,
    ariaLabel: opts.label,
  });
  const iconSpan = menuSpan(`app__visual-tile__icon app__visual-tile__icon--${opts.icon}`, {
    ariaHidden: true,
  });
  const svg = createExcelRibbonSvg(
    VISUAL_TILE_EXCEL_ICON_OVERRIDES[opts.icon] ?? '',
    'app__visual-tile__icon-svg',
  );
  if (svg) {
    iconSpan.classList.add('app__visual-tile__icon--svg');
    iconSpan.append(svg);
  }

  button.append(iconSpan, menuSpan('app__visual-tile__label', { text: opts.label }));
  return button;
};

export const visualMenuGrid = (className?: string): HTMLDivElement => {
  return menuDiv(['app__visual-grid', className].filter(Boolean).join(' '), {
    role: 'presentation',
  });
};

export const visualMenuTileGrid = (
  className: string,
  tiles: readonly VisualMenuTileOptions[],
): HTMLDivElement => {
  const grid = visualMenuGrid(className);
  grid.append(...tiles.map((tile) => visualMenuTile(tile)));
  return grid;
};

export const menuSeparator = (): HTMLDivElement => {
  return menuDiv('app__menu-sep', { role: 'separator' });
};

export const menuScrollBody = (className: string, ariaLabel?: string): HTMLDivElement => {
  return menuDiv(className, { role: 'group', ariaLabel });
};

export const menuSectionHeader = (label: string): HTMLDivElement => {
  return menuDiv('app__menu-heading', { role: 'presentation', text: label });
};

export type MenuLabeledGridOptions = {
  label: string;
  headingClassName: string;
  gridClassName: string;
  children: readonly Node[];
};

export const menuLabeledGrid = (opts: MenuLabeledGridOptions): [HTMLDivElement, HTMLDivElement] => {
  const heading = menuDiv(opts.headingClassName, { text: opts.label });
  const grid = menuDiv(opts.gridClassName, { role: 'group', ariaLabel: opts.label });
  grid.append(...opts.children);

  return [heading, grid];
};

export type SubmenuOptions = {
  id?: string;
  className: string;
  label: string;
  dataset?: Record<string, string>;
};

export const createSubmenu = (opts: SubmenuOptions): HTMLDivElement => {
  const submenu = menuDiv(opts.className, {
    id: opts.id,
    role: 'menu',
    ariaLabel: opts.label,
    hidden: true,
  });
  for (const [key, value] of Object.entries(opts.dataset ?? {})) submenu.dataset[key] = value;
  return submenu;
};

export const submenuItemText = (text: string): HTMLSpanElement => {
  return menuSpan('app__submenu-item__text', { text });
};
