import {
  activateSheetView,
  addAllowedEditRange,
  addSheet,
  aggregateSelection,
  analyzeAccessibilityCells,
  analyzeSpellingCells,
  applyAdvancedFilter,
  applyCellStyle,
  applyConditionalPresetAction,
  applyMerge,
  applyPasteSpecial,
  applyTextScriptToRange,
  applyUnmerge,
  attachFilterDropdown,
  autoSum,
  backstageMenuText,
  boundingRange,
  buildRibbonModel,
  buildTranslationReviewItems,
  bumpDecimals,
  bumpIndent,
  CELL_STYLE_GROUPS,
  CELL_STYLES,
  type CellBorderStyle,
  type CellStyleGroupId,
  type CellStyleId,
  type ClipboardSnapshot,
  type ConditionalDialogOpenOptions,
  type ConditionalPresetAction,
  type ConditionalRule,
  captureSnapshot,
  cellValueIsFormulaError,
  circleInvalidValidationData,
  circleInvalidValidationDataInSheet,
  clearAllowedEditRanges,
  clearComment,
  clearFilter,
  clearFormat,
  clearHyperlink,
  clearPrintArea,
  clearPrintTitles,
  clearSheetBackgroundImage,
  clearTraceArrowsByKind,
  clearValidationCircles,
  clearValidationInRangeWithEngine,
  clearVisualFormat,
  clearWatchedCells,
  colGroupRangeAt,
  colLetter,
  collapseColGroup,
  collapseRowGroup,
  conditionalMenuText,
  copy,
  copyAdvancedFilterResult,
  createColorPalette,
  createDefinedNamesFromSelection,
  createPivotTableFromRange,
  createSessionChart,
  cut,
  cycleCurrency,
  cyclePercent,
  deleteCells,
  deleteCols,
  deleteRows,
  deleteSheetView,
  dictionaries,
  expandColGroup,
  expandRowGroup,
  type FeatureFlags,
  fillRange,
  filterBySelectedCellValue,
  findMatchingCells,
  findNext,
  fluentIconPaths,
  formatAsTable,
  formatCell,
  formatNumber,
  getPageSetup,
  groupCols,
  groupRows,
  hiddenInSelection,
  hideCols,
  hideRows,
  hyperlinkAt,
  ignoreCellError,
  inferAutoFilterRange,
  inferPivotSourceFields,
  inferSortHasHeader,
  insertCells,
  insertCols,
  insertDefinedNameFormula,
  insertManualPageBreak,
  insertRows,
  isCellWritable,
  isWorkbookStructureProtected,
  listComments,
  listDefinedNames,
  type MarginPreset,
  marginPresetOf,
  moveSheet,
  mutators,
  type NumberFormatAction,
  type NumFmt,
  type PageOrientation,
  type PaperSize,
  type PasteOperation,
  type PasteSpecialOptions,
  type PasteWhat,
  PivotAggregation,
  type PivotSourceField,
  pageScaleMenuText,
  parseScriptCommand,
  pasteTSV,
  projectActiveState,
  protectedSheetPassword,
  type Range,
  type ReviewCell,
  RIBBON_KEYSHORTCUTS,
  type RibbonCommand,
  type RibbonReportItem,
  type RibbonTab,
  reapplyFilters,
  recordCommentChange,
  recordConditionalRulesChange,
  recordDefinedNamesChange,
  recordFilterChange,
  recordFormatChange,
  recordIgnoredErrorsChange,
  recordLayoutChange,
  recordPageSetupChange,
  recordProtectionChange,
  recordSheetViewsChange,
  recordTablesChange,
  recordValidationCirclesChange,
  recordWatchesChange,
  removeDuplicates,
  removeManualPageBreak,
  removeSheet,
  renameSheet,
  resetManualPageBreaks,
  reviewCellsFromState,
  ribbonDisplayText,
  rowGroupRangeAt,
  type ScriptCommand,
  type SessionChartKind,
  Spreadsheet,
  type SpreadsheetInstance,
  saveSheetView,
  selectNextFormulaError,
  setAlign,
  setAutoFilter,
  setBorderPreset,
  setCellLocked,
  setFillColor,
  setFont,
  setFontColor,
  setFreezePanes,
  setGridlinesVisible,
  setHeadingsVisible,
  setMarginPreset,
  setNumFmt,
  setPageOrientation,
  setPaperSize,
  setPrintArea,
  setPrintGridlines,
  setPrintHeadings,
  setPrintTitleCols,
  setPrintTitleRows,
  setR1C1ReferenceStyle,
  setRotation,
  setSheetBackgroundImage,
  setSheetHidden,
  setSheetZoom,
  setShowFormulas,
  setVAlign,
  setWorkbookStructureProtected,
  setWorkbookView,
  showCols,
  showRows,
  sortRange,
  summarizeSpreadsheetCompatibility,
  TABLE_STYLE_COLORS,
  type TableStyle,
  tableStyleSwatch,
  textToColumns,
  toggleBold,
  toggleItalic,
  toggleStrike,
  toggleUnderline,
  toggleWrap,
  toolbarMenuText,
  numberFormatForAction as toolbarNumberFormatForAction,
  toolbarText,
  ungroupCols,
  ungroupRows,
  unwatchCell,
  WorkbookHandle,
  warnProtected,
  watchRange,
  workbookStructurePassword,
} from '@libraz/formulon-cell';
import {
  showAdvancedFilterDialog,
  showChoiceDialog,
  showFormatAsTableDialog,
  showMessage,
  showNumberPrompt,
  showPrompt,
  showRemoveDuplicatesDialog,
  showReport,
  showSortDialog,
} from './dialogs.js';
import { applyFixture, isFixtureName } from './fixtures.js';
import { focusMenuItem, handleMenuKeydown, prepareMenu } from './menu-a11y.js';
import { autofitColWidth, autofitRowHeight } from './ribbon/autofit.js';
import { createBackstageFactories } from './ribbon/backstage.js';
import {
  BORDER_PRESETS,
  type BorderPreviewSpec,
  createBorderPreview,
  createLineSamplePreview,
  LINE_STYLES_ALL,
  SVG_NS,
} from './ribbon/border-icons.js';
import { isJapaneseFontName, shouldShowFontOption } from './ribbon/font-availability.js';
import { seedWorkbook } from './seed.js';
import { openSheetTabMenu } from './sheet-tab-menu.js';
import { setupZoomControls } from './zoom-sort.js';

const sheetEl = document.getElementById('sheet');
const autosaveSwitch = document.querySelector<HTMLButtonElement>('.app__autosave-switch');
const titleSearchInput = document.querySelector<HTMLInputElement>(
  '.app__search input[type="search"]',
);
const themeToggle = document.getElementById('theme-toggle') as HTMLButtonElement | null;
const themeLabel = document.getElementById('theme-label');
const docState = document.getElementById('doc-state');
const enginePill = document.getElementById('engine-pill');
const statusState = document.getElementById('status-state');
const statusSelection = document.getElementById('status-selection');
const statusMetric = document.getElementById('status-metric');
const statusEngine = document.getElementById('status-engine');
const statusObjects = document.getElementById('status-objects');
const ribbonRoot = document.getElementById('ribbon-root');

if (!sheetEl) throw new Error('#sheet missing');
if (statusMetric) {
  statusMetric.tabIndex = 0;
  statusMetric.setAttribute('aria-haspopup', 'menu');
}

// `paper` / `ink` / `contrast` are the core theme names; the UI labels them
// Office Light / Office Dark / High Contrast to mirror the ribbon theme menu.
type CoreTheme = 'paper' | 'ink' | 'contrast';
type UiTheme = 'light' | 'dark' | 'contrast';

const html = document.documentElement;
// URL params: `?theme=light|dark` and `?locale=en|ja` let E2E / visual specs
// pin the boot state without scripting the toolbar. They simply override the
// initial values; user toggles still work afterwards.
const bootParams = new URLSearchParams(window.location.search);
const themeParam = bootParams.get('theme');
const localeParam = bootParams.get('locale');
const initialUiTheme: UiTheme =
  themeParam === 'dark' || themeParam === 'light' || themeParam === 'contrast'
    ? themeParam
    : ((html.dataset.theme as UiTheme | undefined) ?? 'light');
let uiTheme: UiTheme = initialUiTheme;
html.dataset.theme = uiTheme;
const toCore = (t: UiTheme): CoreTheme =>
  t === 'dark' ? 'ink' : t === 'contrast' ? 'contrast' : 'paper';
const themeLabelForUi = (theme: UiTheme): string => {
  if (theme === 'contrast') return ribbonMenuText.themeContrast;
  return theme === 'dark' ? ribbonMenuText.themeInk : ribbonMenuText.themePaper;
};

const applyUiTheme = (theme: UiTheme): void => {
  uiTheme = theme;
  html.dataset.theme = uiTheme;
  if (themeLabel) themeLabel.textContent = themeLabelForUi(uiTheme);
  themeToggle?.setAttribute('aria-pressed', uiTheme === 'dark' ? 'true' : 'false');
  // Theme is a UI-only preference; don't let the resulting store update mark the workbook as edited.
  suppressDirty = true;
  inst?.setTheme(toCore(uiTheme));
  suppressDirty = false;
};

let inst: SpreadsheetInstance | null = null;

const seed = seedWorkbook;

type AutomationRun = {
  label: string;
  range: string;
  changed: number;
};

const automationRuns: AutomationRun[] = [];

const ribbonLang = localeParam === 'en' ? 'en' : 'ja';
const ribbonText = toolbarText(ribbonLang);
const ribbonMenuText = toolbarMenuText(ribbonLang);
const pageScaleText = pageScaleMenuText(ribbonLang);
const ribbonReportText = dictionaries[ribbonLang].reviewReports;
const backstageText = backstageMenuText(ribbonLang);
const ribbonDisplayOptionsText = ribbonDisplayText(ribbonLang);
const shellText =
  ribbonLang === 'ja'
    ? {
        addSheet: 'シートの追加',
        autosave: '自動保存',
        autosaveOff: '自動保存はオフです',
        autosaveOn: '自動保存はオンです',
        comments: 'コメント',
        cycleZoom: 'ズームを切り替え',
        edited: '編集中',
        enterFileName: 'ファイル名を入力してください。',
        fileName: 'ファイル名',
        home: 'ホーム',
        loading: '読み込み中...',
        more: 'その他',
        nextSheet: '次のシート',
        openFailed: '読み込み失敗',
        optional: '省略可',
        previousSheet: '前のシート',
        ready: '準備完了',
        redo: 'やり直し',
        save: '保存',
        saveAs: '名前を付けて保存',
        saveFailed: '保存失敗',
        saved: '保存済み',
        search: '検索',
        searchPlaceholder: '検索 (Cmd + Ctrl + U)',
        share: '共有',
        shareReady:
          'このブックを共有する準備ができました。リンク作成は接続された共有サービスで行われます。',
        sheets: 'シート',
        undo: '元に戻す',
        zoom: 'ズーム',
        zoomIn: '拡大',
        zoomOut: '縮小',
      }
    : {
        addSheet: 'Add sheet',
        autosave: 'AutoSave',
        autosaveOff: 'AutoSave is off',
        autosaveOn: 'AutoSave is on',
        comments: 'Comments',
        cycleZoom: 'Cycle zoom',
        edited: 'Edited',
        enterFileName: 'Enter a file name.',
        fileName: 'File name',
        home: 'Home',
        loading: 'Loading...',
        more: 'More',
        nextSheet: 'Next sheet',
        openFailed: 'Open failed',
        optional: 'optional',
        previousSheet: 'Previous sheet',
        ready: 'Ready',
        redo: 'Redo',
        save: 'Save',
        saveAs: 'Save As',
        saveFailed: 'Save failed',
        saved: 'Saved',
        search: 'Search',
        searchPlaceholder: 'Search (Cmd + Ctrl + U)',
        share: 'Share',
        shareReady:
          'This workbook is ready to share. Link creation is handled by the connected sharing service.',
        sheets: 'Sheets',
        undo: 'Undo',
        zoom: 'Zoom',
        zoomIn: 'Zoom in',
        zoomOut: 'Zoom out',
      };
type ShellTextKey = keyof typeof shellText;
let autosaveEnabled = false;
const setShellLabel = (element: Element, value: string): void => {
  element.setAttribute('aria-label', value);
  if (element instanceof HTMLElement) element.title = value;
};
const refreshAutosave = (): void => {
  if (!autosaveSwitch) return;
  autosaveSwitch.setAttribute('aria-pressed', autosaveEnabled ? 'true' : 'false');
  autosaveSwitch.classList.toggle('app__autosave-switch--on', autosaveEnabled);
  autosaveSwitch.title = autosaveEnabled ? shellText.autosaveOn : shellText.autosaveOff;
  autosaveSwitch.setAttribute(
    'aria-label',
    `${shellText.autosave}: ${autosaveEnabled ? shellText.autosaveOn : shellText.autosaveOff}`,
  );
};
const applyShellLocale = (): void => {
  html.lang = ribbonLang === 'ja' ? 'ja' : 'en';
  for (const el of document.querySelectorAll<HTMLElement>('[data-shell-i18n]')) {
    const key = el.dataset.shellI18n as ShellTextKey | undefined;
    if (key && shellText[key]) el.textContent = shellText[key];
  }
  for (const el of document.querySelectorAll<HTMLElement>('[data-shell-i18n-label]')) {
    const key = el.dataset.shellI18nLabel as ShellTextKey | undefined;
    if (key && shellText[key]) setShellLabel(el, shellText[key]);
  }
  for (const el of document.querySelectorAll<HTMLInputElement>('[data-shell-i18n-placeholder]')) {
    const key = el.dataset.shellI18nPlaceholder as ShellTextKey | undefined;
    if (key && shellText[key]) el.placeholder = shellText[key];
  }
  for (const el of document.querySelectorAll<HTMLElement>('[data-shell-i18n-title]')) {
    const key = el.dataset.shellI18nTitle as ShellTextKey | undefined;
    if (key && shellText[key]) el.title = shellText[key];
  }
  refreshAutosave();
};
let activeRibbonTab: RibbonTab = 'home';
let ribbonCollapsed = false;
let ribbonDisplayMenuOpen = false;
let backstageOpen = false;
let backstageReturnTab: RibbonTab = 'home';
let selectedBorderStyle: CellBorderStyle = 'thin';
let selectedBorderColor = '#000000';
let ribbonClipboardSnapshot: ClipboardSnapshot | null = null;
let ribbonClipboardText: string | null = null;
let formulaBarVisible = true;
const playgroundFeatureFlags = (): FeatureFlags => ({
  viewToolbar: false,
  watchWindow: true,
  workbookObjects: true,
  formulaBar: formulaBarVisible,
});

const legacyCommandIds: Record<string, string> = {
  alignC: 'btn-align-center',
  alignL: 'btn-align-left',
  alignR: 'btn-align-right',
  bold: 'btn-bold',
  borders: 'btn-borders',
  currency: 'btn-currency',
  decDown: 'btn-decimals-down',
  decUp: 'btn-decimals-up',
  fontGrow: 'btn-font-grow',
  fontShrink: 'btn-font-shrink',
  formatPainter: 'btn-format-painter',
  freeze: 'btn-freeze',
  italic: 'btn-italic',
  merge: 'btn-merge',
  middle: 'btn-middle',
  percent: 'btn-percent',
  comma: 'btn-comma',
  commentInsert: 'btn-comment',
  hyperlinkInsert: 'btn-hyperlink',
  newCommentReview: 'btn-review-comment',
  pivotTableInsert: 'btn-pivot',
  redoHome: 'btn-redo',
  strike: 'btn-strike',
  top: 'btn-top',
  underline: 'btn-underline',
  undoHome: 'btn-undo',
  wrap: 'btn-wrap',
};

const renderRibbon = (): void => {
  if (!ribbonRoot) return;
  const model = buildRibbonModel(ribbonLang);
  const shell = document.createElement('div');
  shell.className = `demo__ribbon-shell app__ribbon-shell${
    ribbonCollapsed ? ' demo__ribbon-shell--collapsed' : ''
  }`;

  const tabs = document.createElement('div');
  tabs.className = 'demo__ribbon-tabs';
  tabs.setAttribute('role', 'tablist');
  tabs.setAttribute('aria-label', ribbonText.ribbonTabs);
  tabs.dataset.ribbonCollapsed = ribbonCollapsed ? 'true' : 'false';
  for (const tab of model) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = `demo__ribbon-tab${tab.id === 'file' ? ' demo__ribbon-tab--file' : ''}${
      tab.id === activeRibbonTab ? ' demo__ribbon-tab--active' : ''
    }`;
    btn.setAttribute('role', 'tab');
    btn.setAttribute('aria-selected', tab.id === activeRibbonTab ? 'true' : 'false');
    btn.tabIndex = tab.id === activeRibbonTab ? 0 : -1;
    btn.dataset.ribbonTab = tab.id;
    btn.textContent = tab.label;
    tabs.appendChild(btn);
  }
  shell.appendChild(tabs);

  for (const tab of model) {
    const panel = document.createElement('div');
    panel.className = `demo__ribbon${tab.id === 'home' ? ' demo__ribbon--office365-home' : ''}`;
    panel.setAttribute('role', 'toolbar');
    panel.setAttribute('aria-label', `${tab.label} ${ribbonText.ribbon}`);
    panel.dataset.ribbonPanel = tab.id;
    panel.hidden = tab.id !== activeRibbonTab;

    for (const g of tab.groups) {
      const group = document.createElement('section');
      group.className = `demo__ribbon-group${g.variant ? ` demo__ribbon-group--${g.variant}` : ''}`;
      group.setAttribute('aria-label', g.title);

      const tools = document.createElement('div');
      tools.className = 'demo__ribbon-tools';
      for (const c of g.commands) {
        if (c.kind === 'break') {
          const rowBreak = document.createElement('div');
          rowBreak.className = 'demo__rb-break';
          rowBreak.dataset.ribbonCommand = c.id;
          tools.appendChild(rowBreak);
          continue;
        }
        if (c.kind === 'select') {
          tools.appendChild(createRibbonSelect(c));
          continue;
        }
        if (c.kind === 'color') {
          tools.appendChild(createRibbonColor(c));
          continue;
        }
        const b = document.createElement('button');
        b.type = 'button';
        b.className = `demo__rb${c.kind === 'large' ? ' demo__rb--large' : ''}${
          c.kind === 'wide' ? ' demo__rb--wide' : ''
        }${c.kind === 'mono' ? ' demo__rb--mono' : ''}`;
        b.title = c.title;
        b.setAttribute('aria-label', c.title);
        const keyshortcuts = RIBBON_KEYSHORTCUTS[c.id];
        if (keyshortcuts) b.setAttribute('aria-keyshortcuts', keyshortcuts);
        b.dataset.ribbonCommand = c.id;
        const legacyId = legacyCommandIds[c.id];
        if (legacyId) b.id = legacyId;
        b.disabled = !!c.disabled;
        const textOnly = !c.icon || c.kind === 'mono';
        const showLabel = textOnly || c.kind === 'wide' || c.kind === 'large';
        const icon = c.icon && c.kind !== 'mono' ? createRibbonIcon(c.icon) : null;
        if (icon) {
          b.appendChild(icon);
        }
        if (showLabel || (!icon && c.kind !== 'mono')) {
          const label = document.createElement('span');
          label.textContent = c.label;
          b.appendChild(label);
        }
        if (
          c.id === 'paste' ||
          c.id === 'autosum' ||
          c.id === 'autosumFormula' ||
          c.id === 'addIn' ||
          c.id === 'script' ||
          c.id === 'currency' ||
          c.id === 'pageTheme'
        ) {
          b.setAttribute('aria-haspopup', 'menu');
          b.setAttribute('aria-expanded', 'false');
          b.appendChild(makeSvg('0 0 12 12', RIBBON_CHEVRON_PATH, 'demo__rb-split-chevron'));
        }
        tools.appendChild(b);
        if (c.id === 'paste') tools.appendChild(createPasteMenu());
        else if (c.id === 'pivotTableInsert') tools.appendChild(createPivotTableMenu());
        else if (c.id === 'namedRangesInsert')
          tools.appendChild(createDefinedNamesMenu('menu-defined-names-insert'));
        else if (c.id === 'namedRanges')
          tools.appendChild(createDefinedNamesMenu('menu-defined-names'));
        else if (c.id === 'links') tools.appendChild(createLinksMenu('menu-links-file'));
        else if (c.id === 'linksInsert') tools.appendChild(createLinksMenu('menu-links-insert'));
        else if (c.id === 'linksData') tools.appendChild(createLinksMenu('menu-links-data'));
        else if (c.id === 'borders') tools.appendChild(createBordersMenu());
        else if (c.id === 'textOrientation') tools.appendChild(createTextOrientationMenu());
        else if (c.id === 'conditional') tools.appendChild(createConditionalMenu());
        else if (c.id === 'fillHome') tools.appendChild(createFillMenu());
        else if (c.id === 'clearFormat' && g.variant === 'editing')
          tools.appendChild(createClearMenu());
        else if (c.id === 'insertRows') tools.appendChild(createInsertCellsMenu());
        else if (c.id === 'deleteRows') tools.appendChild(createDeleteCellsMenu());
        else if (c.id === 'formatCellsHome') tools.appendChild(createFormatCellsMenu());
        else if (c.id === 'autosum') tools.appendChild(createAutoSumMenu('menu-autosum-home'));
        else if (c.id === 'freeze') tools.appendChild(createFreezeMenu());
        else if (c.id === 'autosumFormula')
          tools.appendChild(createAutoSumMenu('menu-autosum-formulas'));
        else if (c.id === 'clearArrows') tools.appendChild(createClearArrowsMenu());
        else if (c.id === 'errorChecking') tools.appendChild(createErrorCheckingMenu());
        else if (c.id === 'watch') tools.appendChild(createWatchMenu('menu-watch-formulas'));
        else if (c.id === 'watchView') tools.appendChild(createWatchMenu('menu-watch-view'));
        else if (c.id === 'deleteCommentReview') tools.appendChild(createReviewCommentsMenu());
        else if (c.id === 'protectReview')
          tools.appendChild(createProtectMenu('menu-protect-review'));
        else if (c.id === 'protect') tools.appendChild(createProtectMenu('menu-protect-view'));
        else if (c.id === 'calcOptions') tools.appendChild(createCalcOptionsMenu());
        else if (c.id === 'filter') tools.appendChild(createSortMenu('menu-sort'));
        else if (c.id === 'textToColumns') tools.appendChild(createTextToColumnsMenu());
        else if (c.id === 'dataValidation') tools.appendChild(createDataValidationMenu());
        else if (c.id === 'sortFilterHome') tools.appendChild(createSortMenu('menu-sort-home'));
        else if (c.id === 'findHome') tools.appendChild(createFindSelectMenu());
        else if (c.id === 'pictureInsert') tools.appendChild(createPictureInsertMenu());
        else if (c.id === 'shapesInsert') tools.appendChild(createShapesInsertMenu());
        else if (c.id === 'screenshotInsert') tools.appendChild(createScreenshotInsertMenu());
        else if (c.id === 'chartInsert') tools.appendChild(createChartInsertMenu());
        else if (c.id === 'formatTableHome')
          tools.appendChild(createTableStyleMenu('menu-table-style-home'));
        else if (c.id === 'formatTableInsert')
          tools.appendChild(createTableStyleMenu('menu-table-style-insert'));
        else if (c.id === 'cellStyles') tools.appendChild(createCellStylesMenu());
        else if (c.id === 'currency') tools.appendChild(createCurrencyMenu());
        else if (c.id === 'pageTheme') tools.appendChild(createPageThemeMenu());
        else if (c.id === 'printArea') tools.appendChild(createPrintAreaMenu());
        else if (c.id === 'pageBreaks') tools.appendChild(createPageBreaksMenu());
        else if (c.id === 'sheetBackground') tools.appendChild(createSheetBackgroundMenu());
        else if (c.id === 'printTitles') tools.appendChild(createPrintTitlesMenu());
        else if (c.id === 'symbolInsert') tools.appendChild(createSymbolMenu());
        else if (c.id === 'script') tools.appendChild(createScriptMenu());
        else if (c.id === 'addIn') tools.appendChild(createAddInMenu());
        else if (c.id === 'pdf') tools.appendChild(createPdfMenu());
      }

      const label = document.createElement('div');
      label.className = 'demo__ribbon-label';
      label.textContent = g.title;
      group.appendChild(tools);
      group.appendChild(label);
      panel.appendChild(group);
    }

    shell.appendChild(panel);
  }

  if (!backstageOpen) {
    const display = document.createElement('div');
    display.className = 'demo__ribbon-display';
    const toggle = document.createElement('button');
    toggle.type = 'button';
    toggle.className = 'demo__ribbon-toggle';
    toggle.dataset.ribbonToggle = 'true';
    toggle.setAttribute('aria-haspopup', 'menu');
    toggle.setAttribute('aria-expanded', ribbonDisplayMenuOpen ? 'true' : 'false');
    toggle.setAttribute('aria-label', ribbonDisplayOptionsText.label);
    toggle.title = toggle.getAttribute('aria-label') ?? '';
    display.appendChild(toggle);
    if (ribbonDisplayMenuOpen) {
      const menu = document.createElement('div');
      menu.className = 'demo__ribbon-display-menu';
      menu.setAttribute('role', 'menu');
      const options: [string, boolean, string][] = [
        [ribbonDisplayOptionsText.expanded, !ribbonCollapsed, 'expanded'],
        [ribbonDisplayOptionsText.collapsed, ribbonCollapsed, 'collapsed'],
      ];
      for (const [label, checked, option] of options) {
        const item = document.createElement('button');
        item.type = 'button';
        item.className = 'demo__ribbon-display-option';
        item.dataset.ribbonDisplayOption = option;
        item.setAttribute('role', 'menuitemradio');
        item.setAttribute('aria-checked', checked ? 'true' : 'false');
        item.textContent = label;
        menu.appendChild(item);
      }
      display.appendChild(menu);
    }
    shell.appendChild(display);
  }

  ribbonRoot.replaceChildren(shell);
  if (backstageOpen) ribbonRoot.appendChild(createBackstageView());
  projectFormatToolbar();
};

const { createBackstageView } = createBackstageFactories({
  backstageText,
  ribbonText,
  shellSavedText: shellText.saved,
  docName: () => docName,
  docState,
});

const createRibbonIcon = (name: string): SVGSVGElement | null => {
  const paths = fluentIconPaths(name);
  if (!paths) return null;
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.classList.add('demo__rb-icon');
  svg.setAttribute('viewBox', '0 0 24 24');
  svg.setAttribute('fill', 'currentColor');
  svg.setAttribute('focusable', 'false');
  svg.setAttribute('aria-hidden', 'true');
  for (const d of paths) {
    const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    path.setAttribute('d', d);
    svg.appendChild(path);
  }
  return svg;
};

const activeCellFormat = () => {
  if (!inst) return null;
  const s = inst.store.getState();
  const a = s.selection.active;
  return s.format.formats.get(`${a.sheet}:${a.row}:${a.col}`) ?? null;
};

const currentRibbonControlValue = (id: string): string => {
  const f = activeCellFormat();
  const pageSetup =
    inst &&
    (id === 'marginsPreset' ||
      id === 'orientationPreset' ||
      id === 'paperSizePreset' ||
      id === 'scaleWidth' ||
      id === 'scaleHeight' ||
      id === 'scalePercent')
      ? getPageSetup(inst.store.getState(), inst.store.getState().data.sheetIndex)
      : null;
  if (id === 'fontFamily')
    return f?.fontFamily ?? (ribbonLang === 'ja' ? '游ゴシック Regular' : 'Aptos');
  if (id === 'fontSize') return String(f?.fontSize ?? (ribbonLang === 'ja' ? 12 : 11));
  if (id === 'fontColor') return f?.color ?? '#201f1e';
  if (id === 'fillColor') return f?.fill ?? '#ffffff';
  if (id === 'numberFormat') return inst ? projectActiveState(inst).numberFormat : 'general';
  if (id === 'merge') {
    if (!inst) return 'mergeCenter';
    const state = inst.store.getState();
    const r = state.selection.range;
    const anchor = state.merges.byAnchor.get(`${r.sheet}:${r.r0}:${r.c0}`);
    return anchor &&
      anchor.r0 === r.r0 &&
      anchor.c0 === r.c0 &&
      anchor.r1 === r.r1 &&
      anchor.c1 === r.c1
      ? 'unmergeCells'
      : 'mergeCenter';
  }
  if (id === 'marginsPreset')
    return pageSetup ? (marginPresetOf(pageSetup.margins) ?? 'custom') : 'normal';
  if (id === 'orientationPreset') return pageSetup?.orientation ?? 'portrait';
  if (id === 'paperSizePreset') return pageSetup?.paperSize ?? 'A4';
  if (id === 'scaleWidth') return String(pageSetup?.fitWidth ?? 0);
  if (id === 'scaleHeight') return String(pageSetup?.fitHeight ?? 0);
  if (id === 'scalePercent') return String(Math.round((pageSetup?.scale ?? 1) * 100));
  if (id === 'sheetViewSelect') return inst?.store.getState().sheetViews.activeViewId ?? 'current';
  return '';
};

const numberFormatForAction = (action: string): NumFmt | null =>
  toolbarNumberFormatForAction(action as NumberFormatAction, ribbonLang);

function applyRibbonFormat(
  fn: (
    state: ReturnType<SpreadsheetInstance['store']['getState']>,
    store: SpreadsheetInstance['store'],
  ) => void,
): void {
  const i = inst;
  if (!i) return;
  recordFormatChange(i.history, i.store, () => {
    fn(i.store.getState(), i.store);
  });
  (sheetEl as HTMLElement).focus();
}

async function applyCustomPageScaleControl(
  id: 'scaleWidth' | 'scaleHeight' | 'scalePercent',
): Promise<void> {
  const i = inst;
  if (!i) return;
  const sheet = i.store.getState().data.sheetIndex;
  const setup = getPageSetup(i.store.getState(), sheet);
  const isScale = id === 'scalePercent';
  const initial = isScale
    ? String(Math.round((setup.scale ?? 1) * 100))
    : String(id === 'scaleWidth' ? (setup.fitWidth ?? 1) : (setup.fitHeight ?? 1));
  const value = await showPrompt({
    title: isScale
      ? ribbonText.scale
      : id === 'scaleWidth'
        ? pageScaleText.width
        : pageScaleText.height,
    label: isScale ? pageScaleText.customScalePrompt : pageScaleText.customPagesPrompt,
    initial,
    okLabel: 'OK',
    cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
    validate: (raw) => {
      const n = Number.parseInt(raw.trim(), 10);
      if (isScale)
        return Number.isInteger(n) && n >= 10 && n <= 400 ? null : pageScaleText.invalidScale;
      return Number.isInteger(n) && n >= 1 && n <= 99 ? null : pageScaleText.invalidPages;
    },
  });
  if (value === null) {
    focusSheet();
    return;
  }
  const n = Number.parseInt(value.trim(), 10);
  recordPageSetupChange(i.history, i.store, () => {
    if (isScale)
      mutators.setPageSetup(i.store, sheet, { scale: n / 100, fitWidth: 0, fitHeight: 0 });
    else
      mutators.setPageSetup(
        i.store,
        sheet,
        id === 'scaleWidth' ? { fitWidth: n } : { fitHeight: n },
      );
  });
  projectFormatToolbar();
  focusSheet();
}

function applyRibbonControl(id: string, value: string): void {
  if (id === 'fontFamily') {
    applyRibbonFormat((state, store) => setFont(state, store, { fontFamily: value }));
  } else if (id === 'fontSize') {
    applyRibbonFormat((state, store) => setFont(state, store, { fontSize: Number(value) }));
  } else if (id === 'fontColor') {
    applyRibbonFormat((state, store) => setFontColor(state, store, value));
  } else if (id === 'fillColor') {
    applyRibbonFormat((state, store) => setFillColor(state, store, value));
  } else if (id === 'numberFormat') {
    if (value === 'more') {
      inst?.openFormatDialog();
      return;
    }
    const fmt = numberFormatForAction(value);
    if (fmt) applyRibbonFormat((state, store) => setNumFmt(state, store, fmt));
  } else if (id === 'merge') {
    applyMergeControl(value);
  } else if (id === 'marginsPreset') {
    const i = inst;
    if (!i) return;
    if (value === 'custom') {
      i.openPageSetup();
      return;
    }
    const sheet = i.store.getState().data.sheetIndex;
    recordPageSetupChange(i.history, i.store, () =>
      setMarginPreset(i.store, sheet, value as MarginPreset),
    );
    projectFormatToolbar();
    (sheetEl as HTMLElement).focus();
  } else if (id === 'orientationPreset') {
    const i = inst;
    if (!i) return;
    const sheet = i.store.getState().data.sheetIndex;
    recordPageSetupChange(i.history, i.store, () =>
      setPageOrientation(i.store, sheet, value as PageOrientation),
    );
    projectFormatToolbar();
    (sheetEl as HTMLElement).focus();
  } else if (id === 'paperSizePreset') {
    const i = inst;
    if (!i) return;
    const sheet = i.store.getState().data.sheetIndex;
    recordPageSetupChange(i.history, i.store, () =>
      setPaperSize(i.store, sheet, value as PaperSize),
    );
    projectFormatToolbar();
    (sheetEl as HTMLElement).focus();
  } else if (id === 'scaleWidth' || id === 'scaleHeight') {
    const i = inst;
    if (!i) return;
    if (value === 'custom') {
      void applyCustomPageScaleControl(id);
      return;
    }
    const sheet = i.store.getState().data.sheetIndex;
    const n = Math.max(0, Math.min(99, Number.parseInt(value, 10) || 0));
    recordPageSetupChange(i.history, i.store, () => {
      mutators.setPageSetup(
        i.store,
        sheet,
        id === 'scaleWidth' ? { fitWidth: n } : { fitHeight: n },
      );
    });
    projectFormatToolbar();
    (sheetEl as HTMLElement).focus();
  } else if (id === 'scalePercent') {
    const i = inst;
    if (!i) return;
    if (value === 'custom') {
      void applyCustomPageScaleControl('scalePercent');
      return;
    }
    const sheet = i.store.getState().data.sheetIndex;
    const pct = Math.max(10, Math.min(400, Number.parseInt(value, 10) || 100));
    recordPageSetupChange(i.history, i.store, () => {
      mutators.setPageSetup(i.store, sheet, { scale: pct / 100, fitWidth: 0, fitHeight: 0 });
    });
    projectFormatToolbar();
    (sheetEl as HTMLElement).focus();
  } else if (id === 'sheetViewSelect') {
    const i = inst;
    if (!i) return;
    if (value === 'current') {
      i.store.setState((s) => ({ ...s, sheetViews: { ...s.sheetViews, activeViewId: null } }));
      projectFormatToolbar();
      focusSheet();
      return;
    }
    const result = activateSheetView(i.store, value);
    if (result.ok) {
      refreshWorkbookCells();
      projectFormatToolbar();
      focusSheet();
    }
  }
}

function applyMergeControl(value: string): void {
  const i = inst;
  if (!i) return;
  const range = i.store.getState().selection.range;
  if (value === 'unmergeCells') {
    applyUnmerge(i.store, i.workbook, i.history, range);
  } else if (value === 'mergeAcross') {
    i.history.begin();
    try {
      for (let row = range.r0; row <= range.r1; row += 1) {
        applyMerge(i.store, i.workbook, i.history, {
          sheet: range.sheet,
          r0: row,
          c0: range.c0,
          r1: row,
          c1: range.c1,
        });
      }
    } finally {
      i.history.end();
    }
  } else {
    const merged = applyMerge(i.store, i.workbook, i.history, range);
    if (merged && value === 'mergeCenter') {
      applyRibbonFormat((state, store) => setAlign(state, store, 'center'));
    }
  }
  refreshWorkbookCells();
  projectFormatToolbar();
  (sheetEl as HTMLElement).focus();
}

const makeSvg = (viewBox: string, pathData: string, className: string): SVGSVGElement => {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.classList.add(className);
  svg.setAttribute('viewBox', viewBox);
  svg.setAttribute('fill', 'currentColor');
  svg.setAttribute('focusable', 'false');
  svg.setAttribute('aria-hidden', 'true');
  const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
  path.setAttribute('d', pathData);
  svg.appendChild(path);
  return svg;
};

const ribbonMarginDetail = (value: string): string | null => {
  const ja = ribbonLang === 'ja';
  const fmt = (top: string, bottom: string, left: string, right: string): string =>
    ja
      ? `上: ${top}", 下: ${bottom}", 左: ${left}", 右: ${right}"`
      : `Top: ${top}", Bottom: ${bottom}", Left: ${left}", Right: ${right}"`;
  switch (value) {
    case 'normal':
      return fmt('0.75', '0.75', '0.7', '0.7');
    case 'wide':
      return fmt('1', '1', '1', '1');
    case 'narrow':
      return fmt('0.75', '0.75', '0.25', '0.25');
    case 'custom':
      return ja ? 'ユーザー設定の余白...' : 'Custom margins...';
    default:
      return null;
  }
};

const createMarginPresetIcon = (value: string): HTMLSpanElement => {
  const icon = document.createElement('span');
  icon.className = `demo__rb-dd__margin-icon demo__rb-dd__margin-icon--${value}`;
  icon.setAttribute('aria-hidden', 'true');
  icon.append(document.createElement('span'), document.createElement('span'));
  return icon;
};

const numberFormatHasSubtitle = (value: string): boolean => value === 'general';

const numberFormatSubtitle = (value: string): string =>
  value === 'general' ? ribbonText.numberFormatNoSpecific : '';

const themeFontValues = new Set(['Aptos', 'Aptos Display', 'Aptos Narrow']);
const recentFontValues = new Set(['Yu Gothic UI']);
const commonFontValues = new Set(['Calibri', 'Arial', 'Segoe UI', 'Times New Roman', 'Consolas']);
const fontSubmenuFamilies = new Set(['Yu Gothic UI', 'BIZ UDGothic', 'Meiryo UI']);
// Availability probing + Japanese-name detection live in
// ribbon/font-availability.ts. The dropdown threads the current locale
// through via the imported helper.

const ribbonFontSection = (
  value: string,
  options: readonly { value: string; label: string }[],
): string | null => {
  const firstTheme = options.find((option) => themeFontValues.has(option.value))?.value;
  if (value === firstTheme) return ribbonLang === 'ja' ? 'テーマのフォント' : 'Theme Fonts';
  const firstRecent = options.find((option) => recentFontValues.has(option.value))?.value;
  if (value === firstRecent)
    return ribbonLang === 'ja' ? '最近使ったフォント' : 'Recently Used Fonts';
  const firstAll = options.find(
    (option) => !themeFontValues.has(option.value) && !recentFontValues.has(option.value),
  )?.value;
  if (value === firstAll) return ribbonLang === 'ja' ? 'すべてのフォント' : 'All Fonts';
  return null;
};

const ribbonFontRole = (value: string): string | null => {
  switch (value) {
    case 'Aptos Display':
    case '游ゴシック Light':
      return ribbonLang === 'ja' ? '(見出し)' : '(Heading)';
    case 'Aptos Narrow':
    case '游ゴシック Regular':
      return ribbonLang === 'ja' ? '(本文)' : '(Body)';
    default:
      return null;
  }
};

const ribbonOptionsForCommand = (
  command: RibbonCommand,
  current: string,
): readonly { value: string; label: string }[] => {
  const options = command.options ?? [];
  if (command.id === 'sheetViewSelect') {
    const currentLabel = ribbonLang === 'ja' ? '現在の表示' : 'Current view';
    const views =
      inst?.store
        .getState()
        .sheetViews.views.filter((view) => view.sheet === inst?.store.getState().data.sheetIndex)
        .map((view) => ({ value: view.id, label: view.name })) ?? [];
    return [{ value: 'current', label: currentLabel }, ...views];
  }
  if (command.id !== 'fontFamily') return options;
  return options.filter((option) => shouldShowFontOption(option.value, current, ribbonLang));
};

const createRibbonSelect = (command: RibbonCommand): HTMLDivElement => {
  const wrap = document.createElement('div');
  wrap.className = `demo__rb-dd${command.className ? ` ${command.className}` : ''}`;
  wrap.dataset.ribbonCommand = command.id;
  wrap.dataset.ribbonSelect = command.id;
  wrap.dataset.ribbonOptions = JSON.stringify(command.options ?? []);

  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'demo__rb-dd__btn';
  button.title = command.title;
  button.setAttribute('aria-label', command.title);
  button.setAttribute('aria-haspopup', 'listbox');
  button.setAttribute('aria-expanded', 'false');

  const value = document.createElement('span');
  value.className = 'demo__rb-dd__value';
  button.append(
    value,
    makeSvg(
      '0 0 12 12',
      'M2.15 4.65a.5.5 0 0 1 .7 0L6 7.79l3.15-3.14a.5.5 0 1 1 .7.7l-3.5 3.5a.5.5 0 0 1-.7 0l-3.5-3.5a.5.5 0 0 1 0-.7Z',
      'demo__rb-dd__chev',
    ),
  );
  wrap.appendChild(button);

  let detachDocDown: (() => void) | null = null;
  const close = (): void => {
    wrap.classList.remove('demo__rb-dd--open');
    button.setAttribute('aria-expanded', 'false');
    wrap.querySelector('.demo__rb-dd__list')?.remove();
    detachDocDown?.();
    detachDocDown = null;
  };
  const focusListOption = (list: HTMLElement, index: number): void => {
    const options = Array.from(list.querySelectorAll<HTMLButtonElement>('[role="option"]'));
    if (options.length === 0) return;
    const next = ((index % options.length) + options.length) % options.length;
    for (const [idx, option] of options.entries()) option.tabIndex = idx === next ? 0 : -1;
    options[next]?.focus({ preventScroll: true });
    options[next]?.scrollIntoView({ block: 'nearest' });
  };
  const pickOption = (option: HTMLButtonElement): void => {
    const nextValue = option.dataset.value;
    if (nextValue == null) return;
    applyRibbonControl(command.id, nextValue);
    const label = option.querySelector<HTMLElement>('.demo__rb-dd__label')?.textContent;
    if (label) value.textContent = label;
    close();
    button.focus({ preventScroll: true });
  };
  const open = (): void => {
    closeOpenRibbonDropdowns(wrap);
    wrap.classList.add('demo__rb-dd--open');
    button.setAttribute('aria-expanded', 'true');
    const list = document.createElement('div');
    list.className = 'demo__rb-dd__list';
    list.setAttribute('role', 'listbox');
    list.setAttribute('aria-label', command.title);
    list.tabIndex = -1;
    const anchorRect = button.getBoundingClientRect();
    list.style.left = `${Math.round(anchorRect.left)}px`;
    list.style.top = `${Math.round(anchorRect.bottom + 3)}px`;
    list.style.minWidth = `${Math.round(anchorRect.width)}px`;
    const current = currentRibbonControlValue(command.id);
    const options = ribbonOptionsForCommand(command, current);
    for (const option of options) {
      const section = command.id === 'fontFamily' ? ribbonFontSection(option.value, options) : null;
      if (section) {
        const heading = document.createElement('div');
        heading.className = 'demo__rb-dd__section';
        heading.setAttribute('role', 'presentation');
        heading.textContent = section;
        list.appendChild(heading);
      }
      const selected = option.value === current;
      const item = document.createElement('button');
      item.type = 'button';
      item.className = `demo__rb-dd__opt${selected ? ' demo__rb-dd__opt--selected' : ''}`;
      item.setAttribute('role', 'option');
      item.setAttribute('aria-selected', selected ? 'true' : 'false');
      item.tabIndex = -1;
      item.dataset.value = option.value;
      item.dataset.fcValue = option.value;
      const check = document.createElement('span');
      check.className = 'demo__rb-dd__check';
      check.setAttribute('aria-hidden', 'true');
      if (selected) {
        check.appendChild(
          makeSvg(
            '0 0 16 16',
            'M13.36 3.74c.29.28.29.77 0 1.05l-7.01 7.01a.75.75 0 0 1-1.06 0L2.64 9.15a.75.75 0 1 1 1.06-1.06l2.12 2.12 6.48-6.47a.75.75 0 0 1 1.06 0Z',
            'demo__rb-dd__check-icon',
          ),
        );
      }
      const label = document.createElement('span');
      label.className = 'demo__rb-dd__label';
      label.textContent = option.label;
      if (command.id === 'marginsPreset') {
        const text = document.createElement('span');
        text.className = 'demo__rb-dd__margin-text';
        const detail = document.createElement('span');
        detail.className = 'demo__rb-dd__detail';
        detail.textContent = ribbonMarginDetail(option.value) ?? '';
        text.append(label, detail);
        item.append(check, createMarginPresetIcon(option.value), text);
      } else if (command.id === 'fontFamily') {
        const preview = document.createElement('span');
        preview.className = 'demo__rb-dd__font-preview';
        preview.style.fontFamily = `"${option.value}", sans-serif`;
        const role = ribbonFontRole(option.value);
        if (role) {
          const detail = document.createElement('span');
          detail.className = 'demo__rb-dd__font-role';
          detail.textContent = role;
          preview.append(label, detail);
        } else {
          preview.append(label);
        }
        item.append(check, preview);
        if (fontSubmenuFamilies.has(option.value)) {
          const arrow = document.createElement('span');
          arrow.className = 'demo__rb-dd__submenu';
          arrow.setAttribute('aria-hidden', 'true');
          arrow.textContent = '›';
          item.appendChild(arrow);
        }
      } else if (command.id === 'numberFormat' && numberFormatHasSubtitle(option.value)) {
        const text = document.createElement('span');
        text.className = 'demo__rb-dd__numfmt-text';
        const detail = document.createElement('span');
        detail.className = 'demo__rb-dd__numfmt-subtitle';
        detail.textContent = numberFormatSubtitle(option.value);
        text.append(label, detail);
        item.append(check, text);
      } else {
        item.append(check, label);
      }
      item.addEventListener('click', () => pickOption(item));
      list.appendChild(item);
    }
    list.addEventListener('keydown', (event) => {
      const options = Array.from(list.querySelectorAll<HTMLButtonElement>('[role="option"]'));
      const currentIndex = Math.max(
        0,
        options.indexOf(document.activeElement as HTMLButtonElement),
      );
      if (event.key === 'ArrowDown') {
        event.preventDefault();
        focusListOption(list, currentIndex + 1);
      } else if (event.key === 'ArrowUp') {
        event.preventDefault();
        focusListOption(list, currentIndex - 1);
      } else if (event.key === 'Home') {
        event.preventDefault();
        focusListOption(list, 0);
      } else if (event.key === 'End') {
        event.preventDefault();
        focusListOption(list, options.length - 1);
      } else if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        const option = document.activeElement?.closest<HTMLButtonElement>('[role="option"]');
        if (option && list.contains(option)) pickOption(option);
      } else if (event.key === 'Escape') {
        event.preventDefault();
        close();
        button.focus({ preventScroll: true });
      }
    });
    wrap.appendChild(list);
    const selectedIndex = Math.max(
      0,
      Array.from(list.querySelectorAll<HTMLButtonElement>('[role="option"]')).findIndex(
        (option) => option.getAttribute('aria-selected') === 'true',
      ),
    );
    focusListOption(list, selectedIndex);
    setTimeout(() => {
      const onDocDown = (ev: MouseEvent): void => {
        if (ev.target instanceof Node && wrap.contains(ev.target)) return;
        close();
      };
      document.addEventListener('mousedown', onDocDown, true);
      detachDocDown = () => document.removeEventListener('mousedown', onDocDown, true);
    }, 0);
  };

  button.addEventListener('click', () => {
    if (wrap.classList.contains('demo__rb-dd--open')) close();
    else open();
  });
  button.addEventListener('keydown', (event) => {
    if (event.key === 'ArrowDown' || event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      open();
    } else if (event.key === 'Escape') {
      event.preventDefault();
      close();
    }
  });

  updateRibbonSelectDisplay(wrap, command);
  return wrap;
};

const RIBBON_CHEVRON_PATH =
  'M2.15 4.65a.5.5 0 0 1 .7 0L6 7.79l3.15-3.14a.5.5 0 1 1 .7.7l-3.5 3.5a.5.5 0 0 1-.7 0l-3.5-3.5a.5.5 0 0 1 0-.7Z';

// Font / fill color button — an icon with a colored underline bar that opens
// the shared Office-style color palette flyout (theme + standard colors,
// "More Colors…" hands off to the native picker).
const createRibbonColor = (command: RibbonCommand): HTMLDivElement => {
  const wrap = document.createElement('div');
  wrap.className = 'demo__rb-color';
  wrap.dataset.ribbonCommand = command.id;

  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'demo__rb-color__btn';
  button.title = command.title;
  button.setAttribute('aria-label', command.title);
  button.setAttribute('aria-haspopup', 'true');
  button.setAttribute('aria-expanded', 'false');
  if (command.icon) {
    const icon = createRibbonIcon(command.icon);
    if (icon) {
      icon.classList.add('demo__rb-color__icon');
      button.appendChild(icon);
    }
  }
  const swatch = document.createElement('span');
  swatch.className = 'demo__rb-color__swatch';
  swatch.style.background = currentRibbonControlValue(command.id);
  button.append(swatch, makeSvg('0 0 12 12', RIBBON_CHEVRON_PATH, 'demo__rb-color__chev'));
  wrap.appendChild(button);

  // Hidden native picker, reached through the palette's "More Colors…" row.
  const native = document.createElement('input');
  native.type = 'color';
  native.className = 'demo__color-flyout__native';
  native.tabIndex = -1;
  native.setAttribute('aria-hidden', 'true');
  wrap.appendChild(native);

  let detachDocDown: (() => void) | null = null;
  const close = (): void => {
    wrap.classList.remove('demo__rb-color--open');
    button.setAttribute('aria-expanded', 'false');
    wrap.querySelector('.demo__color-flyout')?.remove();
    detachDocDown?.();
    detachDocDown = null;
  };
  const apply = (color: string): void => {
    applyRibbonControl(command.id, color);
    swatch.style.background = color;
  };
  native.addEventListener('input', () => apply(native.value));

  const open = (): void => {
    closeOpenRibbonDropdowns(wrap);
    wrap.classList.add('demo__rb-color--open');
    button.setAttribute('aria-expanded', 'true');
    const flyout = document.createElement('div');
    flyout.className = 'demo__color-flyout';
    const palette = createColorPalette({
      themeLabel: ribbonText.themeColors,
      standardLabel: ribbonText.standardColors,
      moreColorsLabel: ribbonText.moreColors,
      ariaLabel: command.title,
      value: currentRibbonControlValue(command.id),
      automatic:
        command.id === 'fontColor' ? { label: ribbonText.automatic, color: '#000000' } : null,
      onPick: (color) => {
        apply(color);
        close();
        button.focus({ preventScroll: true });
      },
      onMoreColors: () => {
        close();
        native.value = currentRibbonControlValue(command.id);
        native.click();
      },
    });
    flyout.appendChild(palette.el);
    const anchorRect = button.getBoundingClientRect();
    flyout.style.left = `${Math.round(anchorRect.left)}px`;
    flyout.style.top = `${Math.round(anchorRect.bottom + 3)}px`;
    flyout.addEventListener('keydown', (event) => {
      if (event.key !== 'Escape') return;
      event.preventDefault();
      close();
      button.focus({ preventScroll: true });
    });
    wrap.appendChild(flyout);
    palette.focus();
    setTimeout(() => {
      const onDocDown = (ev: MouseEvent): void => {
        if (ev.target instanceof Node && wrap.contains(ev.target)) return;
        close();
      };
      document.addEventListener('mousedown', onDocDown, true);
      detachDocDown = () => document.removeEventListener('mousedown', onDocDown, true);
    }, 0);
  };

  button.addEventListener('click', () => {
    if (wrap.classList.contains('demo__rb-color--open')) close();
    else open();
  });
  button.addEventListener('keydown', (event) => {
    if (event.key === 'ArrowDown' || event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      open();
    } else if (event.key === 'Escape') {
      event.preventDefault();
      close();
    }
  });
  return wrap;
};

const closeOpenRibbonDropdowns = (except?: HTMLElement): void => {
  for (const open of document.querySelectorAll<HTMLElement>('.demo__rb-dd--open')) {
    if (except && open === except) continue;
    open.classList.remove('demo__rb-dd--open');
    open
      .querySelector<HTMLButtonElement>('.demo__rb-dd__btn')
      ?.setAttribute('aria-expanded', 'false');
    open.querySelector('.demo__rb-dd__list')?.remove();
  }
  for (const open of document.querySelectorAll<HTMLElement>('.demo__rb-color--open')) {
    if (except && open === except) continue;
    open.classList.remove('demo__rb-color--open');
    open
      .querySelector<HTMLButtonElement>('.demo__rb-color__btn')
      ?.setAttribute('aria-expanded', 'false');
    open.querySelector('.demo__color-flyout')?.remove();
  }
};

const updateRibbonSelectDisplay = (wrap: HTMLElement, command: RibbonCommand): void => {
  const current = currentRibbonControlValue(command.id);
  const option = ribbonOptionsForCommand(command, current).find(
    (candidate) => candidate.value === current,
  );
  const value = wrap.querySelector<HTMLElement>('.demo__rb-dd__value');
  if (value) {
    const base = option?.label ?? current;
    const role = command.id === 'fontFamily' ? ribbonFontRole(current) : null;
    value.textContent = role ? `${base} ${role}` : base;
  }
};

const ribbonSelectLabel = (wrap: HTMLElement, current: string): string => {
  if (wrap.dataset.ribbonSelect === 'sheetViewSelect') {
    if (current === 'current') return ribbonLang === 'ja' ? '現在の表示' : 'Current view';
    const state = inst?.store.getState();
    return state?.sheetViews.views.find((view) => view.id === current)?.name ?? current;
  }
  try {
    const options = JSON.parse(wrap.dataset.ribbonOptions ?? '[]') as {
      value: string;
      label: string;
    }[];
    const label = options.find((option) => option.value === current)?.label;
    if (label) return label;
    if (wrap.dataset.ribbonSelect === 'scalePercent') return `${current}%`;
    if (wrap.dataset.ribbonSelect === 'scaleWidth' || wrap.dataset.ribbonSelect === 'scaleHeight') {
      if (current === '0') return pageScaleText.automatic;
      return `${current} ${current === '1' ? pageScaleText.page : pageScaleText.pages}`;
    }
    return current;
  } catch {
    return current;
  }
};

const createMenu = (id: string): HTMLDivElement => {
  const menu = document.createElement('div');
  menu.className = 'app__menu';
  menu.id = id;
  menu.hidden = true;
  prepareMenu(menu);
  return menu;
};

const menuButton = (label: string, attr: string, value: string): HTMLButtonElement => {
  const button = document.createElement('button');
  button.className = 'app__menu-item';
  button.type = 'button';
  button.setAttribute('role', 'menuitem');
  button.dataset[attr] = value;
  button.textContent = label;
  return button;
};

const createPrintAreaMenu = (): HTMLDivElement => {
  const t = ribbonMenuText;
  const menu = createMenu('menu-print-area');
  menu.append(
    menuButton(t.printAreaSet, 'printAreaAction', 'set'),
    menuButton(t.printAreaClear, 'printAreaAction', 'clear'),
  );
  return menu;
};

const createPageBreaksMenu = (): HTMLDivElement => {
  const t = ribbonMenuText;
  const menu = createMenu('menu-page-breaks');
  menu.append(
    menuButton(t.pageBreakInsertRow, 'pageBreakAction', 'insert-row'),
    menuButton(t.pageBreakInsertCol, 'pageBreakAction', 'insert-col'),
    menuSeparator(),
    menuButton(t.pageBreakRemoveRow, 'pageBreakAction', 'remove-row'),
    menuButton(t.pageBreakRemoveCol, 'pageBreakAction', 'remove-col'),
    menuSeparator(),
    menuButton(t.pageBreakResetAll, 'pageBreakAction', 'reset-all'),
  );
  return menu;
};

const createSheetBackgroundMenu = (): HTMLDivElement => {
  const t = ribbonMenuText;
  const menu = createMenu('menu-sheet-background');
  menu.append(
    menuButton(t.sheetBackgroundSet, 'sheetBackgroundAction', 'set'),
    menuButton(t.sheetBackgroundClear, 'sheetBackgroundAction', 'clear'),
  );
  return menu;
};

const createPrintTitlesMenu = (): HTMLDivElement => {
  const t = ribbonMenuText;
  const menu = createMenu('menu-print-titles');
  menu.append(
    menuButton(t.printTitleRowsSet, 'printTitlesAction', 'rows'),
    menuButton(t.printTitleColsSet, 'printTitlesAction', 'cols'),
    menuSeparator(),
    menuButton(t.printTitlesClear, 'printTitlesAction', 'clear'),
  );
  return menu;
};

const createPageThemeMenu = (): HTMLDivElement => {
  const t = ribbonMenuText;
  const menu = createMenu('menu-page-theme');
  menu.append(
    menuButton(t.themePaper, 'pageThemeAction', 'light'),
    menuButton(t.themeInk, 'pageThemeAction', 'dark'),
    menuButton(t.themeContrast, 'pageThemeAction', 'contrast'),
  );
  return menu;
};

const SYMBOL_GROUPS = [
  {
    label: ribbonMenuText.symbolMath,
    symbols: ['±', '×', '÷', '≤', '≥', '≠', '≈', '∞', '√', '∑', '∫', 'π'],
  },
  {
    label: ribbonMenuText.symbolGreek,
    symbols: ['Α', 'Β', 'Γ', 'Δ', 'Θ', 'Λ', 'Ξ', 'Π', 'Σ', 'Φ', 'Ψ', 'Ω'],
  },
  { label: ribbonMenuText.symbolCurrency, symbols: ['$', '€', '¥', '£', '¢', '₩', '₹', '₽'] },
  { label: ribbonMenuText.symbolLegal, symbols: ['©', '®', '™', '§', '¶', '†', '‡', '•'] },
] as const;

const symbolMenuHeading = (label: string): HTMLDivElement => {
  const heading = document.createElement('div');
  heading.className = 'app__menu-heading';
  heading.setAttribute('role', 'presentation');
  heading.textContent = label;
  return heading;
};

const createSymbolMenu = (): HTMLDivElement => {
  const menu = createMenu('menu-symbol');
  menu.classList.add('app__menu--symbols');
  for (const group of SYMBOL_GROUPS) {
    menu.append(symbolMenuHeading(group.label));
    for (const symbol of group.symbols) {
      const button = menuButton(symbol, 'symbol', symbol);
      button.title = symbol;
      menu.append(button);
    }
  }
  menu.append(menuSeparator(), menuButton(ribbonMenuText.symbolMore, 'symbolAction', 'more'));
  return menu;
};

const createPivotTableMenu = (): HTMLDivElement => {
  const t = ribbonMenuText;
  const menu = createMenu('menu-pivot-table');
  menu.append(
    menuButton(t.pivotTableFromRange, 'pivotTableAction', 'dialog'),
    menuButton(t.recommendedPivotTables, 'pivotTableAction', 'recommended'),
    menuSeparator(),
    menuButton(t.pivotTableNewSheet, 'pivotTableAction', 'new-sheet'),
    menuButton(t.pivotTableExistingSheet, 'pivotTableAction', 'existing-sheet'),
  );
  return menu;
};

const createDefinedNamesMenu = (id: string): HTMLDivElement => {
  const t = ribbonMenuText;
  const menu = createMenu(id);
  menu.append(
    menuButton(t.defineName, 'definedNameAction', 'define'),
    menuButton(t.nameManager, 'definedNameAction', 'manager'),
    menuSeparator(),
    menuButton(t.createFromSelectionTop, 'definedNameAction', 'create-top-row'),
    menuButton(t.createFromSelectionBottom, 'definedNameAction', 'create-bottom-row'),
    menuButton(t.createFromSelectionLeft, 'definedNameAction', 'create-left-column'),
    menuButton(t.createFromSelectionRight, 'definedNameAction', 'create-right-column'),
    menuSeparator(),
    menuButton(t.useInFormula, 'definedNameAction', 'use-formula'),
  );
  return menu;
};

const updateDefinedNamesMenu = (menu: HTMLElement): void => {
  const t = ribbonMenuText;
  menu.querySelectorAll('[data-defined-name-dynamic]').forEach((node) => node.remove());
  const names = inst ? listDefinedNames(inst.workbook) : [];
  const sep = menuSeparator();
  sep.dataset.definedNameDynamic = 'true';
  menu.appendChild(sep);
  if (names.length === 0) {
    const empty = document.createElement('button');
    empty.className = 'app__menu-item';
    empty.type = 'button';
    empty.disabled = true;
    empty.setAttribute('role', 'menuitem');
    empty.dataset.definedNameDynamic = 'true';
    empty.textContent = t.noDefinedNames;
    menu.appendChild(empty);
    return;
  }
  for (const entry of names) {
    const item = menuButton(entry.name, 'definedNameAction', `insert:${entry.name}`);
    item.dataset.definedNameDynamic = 'true';
    item.title = entry.formula;
    menu.appendChild(item);
  }
};

const createLinksMenu = (id: string): HTMLDivElement => {
  const t = ribbonMenuText;
  const menu = createMenu(id);
  menu.append(
    menuButton(t.linkInsertOrEdit, 'linkAction', 'hyperlink'),
    menuButton(t.linkOpen, 'linkAction', 'open'),
    menuButton(t.linkClear, 'linkAction', 'clear'),
    menuSeparator(),
    menuButton(t.linkExternalLinks, 'linkAction', 'external'),
  );
  return menu;
};

const createDataValidationMenu = (): HTMLDivElement => {
  const t = ribbonMenuText;
  const menu = createMenu('menu-data-validation');
  menu.append(
    menuButton(t.validationSettings, 'validationAction', 'settings'),
    menuButton(t.validationCircleInvalid, 'validationAction', 'circle-invalid'),
    menuButton(t.validationClearCircles, 'validationAction', 'clear-circles'),
    menuSeparator(),
    menuButton(t.validationClearRules, 'validationAction', 'clear-rules'),
  );
  return menu;
};

// ── Borders dropdown (Excel-365 parity) ─────────────────────────────────
// Renders a small SVG cell-outline icon for each border preset. Sides are
// drawn solid in the foreground color (thin/thick/double); the unset sides
// show as a faint dashed cell-outline base so the user can still see the
// cell shape.
// ── Borders dropdown (Excel-365 parity) ─────────────────────────────────
// SVG preview factories + presets live in ribbon/border-icons.ts so the
// menu factory below stays focused on wiring labels and click handlers.

const presetMenuItem = (presetKey: string, label: string): HTMLButtonElement => {
  const btn = document.createElement('button');
  btn.className = 'app__menu-item app__menu-item--preset';
  btn.type = 'button';
  btn.setAttribute('role', 'menuitem');
  btn.dataset.borderPreset = presetKey;
  const spec = BORDER_PRESETS[presetKey];
  if (spec) btn.appendChild(createBorderPreview(spec));
  else {
    const spacer = document.createElement('span');
    spacer.className = 'app__menu-item__icon-spacer';
    btn.appendChild(spacer);
  }
  const text = document.createElement('span');
  text.className = 'app__menu-item__text';
  text.textContent = label;
  btn.appendChild(text);
  return btn;
};

const menuSeparator = (): HTMLDivElement => {
  const sep = document.createElement('div');
  sep.className = 'app__menu-sep';
  sep.setAttribute('role', 'separator');
  return sep;
};

const menuSectionHeader = (label: string): HTMLDivElement => {
  const el = document.createElement('div');
  el.className = 'app__menu-heading';
  el.setAttribute('role', 'presentation');
  el.textContent = label;
  return el;
};

const createPasteMenu = (): HTMLDivElement => {
  const ja = ribbonLang === 'ja';
  const menu = createMenu('menu-paste');
  menu.append(
    menuButton(ja ? '貼り付け' : 'Paste', 'pasteAction', 'all'),
    menuButton(ja ? '数式' : 'Formulas', 'pasteAction', 'formulas'),
    menuButton(
      ja ? '数式と数値の書式' : 'Formulas & Number Formatting',
      'pasteAction',
      'formulas-and-numfmt',
    ),
    menuButton(ja ? '値' : 'Values', 'pasteAction', 'values'),
    menuButton(
      ja ? '値と数値の書式' : 'Values & Number Formatting',
      'pasteAction',
      'values-and-numfmt',
    ),
    menuButton(ja ? '書式設定' : 'Formatting', 'pasteAction', 'formats'),
    menuSeparator(),
    menuButton(ja ? '行/列の入れ替え' : 'Transpose', 'pasteAction', 'transpose'),
    menuButton(ja ? '形式を選択して貼り付け...' : 'Paste Special...', 'pasteAction', 'dialog'),
  );
  return menu;
};

type CfSubmenuKey = 'highlight' | 'topBottom' | 'dataBar' | 'colorScale' | 'iconSet' | 'clear';

const cfMenuText = () =>
  (() => {
    const t = conditionalMenuText(ribbonLang);
    return {
      highlight: t.highlight,
      topBottom: t.topBottom,
      dataBar: t.dataBars,
      colorScale: t.colorScales,
      iconSet: t.iconSets,
      newRule: t.newRule,
      clear: t.clear,
      manage: t.manage,
      greater: t.greater,
      less: t.less,
      between: t.between,
      equal: t.equal,
      text: t.textContains,
      date: t.dateOccurring,
      duplicate: t.duplicates,
      unique: t.unique,
      top10: t.top10,
      bottom10: t.bottom10,
      top10Percent: t.top10Percent,
      bottom10Percent: t.bottom10Percent,
      aboveAvg: t.aboveAvg,
      belowAvg: t.belowAvg,
      textPrompt: t.textPrompt,
      datePrompt: t.datePrompt,
      otherRules: t.otherRules,
      gradient: t.gradientFill,
      solid: t.solidFill,
      direction: t.direction,
      shapes: t.shapes,
      indicators: t.indicators,
      ratings: t.ratings,
      flags: t.flags,
      bars: t.bars,
      clearSelection: t.clearSelection,
      clearSheet: t.clearSheet,
      dataBarGradientBlue: t.dataBarGradientBlue,
      dataBarGradientGreen: t.dataBarGradientGreen,
      dataBarGradientRed: t.dataBarGradientRed,
      dataBarGradientOrange: t.dataBarGradientOrange,
      dataBarGradientPurple: t.dataBarGradientPurple,
      dataBarGradientTeal: t.dataBarGradientTeal,
      dataBarSolidBlue: t.dataBarSolidBlue,
      dataBarSolidGreen: t.dataBarSolidGreen,
      dataBarSolidRed: t.dataBarSolidRed,
      dataBarSolidOrange: t.dataBarSolidOrange,
      dataBarSolidPurple: t.dataBarSolidPurple,
      dataBarSolidGray: t.dataBarSolidGray,
      colorScaleGreenYellowRed: t.colorScaleGreenYellowRed,
      colorScaleRedYellowGreen: t.colorScaleRedYellowGreen,
      colorScaleGreenWhite: t.colorScaleGreenWhite,
      colorScaleRedWhite: t.colorScaleRedWhite,
      colorScaleBlueWhiteRed: t.colorScaleBlueWhiteRed,
      colorScaleRedWhiteBlue: t.colorScaleRedWhiteBlue,
      colorScaleGreenWhiteGreen: t.colorScaleGreenWhiteGreen,
      colorScaleYellowWhiteGreen: t.colorScaleYellowWhiteGreen,
      colorScaleRedWhiteRed: t.colorScaleRedWhiteRed,
      colorScaleBlueWhiteBlue: t.colorScaleBlueWhiteBlue,
      colorScaleYellowRedGreen: t.colorScaleYellowRedGreen,
      colorScaleGreenYellowGreen: t.colorScaleGreenYellowGreen,
      iconArrows3: t.iconArrows3,
      iconArrows5: t.iconArrows5,
      iconTriangles3: t.iconTriangles3,
      iconTraffic3: t.iconTraffic3,
      iconTrafficRim3: t.iconTrafficRim3,
      iconSymbols3: t.iconSymbols3,
      iconFlags3: t.iconFlags3,
      iconStars3: t.iconStars3,
      iconQuarters5: t.iconQuarters5,
      iconRatings5: t.iconRatings5,
      iconBars5: t.iconBars5,
      iconBoxes5: t.iconBoxes5,
    };
  })();

const cfIcon = (kind: 'rule' | 'top' | 'bar' | 'scale' | 'icon' | 'new' | 'clear' | 'manage') => {
  const span = document.createElement('span');
  span.className = `app__cf-icon app__cf-icon--${kind}`;
  span.setAttribute('aria-hidden', 'true');
  return span;
};

const cfMenuItem = (
  label: string,
  action: string,
  icon: Parameters<typeof cfIcon>[0] = 'rule',
): HTMLButtonElement => {
  const btn = document.createElement('button');
  btn.className = 'app__menu-item app__menu-item--preset';
  btn.type = 'button';
  btn.setAttribute('role', 'menuitem');
  btn.dataset.cfAction = action;
  btn.appendChild(cfIcon(icon));
  const text = document.createElement('span');
  text.className = 'app__menu-item__text';
  text.textContent = label;
  btn.appendChild(text);
  return btn;
};

const cfSubmenuTrigger = (
  key: CfSubmenuKey,
  label: string,
  icon: Parameters<typeof cfIcon>[0],
): HTMLButtonElement => {
  const btn = cfMenuItem(label, `submenu-${key}`, icon);
  btn.classList.add('app__menu-item--submenu');
  btn.dataset.cfSubmenu = key;
  const caret = document.createElement('span');
  caret.className = 'app__menu-item__caret';
  caret.textContent = '▶';
  btn.appendChild(caret);
  return btn;
};

const cfSwatchButton = (action: string, colors: string[], title: string): HTMLButtonElement => {
  const btn = document.createElement('button');
  btn.className = 'app__cf-choice';
  btn.type = 'button';
  btn.title = title;
  btn.setAttribute('aria-label', title);
  btn.dataset.cfAction = action;
  const grid = document.createElement('span');
  grid.className = 'app__cf-choice-grid';
  for (const color of colors) {
    const cell = document.createElement('span');
    cell.style.background = color;
    grid.appendChild(cell);
  }
  btn.appendChild(grid);
  return btn;
};

const cfIconChoice = (action: string, glyphs: string[], title: string): HTMLButtonElement => {
  const btn = document.createElement('button');
  btn.className = 'app__cf-icon-choice';
  btn.type = 'button';
  btn.title = title;
  btn.setAttribute('aria-label', title);
  btn.dataset.cfAction = action;
  for (const glyph of glyphs) {
    const span = document.createElement('span');
    span.textContent = glyph;
    btn.appendChild(span);
  }
  return btn;
};

const createCfPanelSubmenu = (key: CfSubmenuKey, label: string): HTMLDivElement => {
  const t = cfMenuText();
  const submenu = document.createElement('div');
  submenu.className = `app__submenu app__submenu--cf app__submenu--cf-${key}`;
  submenu.dataset.cfPanel = key;
  submenu.setAttribute('role', 'menu');
  submenu.setAttribute('aria-label', label);
  submenu.hidden = true;

  if (key === 'highlight') {
    submenu.append(
      cfMenuItem(t.greater, 'cell-gt', 'rule'),
      cfMenuItem(t.less, 'cell-lt', 'rule'),
      cfMenuItem(t.between, 'cell-between', 'rule'),
      cfMenuItem(t.equal, 'cell-eq', 'rule'),
      cfMenuItem(t.text, 'text-contains', 'rule'),
      cfMenuItem(t.date, 'date-occurring', 'rule'),
      cfMenuItem(t.duplicate, 'duplicates', 'rule'),
      cfMenuItem(t.unique, 'unique', 'rule'),
      menuSeparator(),
      cfMenuItem(t.otherRules, 'new-rule', 'manage'),
    );
  } else if (key === 'topBottom') {
    submenu.append(
      cfMenuItem(t.top10, 'top10', 'top'),
      cfMenuItem(t.bottom10, 'bottom10', 'top'),
      cfMenuItem(t.top10Percent, 'top10-percent', 'top'),
      cfMenuItem(t.bottom10Percent, 'bottom10-percent', 'top'),
      cfMenuItem(t.aboveAvg, 'above-avg', 'top'),
      cfMenuItem(t.belowAvg, 'below-avg', 'top'),
      menuSeparator(),
      cfMenuItem(t.otherRules, 'new-rule', 'manage'),
    );
  } else if (key === 'dataBar') {
    submenu.append(menuSectionHeader(t.gradient));
    const gradient = document.createElement('div');
    gradient.className = 'app__cf-choice-row';
    const gradientBars: readonly (readonly [string, string, string])[] = [
      ['data-blue', '#638ec6', t.dataBarGradientBlue],
      ['data-green', '#63a95c', t.dataBarGradientGreen],
      ['data-red', '#c45a5a', t.dataBarGradientRed],
      ['data-orange', '#d6a440', t.dataBarGradientOrange],
      ['data-purple', '#8a74b9', t.dataBarGradientPurple],
      ['data-teal', '#4ba1a8', t.dataBarGradientTeal],
    ];
    gradientBars.forEach(([action, color, label]) => {
      gradient.appendChild(cfSwatchButton(action, ['#fff', color], label));
    });
    submenu.appendChild(gradient);
    submenu.append(menuSectionHeader(t.solid));
    const solid = document.createElement('div');
    solid.className = 'app__cf-choice-row';
    const solidBars: readonly (readonly [string, string, string])[] = [
      ['data-solid-blue', '#4472c4', t.dataBarSolidBlue],
      ['data-solid-green', '#70ad47', t.dataBarSolidGreen],
      ['data-solid-red', '#c00000', t.dataBarSolidRed],
      ['data-solid-orange', '#ed7d31', t.dataBarSolidOrange],
      ['data-solid-purple', '#8064a2', t.dataBarSolidPurple],
      ['data-solid-gray', '#7f7f7f', t.dataBarSolidGray],
    ];
    solidBars.forEach(([action, color, label]) => {
      solid.appendChild(cfSwatchButton(action, [color, color], label));
    });
    submenu.append(solid, menuSeparator(), cfMenuItem(t.otherRules, 'new-rule', 'manage'));
  } else if (key === 'colorScale') {
    const scales = document.createElement('div');
    scales.className = 'app__cf-choice-grid-panel';
    const colorScales: readonly (readonly [string, readonly string[], string])[] = [
      ['scale-gyr', ['#63be7b', '#ffeb84', '#f8696b'], t.colorScaleGreenYellowRed],
      ['scale-ryg', ['#f8696b', '#ffeb84', '#63be7b'], t.colorScaleRedYellowGreen],
      ['scale-gw', ['#63be7b', '#ffffff'], t.colorScaleGreenWhite],
      ['scale-rw', ['#f8696b', '#ffffff'], t.colorScaleRedWhite],
      ['scale-bwr', ['#5a8dee', '#ffffff', '#f8696b'], t.colorScaleBlueWhiteRed],
      ['scale-rwb', ['#f8696b', '#ffffff', '#5a8dee'], t.colorScaleRedWhiteBlue],
      ['scale-gwg', ['#63be7b', '#ffffff', '#00a651'], t.colorScaleGreenWhiteGreen],
      ['scale-ywg', ['#ffeb84', '#ffffff', '#63be7b'], t.colorScaleYellowWhiteGreen],
      ['scale-rwr', ['#f8696b', '#ffffff', '#c00000'], t.colorScaleRedWhiteRed],
      ['scale-bwb', ['#5a8dee', '#ffffff', '#4472c4'], t.colorScaleBlueWhiteBlue],
      ['scale-yry', ['#ffeb84', '#f8696b', '#63be7b'], t.colorScaleYellowRedGreen],
      ['scale-gyg', ['#63be7b', '#ffeb84', '#00a651'], t.colorScaleGreenYellowGreen],
    ];
    colorScales.forEach(([action, colors, label]) => {
      scales.appendChild(cfSwatchButton(action, [...colors], label));
    });
    submenu.append(scales, menuSeparator(), cfMenuItem(t.otherRules, 'new-rule', 'manage'));
  } else if (key === 'iconSet') {
    submenu.append(menuSectionHeader(t.direction));
    const directions = document.createElement('div');
    directions.className = 'app__cf-icon-panel';
    directions.append(
      cfIconChoice('icons-arrows3', ['▲', '▶', '▼'], t.iconArrows3),
      cfIconChoice('icons-arrows5', ['▲', '↗', '▶', '↘', '▼'], t.iconArrows5),
      cfIconChoice('icons-triangles3', ['▲', '▬', '▼'], t.iconTriangles3),
    );
    submenu.appendChild(directions);
    submenu.append(menuSectionHeader(t.shapes));
    const shapes = document.createElement('div');
    shapes.className = 'app__cf-icon-panel';
    shapes.append(
      cfIconChoice('icons-traffic3', ['●', '●', '●'], t.iconTraffic3),
      cfIconChoice('icons-trafficRim3', ['●', '●', '●'], t.iconTrafficRim3),
      cfIconChoice('icons-stars3', ['★', '★', '★'], t.iconStars3),
    );
    submenu.append(shapes, menuSectionHeader(t.indicators));
    const indicators = document.createElement('div');
    indicators.className = 'app__cf-icon-panel';
    indicators.append(
      cfIconChoice('icons-symbols3', ['✓', '!', '×'], t.iconSymbols3),
      cfIconChoice('icons-flags3', ['⚑', '⚑', '⚑'], t.iconFlags3),
    );
    submenu.append(indicators, menuSectionHeader(t.ratings));
    const ratings = document.createElement('div');
    ratings.className = 'app__cf-icon-panel';
    ratings.append(
      cfIconChoice('icons-stars3', ['★', '★', '★'], t.ratings),
      cfIconChoice('icons-quarters5', ['◔', '◑', '◕', '●', '●'], t.iconQuarters5),
      cfIconChoice('icons-ratings5', ['●', '●', '●', '●', '●'], t.iconRatings5),
      cfIconChoice('icons-bars5', ['▮', '▮', '▮', '▮', '▮'], t.iconBars5),
      cfIconChoice('icons-boxes5', ['■', '■', '■', '■', '■'], t.iconBoxes5),
    );
    submenu.append(ratings, menuSeparator(), cfMenuItem(t.otherRules, 'new-rule', 'manage'));
  } else if (key === 'clear') {
    submenu.append(
      cfMenuItem(t.clearSelection, 'clear-selection', 'clear'),
      cfMenuItem(t.clearSheet, 'clear-sheet', 'clear'),
    );
  }
  return submenu;
};

const createConditionalMenu = (): HTMLDivElement => {
  const t = cfMenuText();
  const menu = createMenu('menu-conditional');
  menu.classList.add('app__menu--conditional');
  menu.append(
    cfSubmenuTrigger('highlight', t.highlight, 'rule'),
    cfSubmenuTrigger('topBottom', t.topBottom, 'top'),
    menuSeparator(),
    cfSubmenuTrigger('dataBar', t.dataBar, 'bar'),
    cfSubmenuTrigger('colorScale', t.colorScale, 'scale'),
    cfSubmenuTrigger('iconSet', t.iconSet, 'icon'),
    menuSeparator(),
    cfMenuItem(t.newRule, 'new-rule', 'new'),
    cfSubmenuTrigger('clear', t.clear, 'clear'),
    cfMenuItem(t.manage, 'manage', 'manage'),
  );
  menu.append(
    createCfPanelSubmenu('highlight', t.highlight),
    createCfPanelSubmenu('topBottom', t.topBottom),
    createCfPanelSubmenu('dataBar', t.dataBar),
    createCfPanelSubmenu('colorScale', t.colorScale),
    createCfPanelSubmenu('iconSet', t.iconSet),
    createCfPanelSubmenu('clear', t.clear),
  );
  return menu;
};

const drawActionItem = (
  action: string,
  label: string,
  icon?: BorderPreviewSpec,
): HTMLButtonElement => {
  const btn = document.createElement('button');
  btn.className = 'app__menu-item app__menu-item--preset';
  btn.type = 'button';
  btn.setAttribute('role', 'menuitemcheckbox');
  btn.setAttribute('aria-checked', 'false');
  btn.dataset.borderDraw = action;
  if (icon) btn.appendChild(createBorderPreview(icon));
  else {
    const spacer = document.createElement('span');
    spacer.className = 'app__menu-item__icon-spacer';
    btn.appendChild(spacer);
  }
  const text = document.createElement('span');
  text.className = 'app__menu-item__text';
  text.textContent = label;
  btn.appendChild(text);
  return btn;
};

const submenuTrigger = (
  submenuKey: 'lineColor' | 'lineStyle',
  label: string,
): HTMLButtonElement => {
  const btn = document.createElement('button');
  btn.className = 'app__menu-item app__menu-item--preset app__menu-item--submenu';
  btn.type = 'button';
  btn.setAttribute('role', 'menuitem');
  btn.setAttribute('aria-haspopup', 'menu');
  btn.setAttribute('aria-expanded', 'false');
  btn.dataset.borderSubmenu = submenuKey;
  const spacer = document.createElement('span');
  spacer.className = 'app__menu-item__icon-spacer';
  btn.appendChild(spacer);
  const text = document.createElement('span');
  text.className = 'app__menu-item__text';
  text.textContent = label;
  btn.appendChild(text);
  const caret = document.createElement('span');
  caret.className = 'app__menu-item__caret';
  caret.setAttribute('aria-hidden', 'true');
  caret.textContent = '▶';
  btn.appendChild(caret);
  return btn;
};

// Line sample preview + LINE_STYLES_ALL live in ribbon/border-icons.ts.

const createLineStyleSubmenu = (label: string): HTMLDivElement => {
  const submenu = document.createElement('div');
  submenu.className = 'app__submenu app__submenu--line-style';
  submenu.id = 'menu-borders-line-style';
  submenu.setAttribute('role', 'menu');
  submenu.setAttribute('aria-label', label);
  submenu.hidden = true;
  for (const value of LINE_STYLES_ALL) {
    const btn = document.createElement('button');
    btn.className = 'app__submenu-item';
    btn.type = 'button';
    btn.setAttribute('role', 'menuitemradio');
    btn.setAttribute('aria-checked', value === 'thin' ? 'true' : 'false');
    btn.dataset.borderLineStyle = value;
    if (value === 'none') {
      const span = document.createElement('span');
      span.textContent = ribbonText.lineStyleNone;
      span.className = 'app__submenu-item__text';
      btn.appendChild(span);
    } else {
      btn.appendChild(createLineSamplePreview(value));
    }
    submenu.appendChild(btn);
  }
  return submenu;
};

const createLineColorSubmenu = (label: string): HTMLDivElement => {
  const submenu = document.createElement('div');
  submenu.className = 'app__submenu app__submenu--line-color';
  submenu.id = 'menu-borders-line-color';
  submenu.setAttribute('role', 'menu');
  submenu.setAttribute('aria-label', label);
  submenu.hidden = true;
  const palette = createColorPalette({
    themeLabel: ribbonText.themeColors,
    standardLabel: ribbonText.standardColors,
    ariaLabel: label,
    value: selectedBorderColor,
    automatic: { label: ribbonText.automatic, color: '#000000' },
    onPick: (color) => {
      selectedBorderColor = color;
      inst?.borderDraw?.setColor(color);
      closeBorderSubmenus();
    },
  });
  submenu.appendChild(palette.el);
  return submenu;
};

const createBordersMenu = (): HTMLDivElement => {
  const t = ribbonText;
  const menu = createMenu('menu-borders');
  menu.classList.add('app__menu--borders');
  menu.append(
    // Section 1: single-side edges (image 1: 下罫線 / 上罫線 / 左罫線 / 右罫線)
    presetMenuItem('bottom', t.bottomBorder),
    presetMenuItem('top', t.topBorder),
    presetMenuItem('left', t.leftBorder),
    presetMenuItem('right', t.rightBorder),
    menuSeparator(),
    // Section 2: frame presets
    presetMenuItem('clear', t.noBorder),
    presetMenuItem('all', t.allBorders),
    presetMenuItem('outline', t.outsideBorders),
    presetMenuItem('thickOutline', t.thickOutsideBorders),
    menuSeparator(),
    // Section 3: combined
    presetMenuItem('doubleBottom', t.doubleBottomBorder),
    presetMenuItem('thickBottom', t.thickBottomBorder),
    presetMenuItem('topAndBottom', t.topAndBottomBorder),
    presetMenuItem('topAndThickBottom', t.topAndThickBottomBorder),
    presetMenuItem('topAndDoubleBottom', t.topAndDoubleBottomBorder),
    // Section 4 header + draw actions
    menuSectionHeader(t.drawBordersHeading),
    drawActionItem('draw', t.drawBorder, { bottom: 'thin' }),
    drawActionItem('grid', t.drawBorderGrid, {
      top: 'thin',
      right: 'thin',
      bottom: 'thin',
      left: 'thin',
      innerGrid: true,
      showBase: false,
    }),
    drawActionItem('erase', t.eraseBorder),
    submenuTrigger('lineColor', t.lineColor),
    submenuTrigger('lineStyle', t.lineStyle),
    menuSeparator(),
    // Footer
    presetMenuItem('format', t.moreBorders),
  );
  // Submenus sit beside the main dropdown.
  menu.appendChild(createLineColorSubmenu(t.lineColor));
  menu.appendChild(createLineStyleSubmenu(t.lineStyle));
  return menu;
};

const createFreezeMenu = (): HTMLDivElement => {
  const ja = ribbonLang === 'ja';
  const menu = createMenu('menu-freeze');
  menu.append(
    menuButton(ja ? '先頭行の固定' : 'Freeze Top Row', 'freeze', 'row'),
    menuButton(ja ? '先頭列の固定' : 'Freeze First Column', 'freeze', 'col'),
    menuButton(ja ? 'ウィンドウ枠の固定' : 'Freeze Panes', 'freeze', 'selection'),
    menuButton(ja ? 'ウィンドウ枠固定の解除' : 'Unfreeze Panes', 'freeze', 'off'),
  );
  return menu;
};

const createFillMenu = (): HTMLDivElement => {
  const t = toolbarMenuText(ribbonLang);
  const menu = createMenu('menu-fill');
  menu.append(
    menuButton(t.fillDown, 'fill', 'down'),
    menuButton(t.fillRight, 'fill', 'right'),
    menuButton(t.fillUp, 'fill', 'up'),
    menuButton(t.fillLeft, 'fill', 'left'),
    menuSeparator(),
    menuButton(t.series, 'fill', 'series'),
    menuSeparator(),
    menuButton(t.fillDays, 'fill', 'days'),
    menuButton(t.fillWeekdays, 'fill', 'weekdays'),
    menuButton(t.fillMonths, 'fill', 'months'),
    menuButton(t.fillYears, 'fill', 'years'),
  );
  return menu;
};

const createClearMenu = (): HTMLDivElement => {
  const t = toolbarMenuText(ribbonLang);
  const menu = createMenu('menu-clear');
  menu.append(
    menuButton(t.clearAll, 'clear', 'all'),
    menuButton(t.clearFormats, 'clear', 'formats'),
    menuButton(t.clearContents, 'clear', 'contents'),
    menuButton(t.clearComments, 'clear', 'comments'),
    menuButton(t.clearHyperlinks, 'clear', 'hyperlinks'),
    menuButton(t.removeHyperlinks, 'clear', 'remove-hyperlinks'),
    menuButton(t.clearConditional, 'clear', 'conditional'),
  );
  return menu;
};

type TextOrientationGlyph = 'ccw' | 'cw' | 'vertical' | 'up' | 'down' | 'format';

const createTextOrientationIcon = (glyph: TextOrientationGlyph): SVGSVGElement => {
  const svg = document.createElementNS(SVG_NS, 'svg');
  svg.setAttribute('viewBox', '0 0 16 16');
  svg.setAttribute('width', '16');
  svg.setAttribute('height', '16');
  svg.setAttribute('focusable', 'false');
  svg.setAttribute('aria-hidden', 'true');
  svg.setAttribute('fill', 'none');
  svg.setAttribute('stroke', 'currentColor');
  svg.setAttribute('stroke-width', '1.2');
  svg.setAttribute('stroke-linecap', 'round');
  svg.setAttribute('stroke-linejoin', 'round');
  const baseline = document.createElementNS(SVG_NS, 'line');
  baseline.setAttribute('x1', '2');
  baseline.setAttribute('y1', '13');
  baseline.setAttribute('x2', '14');
  baseline.setAttribute('y2', '13');
  svg.appendChild(baseline);
  if (glyph === 'ccw' || glyph === 'cw') {
    const angle = glyph === 'ccw' ? -35 : 35;
    const text = document.createElementNS(SVG_NS, 'text');
    text.setAttribute('x', '4');
    text.setAttribute('y', '11');
    text.setAttribute('transform', `rotate(${angle} 8 11)`);
    text.setAttribute('font-family', 'system-ui, sans-serif');
    text.setAttribute('font-size', '7');
    text.setAttribute('font-weight', '700');
    text.setAttribute('fill', 'currentColor');
    text.setAttribute('stroke', 'none');
    text.textContent = 'ab';
    svg.appendChild(text);
  } else if (glyph === 'vertical') {
    for (let i = 0; i < 3; i += 1) {
      const ch = document.createElementNS(SVG_NS, 'text');
      ch.setAttribute('x', '8');
      ch.setAttribute('y', String(4 + i * 3));
      ch.setAttribute('text-anchor', 'middle');
      ch.setAttribute('font-family', 'system-ui, sans-serif');
      ch.setAttribute('font-size', '3');
      ch.setAttribute('font-weight', '700');
      ch.setAttribute('fill', 'currentColor');
      ch.setAttribute('stroke', 'none');
      ch.textContent = 'a';
      svg.appendChild(ch);
    }
  } else if (glyph === 'up' || glyph === 'down') {
    const text = document.createElementNS(SVG_NS, 'text');
    text.setAttribute('x', '0');
    text.setAttribute('y', '0');
    const rotate = glyph === 'up' ? -90 : 90;
    text.setAttribute('transform', `translate(8 11) rotate(${rotate})`);
    text.setAttribute('text-anchor', 'middle');
    text.setAttribute('font-family', 'system-ui, sans-serif');
    text.setAttribute('font-size', '7');
    text.setAttribute('font-weight', '700');
    text.setAttribute('fill', 'currentColor');
    text.setAttribute('stroke', 'none');
    text.textContent = 'ab';
    svg.appendChild(text);
  } else if (glyph === 'format') {
    const grid = document.createElementNS(SVG_NS, 'rect');
    grid.setAttribute('x', '2.5');
    grid.setAttribute('y', '3.5');
    grid.setAttribute('width', '11');
    grid.setAttribute('height', '7');
    svg.appendChild(grid);
    const hLine = document.createElementNS(SVG_NS, 'line');
    hLine.setAttribute('x1', '2.5');
    hLine.setAttribute('y1', '7');
    hLine.setAttribute('x2', '13.5');
    hLine.setAttribute('y2', '7');
    svg.appendChild(hLine);
    const vLine = document.createElementNS(SVG_NS, 'line');
    vLine.setAttribute('x1', '8');
    vLine.setAttribute('y1', '3.5');
    vLine.setAttribute('x2', '8');
    vLine.setAttribute('y2', '10.5');
    svg.appendChild(vLine);
  }
  return svg;
};

const textOrientationMenuItem = (
  glyph: TextOrientationGlyph,
  label: string,
  value: string,
): HTMLButtonElement => {
  const btn = document.createElement('button');
  btn.className = 'app__menu-item app__menu-item--preset';
  btn.type = 'button';
  btn.setAttribute('role', 'menuitem');
  btn.dataset.textOrientation = value;
  btn.appendChild(createTextOrientationIcon(glyph));
  const text = document.createElement('span');
  text.className = 'app__menu-item__text';
  text.textContent = label;
  btn.appendChild(text);
  return btn;
};

const createTextOrientationMenu = (): HTMLDivElement => {
  const t = ribbonMenuText;
  const menu = createMenu('menu-text-orientation');
  menu.append(
    textOrientationMenuItem('ccw', t.orientationAngleCounterclockwise, 'ccw'),
    textOrientationMenuItem('cw', t.orientationAngleClockwise, 'cw'),
    textOrientationMenuItem('vertical', t.orientationVerticalText, 'vertical'),
    textOrientationMenuItem('up', t.orientationRotateTextUp, 'up'),
    textOrientationMenuItem('down', t.orientationRotateTextDown, 'down'),
    menuSeparator(),
    textOrientationMenuItem('format', t.orientationFormatAlignment, 'format'),
  );
  return menu;
};

const createInsertCellsMenu = (): HTMLDivElement => {
  const ja = ribbonLang === 'ja';
  const t = toolbarMenuText(ribbonLang);
  const sheetTabs = dictionaries[ribbonLang].sheetTabs;
  const menu = createMenu('menu-insert-cells');
  menu.append(
    menuButton(ja ? 'セルを挿入...' : 'Insert Cells...', 'cellInsert', 'cells'),
    menuButton(t.insertShiftDown, 'cellInsert', 'shift-down'),
    menuButton(t.insertShiftRight, 'cellInsert', 'shift-right'),
    menuSeparator(),
    menuButton(ja ? 'シートの行を挿入' : 'Insert Sheet Rows', 'cellInsert', 'rows'),
    menuButton(ja ? 'シートの列を挿入' : 'Insert Sheet Columns', 'cellInsert', 'cols'),
    menuSeparator(),
    menuButton(sheetTabs.insertSheet, 'cellInsert', 'sheet'),
  );
  return menu;
};

const createDeleteCellsMenu = (): HTMLDivElement => {
  const ja = ribbonLang === 'ja';
  const t = toolbarMenuText(ribbonLang);
  const sheetTabs = dictionaries[ribbonLang].sheetTabs;
  const menu = createMenu('menu-delete-cells');
  menu.append(
    menuButton(ja ? 'セルを削除...' : 'Delete Cells...', 'cellDelete', 'cells'),
    menuButton(t.deleteShiftUp, 'cellDelete', 'shift-up'),
    menuButton(t.deleteShiftLeft, 'cellDelete', 'shift-left'),
    menuSeparator(),
    menuButton(ja ? 'シートの行を削除' : 'Delete Sheet Rows', 'cellDelete', 'rows'),
    menuButton(ja ? 'シートの列を削除' : 'Delete Sheet Columns', 'cellDelete', 'cols'),
    menuSeparator(),
    menuButton(sheetTabs.deleteSheet, 'cellDelete', 'sheet'),
  );
  return menu;
};

const createFormatCellsMenu = (): HTMLDivElement => {
  const t = toolbarMenuText(ribbonLang);
  const sheetTabs = dictionaries[ribbonLang].sheetTabs;
  const menu = createMenu('menu-format-cells');
  menu.append(
    menuButton(t.formatCells, 'cellFormat', 'dialog'),
    menuSeparator(),
    menuButton(t.rowHeight, 'cellFormat', 'row-height'),
    menuButton(t.autoFitRowHeight, 'cellFormat', 'row-autofit'),
    menuButton(t.colWidth, 'cellFormat', 'col-width'),
    menuButton(t.autoFitColWidth, 'cellFormat', 'col-autofit'),
    menuSeparator(),
    menuButton(t.hideRows, 'cellFormat', 'hide-rows'),
    menuButton(t.showRows, 'cellFormat', 'show-rows'),
    menuButton(t.hideCols, 'cellFormat', 'hide-cols'),
    menuButton(t.showCols, 'cellFormat', 'show-cols'),
    menuSeparator(),
    menuButton(sheetTabs.rename, 'cellFormat', 'rename-sheet'),
    menuButton(sheetTabs.moveLeft, 'cellFormat', 'move-sheet-left'),
    menuButton(sheetTabs.moveRight, 'cellFormat', 'move-sheet-right'),
    menuButton(sheetTabs.hideSheet, 'cellFormat', 'hide-sheet'),
    menuButton(sheetTabs.unhideSheet, 'cellFormat', 'unhide-sheet'),
    menuSeparator(),
    menuButton(`${sheetTabs.tabColor}: ${sheetTabs.noColor}`, 'cellFormat', 'tab-color-none'),
    menuButton(`${sheetTabs.tabColor}: ${sheetTabs.tabColorRed}`, 'cellFormat', 'tab-color-red'),
    menuButton(
      `${sheetTabs.tabColor}: ${sheetTabs.tabColorOrange}`,
      'cellFormat',
      'tab-color-orange',
    ),
    menuButton(
      `${sheetTabs.tabColor}: ${sheetTabs.tabColorYellow}`,
      'cellFormat',
      'tab-color-yellow',
    ),
    menuButton(
      `${sheetTabs.tabColor}: ${sheetTabs.tabColorGreen}`,
      'cellFormat',
      'tab-color-green',
    ),
    menuButton(`${sheetTabs.tabColor}: ${sheetTabs.tabColorBlue}`, 'cellFormat', 'tab-color-blue'),
    menuButton(
      `${sheetTabs.tabColor}: ${sheetTabs.tabColorPurple}`,
      'cellFormat',
      'tab-color-purple',
    ),
    menuButton(`${sheetTabs.tabColor}: ${sheetTabs.tabColorGray}`, 'cellFormat', 'tab-color-gray'),
    menuSeparator(),
    menuButton(ribbonMenuText.lockCell, 'cellFormat', 'lock-cell'),
    menuButton(ribbonMenuText.unlockCell, 'cellFormat', 'unlock-cell'),
    menuButton(t.protectSheet, 'cellFormat', 'protect-sheet'),
  );
  return menu;
};

type AutoSumFormulaName = 'SUM' | 'AVERAGE' | 'COUNT' | 'MAX' | 'MIN' | 'MORE';

const createAutoSumMenu = (id: string): HTMLDivElement => {
  const t = toolbarMenuText(ribbonLang);
  const menu = createMenu(id);
  menu.append(
    menuButton(t.autosumSum, 'autosumFn', 'SUM'),
    menuButton(t.autosumAverage, 'autosumFn', 'AVERAGE'),
    menuButton(t.autosumCount, 'autosumFn', 'COUNT'),
    menuButton(t.autosumMax, 'autosumFn', 'MAX'),
    menuButton(t.autosumMin, 'autosumFn', 'MIN'),
    menuSeparator(),
    menuButton(t.autosumMoreFunctions, 'autosumFn', 'MORE'),
  );
  return menu;
};

const calcOptionButton = (label: string, value: string): HTMLButtonElement => {
  const button = menuButton(label, 'calcOption', value);
  if (value === 'auto' || value === 'manual' || value === 'auto-no-table') {
    button.setAttribute('role', 'menuitemradio');
    button.setAttribute('aria-checked', 'false');
  }
  return button;
};

const createCalcOptionsMenu = (): HTMLDivElement => {
  const ja = ribbonLang === 'ja';
  const menu = createMenu('menu-calc-options');
  menu.append(
    calcOptionButton(ja ? '自動' : 'Automatic', 'auto'),
    calcOptionButton(
      ja ? 'データ テーブル以外自動' : 'Automatic Except for Data Tables',
      'auto-no-table',
    ),
    calcOptionButton(ja ? '手動' : 'Manual', 'manual'),
    menuSeparator(),
    calcOptionButton(ja ? '再計算実行' : 'Calculate Now', 'calculate-now'),
    calcOptionButton(ja ? 'シート再計算' : 'Calculate Sheet', 'calculate-sheet'),
    menuSeparator(),
    calcOptionButton(ja ? '反復計算...' : 'Iterative Calculation...', 'iterative'),
  );
  return menu;
};

const createClearArrowsMenu = (): HTMLDivElement => {
  const t = ribbonMenuText;
  const menu = createMenu('menu-clear-arrows');
  menu.append(
    menuButton(t.removeArrowsAll, 'formulaAuditAction', 'clear-all'),
    menuButton(t.removePrecedentArrows, 'formulaAuditAction', 'clear-precedents'),
    menuButton(t.removeDependentArrows, 'formulaAuditAction', 'clear-dependents'),
  );
  return menu;
};

const createErrorCheckingMenu = (): HTMLDivElement => {
  const t = ribbonMenuText;
  const menu = createMenu('menu-error-checking');
  menu.append(
    menuButton(t.errorChecking, 'formulaAuditAction', 'error-checking'),
    menuButton(t.traceError, 'formulaAuditAction', 'trace-error'),
    menuSeparator(),
    menuButton(t.ignoreError, 'formulaAuditAction', 'ignore-error'),
  );
  return menu;
};

const createWatchMenu = (id: string): HTMLDivElement => {
  const t = ribbonMenuText;
  const menu = createMenu(id);
  menu.append(
    menuButton(t.watchWindow, 'watchAction', 'open'),
    menuButton(t.watchAdd, 'watchAction', 'add'),
    menuButton(t.watchDelete, 'watchAction', 'delete'),
    menuSeparator(),
    menuButton(t.watchDeleteAll, 'watchAction', 'delete-all'),
  );
  return menu;
};

const createReviewCommentsMenu = (): HTMLDivElement => {
  const t = ribbonMenuText;
  const menu = createMenu('menu-review-comments');
  menu.append(
    menuButton(t.commentDelete, 'commentAction', 'delete-active'),
    menuButton(t.commentDeleteAll, 'commentAction', 'delete-all'),
  );
  return menu;
};

const createProtectMenu = (id: string): HTMLDivElement => {
  const t = ribbonMenuText;
  const menu = createMenu(id);
  menu.append(
    menuButton(t.protectSheetCommand, 'protectAction', 'protect-sheet'),
    menuButton(t.unprotectSheetCommand, 'protectAction', 'unprotect-sheet'),
    menuSeparator(),
    menuButton(t.lockCell, 'protectAction', 'lock-cell'),
    menuButton(t.unlockCell, 'protectAction', 'unlock-cell'),
    menuSeparator(),
    menuButton(t.protectWorkbookCommand, 'protectAction', 'protect-workbook'),
    menuButton(t.unprotectWorkbookCommand, 'protectAction', 'unprotect-workbook'),
    menuButton(t.allowEditRangesCommand, 'protectAction', 'allow-edit-ranges'),
    menuButton(t.allowEditRangesClearCommand, 'protectAction', 'clear-allowed-edit-ranges'),
  );
  return menu;
};

const createChartInsertMenu = (): HTMLDivElement => {
  const t = ribbonMenuText;
  const menu = createMenu('menu-chart-insert');
  menu.append(
    menuButton(t.chartColumn, 'chartInsert', 'column'),
    menuButton(t.chartBar, 'chartInsert', 'bar'),
    menuButton(t.chartLine, 'chartInsert', 'line'),
    menuButton(t.chartArea, 'chartInsert', 'area'),
    menuButton(t.chartPie, 'chartInsert', 'pie'),
    menuButton(t.chartScatter, 'chartInsert', 'scatter'),
    menuSeparator(),
    menuButton(t.recommendedCharts, 'chartInsert', 'recommended'),
  );
  return menu;
};

const createPictureInsertMenu = (): HTMLDivElement => {
  const t = ribbonMenuText;
  const menu = createMenu('menu-picture-insert');
  menu.append(
    menuButton(t.pictureThisDevice, 'pictureInsert', 'device'),
    menuButton(t.pictureOnline, 'pictureInsert', 'online'),
  );
  return menu;
};

const createShapesInsertMenu = (): HTMLDivElement => {
  const t = ribbonMenuText;
  const menu = createMenu('menu-shapes-insert');
  menu.append(
    menuButton(t.shapeRectangle, 'shapeInsert', 'rectangle'),
    menuButton(t.shapeRoundedRectangle, 'shapeInsert', 'rounded-rectangle'),
    menuButton(t.shapeOval, 'shapeInsert', 'oval'),
    menuSeparator(),
    menuButton(t.shapeLine, 'shapeInsert', 'line'),
    menuButton(t.shapeArrow, 'shapeInsert', 'arrow'),
  );
  return menu;
};

const createScreenshotInsertMenu = (): HTMLDivElement => {
  const t = ribbonMenuText;
  const menu = createMenu('menu-screenshot-insert');
  menu.append(
    menuButton(t.screenshotCurrentView, 'screenshotInsert', 'current-view'),
    menuButton(t.screenshotScreenClipping, 'screenshotInsert', 'screen-clipping'),
  );
  return menu;
};

const createScriptMenu = (): HTMLDivElement => {
  const t = ribbonMenuText;
  const menu = createMenu('menu-script');
  menu.append(
    menuButton(t.scriptCommandUppercase, 'scriptAction', 'uppercase'),
    menuButton(t.scriptCommandLowercase, 'scriptAction', 'lowercase'),
    menuButton(t.scriptCommandTrim, 'scriptAction', 'trim'),
    menuButton(t.scriptCommandClear, 'scriptAction', 'clear'),
    menuSeparator(),
    menuButton(t.scriptRunCustom, 'scriptAction', 'custom'),
  );
  return menu;
};

const createAddInMenu = (): HTMLDivElement => {
  const t = ribbonMenuText;
  const menu = createMenu('menu-add-ins');
  menu.append(
    menuButton(t.addInGet, 'addInAction', 'get'),
    menuButton(t.addInMy, 'addInAction', 'my'),
    menuSeparator(),
    menuButton(t.addInManage, 'addInAction', 'manage'),
  );
  return menu;
};

const createPdfMenu = (): HTMLDivElement => {
  const t = ribbonMenuText;
  const menu = createMenu('menu-pdf');
  menu.append(
    menuButton(t.pdfCreate, 'pdfAction', 'create'),
    menuButton(t.pdfShare, 'pdfAction', 'share'),
    menuSeparator(),
    menuButton(t.pdfPreferences, 'pdfAction', 'preferences'),
  );
  return menu;
};

type TableVariantId = 'plain' | 'banded' | 'firstCol' | 'bandedFirstCol';

const TABLE_VARIANTS_LIGHT_MEDIUM: TableVariantId[] = [
  'plain',
  'banded',
  'firstCol',
  'bandedFirstCol',
];
const TABLE_VARIANTS_DARK: TableVariantId[] = ['banded'];

const tableVariantOptions = (variant: TableVariantId): { banded: boolean; firstCol: boolean } => {
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
): HTMLButtonElement => {
  const swatch = tableStyleSwatch(style, color);
  const { banded, firstCol } = tableVariantOptions(variant);
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'app__tablestyle-swatch';
  btn.setAttribute('role', 'menuitem');
  btn.dataset.tableStyle = style;
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

const createTableStyleMenu = (id: string): HTMLDivElement => {
  const t = toolbarMenuText(ribbonLang);
  const menu = createMenu(id);
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
  menu.appendChild(menuSeparator());
  menu.appendChild(tableStyleFooterButton(t.tableStyleNew, 'new-table-style'));
  menu.appendChild(tableStyleFooterButton(t.tableStyleNewPivot, 'new-pivot-style'));
  return menu;
};

const cellStyleGalleryLabel = (id: CellStyleId): string => {
  const strings = dictionaries[ribbonLang].cellStylesGallery.styles;
  return strings[id] ?? CELL_STYLES.find((s) => s.id === id)?.label ?? id;
};

const cellStyleGroupLabel = (id: CellStyleGroupId): string =>
  dictionaries[ribbonLang].cellStylesGallery.groups[id];

const createCellStyleChip = (id: CellStyleId): HTMLButtonElement => {
  const def = CELL_STYLES.find((s) => s.id === id);
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'app__menu-item app__cellstyle-chip';
  btn.setAttribute('role', 'menuitem');
  btn.dataset.cellStyle = id;
  const label = cellStyleGalleryLabel(id);
  btn.title = label;
  btn.setAttribute('aria-label', label);
  btn.textContent = label;
  const fmt = def?.format ?? {};
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

const createCellStylesMenu = (): HTMLDivElement => {
  const t = toolbarMenuText(ribbonLang);
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
  menu.appendChild(menuSeparator());
  menu.appendChild(cellStyleFooterButton(t.cellStyleNew, 'new-cell-style'));
  menu.appendChild(cellStyleFooterButton(t.cellStyleMerge, 'merge-cell-style'));
  return menu;
};

const createSortMenu = (id: string): HTMLDivElement => {
  const t = ribbonMenuText;
  const menu = createMenu(id);
  menu.append(
    menuButton(t.sortAscendingMenu, 'sort', 'asc'),
    menuButton(t.sortDescendingMenu, 'sort', 'desc'),
    menuButton(t.sortCustom, 'sort', 'custom'),
    menuSeparator(),
    menuButton(t.filterToggle, 'sort', 'filter'),
    menuButton(t.filterBySelectedCellValue, 'sort', 'filter-by-value'),
    menuButton(t.filterClearAll, 'sort', 'filter-clear'),
    menuButton(t.filterReapply, 'sort', 'filter-reapply'),
    menuButton(t.filterAdvanced, 'sort', 'filter-advanced'),
    menuSeparator(),
    menuButton(ribbonText.removeDuplicates, 'sort', 'dedupe'),
    menuButton(ribbonText.conditionalFormatting, 'sort', 'conditional'),
    menuButton(t.nameManager, 'sort', 'named'),
  );
  return menu;
};

const createTextToColumnsMenu = (): HTMLDivElement => {
  const t = ribbonMenuText;
  const menu = createMenu('menu-text-to-columns');
  menu.append(
    menuButton(t.textToColumnsComma, 'textToColumnsDelimiter', ','),
    menuButton(t.textToColumnsTab, 'textToColumnsDelimiter', '\\t'),
    menuButton(t.textToColumnsSemicolon, 'textToColumnsDelimiter', ';'),
    menuButton(t.textToColumnsSpace, 'textToColumnsDelimiter', ' '),
    menuSeparator(),
    menuButton(t.textToColumnsCustom, 'textToColumnsDelimiter', 'custom'),
  );
  return menu;
};

const createFindSelectMenu = (): HTMLDivElement => {
  const t = toolbarMenuText(ribbonLang);
  const menu = createMenu('menu-find-select');
  menu.append(
    menuButton(t.find, 'findSelect', 'find'),
    menuButton(t.replace, 'findSelect', 'replace'),
    menuButton(t.goTo, 'findSelect', 'go-to'),
    menuButton(t.goToSpecial, 'findSelect', 'go-to-special'),
    menuSeparator(),
    menuButton(t.findFormulas, 'findSelect', 'formulas'),
    menuButton(t.findConstants, 'findSelect', 'constants'),
    menuButton(t.findConditionalFormatting, 'findSelect', 'conditional-format'),
    menuButton(t.findDataValidation, 'findSelect', 'data-validation'),
    menuButton(t.comments, 'findSelect', 'comments'),
  );
  return menu;
};

const selectMatchingAddresses = (
  matches: readonly { sheet: number; row: number; col: number }[],
) => {
  const i = inst;
  if (!i) return;
  if (matches.length === 0) {
    void showMessage({
      title: ribbonMenuText.findSelect,
      message: ribbonMenuText.findNoMatches,
    });
    return;
  }
  const range = boundingRange(matches);
  const active = matches[0];
  if (!active) return;
  i.store.setState((state) => ({
    ...state,
    selection: {
      active,
      anchor: active,
      range,
      extraRanges: [],
    },
  }));
  focusSheet();
};

const applyFindSelectAction = (action: string): void => {
  const i = inst;
  if (!i) return;
  if (action === 'find') {
    i.openFindReplace('find');
    return;
  }
  if (action === 'replace') {
    i.openFindReplace('replace');
    return;
  }
  if (action === 'go-to') {
    i.openGoTo();
    return;
  }
  if (action === 'go-to-special') {
    i.openGoToSpecial();
    return;
  }
  if (action === 'conditional-format') {
    selectMatchingAddresses(findMatchingCells(i.workbook, i.store, 'sheet', 'conditional-format'));
    return;
  }
  if (action === 'formulas' || action === 'constants' || action === 'data-validation') {
    selectMatchingAddresses(findMatchingCells(i.workbook, i.store, 'sheet', action));
    return;
  }
  if (action === 'comments') {
    selectMatchingAddresses(listComments(i.store.getState()).map((entry) => entry.addr));
  }
};

const applyAutoSumFormula = (fn: AutoSumFormulaName = 'SUM'): void => {
  const i = inst;
  if (!i) return;
  if (fn === 'MORE') {
    i.openFunctionArguments();
    return;
  }
  i.history.begin();
  let result: ReturnType<typeof autoSum> = null;
  try {
    result = autoSum(i.store.getState(), i.workbook, fn);
  } finally {
    i.history.end();
  }
  if (result) {
    refreshWorkbookCells();
    mutators.setActive(i.store, result.addr);
  }
  focusSheet();
};

const cfSelectionRange = () => {
  const i = inst;
  if (!i) return null;
  const r = i.store.getState().selection.range;
  return {
    sheet: r.sheet,
    r0: Math.min(r.r0, r.r1),
    c0: Math.min(r.c0, r.c1),
    r1: Math.max(r.r0, r.r1),
    c1: Math.max(r.c0, r.c1),
  };
};

const normalizedSelectionRange = () => {
  const i = inst;
  if (!i) return null;
  const r = i.store.getState().selection.range;
  return {
    sheet: r.sheet,
    r0: Math.min(r.r0, r.r1),
    c0: Math.min(r.c0, r.c1),
    r1: Math.max(r.r0, r.r1),
    c1: Math.max(r.c0, r.c1),
  };
};

const clearSelectionContents = (): void => {
  const i = inst;
  const range = normalizedSelectionRange();
  if (!i || !range) return;
  i.history.begin();
  try {
    for (let row = range.r0; row <= range.r1; row += 1) {
      for (let col = range.c0; col <= range.c1; col += 1) {
        i.workbook.setBlank({ sheet: range.sheet, row, col });
      }
    }
  } finally {
    i.history.end();
  }
  refreshWorkbookCells();
  (sheetEl as HTMLElement).focus();
};

const applyFillDirection = (direction: 'down' | 'right' | 'up' | 'left'): void => {
  const i = inst;
  const range = normalizedSelectionRange();
  if (!i || !range) return;
  let src = range;
  if (direction === 'down') src = { ...range, r1: range.r0 };
  else if (direction === 'up') src = { ...range, r0: range.r1 };
  else if (direction === 'right') src = { ...range, c1: range.c0 };
  else src = { ...range, c0: range.c1 };
  if (src.r0 === range.r0 && src.r1 === range.r1 && src.c0 === range.c0 && src.c1 === range.c1) {
    return;
  }
  i.history.begin();
  try {
    recordFormatChange(i.history, i.store, () => {
      fillRange(i.store.getState(), i.workbook, src, range, {
        formatting: 'with',
        store: i.store,
      });
    });
  } finally {
    i.history.end();
  }
  refreshWorkbookCells();
  (sheetEl as HTMLElement).focus();
};

type RibbonFillDirection = 'down' | 'right' | 'up' | 'left';
type RibbonFillSeriesMode = 'auto' | 'copy' | 'days' | 'weekdays' | 'months' | 'years';

const fillSeriesSourceRange = (range: Range, direction: RibbonFillDirection): Range => {
  if (direction === 'down') return { ...range, r1: range.r0 };
  if (direction === 'up') return { ...range, r0: range.r1 };
  if (direction === 'right') return { ...range, c1: range.c0 };
  return { ...range, c0: range.c1 };
};

const runFillSeries = (
  range: Range,
  direction: RibbonFillDirection,
  mode: RibbonFillSeriesMode,
): void => {
  const i = inst;
  if (!i) return;
  const src = fillSeriesSourceRange(range, direction);
  if (src.r0 === range.r0 && src.r1 === range.r1 && src.c0 === range.c0 && src.c1 === range.c1) {
    return;
  }
  const dateUnit =
    mode === 'days' || mode === 'weekdays' || mode === 'months' || mode === 'years'
      ? mode
      : undefined;
  i.history.begin();
  try {
    recordFormatChange(i.history, i.store, () => {
      fillRange(i.store.getState(), i.workbook, src, range, {
        copyOnly: mode === 'copy',
        dateUnit,
        formatting: 'with',
        store: i.store,
      });
    });
  } finally {
    i.history.end();
  }
  refreshWorkbookCells();
  (sheetEl as HTMLElement).focus();
};

const inferFillSeriesDirection = (range: Range): RibbonFillDirection =>
  range.r1 > range.r0 ? 'down' : 'right';

const makeFillSeriesRadio = <T extends string>(
  name: string,
  value: T,
  label: string,
  checked: boolean,
): HTMLLabelElement => {
  const wrap = document.createElement('label');
  wrap.className = 'fc-fmtdlg__radio';
  const input = document.createElement('input');
  input.type = 'radio';
  input.name = name;
  input.value = value;
  input.checked = checked;
  const span = document.createElement('span');
  span.textContent = label;
  wrap.append(input, span);
  return wrap;
};

const selectedFillSeriesRadio = <T extends string>(
  root: HTMLElement,
  name: string,
  fallback: T,
): T =>
  (root.querySelector<HTMLInputElement>(`input[name="${name}"]:checked`)?.value as T | undefined) ??
  fallback;

const showFillSeriesDialog = (
  range: Range,
): Promise<{ direction: RibbonFillDirection; mode: RibbonFillSeriesMode } | null> => {
  return new Promise((resolve) => {
    const ja = ribbonLang === 'ja';
    const title = ja ? '連続データ' : 'Series';
    const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    const overlay = document.createElement('div');
    overlay.className = 'fc-fmtdlg app__dlg';
    overlay.setAttribute('role', 'dialog');
    overlay.setAttribute('aria-modal', 'true');
    overlay.setAttribute('aria-label', title);

    const panel = document.createElement('div');
    panel.className = 'fc-fmtdlg__panel app__dlg__panel fc-pastesp__panel';
    overlay.appendChild(panel);

    const header = document.createElement('div');
    header.className = 'fc-fmtdlg__header';
    header.textContent = title;
    panel.appendChild(header);

    const body = document.createElement('div');
    body.className = 'fc-fmtdlg__body fc-pastesp__body';
    panel.appendChild(body);

    const cols = document.createElement('div');
    cols.className = 'fc-pastesp__cols';
    body.appendChild(cols);

    const dirName = `app-fill-series-dir-${Math.random().toString(36).slice(2)}`;
    const dirGroup = document.createElement('div');
    dirGroup.className = 'fc-pastesp__group';
    const dirLegend = document.createElement('div');
    dirLegend.className = 'fc-pastesp__legend';
    dirLegend.textContent = ja ? '範囲' : 'Series in';
    const dirList = document.createElement('div');
    dirList.className = 'fc-pastesp__list';
    dirList.setAttribute('role', 'radiogroup');
    dirList.setAttribute('aria-label', dirLegend.textContent);
    const initialDirection = inferFillSeriesDirection(range);
    const directionOptions: Array<{ value: RibbonFillDirection; label: string }> = [
      { value: 'down', label: ja ? '列' : 'Columns' },
      { value: 'right', label: ja ? '行' : 'Rows' },
      { value: 'up', label: ja ? '上方向' : 'Up' },
      { value: 'left', label: ja ? '左方向' : 'Left' },
    ];
    for (const option of directionOptions) {
      dirList.appendChild(
        makeFillSeriesRadio(dirName, option.value, option.label, option.value === initialDirection),
      );
    }
    dirGroup.append(dirLegend, dirList);
    cols.appendChild(dirGroup);

    const modeName = `app-fill-series-mode-${Math.random().toString(36).slice(2)}`;
    const modeGroup = document.createElement('div');
    modeGroup.className = 'fc-pastesp__group';
    const modeLegend = document.createElement('div');
    modeLegend.className = 'fc-pastesp__legend';
    modeLegend.textContent = ja ? '種類' : 'Type';
    const modeList = document.createElement('div');
    modeList.className = 'fc-pastesp__list';
    modeList.setAttribute('role', 'radiogroup');
    modeList.setAttribute('aria-label', modeLegend.textContent);
    const modeOptions: Array<{ value: RibbonFillSeriesMode; label: string }> = [
      { value: 'auto', label: ja ? 'オートフィル' : 'AutoFill' },
      { value: 'copy', label: ja ? 'コピー' : 'Copy' },
      { value: 'days', label: ja ? '日' : 'Day' },
      { value: 'weekdays', label: ja ? '週日' : 'Weekday' },
      { value: 'months', label: ja ? '月' : 'Month' },
      { value: 'years', label: ja ? '年' : 'Year' },
    ];
    for (const option of modeOptions) {
      modeList.appendChild(
        makeFillSeriesRadio(modeName, option.value, option.label, option.value === 'auto'),
      );
    }
    modeGroup.append(modeLegend, modeList);
    cols.appendChild(modeGroup);

    const footer = document.createElement('div');
    footer.className = 'fc-fmtdlg__footer';
    panel.appendChild(footer);
    const cancelBtn = document.createElement('button');
    cancelBtn.type = 'button';
    cancelBtn.className = 'fc-fmtdlg__btn';
    cancelBtn.textContent = ja ? 'キャンセル' : 'Cancel';
    const okBtn = document.createElement('button');
    okBtn.type = 'button';
    okBtn.className = 'fc-fmtdlg__btn fc-fmtdlg__btn--primary';
    okBtn.textContent = 'OK';
    footer.append(cancelBtn, okBtn);

    let done = false;
    const finish = (
      value: { direction: RibbonFillDirection; mode: RibbonFillSeriesMode } | null,
    ): void => {
      if (done) return;
      done = true;
      overlay.removeEventListener('keydown', onKey);
      overlay.remove();
      opener?.focus({ preventScroll: true });
      resolve(value);
    };
    const apply = (): void => {
      finish({
        direction: selectedFillSeriesRadio<RibbonFillDirection>(overlay, dirName, initialDirection),
        mode: selectedFillSeriesRadio<RibbonFillSeriesMode>(overlay, modeName, 'auto'),
      });
    };
    const onKey = (event: KeyboardEvent): void => {
      event.stopPropagation();
      if (event.key === 'Escape') {
        event.preventDefault();
        finish(null);
      } else if (event.key === 'Enter') {
        event.preventDefault();
        apply();
      }
    };
    cancelBtn.addEventListener('click', () => finish(null));
    okBtn.addEventListener('click', apply);
    overlay.addEventListener('keydown', onKey);
    overlay.addEventListener('click', (event) => {
      if (event.target === overlay) finish(null);
    });
    document.body.appendChild(overlay);
    requestAnimationFrame(() => {
      dirList.querySelector<HTMLInputElement>('input[type="radio"]')?.focus();
    });
  });
};

const applyFillSeries = async (modeOverride?: RibbonFillSeriesMode): Promise<void> => {
  const i = inst;
  const range = normalizedSelectionRange();
  if (!i || !range) return;
  if (modeOverride) {
    runFillSeries(range, inferFillSeriesDirection(range), modeOverride);
    return;
  }
  const choice = await showFillSeriesDialog(range);
  if (!choice) return;
  runFillSeries(range, choice.direction, choice.mode);
};

const applyClearAction = (action: string): void => {
  const i = inst;
  const range = normalizedSelectionRange();
  if (!i || !range) return;
  if (action === 'contents') {
    clearSelectionContents();
    return;
  }
  if (action === 'formats') {
    applyRibbonFormat((s, store) => clearVisualFormat(s, store));
    return;
  }
  if (action === 'conditional') {
    recordConditionalRulesChange(i.history, i.store, () => {
      mutators.clearConditionalRulesInRange(i.store, range);
    });
    refreshWorkbookCells();
    (sheetEl as HTMLElement).focus();
    return;
  }
  if (action === 'comments') {
    const addrs: Array<{ sheet: number; row: number; col: number }> = [];
    for (let row = range.r0; row <= range.r1; row += 1) {
      for (let col = range.c0; col <= range.c1; col += 1) {
        addrs.push({ sheet: range.sheet, row, col });
      }
    }
    recordCommentChange(i.history, i.store, i.workbook, addrs, () => {
      for (let row = range.r0; row <= range.r1; row += 1) {
        for (let col = range.c0; col <= range.c1; col += 1) {
          clearComment(i.store, { sheet: range.sheet, row, col }, i.workbook);
        }
      }
    });
    refreshWorkbookCells();
    (sheetEl as HTMLElement).focus();
    return;
  }
  if (action === 'hyperlinks') {
    clearHyperlinksInSelection('clear');
    return;
  }
  if (action === 'remove-hyperlinks') {
    clearHyperlinksInSelection('remove');
    return;
  }
  if (action === 'all') {
    i.history.begin();
    try {
      recordFormatChange(i.history, i.store, () => {
        clearFormat(i.store.getState(), i.store);
      });
      recordConditionalRulesChange(i.history, i.store, () => {
        mutators.clearConditionalRulesInRange(i.store, range);
      });
      for (let row = range.r0; row <= range.r1; row += 1) {
        for (let col = range.c0; col <= range.c1; col += 1) {
          i.workbook.setBlank({ sheet: range.sheet, row, col });
        }
      }
    } finally {
      i.history.end();
    }
    refreshWorkbookCells();
    (sheetEl as HTMLElement).focus();
  }
};

const promptDimension = async (
  title: string,
  label: string,
  initial: number,
  max: number,
): Promise<number | null> => {
  return showNumberPrompt({
    title,
    label,
    initial,
    min: 1,
    max,
    step: 1,
    okLabel: 'OK',
    cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
    invalidMessage:
      ribbonLang === 'ja'
        ? `1 から ${max} までの数値を入力してください`
        : `Enter a number from 1 to ${max}.`,
  });
};

// Off-screen autofit measurement lives in ribbon/autofit.ts. The two
// exported helpers take the SpreadsheetInstance + locale so they can be
// reused outside this file (and so they don't capture the playground's
// closure-scoped state).

const applyCellInsertAction = async (action: string): Promise<void> => {
  const i = inst;
  const range = normalizedSelectionRange();
  if (!i || !range) return;
  if (action === 'rows') {
    insertRows(i.store, i.workbook, i.history, range.r0, selectedRowCount());
  } else if (action === 'cols') {
    insertCols(i.store, i.workbook, i.history, range.c0, selectedColCount());
  } else if (action === 'sheet') {
    const added = addSheet(i.store, i.workbook, i.history);
    if (added >= 0) {
      renderSheetTabs();
      switchSheet(added);
    } else if (statusMetric && isWorkbookStructureProtected(i.store.getState())) {
      statusMetric.textContent = ribbonMenuText.workbookStructureProtectedBlocked;
    }
  } else if (action === 'shift-down' || action === 'shift-right') {
    insertCells(i.store, i.workbook, i.history, range, action === 'shift-down' ? 'down' : 'right');
  } else {
    const choice = await showChoiceDialog<'down' | 'right'>({
      title: ribbonLang === 'ja' ? 'セルを挿入' : 'Insert Cells',
      label: ribbonLang === 'ja' ? '挿入後のセルの移動方向' : 'Shift cells',
      initial: 'down',
      okLabel: 'OK',
      cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
      options: [
        { value: 'down', label: ribbonLang === 'ja' ? '下方向にシフト' : 'Shift cells down' },
        { value: 'right', label: ribbonLang === 'ja' ? '右方向にシフト' : 'Shift cells right' },
      ],
    });
    if (choice === 'down' || choice === 'right') {
      insertCells(i.store, i.workbook, i.history, range, choice);
    }
  }
  refreshWorkbookCells();
  (sheetEl as HTMLElement).focus();
};

const applyCellDeleteAction = async (action: string): Promise<void> => {
  const i = inst;
  const range = normalizedSelectionRange();
  if (!i || !range) return;
  if (action === 'rows') {
    deleteRows(i.store, i.workbook, i.history, range.r0, selectedRowCount());
  } else if (action === 'cols') {
    deleteCols(i.store, i.workbook, i.history, range.c0, selectedColCount());
  } else if (action === 'sheet') {
    const before = i.store.getState().data.sheetIndex;
    if (removeSheet(i.store, i.workbook, before)) {
      renderSheetTabs();
      switchSheet(i.store.getState().data.sheetIndex);
    }
  } else if (action === 'shift-up' || action === 'shift-left') {
    deleteCells(i.store, i.workbook, i.history, range, action === 'shift-up' ? 'up' : 'left');
  } else {
    const choice = await showChoiceDialog<'up' | 'left'>({
      title: ribbonLang === 'ja' ? 'セルを削除' : 'Delete Cells',
      label: ribbonLang === 'ja' ? '削除後のセルの移動方向' : 'Shift cells',
      initial: 'up',
      okLabel: 'OK',
      cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
      options: [
        { value: 'up', label: ribbonLang === 'ja' ? '上方向にシフト' : 'Shift cells up' },
        { value: 'left', label: ribbonLang === 'ja' ? '左方向にシフト' : 'Shift cells left' },
      ],
    });
    if (choice === 'up' || choice === 'left') {
      deleteCells(i.store, i.workbook, i.history, range, choice);
    }
  }
  refreshWorkbookCells();
  (sheetEl as HTMLElement).focus();
};

const sheetTabColorByAction = (action: string): string | null | undefined => {
  const colors: Record<string, string | null> = {
    'tab-color-none': null,
    'tab-color-red': '#c00000',
    'tab-color-orange': '#ed7d31',
    'tab-color-yellow': '#ffc000',
    'tab-color-green': '#70ad47',
    'tab-color-blue': '#4472c4',
    'tab-color-purple': '#7030a0',
    'tab-color-gray': '#808080',
  };
  return Object.hasOwn(colors, action) ? colors[action] : undefined;
};

const applyCellFormatAction = async (action: string): Promise<void> => {
  const i = inst;
  const range = normalizedSelectionRange();
  if (!i || !range) return;
  if (action === 'dialog') {
    i.openFormatDialog();
    return;
  }
  if (action === 'protect-sheet') {
    await runSheetProtectionFlow();
    return;
  }
  if (action === 'rename-sheet') {
    const sheet = i.store.getState().data.sheetIndex;
    const current = i.workbook.sheetName(sheet);
    const name = await showPrompt({
      title: dictionaries[ribbonLang].sheetTabs.rename,
      label: ribbonLang === 'ja' ? 'シート名' : 'Sheet name',
      initial: current,
      validate: (raw) =>
        raw.trim()
          ? null
          : ribbonLang === 'ja'
            ? 'シート名を入力してください。'
            : 'Enter a sheet name.',
    });
    if (name !== null && renameSheet(i.workbook, sheet, name.trim(), i.store, i.history)) {
      renderSheetTabs();
    }
    return;
  }
  if (action === 'move-sheet-left' || action === 'move-sheet-right') {
    const sheet = i.store.getState().data.sheetIndex;
    const target = action === 'move-sheet-left' ? sheet - 1 : sheet + 1;
    if (target >= 0 && target < i.workbook.sheetCount) {
      moveSheet(i.store, i.workbook, sheet, target, i.history);
      renderSheetTabs();
    }
    return;
  }
  if (action === 'hide-sheet') {
    const sheet = i.store.getState().data.sheetIndex;
    if (setSheetHidden(i.store, i.workbook, i.history, sheet, true)) {
      renderSheetTabs();
      switchSheet(i.store.getState().data.sheetIndex);
      refreshWorkbookCells();
    }
    return;
  }
  if (action === 'unhide-sheet') {
    const hidden = [...i.store.getState().layout.hiddenSheets].sort((a, b) => a - b)[0];
    if (hidden != null && setSheetHidden(i.store, i.workbook, i.history, hidden, false)) {
      renderSheetTabs();
    }
    return;
  }
  const tabColor = sheetTabColorByAction(action);
  if (tabColor !== undefined) {
    const sheet = i.store.getState().data.sheetIndex;
    recordLayoutChange(i.history, i.store, () => {
      mutators.setSheetTabColor(i.store, sheet, tabColor);
    });
    renderSheetTabs();
    return;
  }
  if (action === 'lock-cell' || action === 'unlock-cell') {
    const locked = action === 'lock-cell';
    recordFormatChange(i.history, i.store, () => {
      setCellLocked(i.store, range, locked);
    });
    if (statusMetric) {
      statusMetric.textContent = locked
        ? ribbonMenuText.cellsLockedStatus
        : ribbonMenuText.cellsUnlockedStatus;
    }
    projectFormatToolbar();
    focusSheet();
    return;
  }
  if (action === 'hide-rows') {
    hideRows(i.store, i.history, range.r0, range.r1, i.workbook);
    return;
  }
  if (action === 'hide-cols') {
    hideCols(i.store, i.history, range.c0, range.c1, i.workbook);
    return;
  }
  if (action === 'show-rows') {
    const targets = hiddenInSelection(i.store.getState().layout, 'row', range.r0, range.r1);
    if (targets.length > 0)
      showRows(i.store, i.history, targets[0] ?? range.r0, targets.at(-1) ?? range.r1, i.workbook);
    return;
  }
  if (action === 'show-cols') {
    const targets = hiddenInSelection(i.store.getState().layout, 'col', range.c0, range.c1);
    if (targets.length > 0)
      showCols(i.store, i.history, targets[0] ?? range.c0, targets.at(-1) ?? range.c1, i.workbook);
    return;
  }
  if (action === 'row-height') {
    const n = await promptDimension(
      ribbonLang === 'ja' ? '行の高さ' : 'Row Height',
      ribbonLang === 'ja' ? '高さ (px)' : 'Height (px)',
      i.store.getState().layout.defaultRowHeight,
      409,
    );
    if (n === null) return;
    recordLayoutChange(i.history, i.store, () => {
      for (let row = range.r0; row <= range.r1; row += 1) {
        mutators.setRowHeight(i.store, row, n);
        i.workbook.setRowHeight(range.sheet, row, n);
      }
    });
    return;
  }
  if (action === 'row-autofit') {
    recordLayoutChange(i.history, i.store, () => {
      for (let row = range.r0; row <= range.r1; row += 1) {
        const height = autofitRowHeight(i, row, range.c0, range.c1, ribbonLang);
        mutators.setRowHeight(i.store, row, height);
        i.workbook.setRowHeight(range.sheet, row, height);
      }
    });
    return;
  }
  if (action === 'col-width') {
    const n = await promptDimension(
      ribbonLang === 'ja' ? '列の幅' : 'Column Width',
      ribbonLang === 'ja' ? '幅 (px)' : 'Width (px)',
      i.store.getState().layout.defaultColWidth,
      2048,
    );
    if (n === null) return;
    recordLayoutChange(i.history, i.store, () => {
      for (let col = range.c0; col <= range.c1; col += 1) {
        mutators.setColWidth(i.store, col, n);
        i.workbook.setColumnWidth(range.sheet, col, col, n);
      }
    });
  }
  if (action === 'col-autofit') {
    recordLayoutChange(i.history, i.store, () => {
      for (let col = range.c0; col <= range.c1; col += 1) {
        const width = autofitColWidth(i, col, range.r0, range.r1, ribbonLang);
        mutators.setColWidth(i.store, col, width);
        i.workbook.setColumnWidth(range.sheet, col, col, width);
      }
    });
  }
};

const applyTextOrientationAction = (action: string): void => {
  const i = inst;
  if (!i) return;
  if (action === 'format') {
    i.openFormatDialog();
    return;
  }
  const rotations: Record<string, number> = {
    horizontal: 0,
    ccw: 45,
    cw: -45,
    vertical: 90,
    up: 90,
    down: -90,
  };
  const rotation = rotations[action];
  if (typeof rotation !== 'number') return;
  applyRibbonFormat((s, store) => setRotation(s, store, rotation));
};

const addConditionalRuleFromRibbon = (rule: ConditionalRule): void => {
  const i = inst;
  if (!i) return;
  recordConditionalRulesChange(i.history, i.store, () => {
    mutators.addConditionalRule(i.store, rule);
  });
  refreshWorkbookCells();
  (sheetEl as HTMLElement).focus();
};

const promptCfNumber = async (
  title: string,
  initial = 0,
  options: { min?: number; max?: number; step?: number } = {},
): Promise<number | null> => {
  return showNumberPrompt({
    title,
    label: ribbonLang === 'ja' ? '値' : 'Value',
    initial,
    min: options.min,
    max: options.max,
    step: options.step,
    cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
    invalidMessage: ribbonLang === 'ja' ? '数値を入力してください' : 'Enter a number',
  });
};

const promptCfText = async (title: string, label: string, initial = ''): Promise<string | null> => {
  const value = await showPrompt({
    title,
    label,
    initial,
    validate: (raw) =>
      raw.trim() ? null : ribbonLang === 'ja' ? '値を入力してください' : 'Enter a value',
  });
  return value === null ? null : value.trim();
};

const DATE_PERIODS = [
  'yesterday',
  'today',
  'tomorrow',
  'last7',
  'last-week',
  'this-week',
  'next-week',
  'last-month',
  'this-month',
  'next-month',
] as const;

type DatePeriod = (typeof DATE_PERIODS)[number];

const isDatePeriod = (value: string): value is DatePeriod =>
  DATE_PERIODS.includes(value as DatePeriod);

const cfDatePeriodOptions = (): Array<{ value: DatePeriod; label: string }> =>
  ribbonLang === 'ja'
    ? [
        { value: 'yesterday', label: '昨日' },
        { value: 'today', label: '今日' },
        { value: 'tomorrow', label: '明日' },
        { value: 'last7', label: '過去 7 日間' },
        { value: 'last-week', label: '先週' },
        { value: 'this-week', label: '今週' },
        { value: 'next-week', label: '来週' },
        { value: 'last-month', label: '先月' },
        { value: 'this-month', label: '今月' },
        { value: 'next-month', label: '来月' },
      ]
    : [
        { value: 'yesterday', label: 'Yesterday' },
        { value: 'today', label: 'Today' },
        { value: 'tomorrow', label: 'Tomorrow' },
        { value: 'last7', label: 'In the last 7 days' },
        { value: 'last-week', label: 'Last week' },
        { value: 'this-week', label: 'This week' },
        { value: 'next-week', label: 'Next week' },
        { value: 'last-month', label: 'Last month' },
        { value: 'this-month', label: 'This month' },
        { value: 'next-month', label: 'Next month' },
      ];

const cfFill = { fill: '#ffc7ce', color: '#9c0006' } as const;
const cfTopFill = { fill: '#c6efce', color: '#006100' } as const;

const conditionalPresetActions = new Set<ConditionalPresetAction>([
  'clear-selection',
  'clear-sheet',
  'duplicates',
  'unique',
  'above-avg',
  'below-avg',
  'data-blue',
  'data-green',
  'data-red',
  'data-orange',
  'data-purple',
  'data-teal',
  'data-solid-blue',
  'data-solid-green',
  'data-solid-red',
  'data-solid-orange',
  'data-solid-purple',
  'data-solid-gray',
  'scale-gyr',
  'scale-ryg',
  'scale-gw',
  'scale-rw',
  'scale-bwr',
  'scale-rwb',
  'scale-gwg',
  'scale-ywg',
  'scale-rwr',
  'scale-bwb',
  'scale-yry',
  'scale-gyg',
  'icons-arrows3',
  'icons-arrows5',
  'icons-triangles3',
  'icons-traffic3',
  'icons-trafficRim3',
  'icons-symbols3',
  'icons-flags3',
  'icons-stars3',
  'icons-quarters5',
  'icons-ratings5',
  'icons-bars5',
  'icons-boxes5',
]);

const isConditionalPresetAction = (action: string): action is ConditionalPresetAction =>
  conditionalPresetActions.has(action as ConditionalPresetAction);

const conditionalRuleKindForPanel = (
  panel: string | undefined,
): ConditionalDialogOpenOptions['kind'] | undefined => {
  if (panel === 'dataBar') return 'data-bar';
  if (panel === 'colorScale') return 'color-scale';
  if (panel === 'iconSet') return 'icon-set';
  if (panel === 'topBottom') return 'top-bottom';
  if (panel === 'highlight') return 'cell-value';
  return undefined;
};

const applyConditionalMenuAction = async (action: string, panel?: string): Promise<void> => {
  const i = inst;
  const range = cfSelectionRange();
  if (!i || !range) return;
  const title = cfMenuText();
  if (action === 'new-rule') {
    i.openConditionalDialog({ mode: 'new', kind: conditionalRuleKindForPanel(panel) });
    return;
  }
  if (action === 'manage') {
    i.openCfRulesDialog();
    return;
  }
  if (action === 'cell-gt' || action === 'cell-lt' || action === 'cell-eq') {
    const n = await promptCfNumber(
      action === 'cell-gt' ? title.greater : action === 'cell-lt' ? title.less : title.equal,
      0,
    );
    if (n === null) return;
    addConditionalRuleFromRibbon({
      kind: 'cell-value',
      range,
      op: action === 'cell-gt' ? '>' : action === 'cell-lt' ? '<' : '=',
      a: n,
      apply: cfFill,
    });
    return;
  }
  if (action === 'cell-between') {
    const a = await promptCfNumber(title.between, 0);
    if (a === null) return;
    const b = await promptCfNumber(title.between, 100);
    if (b === null) return;
    addConditionalRuleFromRibbon({
      kind: 'cell-value',
      range,
      op: 'between',
      a: Math.min(a, b),
      b: Math.max(a, b),
      apply: cfFill,
    });
    return;
  }
  if (action === 'text-contains') {
    const text = await promptCfText(title.text, title.textPrompt);
    if (text === null) return;
    addConditionalRuleFromRibbon({ kind: 'text-contains', range, text, apply: cfFill });
    return;
  }
  if (action === 'date-occurring') {
    const period = await showChoiceDialog<DatePeriod>({
      title: title.date,
      label: title.datePrompt,
      options: cfDatePeriodOptions(),
      initial: 'today',
      cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
    });
    if (period === null) return;
    if (!isDatePeriod(period)) {
      void showMessage({
        title: title.date,
        message:
          ribbonLang === 'ja'
            ? '指定できる日付条件を入力してください。'
            : 'Enter one of the supported date conditions.',
      });
      return;
    }
    addConditionalRuleFromRibbon({ kind: 'date-occurring', range, period, apply: cfFill });
    return;
  }
  if (
    action === 'top10' ||
    action === 'bottom10' ||
    action === 'top10-percent' ||
    action === 'bottom10-percent'
  ) {
    const isPercent = action.endsWith('-percent');
    const n = await promptCfNumber(
      action.startsWith('top')
        ? isPercent
          ? title.top10Percent
          : title.top10
        : isPercent
          ? title.bottom10Percent
          : title.bottom10,
      10,
      { min: 1, max: isPercent ? 100 : undefined, step: 1 },
    );
    if (n === null) return;
    addConditionalRuleFromRibbon({
      kind: 'top-bottom',
      range,
      mode: action.startsWith('top') ? 'top' : 'bottom',
      n: Math.max(1, Math.floor(n)),
      percent: isPercent,
      apply: cfTopFill,
    });
    return;
  }
  if (isConditionalPresetAction(action)) {
    let changed = false;
    recordConditionalRulesChange(i.history, i.store, () => {
      changed = applyConditionalPresetAction(i.store, action, range);
    });
    if (changed) refreshWorkbookCells();
  }
};

applyShellLocale();
renderRibbon();

const pasteBtn = document.querySelector<HTMLButtonElement>('button[data-ribbon-command="paste"]');
const pasteMenu = document.getElementById('menu-paste') as HTMLDivElement | null;
pasteBtn?.setAttribute('aria-haspopup', 'menu');
pasteBtn?.setAttribute('aria-expanded', 'false');
const closePasteMenu = (restoreFocus = false): void => {
  if (!pasteMenu) return;
  pasteMenu.hidden = true;
  pasteBtn?.setAttribute('aria-expanded', 'false');
  if (restoreFocus) pasteBtn?.focus();
};
const openPasteMenu = (): void => {
  if (!pasteMenu || !pasteBtn) return;
  closeBorderMenu();
  closeFreezeMenu();
  closeConditionalMenu();
  closeTextOrientationMenu();
  closeFillMenu();
  closeClearMenu();
  closeCellsMenus();
  closeSortFilterHomeMenu();
  closeFindSelectMenu();
  pasteMenu.hidden = false;
  pasteBtn.setAttribute('aria-haspopup', 'menu');
  pasteBtn.setAttribute('aria-expanded', 'true');
  focusMenuItem(pasteMenu, 'first');
};
pasteBtn?.addEventListener('click', (event) => {
  event.preventDefault();
  event.stopPropagation();
  const target = event.target as Element | null;
  if (event.altKey || event.shiftKey || target?.closest('.demo__rb-split-chevron')) {
    if (!pasteMenu) return;
    if (pasteMenu.hidden) openPasteMenu();
    else closePasteMenu(true);
    return;
  }
  closePasteMenu();
  void pasteClipboardIntoSelection();
});
pasteBtn?.addEventListener('keydown', (event) => {
  if (event.key !== 'ArrowDown') return;
  event.preventDefault();
  event.stopPropagation();
  if (pasteMenu?.hidden) openPasteMenu();
});
pasteMenu?.addEventListener('click', (event) => {
  const item = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-paste-action]');
  const action = item?.dataset.pasteAction;
  if (!action) return;
  event.preventDefault();
  event.stopPropagation();
  closePasteMenu();
  void applyRibbonPasteAction(action);
});
pasteMenu?.addEventListener('keydown', (event) => {
  handleMenuKeydown(event, pasteMenu, { close: closePasteMenu, restoreFocusTo: pasteBtn });
});
document.addEventListener('mousedown', (event) => {
  const target = event.target as Node | null;
  if (
    pasteMenu?.hidden === false &&
    target &&
    !pasteMenu.contains(target) &&
    !pasteBtn?.contains(target)
  ) {
    closePasteMenu();
  }
});
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && pasteMenu?.hidden === false) closePasteMenu(true);
});

const conditionalBtn = document.querySelector<HTMLButtonElement>(
  'button[data-ribbon-command="conditional"]',
);
const conditionalMenu = document.getElementById('menu-conditional') as HTMLDivElement | null;
let conditionalSubmenuCloseTimer: number | null = null;
const cancelConditionalSubmenuClose = (): void => {
  if (conditionalSubmenuCloseTimer === null) return;
  window.clearTimeout(conditionalSubmenuCloseTimer);
  conditionalSubmenuCloseTimer = null;
};
const scheduleConditionalSubmenuClose = (): void => {
  cancelConditionalSubmenuClose();
  conditionalSubmenuCloseTimer = window.setTimeout(() => {
    closeConditionalSubmenus();
  }, 180);
};
const closeConditionalSubmenus = (): void => {
  cancelConditionalSubmenuClose();
  conditionalMenu?.querySelectorAll<HTMLElement>('[data-cf-panel]').forEach((panel) => {
    panel.hidden = true;
  });
  conditionalMenu?.querySelectorAll<HTMLElement>('[data-cf-submenu]').forEach((trigger) => {
    trigger.classList.remove('app__menu-item--active');
  });
};
const closeConditionalMenu = (restoreFocus = false): void => {
  if (!conditionalMenu) return;
  conditionalMenu.hidden = true;
  closeConditionalSubmenus();
  conditionalBtn?.setAttribute('aria-expanded', 'false');
  if (restoreFocus) conditionalBtn?.focus();
};
const openConditionalMenu = (): void => {
  if (!conditionalMenu || !conditionalBtn) return;
  closeBorderMenu();
  closeFreezeMenu();
  closeTextOrientationMenu();
  closeFillMenu();
  closeClearMenu();
  conditionalMenu.hidden = false;
  conditionalBtn.setAttribute('aria-haspopup', 'menu');
  conditionalBtn.setAttribute('aria-expanded', 'true');
  focusMenuItem(conditionalMenu, 'first');
};
const openConditionalSubmenu = (key: string, trigger: HTMLElement): void => {
  if (!conditionalMenu) return;
  cancelConditionalSubmenuClose();
  closeConditionalSubmenus();
  const panel = conditionalMenu.querySelector<HTMLElement>(`[data-cf-panel="${key}"]`);
  if (!panel) return;
  const menuRect = conditionalMenu.getBoundingClientRect();
  const triggerRect = trigger.getBoundingClientRect();
  panel.style.top = `${Math.max(0, triggerRect.top - menuRect.top - 4)}px`;
  panel.hidden = false;
  trigger.classList.add('app__menu-item--active');
};

conditionalBtn?.addEventListener('click', (event) => {
  event.preventDefault();
  event.stopPropagation();
  if (!conditionalMenu) return;
  if (conditionalMenu.hidden) openConditionalMenu();
  else closeConditionalMenu(true);
});
conditionalMenu?.addEventListener('click', (event) => {
  const target = event.target as Element | null;
  const submenu = target?.closest<HTMLElement>('[data-cf-submenu]');
  if (submenu) {
    event.preventDefault();
    event.stopPropagation();
    openConditionalSubmenu(submenu.dataset.cfSubmenu ?? '', submenu);
    return;
  }
  const item = target?.closest<HTMLButtonElement>('[data-cf-action]');
  const action = item?.dataset.cfAction;
  if (!item || !action || action.startsWith('submenu-')) return;
  event.preventDefault();
  event.stopPropagation();
  const panel = item.closest<HTMLElement>('[data-cf-panel]')?.dataset.cfPanel;
  closeConditionalMenu();
  void applyConditionalMenuAction(action, panel);
});
conditionalMenu?.querySelectorAll<HTMLElement>('[data-cf-submenu]').forEach((trigger) => {
  trigger.addEventListener('mouseenter', () =>
    openConditionalSubmenu(trigger.dataset.cfSubmenu ?? '', trigger),
  );
});
conditionalMenu
  ?.querySelectorAll<HTMLElement>('.app__menu-item:not([data-cf-submenu])')
  .forEach((item) => {
    item.addEventListener('mouseenter', scheduleConditionalSubmenuClose);
  });
conditionalMenu?.querySelectorAll<HTMLElement>('[data-cf-panel]').forEach((panel) => {
  panel.addEventListener('mouseenter', cancelConditionalSubmenuClose);
  panel.addEventListener('mouseleave', scheduleConditionalSubmenuClose);
});
document.addEventListener('click', (event) => {
  if (conditionalMenu?.hidden !== false) return;
  const target = event.target as Node | null;
  if (target && (conditionalMenu.contains(target) || conditionalBtn?.contains(target))) return;
  closeConditionalMenu();
});
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && conditionalMenu?.hidden === false) closeConditionalMenu(true);
});

const fillBtn = document.querySelector<HTMLButtonElement>('button[data-ribbon-command="fillHome"]');
const fillMenu = document.getElementById('menu-fill') as HTMLDivElement | null;
const closeFillMenu = (restoreFocus = false): void => {
  if (!fillMenu) return;
  fillMenu.hidden = true;
  fillBtn?.setAttribute('aria-expanded', 'false');
  if (restoreFocus) fillBtn?.focus();
};
const openFillMenu = (): void => {
  if (!fillMenu || !fillBtn) return;
  closeBorderMenu();
  closeFreezeMenu();
  closeConditionalMenu();
  closeTextOrientationMenu();
  closeClearMenu();
  fillMenu.hidden = false;
  fillBtn.setAttribute('aria-haspopup', 'menu');
  fillBtn.setAttribute('aria-expanded', 'true');
  focusMenuItem(fillMenu, 'first');
};
fillBtn?.addEventListener('click', (event) => {
  event.preventDefault();
  event.stopPropagation();
  if (!fillMenu) return;
  if (fillMenu.hidden) openFillMenu();
  else closeFillMenu(true);
});
fillMenu?.addEventListener('click', (event) => {
  const item = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-fill]');
  const action = item?.dataset.fill;
  if (!action) return;
  event.preventDefault();
  event.stopPropagation();
  closeFillMenu();
  if (action === 'series') {
    void applyFillSeries();
    return;
  }
  if (action === 'days' || action === 'weekdays' || action === 'months' || action === 'years') {
    void applyFillSeries(action);
    return;
  }
  applyFillDirection(action as 'down' | 'right' | 'up' | 'left');
});
fillMenu?.addEventListener('keydown', (event) => {
  handleMenuKeydown(event, fillMenu, { close: closeFillMenu, restoreFocusTo: fillBtn });
});

const clearBtn = document.querySelector<HTMLButtonElement>(
  '.demo__ribbon-group--editing button[data-ribbon-command="clearFormat"]',
);
const clearMenu = document.getElementById('menu-clear') as HTMLDivElement | null;
const closeClearMenu = (restoreFocus = false): void => {
  if (!clearMenu) return;
  clearMenu.hidden = true;
  clearBtn?.setAttribute('aria-expanded', 'false');
  if (restoreFocus) clearBtn?.focus();
};
const openClearMenu = (): void => {
  if (!clearMenu || !clearBtn) return;
  closeBorderMenu();
  closeFreezeMenu();
  closeConditionalMenu();
  closeTextOrientationMenu();
  closeFillMenu();
  clearMenu.hidden = false;
  clearBtn.setAttribute('aria-haspopup', 'menu');
  clearBtn.setAttribute('aria-expanded', 'true');
  focusMenuItem(clearMenu, 'first');
};
clearBtn?.addEventListener('click', (event) => {
  event.preventDefault();
  event.stopPropagation();
  if (!clearMenu) return;
  if (clearMenu.hidden) openClearMenu();
  else closeClearMenu(true);
});
clearMenu?.addEventListener('click', (event) => {
  const item = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-clear]');
  const action = item?.dataset.clear;
  if (!action) return;
  event.preventDefault();
  event.stopPropagation();
  closeClearMenu();
  applyClearAction(action);
});
clearMenu?.addEventListener('keydown', (event) => {
  handleMenuKeydown(event, clearMenu, { close: closeClearMenu, restoreFocusTo: clearBtn });
});

const getPrintAreaBtn = (): HTMLButtonElement | null =>
  document.querySelector<HTMLButtonElement>('button[data-ribbon-command="printArea"]');
const getPrintAreaMenu = (): HTMLDivElement | null =>
  document.getElementById('menu-print-area') as HTMLDivElement | null;
const getSymbolBtn = (): HTMLButtonElement | null =>
  document.querySelector<HTMLButtonElement>('button[data-ribbon-command="symbolInsert"]');
const getSymbolMenu = (): HTMLDivElement | null =>
  document.getElementById('menu-symbol') as HTMLDivElement | null;
const closePrintAreaMenu = (restoreFocus = false): void => {
  const printAreaMenu = getPrintAreaMenu();
  const printAreaBtn = getPrintAreaBtn();
  if (!printAreaMenu) return;
  printAreaMenu.hidden = true;
  printAreaBtn?.setAttribute('aria-expanded', 'false');
  if (restoreFocus) printAreaBtn?.focus();
};
const closeSymbolMenu = (restoreFocus = false): void => {
  const symbolMenu = getSymbolMenu();
  const symbolBtn = getSymbolBtn();
  if (!symbolMenu) return;
  symbolMenu.hidden = true;
  symbolBtn?.setAttribute('aria-expanded', 'false');
  if (restoreFocus) symbolBtn?.focus();
};
const openPrintAreaMenu = (printAreaBtn = getPrintAreaBtn()): void => {
  const printAreaMenu = getPrintAreaMenu();
  if (!printAreaMenu || !printAreaBtn) return;
  closeBorderMenu();
  closeFreezeMenu();
  closeConditionalMenu();
  closeTextOrientationMenu();
  closeFillMenu();
  closeClearMenu();
  printAreaMenu.hidden = false;
  printAreaBtn.setAttribute('aria-haspopup', 'menu');
  printAreaBtn.setAttribute('aria-expanded', 'true');
  focusMenuItem(printAreaMenu, 'first');
};
const openSymbolMenu = (symbolBtn = getSymbolBtn()): void => {
  const symbolMenu = getSymbolMenu();
  if (!symbolMenu || !symbolBtn) return;
  closeBorderMenu();
  closeFreezeMenu();
  closeConditionalMenu();
  closeTextOrientationMenu();
  closeFillMenu();
  closeClearMenu();
  closePrintAreaMenu();
  symbolMenu.hidden = false;
  symbolBtn.setAttribute('aria-haspopup', 'menu');
  symbolBtn.setAttribute('aria-expanded', 'true');
  focusMenuItem(symbolMenu, 'first');
};
document.addEventListener('click', (event) => {
  const item = (event.target as Element | null)?.closest<HTMLButtonElement>(
    '[data-print-area-action]',
  );
  const action = item?.dataset.printAreaAction;
  if (!action) return;
  event.preventDefault();
  event.stopPropagation();
  closePrintAreaMenu();
  applyPrintAreaAction(action as 'set' | 'clear');
});
document.addEventListener('click', (event) => {
  const actionItem = (event.target as Element | null)?.closest<HTMLButtonElement>(
    '[data-symbol-action]',
  );
  if (actionItem?.dataset.symbolAction === 'more') {
    event.preventDefault();
    event.stopPropagation();
    closeSymbolMenu();
    void insertCustomSymbolIntoActiveCell();
    return;
  }
  const item = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-symbol]');
  const symbol = item?.dataset.symbol;
  if (!symbol) return;
  event.preventDefault();
  event.stopPropagation();
  closeSymbolMenu();
  insertSymbolIntoActiveCell(symbol);
});
document.addEventListener('keydown', (event) => {
  const printAreaMenu = getPrintAreaMenu();
  if (!printAreaMenu || printAreaMenu.hidden) return;
  handleMenuKeydown(event, printAreaMenu, {
    close: closePrintAreaMenu,
    restoreFocusTo: getPrintAreaBtn(),
  });
});
document.addEventListener('keydown', (event) => {
  const symbolMenu = getSymbolMenu();
  if (!symbolMenu || symbolMenu.hidden) return;
  handleMenuKeydown(event, symbolMenu, {
    close: closeSymbolMenu,
    restoreFocusTo: getSymbolBtn(),
  });
});
document.addEventListener('mousedown', (event) => {
  const target = event.target as Node | null;
  const printAreaMenu = getPrintAreaMenu();
  const printAreaBtn = getPrintAreaBtn();
  const symbolMenu = getSymbolMenu();
  const symbolBtn = getSymbolBtn();
  if (
    fillMenu?.hidden === false &&
    target &&
    !fillMenu.contains(target) &&
    !fillBtn?.contains(target)
  ) {
    closeFillMenu();
  }
  if (
    clearMenu?.hidden === false &&
    target &&
    !clearMenu.contains(target) &&
    !clearBtn?.contains(target)
  ) {
    closeClearMenu();
  }
  if (
    printAreaMenu?.hidden === false &&
    target &&
    !printAreaMenu.contains(target) &&
    !printAreaBtn?.contains(target)
  ) {
    closePrintAreaMenu();
  }
  if (
    symbolMenu?.hidden === false &&
    target &&
    !symbolMenu.contains(target) &&
    !symbolBtn?.contains(target)
  ) {
    closeSymbolMenu();
  }
});
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && fillMenu?.hidden === false) closeFillMenu(true);
  if (event.key === 'Escape' && clearMenu?.hidden === false) closeClearMenu(true);
  if (event.key === 'Escape' && getPrintAreaMenu()?.hidden === false) closePrintAreaMenu(true);
  if (event.key === 'Escape' && getSymbolMenu()?.hidden === false) closeSymbolMenu(true);
});

const textOrientationBtn = document.querySelector<HTMLButtonElement>(
  'button[data-ribbon-command="textOrientation"]',
);
const textOrientationMenu = document.getElementById(
  'menu-text-orientation',
) as HTMLDivElement | null;
const closeTextOrientationMenu = (restoreFocus = false): void => {
  if (!textOrientationMenu) return;
  textOrientationMenu.hidden = true;
  textOrientationBtn?.setAttribute('aria-expanded', 'false');
  if (restoreFocus) textOrientationBtn?.focus();
};
const openTextOrientationMenu = (): void => {
  if (!textOrientationMenu || !textOrientationBtn) return;
  closeBorderMenu();
  closeFreezeMenu();
  closeConditionalMenu();
  closeFillMenu();
  closeClearMenu();
  closeTextOrientationMenu();
  closeCellsMenus();
  textOrientationMenu.hidden = false;
  textOrientationBtn.setAttribute('aria-haspopup', 'menu');
  textOrientationBtn.setAttribute('aria-expanded', 'true');
  focusMenuItem(textOrientationMenu, 'first');
};
textOrientationBtn?.addEventListener('click', (event) => {
  event.preventDefault();
  event.stopPropagation();
  if (!textOrientationMenu) return;
  if (textOrientationMenu.hidden) openTextOrientationMenu();
  else closeTextOrientationMenu(true);
});
textOrientationMenu?.addEventListener('click', (event) => {
  const item = (event.target as Element | null)?.closest<HTMLButtonElement>(
    '[data-text-orientation]',
  );
  const action = item?.dataset.textOrientation;
  if (!action) return;
  event.preventDefault();
  event.stopPropagation();
  closeTextOrientationMenu();
  applyTextOrientationAction(action);
});
textOrientationMenu?.addEventListener('keydown', (event) => {
  handleMenuKeydown(event, textOrientationMenu, {
    close: closeTextOrientationMenu,
    restoreFocusTo: textOrientationBtn,
  });
});
document.addEventListener('mousedown', (event) => {
  const target = event.target as Node | null;
  if (
    textOrientationMenu?.hidden === false &&
    target &&
    !textOrientationMenu.contains(target) &&
    !textOrientationBtn?.contains(target)
  ) {
    closeTextOrientationMenu();
  }
});
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && textOrientationMenu?.hidden === false)
    closeTextOrientationMenu(true);
});

const insertCellsBtn = document.querySelector<HTMLButtonElement>(
  'button[data-ribbon-command="insertRows"]',
);
const deleteCellsBtn = document.querySelector<HTMLButtonElement>(
  'button[data-ribbon-command="deleteRows"]',
);
const formatCellsBtn = document.querySelector<HTMLButtonElement>(
  'button[data-ribbon-command="formatCellsHome"]',
);
const insertCellsMenu = document.getElementById('menu-insert-cells') as HTMLDivElement | null;
const deleteCellsMenu = document.getElementById('menu-delete-cells') as HTMLDivElement | null;
const formatCellsMenu = document.getElementById('menu-format-cells') as HTMLDivElement | null;

const closeCellsMenus = (restoreFocusTo?: HTMLElement | null): void => {
  for (const [menu, btn] of [
    [insertCellsMenu, insertCellsBtn],
    [deleteCellsMenu, deleteCellsBtn],
    [formatCellsMenu, formatCellsBtn],
  ] as const) {
    if (!menu) continue;
    menu.hidden = true;
    btn?.setAttribute('aria-expanded', 'false');
  }
  restoreFocusTo?.focus();
};

const openCellsMenu = (menu: HTMLDivElement | null, button: HTMLButtonElement | null): void => {
  if (!menu || !button) return;
  closeBorderMenu();
  closeFreezeMenu();
  closeConditionalMenu();
  closeFillMenu();
  closeClearMenu();
  closeCellsMenus();
  menu.hidden = false;
  button.setAttribute('aria-haspopup', 'menu');
  button.setAttribute('aria-expanded', 'true');
  focusMenuItem(menu, 'first');
};

insertCellsBtn?.addEventListener('click', (event) => {
  event.preventDefault();
  event.stopPropagation();
  if (!insertCellsMenu) return;
  if (insertCellsMenu.hidden) openCellsMenu(insertCellsMenu, insertCellsBtn);
  else closeCellsMenus(insertCellsBtn);
});
deleteCellsBtn?.addEventListener('click', (event) => {
  event.preventDefault();
  event.stopPropagation();
  if (!deleteCellsMenu) return;
  if (deleteCellsMenu.hidden) openCellsMenu(deleteCellsMenu, deleteCellsBtn);
  else closeCellsMenus(deleteCellsBtn);
});
formatCellsBtn?.addEventListener('click', (event) => {
  event.preventDefault();
  event.stopPropagation();
  if (!formatCellsMenu) return;
  if (formatCellsMenu.hidden) openCellsMenu(formatCellsMenu, formatCellsBtn);
  else closeCellsMenus(formatCellsBtn);
});

insertCellsMenu?.addEventListener('click', (event) => {
  const item = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-cell-insert]');
  const action = item?.dataset.cellInsert;
  if (!action) return;
  event.preventDefault();
  event.stopPropagation();
  closeCellsMenus();
  void applyCellInsertAction(action);
});
deleteCellsMenu?.addEventListener('click', (event) => {
  const item = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-cell-delete]');
  const action = item?.dataset.cellDelete;
  if (!action) return;
  event.preventDefault();
  event.stopPropagation();
  closeCellsMenus();
  void applyCellDeleteAction(action);
});
formatCellsMenu?.addEventListener('click', (event) => {
  const item = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-cell-format]');
  const action = item?.dataset.cellFormat;
  if (!action) return;
  event.preventDefault();
  event.stopPropagation();
  closeCellsMenus();
  void applyCellFormatAction(action);
});
insertCellsMenu?.addEventListener('keydown', (event) => {
  handleMenuKeydown(event, insertCellsMenu, {
    close: () => closeCellsMenus(insertCellsBtn),
    restoreFocusTo: insertCellsBtn,
  });
});
deleteCellsMenu?.addEventListener('keydown', (event) => {
  handleMenuKeydown(event, deleteCellsMenu, {
    close: () => closeCellsMenus(deleteCellsBtn),
    restoreFocusTo: deleteCellsBtn,
  });
});
formatCellsMenu?.addEventListener('keydown', (event) => {
  handleMenuKeydown(event, formatCellsMenu, {
    close: () => closeCellsMenus(formatCellsBtn),
    restoreFocusTo: formatCellsBtn,
  });
});
document.addEventListener('mousedown', (event) => {
  const target = event.target as Node | null;
  if (!target) return;
  const inside =
    insertCellsMenu?.contains(target) ||
    deleteCellsMenu?.contains(target) ||
    formatCellsMenu?.contains(target) ||
    insertCellsBtn?.contains(target) ||
    deleteCellsBtn?.contains(target) ||
    formatCellsBtn?.contains(target);
  if (!inside) closeCellsMenus();
});
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape') closeCellsMenus();
});

const selectionToA1Range = (): string | null => {
  const i = inst;
  if (!i) return null;
  const r = i.store.getState().selection.range;
  const start = `${colLetter(r.c0)}${r.r0 + 1}`;
  const end = `${colLetter(r.c1)}${r.r1 + 1}`;
  return start === end ? start : `${start}:${end}`;
};

const applyPrintAreaAction = (action: 'set' | 'clear'): void => {
  const i = inst;
  if (!i) return;
  const sheet = i.store.getState().data.sheetIndex;
  recordPageSetupChange(i.history, i.store, () => {
    if (action === 'clear') {
      clearPrintArea(i.store, sheet);
      return;
    }
    const area = selectionToA1Range();
    if (area) setPrintArea(i.store, sheet, area);
  });
  const setup = getPageSetup(i.store.getState(), sheet);
  const message =
    action === 'clear'
      ? ribbonMenuText.printAreaStatusClear
      : ribbonMenuText.printAreaStatusSet.replace('{range}', setup.printArea ?? '');
  showMessage({
    title: ribbonText.printArea,
    message,
  });
  projectFormatToolbar();
  focusSheet();
};

type PageBreakAction =
  | 'insert-auto'
  | 'insert-row'
  | 'insert-col'
  | 'remove-row'
  | 'remove-col'
  | 'reset-all';

const applyPageBreakAction = (action: string = 'insert-auto'): void => {
  const i = inst;
  if (!i) return;
  const state = i.store.getState();
  const sheet = state.data.sheetIndex;
  const range = state.selection.range;
  recordPageSetupChange(i.history, i.store, () => {
    const pageBreakAction = action as PageBreakAction;
    if (pageBreakAction === 'reset-all') {
      resetManualPageBreaks(i.store, sheet);
      return;
    }
    if (pageBreakAction === 'remove-row') {
      removeManualPageBreak(i.store, sheet, 'row', range.r0);
      return;
    }
    if (pageBreakAction === 'remove-col') {
      removeManualPageBreak(i.store, sheet, 'col', range.c0);
      return;
    }
    if (pageBreakAction === 'insert-col') {
      if (range.c0 > 0) insertManualPageBreak(i.store, sheet, 'col', range.c0);
      return;
    }
    if (pageBreakAction === 'insert-row') {
      if (range.r0 > 0) insertManualPageBreak(i.store, sheet, 'row', range.r0);
      return;
    }
    if (range.r0 > 0) insertManualPageBreak(i.store, sheet, 'row', range.r0);
    else if (range.c0 > 0) insertManualPageBreak(i.store, sheet, 'col', range.c0);
  });
  projectFormatToolbar();
  focusSheet();
};

const DEFAULT_SHEET_BACKGROUND =
  'linear-gradient(135deg, rgba(33,115,70,0.12), rgba(0,120,212,0.08)), repeating-linear-gradient(45deg, rgba(255,255,255,0.36) 0 12px, rgba(255,255,255,0.12) 12px 24px)';

const applySheetBackgroundAction = async (action: 'set' | 'clear' = 'set'): Promise<void> => {
  const i = inst;
  if (!i) return;
  const sheet = i.store.getState().data.sheetIndex;
  if (action === 'clear') {
    clearSheetBackgroundImage(i.store, sheet, i.history);
    projectFormatToolbar();
    focusSheet();
    return;
  }
  const current = i.store.getState().ui.sheetBackgroundImages.get(sheet);
  const value = await showPrompt({
    title: ribbonMenuText.sheetBackgroundSet,
    label: ribbonMenuText.sheetBackgroundPrompt,
    initial: current ?? DEFAULT_SHEET_BACKGROUND,
    okLabel: 'OK',
    cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
    validate: (raw) =>
      raw.trim()
        ? null
        : ribbonLang === 'ja'
          ? '背景画像のURLを入力してください。'
          : 'Enter a background image URL.',
  });
  if (value === null) {
    focusSheet();
    return;
  }
  setSheetBackgroundImage(i.store, sheet, value.trim(), i.history);
  projectFormatToolbar();
  focusSheet();
};

type PrintTitlesAction = 'rows' | 'cols' | 'clear';

const applyPrintTitlesAction = (action: PrintTitlesAction): void => {
  const i = inst;
  if (!i) return;
  const state = i.store.getState();
  const sheet = state.data.sheetIndex;
  const range = state.selection.range;
  if (action === 'clear') {
    clearPrintTitles(i.store, sheet, i.history);
  } else if (action === 'rows') {
    setPrintTitleRows(i.store, sheet, `${range.r0 + 1}:${range.r1 + 1}`, i.history);
  } else {
    setPrintTitleCols(i.store, sheet, `${colLetter(range.c0)}:${colLetter(range.c1)}`, i.history);
  }
  projectFormatToolbar();
  focusSheet();
};

const selectionOutlineAxis = (): 'row' | 'col' => {
  const i = inst;
  if (!i) return 'row';
  const r = i.store.getState().selection.range;
  const rowSpan = r.r1 - r.r0;
  const colSpan = r.c1 - r.c0;
  return rowSpan >= colSpan ? 'row' : 'col';
};

const selectionDetailOutlineAxis = (): 'row' | 'col' => {
  const i = inst;
  if (!i) return 'row';
  const state = i.store.getState();
  const activeRowLevel = state.layout.outlineRows.get(state.selection.active.row) ?? 0;
  const activeColLevel = state.layout.outlineCols.get(state.selection.active.col) ?? 0;
  if (activeRowLevel > 0 && activeColLevel === 0) return 'row';
  if (activeColLevel > 0 && activeRowLevel === 0) return 'col';
  return selectionOutlineAxis();
};

const selectedRowOutlineRange = (): { r0: number; r1: number } | null => {
  const i = inst;
  if (!i) return null;
  const state = i.store.getState();
  const range = state.selection.range;
  let bestRow = state.selection.active.row;
  let bestLevel = state.layout.outlineRows.get(bestRow) ?? 0;
  for (let row = range.r0; row <= range.r1; row += 1) {
    const level = state.layout.outlineRows.get(row) ?? 0;
    if (level > bestLevel) {
      bestRow = row;
      bestLevel = level;
    }
  }
  return bestLevel > 0 ? rowGroupRangeAt(state.layout, bestRow, bestLevel) : null;
};

const selectedColOutlineRange = (): { c0: number; c1: number } | null => {
  const i = inst;
  if (!i) return null;
  const state = i.store.getState();
  const range = state.selection.range;
  let bestCol = state.selection.active.col;
  let bestLevel = state.layout.outlineCols.get(bestCol) ?? 0;
  for (let col = range.c0; col <= range.c1; col += 1) {
    const level = state.layout.outlineCols.get(col) ?? 0;
    if (level > bestLevel) {
      bestCol = col;
      bestLevel = level;
    }
  }
  return bestLevel > 0 ? colGroupRangeAt(state.layout, bestCol, bestLevel) : null;
};

const applyOutlineAction = (action: 'group' | 'ungroup' | 'show-detail' | 'hide-detail'): void => {
  const i = inst;
  if (!i) return;
  const range = i.store.getState().selection.range;
  const axis =
    action === 'show-detail' || action === 'hide-detail'
      ? selectionDetailOutlineAxis()
      : selectionOutlineAxis();
  if (axis === 'row') {
    if (action === 'group') groupRows(i.store, i.history, range.r0, range.r1, i.workbook);
    else if (action === 'ungroup') ungroupRows(i.store, i.history, range.r0, range.r1, i.workbook);
    else {
      const group = selectedRowOutlineRange();
      if (!group) return;
      if (action === 'show-detail')
        expandRowGroup(i.store, i.history, group.r0, group.r1, i.workbook);
      else collapseRowGroup(i.store, i.history, group.r0, group.r1, i.workbook);
    }
  } else {
    if (action === 'group') groupCols(i.store, i.history, range.c0, range.c1, i.workbook);
    else if (action === 'ungroup') ungroupCols(i.store, i.history, range.c0, range.c1, i.workbook);
    else {
      const group = selectedColOutlineRange();
      if (!group) return;
      if (action === 'show-detail')
        expandColGroup(i.store, i.history, group.c0, group.c1, i.workbook);
      else collapseColGroup(i.store, i.history, group.c0, group.c1, i.workbook);
    }
  }
  refreshWorkbookCells();
  projectFormatToolbar();
  focusSheet();
};

const selectReviewComment = (direction: 1 | -1): void => {
  const i = inst;
  if (!i) return;
  const comments = listComments(i.store.getState());
  const ja = ribbonLang === 'ja';
  if (comments.length === 0) {
    void showMessage({
      title: ja ? 'コメント' : 'Comments',
      message: ja ? 'コメントまたはメモが見つかりません。' : 'No comments or notes were found.',
    });
    return;
  }
  const active = i.store.getState().selection.active;
  const current = comments.findIndex(
    (entry) => entry.addr.row === active.row && entry.addr.col === active.col,
  );
  const nextIndex =
    current >= 0
      ? (current + direction + comments.length) % comments.length
      : direction > 0
        ? 0
        : comments.length - 1;
  const next = comments[nextIndex]?.addr;
  if (!next) return;
  mutators.setActive(i.store, next);
  i.openCommentDialog();
};

const deleteActiveReviewComment = (): void => {
  const i = inst;
  if (!i) return;
  const addr = i.store.getState().selection.active;
  recordCommentChange(i.history, i.store, i.workbook, [addr], () => {
    clearComment(i.store, addr, i.workbook);
  });
  projectFormatToolbar();
  focusSheet();
};

const applyReviewCommentAction = (action: string): void => {
  const i = inst;
  if (!i) return;
  if (action === 'delete-active') {
    deleteActiveReviewComment();
    return;
  }
  if (action !== 'delete-all') return;
  const comments = listComments(i.store.getState());
  if (comments.length === 0) {
    void showMessage({
      title: ribbonText.comments,
      message: ribbonMenuText.commentNone,
    });
    return;
  }
  recordCommentChange(
    i.history,
    i.store,
    i.workbook,
    comments.map((entry) => entry.addr),
    () => {
      for (const entry of comments) clearComment(i.store, entry.addr, i.workbook);
    },
  );
  if (statusMetric) {
    statusMetric.textContent = ribbonMenuText.commentsDeleted.replace(
      '{count}',
      String(comments.length),
    );
  }
  projectFormatToolbar();
  focusSheet();
};

const insertSymbolIntoActiveCell = (symbol: string): void => {
  const i = inst;
  if (!i) return;
  const addr = i.store.getState().selection.active;
  if (i.workbook.cellFormula(addr)) {
    void showMessage({
      title: ribbonLang === 'ja' ? '記号' : 'Symbol',
      message:
        ribbonLang === 'ja'
          ? '数式セルには記号を直接挿入できません。'
          : 'Symbols cannot be inserted directly into a formula cell.',
    });
    return;
  }
  if (!isCellWritable(i.store.getState(), addr)) {
    warnProtected(addr);
    void showMessage({
      title: ribbonLang === 'ja' ? '記号' : 'Symbol',
      message:
        ribbonLang === 'ja'
          ? '保護されたセルには記号を挿入できません。'
          : 'Symbols cannot be inserted into a protected cell.',
    });
    return;
  }
  const value = i.workbook.getValue(addr);
  const current = value.kind === 'text' ? value.value : '';
  i.history.begin();
  try {
    i.workbook.setText(addr, `${current}${symbol}`);
  } finally {
    i.history.end();
  }
  refreshWorkbookCells();
  focusSheet();
};

const insertCustomSymbolIntoActiveCell = async (): Promise<void> => {
  const value = await showPrompt({
    title: ribbonMenuText.symbolMore,
    label: ribbonMenuText.symbolPrompt,
    okLabel: 'OK',
    cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
    validate: (raw) => (raw.trim() ? null : ribbonMenuText.symbolInvalid),
  });
  if (value === null) {
    focusSheet();
    return;
  }
  insertSymbolIntoActiveCell(value.trim());
};

const sortFilterHomeBtn = document.querySelector<HTMLButtonElement>(
  'button[data-ribbon-command="sortFilterHome"]',
);
const sortFilterHomeMenu = document.getElementById('menu-sort-home') as HTMLDivElement | null;
const closeSortFilterHomeMenu = (restoreFocus = false): void => {
  if (!sortFilterHomeMenu) return;
  sortFilterHomeMenu.hidden = true;
  sortFilterHomeBtn?.setAttribute('aria-expanded', 'false');
  if (restoreFocus) sortFilterHomeBtn?.focus();
};
const openSortFilterHomeMenu = (): void => {
  if (!sortFilterHomeMenu || !sortFilterHomeBtn) return;
  closeBorderMenu();
  closeFreezeMenu();
  closeConditionalMenu();
  closeFillMenu();
  closeClearMenu();
  closeCellsMenus();
  closeTextOrientationMenu();
  closeFindSelectMenu();
  sortFilterHomeMenu.hidden = false;
  sortFilterHomeBtn.setAttribute('aria-haspopup', 'menu');
  sortFilterHomeBtn.setAttribute('aria-expanded', 'true');
  focusMenuItem(sortFilterHomeMenu, 'first');
};
const applySortMenuAction = (action: string): void => {
  const i = inst;
  if (!i) return;
  if (action === 'asc' || action === 'desc') sortSelection(action);
  else if (action === 'custom') void customSortSelection();
  else if (action === 'filter') openFilterForSelection();
  else if (action === 'filter-by-value') {
    const state = i.store.getState();
    const range = state.ui.filterRange ?? inferAutoFilterRange(state);
    recordFilterChange(i.history, i.store, () => {
      filterBySelectedCellValue(i.store.getState(), i.store, range);
    });
    focusSheet();
  } else if (action === 'filter-clear') {
    const state = i.store.getState();
    const range = state.ui.filterRange ?? inferAutoFilterRange(state);
    recordFilterChange(i.history, i.store, () => {
      clearFilter(i.store.getState(), i.store, range);
    });
    focusSheet();
  } else if (action === 'filter-reapply') {
    recordFilterChange(i.history, i.store, () => {
      reapplyFilters(i.store.getState(), i.store);
    });
    focusSheet();
  } else if (action === 'filter-advanced') {
    void applyAdvancedFilterAction();
  } else if (action === 'dedupe') removeDuplicateRows();
  else if (action === 'conditional') i.openConditionalDialog();
  else if (action === 'named') i.openNamedRangeDialog();
};
sortFilterHomeBtn?.addEventListener('click', (event) => {
  event.preventDefault();
  event.stopPropagation();
  if (!sortFilterHomeMenu) return;
  if (sortFilterHomeMenu.hidden) openSortFilterHomeMenu();
  else closeSortFilterHomeMenu(true);
});
sortFilterHomeMenu?.addEventListener('click', (event) => {
  const item = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-sort]');
  const action = item?.dataset.sort;
  if (!action || !inst) return;
  event.preventDefault();
  event.stopPropagation();
  closeSortFilterHomeMenu();
  applySortMenuAction(action);
});
sortFilterHomeMenu?.addEventListener('keydown', (event) => {
  handleMenuKeydown(event, sortFilterHomeMenu, {
    close: closeSortFilterHomeMenu,
    restoreFocusTo: sortFilterHomeBtn,
  });
});
document.addEventListener('mousedown', (event) => {
  const target = event.target as Node | null;
  if (
    sortFilterHomeMenu?.hidden === false &&
    target &&
    !sortFilterHomeMenu.contains(target) &&
    !sortFilterHomeBtn?.contains(target)
  ) {
    closeSortFilterHomeMenu();
  }
});
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && sortFilterHomeMenu?.hidden === false) closeSortFilterHomeMenu(true);
});

const dataSortBtn = document.querySelector<HTMLButtonElement>(
  'button[data-ribbon-command="filter"]',
);
const dataSortMenu = document.getElementById('menu-sort') as HTMLDivElement | null;
const closeDataSortMenu = (restoreFocus = false): void => {
  if (!dataSortMenu) return;
  dataSortMenu.hidden = true;
  dataSortBtn?.setAttribute('aria-expanded', 'false');
  if (restoreFocus) dataSortBtn?.focus();
};
const openDataSortMenu = (): void => {
  if (!dataSortMenu || !dataSortBtn) return;
  closeBorderMenu();
  closeFreezeMenu();
  closeConditionalMenu();
  closeFillMenu();
  closeClearMenu();
  closeCellsMenus();
  closeTextOrientationMenu();
  closeSortFilterHomeMenu();
  closeFindSelectMenu();
  closePasteMenu();
  dataSortMenu.hidden = false;
  dataSortBtn.setAttribute('aria-haspopup', 'menu');
  dataSortBtn.setAttribute('aria-expanded', 'true');
  focusMenuItem(dataSortMenu, 'first');
};
dataSortBtn?.addEventListener('click', (event) => {
  event.preventDefault();
  event.stopPropagation();
  if (!dataSortMenu) return;
  if (dataSortMenu.hidden) openDataSortMenu();
  else closeDataSortMenu(true);
});
dataSortMenu?.addEventListener('click', (event) => {
  const item = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-sort]');
  const action = item?.dataset.sort;
  if (!action) return;
  event.preventDefault();
  event.stopPropagation();
  closeDataSortMenu();
  applySortMenuAction(action);
});
dataSortMenu?.addEventListener('keydown', (event) => {
  handleMenuKeydown(event, dataSortMenu, { close: closeDataSortMenu, restoreFocusTo: dataSortBtn });
});
document.addEventListener('mousedown', (event) => {
  const target = event.target as Node | null;
  if (
    dataSortMenu?.hidden === false &&
    target &&
    !dataSortMenu.contains(target) &&
    !dataSortBtn?.contains(target)
  ) {
    closeDataSortMenu();
  }
});
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && dataSortMenu?.hidden === false) closeDataSortMenu(true);
});

const findSelectBtn = document.querySelector<HTMLButtonElement>(
  'button[data-ribbon-command="findHome"]',
);
const findSelectMenu = document.getElementById('menu-find-select') as HTMLDivElement | null;
const closeFindSelectMenu = (restoreFocus = false): void => {
  if (!findSelectMenu) return;
  findSelectMenu.hidden = true;
  findSelectBtn?.setAttribute('aria-expanded', 'false');
  if (restoreFocus) findSelectBtn?.focus();
};
const openFindSelectMenu = (): void => {
  if (!findSelectMenu || !findSelectBtn) return;
  closeBorderMenu();
  closeFreezeMenu();
  closeConditionalMenu();
  closeFillMenu();
  closeClearMenu();
  closeCellsMenus();
  closeTextOrientationMenu();
  closeSortFilterHomeMenu();
  findSelectMenu.hidden = false;
  findSelectBtn.setAttribute('aria-haspopup', 'menu');
  findSelectBtn.setAttribute('aria-expanded', 'true');
  focusMenuItem(findSelectMenu, 'first');
};
findSelectBtn?.addEventListener('click', (event) => {
  event.preventDefault();
  event.stopPropagation();
  if (!findSelectMenu) return;
  if (findSelectMenu.hidden) openFindSelectMenu();
  else closeFindSelectMenu(true);
});
findSelectMenu?.addEventListener('click', (event) => {
  const item = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-find-select]');
  const action = item?.dataset.findSelect;
  if (!action) return;
  event.preventDefault();
  event.stopPropagation();
  closeFindSelectMenu();
  applyFindSelectAction(action);
});
findSelectMenu?.addEventListener('keydown', (event) => {
  handleMenuKeydown(event, findSelectMenu, {
    close: closeFindSelectMenu,
    restoreFocusTo: findSelectBtn,
  });
});
document.addEventListener('mousedown', (event) => {
  const target = event.target as Node | null;
  if (
    findSelectMenu?.hidden === false &&
    target &&
    !findSelectMenu.contains(target) &&
    !findSelectBtn?.contains(target)
  ) {
    closeFindSelectMenu();
  }
});
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && findSelectMenu?.hidden === false) closeFindSelectMenu(true);
});

const setupAutoSumMenu = (command: 'autosum' | 'autosumFormula', menuId: string): void => {
  const button = document.querySelector<HTMLButtonElement>(
    `button[data-ribbon-command="${command}"]`,
  );
  const menu = document.getElementById(menuId) as HTMLDivElement | null;
  const close = (restoreFocus = false): void => {
    if (!menu) return;
    menu.hidden = true;
    button?.setAttribute('aria-expanded', 'false');
    if (restoreFocus) button?.focus();
  };
  const open = (): void => {
    if (!button || !menu) return;
    closeBorderMenu();
    closeFreezeMenu();
    closeConditionalMenu();
    closeFillMenu();
    closeClearMenu();
    closeCellsMenus();
    closeTextOrientationMenu();
    closeSortFilterHomeMenu();
    closeFindSelectMenu();
    menu.hidden = false;
    button.setAttribute('aria-haspopup', 'menu');
    button.setAttribute('aria-expanded', 'true');
    focusMenuItem(menu, 'first');
  };
  button?.addEventListener('click', (event) => {
    event.preventDefault();
    event.stopPropagation();
    const target = event.target as Element | null;
    if (!menu) return;
    if (event.altKey || event.shiftKey || target?.closest('.demo__rb-split-chevron')) {
      if (menu.hidden) open();
      else close(true);
      return;
    }
    close();
    applyAutoSumFormula('SUM');
  });
  button?.addEventListener('keydown', (event) => {
    if (event.key !== 'ArrowDown') return;
    event.preventDefault();
    event.stopPropagation();
    if (menu?.hidden) open();
  });
  menu?.addEventListener('click', (event) => {
    const item = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-autosum-fn]');
    const fn = item?.dataset.autosumFn as AutoSumFormulaName | undefined;
    if (!fn) return;
    event.preventDefault();
    event.stopPropagation();
    close();
    applyAutoSumFormula(fn);
  });
  menu?.addEventListener('keydown', (event) => {
    handleMenuKeydown(event, menu, { close, restoreFocusTo: button });
  });
  document.addEventListener('mousedown', (event) => {
    const target = event.target as Node | null;
    if (menu?.hidden === false && target && !menu.contains(target) && !button?.contains(target)) {
      close();
    }
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && menu?.hidden === false) close(true);
  });
};

setupAutoSumMenu('autosum', 'menu-autosum-home');
setupAutoSumMenu('autosumFormula', 'menu-autosum-formulas');

const applyCalcOptionAction = (action: string): void => {
  const i = inst;
  if (!i) return;
  if (action === 'auto' || action === 'manual' || action === 'auto-no-table') {
    const mode = action === 'auto' ? 0 : action === 'manual' ? 1 : 2;
    const ok = i.workbook.setCalcMode(mode as 0 | 1 | 2);
    if (!ok) {
      void showMessage({
        title: ribbonLang === 'ja' ? '計算方法' : 'Calculation Options',
        message:
          ribbonLang === 'ja'
            ? 'このエンジンでは計算モードの保存に対応していません。'
            : 'This engine does not support saving calculation mode.',
      });
      return;
    }
    focusSheet();
    updateCalcOptionsMenu();
    return;
  }
  if (action === 'calculate-now' || action === 'calculate-sheet') {
    i.recalc();
    refreshWorkbookCells();
    focusSheet();
    return;
  }
  if (action === 'iterative') {
    i.openIterativeDialog();
  }
};

const updateCalcOptionsMenu = (menu: HTMLElement = document.body): void => {
  const mode = inst?.workbook.calcMode();
  const active = mode === 0 ? 'auto' : mode === 1 ? 'manual' : mode === 2 ? 'auto-no-table' : null;
  for (const item of menu.querySelectorAll<HTMLElement>('[data-calc-option]')) {
    const isModeItem =
      item.dataset.calcOption === 'auto' ||
      item.dataset.calcOption === 'manual' ||
      item.dataset.calcOption === 'auto-no-table';
    if (!isModeItem) continue;
    const selected = item.dataset.calcOption === active;
    item.setAttribute('aria-checked', selected ? 'true' : 'false');
    item.classList.toggle('app__menu-item--checked', selected);
  }
};

const calcOptionsBtn = document.querySelector<HTMLButtonElement>(
  'button[data-ribbon-command="calcOptions"]',
);
const calcOptionsMenu = document.getElementById('menu-calc-options') as HTMLDivElement | null;
const closeCalcOptionsMenu = (restoreFocus = false): void => {
  if (!calcOptionsMenu) return;
  calcOptionsMenu.hidden = true;
  calcOptionsBtn?.setAttribute('aria-expanded', 'false');
  if (restoreFocus) calcOptionsBtn?.focus();
};
const openCalcOptionsMenu = (): void => {
  if (!calcOptionsMenu || !calcOptionsBtn) return;
  closeBorderMenu();
  closeFreezeMenu();
  closeConditionalMenu();
  closeFillMenu();
  closeClearMenu();
  closeCellsMenus();
  closeTextOrientationMenu();
  closeSortFilterHomeMenu();
  closeFindSelectMenu();
  closePasteMenu();
  calcOptionsMenu.hidden = false;
  calcOptionsBtn.setAttribute('aria-haspopup', 'menu');
  calcOptionsBtn.setAttribute('aria-expanded', 'true');
  focusMenuItem(calcOptionsMenu, 'first');
};
calcOptionsBtn?.addEventListener('click', (event) => {
  event.preventDefault();
  event.stopPropagation();
  if (!calcOptionsMenu) return;
  if (calcOptionsMenu.hidden) openCalcOptionsMenu();
  else closeCalcOptionsMenu(true);
});
calcOptionsMenu?.addEventListener('click', (event) => {
  const item = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-calc-option]');
  const action = item?.dataset.calcOption;
  if (!action) return;
  event.preventDefault();
  event.stopPropagation();
  closeCalcOptionsMenu();
  applyCalcOptionAction(action);
});
calcOptionsMenu?.addEventListener('keydown', (event) => {
  handleMenuKeydown(event, calcOptionsMenu, {
    close: closeCalcOptionsMenu,
    restoreFocusTo: calcOptionsBtn,
  });
});
document.addEventListener('mousedown', (event) => {
  const target = event.target as Node | null;
  if (
    calcOptionsMenu?.hidden === false &&
    target &&
    !calcOptionsMenu.contains(target) &&
    !calcOptionsBtn?.contains(target)
  ) {
    closeCalcOptionsMenu();
  }
});
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && calcOptionsMenu?.hidden === false) closeCalcOptionsMenu(true);
});

const chartInsertBtn = document.querySelector<HTMLButtonElement>(
  'button[data-ribbon-command="chartInsert"]',
);
const chartInsertMenu = document.getElementById('menu-chart-insert') as HTMLDivElement | null;
const closeChartInsertMenu = (restoreFocus = false): void => {
  if (!chartInsertMenu) return;
  chartInsertMenu.hidden = true;
  chartInsertBtn?.setAttribute('aria-expanded', 'false');
  if (restoreFocus) chartInsertBtn?.focus();
};
const openChartInsertMenu = (): void => {
  if (!chartInsertMenu || !chartInsertBtn) return;
  closeBorderMenu();
  closeFreezeMenu();
  closeConditionalMenu();
  closeFillMenu();
  closeClearMenu();
  closeCellsMenus();
  closeTextOrientationMenu();
  closeSortFilterHomeMenu();
  closeFindSelectMenu();
  closePasteMenu();
  closeCalcOptionsMenu();
  chartInsertMenu.hidden = false;
  chartInsertBtn.setAttribute('aria-haspopup', 'menu');
  chartInsertBtn.setAttribute('aria-expanded', 'true');
  focusMenuItem(chartInsertMenu, 'first');
};
chartInsertBtn?.addEventListener('click', (event) => {
  event.preventDefault();
  event.stopPropagation();
  if (!chartInsertMenu) return;
  if (chartInsertMenu.hidden) openChartInsertMenu();
  else closeChartInsertMenu(true);
});
chartInsertMenu?.addEventListener('click', (event) => {
  const item = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-chart-insert]');
  const action = item?.dataset.chartInsert;
  if (!action) return;
  event.preventDefault();
  event.stopPropagation();
  closeChartInsertMenu();
  if (action === 'recommended') void createRecommendedChartFromSelection();
  else createChartFromSelection(chartKindFromAction(action));
});
chartInsertMenu?.addEventListener('keydown', (event) => {
  handleMenuKeydown(event, chartInsertMenu, {
    close: closeChartInsertMenu,
    restoreFocusTo: chartInsertBtn,
  });
});
document.addEventListener('mousedown', (event) => {
  const target = event.target as Node | null;
  if (
    chartInsertMenu?.hidden === false &&
    target &&
    !chartInsertMenu.contains(target) &&
    !chartInsertBtn?.contains(target)
  ) {
    closeChartInsertMenu();
  }
});
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && chartInsertMenu?.hidden === false) closeChartInsertMenu(true);
});

type RibbonDropdownSpec = {
  menuId: string;
  command: string;
};

const DYNAMIC_RIBBON_DROPDOWN_IDS = new Set([
  'menu-paste',
  'menu-pivot-table',
  'menu-defined-names-insert',
  'menu-defined-names',
  'menu-links-file',
  'menu-links-insert',
  'menu-links-data',
  'menu-conditional',
  'menu-fill',
  'menu-clear',
  'menu-text-orientation',
  'menu-insert-cells',
  'menu-delete-cells',
  'menu-format-cells',
  'menu-page-theme',
  'menu-page-breaks',
  'menu-sheet-background',
  'menu-print-titles',
  'menu-sort-home',
  'menu-sort',
  'menu-find-select',
  'menu-autosum-home',
  'menu-autosum-formulas',
  'menu-clear-arrows',
  'menu-error-checking',
  'menu-watch-formulas',
  'menu-watch-view',
  'menu-review-comments',
  'menu-protect-review',
  'menu-protect-view',
  'menu-calc-options',
  'menu-chart-insert',
  'menu-picture-insert',
  'menu-shapes-insert',
  'menu-screenshot-insert',
  'menu-script',
  'menu-table-style-home',
  'menu-table-style-insert',
  'menu-cell-styles-home',
  'menu-currency-home',
  'menu-text-to-columns',
  'menu-data-validation',
  'menu-add-ins',
  'menu-pdf',
]);

const dynamicDropdownSpecForButton = (button: HTMLButtonElement): RibbonDropdownSpec | null => {
  const command = button.dataset.ribbonCommand ?? '';
  if (command === 'paste') return { command, menuId: 'menu-paste' };
  if (command === 'pivotTableInsert') return { command, menuId: 'menu-pivot-table' };
  if (command === 'namedRangesInsert') return { command, menuId: 'menu-defined-names-insert' };
  if (command === 'namedRanges') return { command, menuId: 'menu-defined-names' };
  if (command === 'links') return { command, menuId: 'menu-links-file' };
  if (command === 'linksInsert') return { command, menuId: 'menu-links-insert' };
  if (command === 'linksData') return { command, menuId: 'menu-links-data' };
  if (command === 'conditional') return { command, menuId: 'menu-conditional' };
  if (command === 'fillHome') return { command, menuId: 'menu-fill' };
  if (command === 'clearFormat' && button.closest<HTMLElement>('.demo__ribbon-group--editing')) {
    return { command, menuId: 'menu-clear' };
  }
  if (command === 'textOrientation') return { command, menuId: 'menu-text-orientation' };
  if (command === 'insertRows') return { command, menuId: 'menu-insert-cells' };
  if (command === 'deleteRows') return { command, menuId: 'menu-delete-cells' };
  if (command === 'formatCellsHome') return { command, menuId: 'menu-format-cells' };
  if (command === 'pageTheme') return { command, menuId: 'menu-page-theme' };
  if (command === 'pageBreaks') return { command, menuId: 'menu-page-breaks' };
  if (command === 'sheetBackground') return { command, menuId: 'menu-sheet-background' };
  if (command === 'printTitles') return { command, menuId: 'menu-print-titles' };
  if (command === 'sortFilterHome') return { command, menuId: 'menu-sort-home' };
  if (command === 'filter') return { command, menuId: 'menu-sort' };
  if (command === 'findHome') return { command, menuId: 'menu-find-select' };
  if (command === 'autosum') return { command, menuId: 'menu-autosum-home' };
  if (command === 'autosumFormula') return { command, menuId: 'menu-autosum-formulas' };
  if (command === 'clearArrows') return { command, menuId: 'menu-clear-arrows' };
  if (command === 'errorChecking') return { command, menuId: 'menu-error-checking' };
  if (command === 'watch') return { command, menuId: 'menu-watch-formulas' };
  if (command === 'watchView') return { command, menuId: 'menu-watch-view' };
  if (command === 'deleteCommentReview') return { command, menuId: 'menu-review-comments' };
  if (command === 'protectReview') return { command, menuId: 'menu-protect-review' };
  if (command === 'protect') return { command, menuId: 'menu-protect-view' };
  if (command === 'calcOptions') return { command, menuId: 'menu-calc-options' };
  if (command === 'chartInsert') return { command, menuId: 'menu-chart-insert' };
  if (command === 'pictureInsert') return { command, menuId: 'menu-picture-insert' };
  if (command === 'shapesInsert') return { command, menuId: 'menu-shapes-insert' };
  if (command === 'screenshotInsert') return { command, menuId: 'menu-screenshot-insert' };
  if (command === 'script') return { command, menuId: 'menu-script' };
  if (command === 'formatTableHome') return { command, menuId: 'menu-table-style-home' };
  if (command === 'formatTableInsert') return { command, menuId: 'menu-table-style-insert' };
  if (command === 'cellStyles') return { command, menuId: 'menu-cell-styles-home' };
  if (command === 'currency') return { command, menuId: 'menu-currency-home' };
  if (command === 'textToColumns') return { command, menuId: 'menu-text-to-columns' };
  if (command === 'dataValidation') return { command, menuId: 'menu-data-validation' };
  if (command === 'addIn') return { command, menuId: 'menu-add-ins' };
  if (command === 'pdf') return { command, menuId: 'menu-pdf' };
  return null;
};

const dynamicDropdownSpecForMenu = (menu: HTMLElement): RibbonDropdownSpec | null => {
  switch (menu.id) {
    case 'menu-paste':
      return { command: 'paste', menuId: menu.id };
    case 'menu-pivot-table':
      return { command: 'pivotTableInsert', menuId: menu.id };
    case 'menu-defined-names-insert':
      return { command: 'namedRangesInsert', menuId: menu.id };
    case 'menu-defined-names':
      return { command: 'namedRanges', menuId: menu.id };
    case 'menu-links-file':
      return { command: 'links', menuId: menu.id };
    case 'menu-links-insert':
      return { command: 'linksInsert', menuId: menu.id };
    case 'menu-links-data':
      return { command: 'linksData', menuId: menu.id };
    case 'menu-conditional':
      return { command: 'conditional', menuId: menu.id };
    case 'menu-fill':
      return { command: 'fillHome', menuId: menu.id };
    case 'menu-clear':
      return { command: 'clearFormat', menuId: menu.id };
    case 'menu-text-orientation':
      return { command: 'textOrientation', menuId: menu.id };
    case 'menu-insert-cells':
      return { command: 'insertRows', menuId: menu.id };
    case 'menu-delete-cells':
      return { command: 'deleteRows', menuId: menu.id };
    case 'menu-format-cells':
      return { command: 'formatCellsHome', menuId: menu.id };
    case 'menu-page-theme':
      return { command: 'pageTheme', menuId: menu.id };
    case 'menu-page-breaks':
      return { command: 'pageBreaks', menuId: menu.id };
    case 'menu-sheet-background':
      return { command: 'sheetBackground', menuId: menu.id };
    case 'menu-print-titles':
      return { command: 'printTitles', menuId: menu.id };
    case 'menu-sort-home':
      return { command: 'sortFilterHome', menuId: menu.id };
    case 'menu-sort':
      return { command: 'filter', menuId: menu.id };
    case 'menu-find-select':
      return { command: 'findHome', menuId: menu.id };
    case 'menu-autosum-home':
      return { command: 'autosum', menuId: menu.id };
    case 'menu-autosum-formulas':
      return { command: 'autosumFormula', menuId: menu.id };
    case 'menu-clear-arrows':
      return { command: 'clearArrows', menuId: menu.id };
    case 'menu-error-checking':
      return { command: 'errorChecking', menuId: menu.id };
    case 'menu-watch-formulas':
      return { command: 'watch', menuId: menu.id };
    case 'menu-watch-view':
      return { command: 'watchView', menuId: menu.id };
    case 'menu-review-comments':
      return { command: 'deleteCommentReview', menuId: menu.id };
    case 'menu-protect-review':
      return { command: 'protectReview', menuId: menu.id };
    case 'menu-protect-view':
      return { command: 'protect', menuId: menu.id };
    case 'menu-calc-options':
      return { command: 'calcOptions', menuId: menu.id };
    case 'menu-chart-insert':
      return { command: 'chartInsert', menuId: menu.id };
    case 'menu-picture-insert':
      return { command: 'pictureInsert', menuId: menu.id };
    case 'menu-shapes-insert':
      return { command: 'shapesInsert', menuId: menu.id };
    case 'menu-screenshot-insert':
      return { command: 'screenshotInsert', menuId: menu.id };
    case 'menu-script':
      return { command: 'script', menuId: menu.id };
    case 'menu-table-style-home':
      return { command: 'formatTableHome', menuId: menu.id };
    case 'menu-table-style-insert':
      return { command: 'formatTableInsert', menuId: menu.id };
    case 'menu-cell-styles-home':
      return { command: 'cellStyles', menuId: menu.id };
    case 'menu-currency-home':
      return { command: 'currency', menuId: menu.id };
    case 'menu-text-to-columns':
      return { command: 'textToColumns', menuId: menu.id };
    case 'menu-data-validation':
      return { command: 'dataValidation', menuId: menu.id };
    case 'menu-add-ins':
      return { command: 'addIn', menuId: menu.id };
    case 'menu-pdf':
      return { command: 'pdf', menuId: menu.id };
    default:
      return null;
  }
};

const dynamicDropdownButtonForSpec = (spec: RibbonDropdownSpec): HTMLButtonElement | null => {
  if (spec.menuId === 'menu-clear') {
    return document.querySelector<HTMLButtonElement>(
      '.demo__ribbon-group--editing button[data-ribbon-command="clearFormat"]',
    );
  }
  return document.querySelector<HTMLButtonElement>(`button[data-ribbon-command="${spec.command}"]`);
};

const closeDynamicConditionalSubmenus = (menu: HTMLElement): void => {
  menu.querySelectorAll<HTMLElement>('[data-cf-panel]').forEach((panel) => {
    panel.hidden = true;
  });
  menu.querySelectorAll<HTMLElement>('[data-cf-submenu]').forEach((trigger) => {
    trigger.classList.remove('app__menu-item--active');
  });
};

const closeDynamicRibbonDropdown = (spec: RibbonDropdownSpec, restoreFocus = false): void => {
  const menu = document.getElementById(spec.menuId) as HTMLDivElement | null;
  const button = dynamicDropdownButtonForSpec(spec);
  if (!menu) return;
  menu.hidden = true;
  if (menu.id === 'menu-conditional') closeDynamicConditionalSubmenus(menu);
  button?.setAttribute('aria-expanded', 'false');
  if (restoreFocus) button?.focus();
};

const closeAllDynamicRibbonDropdowns = (exceptMenuId?: string): void => {
  for (const menu of document.querySelectorAll<HTMLDivElement>('.app__menu')) {
    if (!DYNAMIC_RIBBON_DROPDOWN_IDS.has(menu.id) || menu.id === exceptMenuId) continue;
    const spec = dynamicDropdownSpecForMenu(menu);
    if (spec) closeDynamicRibbonDropdown(spec);
  }
};

const openDynamicRibbonDropdown = (
  spec: RibbonDropdownSpec,
  button = dynamicDropdownButtonForSpec(spec),
): void => {
  const menu = document.getElementById(spec.menuId) as HTMLDivElement | null;
  if (!menu || !button) return;
  if (spec.menuId === 'menu-calc-options') updateCalcOptionsMenu(menu);
  if (spec.menuId === 'menu-defined-names' || spec.menuId === 'menu-defined-names-insert') {
    updateDefinedNamesMenu(menu);
  }
  closeAllDynamicRibbonDropdowns(spec.menuId);
  closeBorderMenu();
  closeFreezeMenu();
  closePrintAreaMenu();
  closeSymbolMenu();
  menu.hidden = false;
  button.setAttribute('aria-haspopup', 'menu');
  button.setAttribute('aria-expanded', 'true');
  focusMenuItem(menu);
};

const openDynamicConditionalSubmenu = (
  menu: HTMLElement,
  key: string,
  trigger: HTMLElement,
): void => {
  closeDynamicConditionalSubmenus(menu);
  const panel = menu.querySelector<HTMLElement>(`[data-cf-panel="${key}"]`);
  if (!panel) return;
  const menuRect = menu.getBoundingClientRect();
  const triggerRect = trigger.getBoundingClientRect();
  panel.style.top = `${Math.max(0, triggerRect.top - menuRect.top - 4)}px`;
  panel.hidden = false;
  trigger.classList.add('app__menu-item--active');
};

const dynamicRibbonDropdownClick = (event: MouseEvent): boolean => {
  const target = event.target as Element | null;
  const menu = target?.closest<HTMLElement>('.app__menu');
  if (!menu || !DYNAMIC_RIBBON_DROPDOWN_IDS.has(menu.id)) return false;
  const spec = dynamicDropdownSpecForMenu(menu);
  if (!spec) return false;

  const paste = target?.closest<HTMLButtonElement>('[data-paste-action]');
  if (paste?.dataset.pasteAction) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    void applyRibbonPasteAction(paste.dataset.pasteAction);
    return true;
  }

  const pivot = target?.closest<HTMLButtonElement>('[data-pivot-table-action]');
  if (pivot?.dataset.pivotTableAction) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    applyPivotTableAction(pivot.dataset.pivotTableAction);
    return true;
  }

  const definedName = target?.closest<HTMLButtonElement>('[data-defined-name-action]');
  if (definedName?.dataset.definedNameAction) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    void applyDefinedNameAction(definedName.dataset.definedNameAction);
    return true;
  }

  const link = target?.closest<HTMLButtonElement>('[data-link-action]');
  if (link?.dataset.linkAction) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    applyLinksAction(link.dataset.linkAction);
    return true;
  }

  const cfSubmenu = target?.closest<HTMLElement>('[data-cf-submenu]');
  if (cfSubmenu && menu.id === 'menu-conditional') {
    event.preventDefault();
    event.stopPropagation();
    openDynamicConditionalSubmenu(menu, cfSubmenu.dataset.cfSubmenu ?? '', cfSubmenu);
    return true;
  }
  const cfItem = target?.closest<HTMLButtonElement>('[data-cf-action]');
  const cfAction = cfItem?.dataset.cfAction;
  if (cfAction && menu.id === 'menu-conditional' && !cfAction.startsWith('submenu-')) {
    event.preventDefault();
    event.stopPropagation();
    const panel = cfItem?.closest<HTMLElement>('[data-cf-panel]')?.dataset.cfPanel;
    closeDynamicRibbonDropdown(spec);
    void applyConditionalMenuAction(cfAction, panel);
    return true;
  }

  const fill = target?.closest<HTMLButtonElement>('[data-fill]');
  if (fill?.dataset.fill) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    if (fill.dataset.fill === 'series') void applyFillSeries();
    else if (
      fill.dataset.fill === 'days' ||
      fill.dataset.fill === 'weekdays' ||
      fill.dataset.fill === 'months' ||
      fill.dataset.fill === 'years'
    ) {
      void applyFillSeries(fill.dataset.fill);
    } else applyFillDirection(fill.dataset.fill as 'down' | 'right' | 'up' | 'left');
    return true;
  }

  const clear = target?.closest<HTMLButtonElement>('[data-clear]');
  if (clear?.dataset.clear) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    applyClearAction(clear.dataset.clear);
    return true;
  }

  const orientation = target?.closest<HTMLButtonElement>('[data-text-orientation]');
  if (orientation?.dataset.textOrientation) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    applyTextOrientationAction(orientation.dataset.textOrientation);
    return true;
  }

  const insert = target?.closest<HTMLButtonElement>('[data-cell-insert]');
  if (insert?.dataset.cellInsert) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    void applyCellInsertAction(insert.dataset.cellInsert);
    return true;
  }

  const del = target?.closest<HTMLButtonElement>('[data-cell-delete]');
  if (del?.dataset.cellDelete) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    void applyCellDeleteAction(del.dataset.cellDelete);
    return true;
  }

  const format = target?.closest<HTMLButtonElement>('[data-cell-format]');
  if (format?.dataset.cellFormat) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    void applyCellFormatAction(format.dataset.cellFormat);
    return true;
  }

  const pageBreak = target?.closest<HTMLButtonElement>('[data-page-break-action]');
  if (pageBreak?.dataset.pageBreakAction) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    applyPageBreakAction(pageBreak.dataset.pageBreakAction);
    return true;
  }

  const sheetBackground = target?.closest<HTMLButtonElement>('[data-sheet-background-action]');
  if (sheetBackground?.dataset.sheetBackgroundAction) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    const action = sheetBackground.dataset.sheetBackgroundAction === 'clear' ? 'clear' : 'set';
    void applySheetBackgroundAction(action);
    return true;
  }

  const printTitles = target?.closest<HTMLButtonElement>('[data-print-titles-action]');
  if (printTitles?.dataset.printTitlesAction) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    applyPrintTitlesAction(printTitles.dataset.printTitlesAction as PrintTitlesAction);
    return true;
  }

  const pageTheme = target?.closest<HTMLButtonElement>('[data-page-theme-action]');
  const pageThemeAction = pageTheme?.dataset.pageThemeAction as UiTheme | undefined;
  if (pageThemeAction) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    applyUiTheme(pageThemeAction);
    focusSheet();
    return true;
  }

  const sort = target?.closest<HTMLButtonElement>('[data-sort]');
  if (sort?.dataset.sort) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    applySortMenuAction(sort.dataset.sort);
    return true;
  }

  const find = target?.closest<HTMLButtonElement>('[data-find-select]');
  if (find?.dataset.findSelect) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    applyFindSelectAction(find.dataset.findSelect);
    return true;
  }

  const autosum = target?.closest<HTMLButtonElement>('[data-autosum-fn]');
  const fn = autosum?.dataset.autosumFn as AutoSumFormulaName | undefined;
  if (fn) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    applyAutoSumFormula(fn);
    return true;
  }

  const formulaAudit = target?.closest<HTMLButtonElement>('[data-formula-audit-action]');
  if (formulaAudit?.dataset.formulaAuditAction) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    applyFormulaAuditAction(formulaAudit.dataset.formulaAuditAction);
    return true;
  }

  const watch = target?.closest<HTMLButtonElement>('[data-watch-action]');
  if (watch?.dataset.watchAction) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    applyWatchAction(watch.dataset.watchAction);
    return true;
  }

  const comment = target?.closest<HTMLButtonElement>('[data-comment-action]');
  if (comment?.dataset.commentAction) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    applyReviewCommentAction(comment.dataset.commentAction);
    return true;
  }

  const protect = target?.closest<HTMLButtonElement>('[data-protect-action]');
  if (protect?.dataset.protectAction) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    void applyProtectAction(protect.dataset.protectAction);
    return true;
  }

  const calc = target?.closest<HTMLButtonElement>('[data-calc-option]');
  if (calc?.dataset.calcOption) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    applyCalcOptionAction(calc.dataset.calcOption);
    return true;
  }

  const chart = target?.closest<HTMLButtonElement>('[data-chart-insert]');
  if (chart?.dataset.chartInsert) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    if (chart.dataset.chartInsert === 'recommended') void createRecommendedChartFromSelection();
    else createChartFromSelection(chartKindFromAction(chart.dataset.chartInsert));
    return true;
  }

  const picture = target?.closest<HTMLButtonElement>('[data-picture-insert]');
  if (picture?.dataset.pictureInsert) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    void insertPictureFromRibbon(picture.dataset.pictureInsert);
    return true;
  }

  const shape = target?.closest<HTMLButtonElement>('[data-shape-insert]');
  if (shape?.dataset.shapeInsert) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    insertShapeFromRibbon(shape.dataset.shapeInsert as SessionShapeKind);
    return true;
  }

  const screenshot = target?.closest<HTMLButtonElement>('[data-screenshot-insert]');
  if (screenshot?.dataset.screenshotInsert) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    insertScreenshotFromRibbon();
    return true;
  }

  const script = target?.closest<HTMLButtonElement>('[data-script-action]');
  if (script?.dataset.scriptAction) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    void applyScriptAction(script.dataset.scriptAction);
    return true;
  }

  const pdf = target?.closest<HTMLButtonElement>('[data-pdf-action]');
  if (pdf?.dataset.pdfAction) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    applyPdfAction(pdf.dataset.pdfAction);
    return true;
  }

  const tableStyle = target?.closest<HTMLButtonElement>('[data-table-style]');
  const style = tableStyle?.dataset.tableStyle as TableStyle | undefined;
  if (style) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    const variant = (tableStyle?.dataset.tableVariant as TableVariantId | undefined) ?? 'banded';
    void createTableFromSelection(style, tableStyle?.dataset.tableColor, variant);
    return true;
  }

  const tableFooter = target?.closest<HTMLButtonElement>('[data-table-style-footer]');
  if (tableFooter?.dataset.tableStyleFooter) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    void openTableStyleFooterAction(tableFooter.dataset.tableStyleFooter);
    return true;
  }

  const cellStyleChip = target?.closest<HTMLButtonElement>('[data-cell-style]');
  if (cellStyleChip?.dataset.cellStyle) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    applyCellStyleFromRibbon(cellStyleChip.dataset.cellStyle as CellStyleId);
    return true;
  }

  const cellStyleFooter = target?.closest<HTMLButtonElement>('[data-cell-style-footer]');
  if (cellStyleFooter?.dataset.cellStyleFooter) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    void openCellStyleFooterAction(cellStyleFooter.dataset.cellStyleFooter);
    return true;
  }

  const currencyPreset = target?.closest<HTMLButtonElement>('[data-currency-preset]');
  if (currencyPreset?.dataset.currencyPreset) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    applyCurrencyPreset(currencyPreset.dataset.currencyPreset);
    return true;
  }

  const currencyFooter = target?.closest<HTMLButtonElement>('[data-currency-footer]');
  if (currencyFooter?.dataset.currencyFooter) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    openCurrencyFooterAction(currencyFooter.dataset.currencyFooter);
    return true;
  }

  const textToColumnsItem = target?.closest<HTMLButtonElement>('[data-text-to-columns-delimiter]');
  const delimiter = textToColumnsItem?.dataset.textToColumnsDelimiter;
  if (delimiter !== undefined) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    if (delimiter === 'custom') void splitTextToColumnsCustom();
    else splitTextToColumns(delimiter === '\\t' ? '\t' : delimiter);
    return true;
  }

  const validation = target?.closest<HTMLButtonElement>('[data-validation-action]');
  if (validation?.dataset.validationAction) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    applyDataValidationAction(validation.dataset.validationAction);
    return true;
  }

  const addIn = target?.closest<HTMLButtonElement>('[data-add-in-action]');
  if (addIn?.dataset.addInAction) {
    event.preventDefault();
    event.stopPropagation();
    closeDynamicRibbonDropdown(spec);
    applyAddInAction(addIn.dataset.addInAction);
    return true;
  }

  return false;
};

document.addEventListener('click', (event) => {
  dynamicRibbonDropdownClick(event);
});

document.addEventListener('mouseover', (event) => {
  const target = event.target as Element | null;
  const menu = target?.closest<HTMLElement>('#menu-conditional');
  if (!menu || menu === conditionalMenu) return;
  const trigger = target?.closest<HTMLElement>('[data-cf-submenu]');
  if (trigger) {
    openDynamicConditionalSubmenu(menu, trigger.dataset.cfSubmenu ?? '', trigger);
    return;
  }
  if (target?.closest('.app__menu-item:not([data-cf-submenu])')) {
    closeDynamicConditionalSubmenus(menu);
  }
});

document.addEventListener('keydown', (event) => {
  const menu = Array.from(document.querySelectorAll<HTMLDivElement>('.app__menu')).find(
    (candidate) => !candidate.hidden && DYNAMIC_RIBBON_DROPDOWN_IDS.has(candidate.id),
  );
  if (!menu) return;
  const spec = dynamicDropdownSpecForMenu(menu);
  if (!spec) return;
  handleMenuKeydown(event, menu, {
    close: (restoreFocus) => closeDynamicRibbonDropdown(spec, restoreFocus),
    restoreFocusTo: dynamicDropdownButtonForSpec(spec),
  });
});

document.addEventListener('mousedown', (event) => {
  const target = event.target as Node | null;
  if (!target) return;
  for (const menu of document.querySelectorAll<HTMLDivElement>('.app__menu')) {
    if (menu.hidden || !DYNAMIC_RIBBON_DROPDOWN_IDS.has(menu.id)) continue;
    const spec = dynamicDropdownSpecForMenu(menu);
    const button = spec ? dynamicDropdownButtonForSpec(spec) : null;
    if (menu.contains(target) || button?.contains(target)) continue;
    if (spec) closeDynamicRibbonDropdown(spec);
  }
});

const colLabel = (n: number): string => {
  let out = '';
  let v = n;
  do {
    out = String.fromCharCode(65 + (v % 26)) + out;
    v = Math.floor(v / 26) - 1;
  } while (v >= 0);
  return out;
};

const fmt = (n: number): string => {
  if (!Number.isFinite(n)) return '—';
  const abs = Math.abs(n);
  if (abs !== 0 && (abs < 0.01 || abs >= 1e9)) return n.toExponential(3);
  return n.toLocaleString('en-US', { maximumFractionDigits: 4 });
};

type StatKey = 'sum' | 'avg' | 'count' | 'min' | 'max';
const STAT_KEYS: StatKey[] = ['sum', 'avg', 'count', 'min', 'max'];
const activeStats: Set<StatKey> = (() => {
  try {
    const saved = localStorage.getItem('fc-status-stats');
    if (saved) return new Set(JSON.parse(saved) as StatKey[]);
  } catch {}
  return new Set<StatKey>(['sum', 'avg', 'count']);
})();
const persistStats = (): void => {
  try {
    localStorage.setItem('fc-status-stats', JSON.stringify(Array.from(activeStats)));
  } catch {}
};

// Composite badge showing both passthrough OOXML parts and spreadsheet Tables.
// We accumulate the latest snapshot from each event and render together so
// switching workbooks doesn't leak stale numbers from the previous one.
const objectCounts = { passthroughs: 0, tables: 0, passByCat: {} as Record<string, number> };
function refreshObjectsBadge(
  source: 'passthroughs' | 'tables',
  detail: { count: number; byCategory?: Record<string, number> },
): void {
  if (source === 'passthroughs') {
    objectCounts.passthroughs = detail.count;
    objectCounts.passByCat = detail.byCategory ?? {};
  } else {
    objectCounts.tables = detail.count;
  }
  if (!statusObjects) return;
  const parts: string[] = [];
  if (objectCounts.tables > 0)
    parts.push(`${objectCounts.tables} table${objectCounts.tables === 1 ? '' : 's'}`);
  const charts = objectCounts.passByCat.charts ?? 0;
  const drawings = objectCounts.passByCat.drawings ?? 0;
  const pivots = objectCounts.passByCat.pivotTables ?? 0;
  if (charts > 0) parts.push(`${charts} chart${charts === 1 ? '' : 's'}`);
  if (drawings > 0) parts.push(`${drawings} drawing${drawings === 1 ? '' : 's'}`);
  if (pivots > 0) parts.push(`${pivots} pivot${pivots === 1 ? '' : 's'}`);
  if (parts.length === 0) {
    statusObjects.hidden = true;
    statusObjects.textContent = '';
    return;
  }
  statusObjects.hidden = false;
  statusObjects.textContent = `objects · ${parts.join(', ')}`;
  statusObjects.title = 'Read-only — loaded from .xlsx but not editable in formulon-cell';
}

function projectStatus(): void {
  if (!inst) return;
  const s = inst.store.getState();
  const a = s.selection.active;
  const r = s.selection.range;

  if (statusSelection) {
    if (r.r0 === r.r1 && r.c0 === r.c1) {
      statusSelection.textContent = `${colLabel(a.col)}${a.row + 1}`;
    } else {
      const tl = `${colLabel(r.c0)}${r.r0 + 1}`;
      const br = `${colLabel(r.c1)}${r.r1 + 1}`;
      const cells = (r.r1 - r.r0 + 1) * (r.c1 - r.c0 + 1);
      statusSelection.textContent = `${tl}:${br} · ${cells} cells`;
    }
  }

  if (statusMetric) {
    const stats = aggregateSelection(s);
    if (stats.numericCount === 0) {
      statusMetric.textContent = '';
    } else {
      const parts: string[] = [];
      if (activeStats.has('sum')) parts.push(`Sum ${fmt(stats.sum)}`);
      if (activeStats.has('avg')) parts.push(`Avg ${fmt(stats.avg)}`);
      if (activeStats.has('count')) parts.push(`Count ${stats.numericCount}`);
      if (activeStats.has('min')) parts.push(`Min ${fmt(stats.min)}`);
      if (activeStats.has('max')) parts.push(`Max ${fmt(stats.max)}`);
      statusMetric.textContent = parts.join(' · ');
    }
  }
}

// Right-click on the status metric → checkbox menu to toggle stats.
statusMetric?.addEventListener('contextmenu', (e) => {
  e.preventDefault();
  const opener =
    document.activeElement instanceof HTMLElement ? document.activeElement : statusMetric;
  const menu = document.createElement('div');
  menu.className = 'app__dropdown';
  prepareMenu(menu, 'Selection summary');
  menu.style.position = 'fixed';
  menu.style.left = `${e.clientX}px`;
  menu.style.bottom = `${window.innerHeight - e.clientY + 4}px`;
  menu.style.top = '';
  let cleanupMenuListeners = (): void => {};
  const closeMenu = (restoreFocus = false): void => {
    menu.remove();
    cleanupMenuListeners();
    if (restoreFocus) opener?.focus();
  };
  for (const key of STAT_KEYS) {
    const item = document.createElement('button');
    item.type = 'button';
    item.className = 'app__menu-item';
    item.setAttribute('role', 'menuitemcheckbox');
    item.setAttribute('aria-checked', activeStats.has(key) ? 'true' : 'false');
    item.tabIndex = -1;
    item.textContent = `${activeStats.has(key) ? '✓ ' : '  '}${key.toUpperCase()}`;
    item.addEventListener('click', () => {
      if (activeStats.has(key)) activeStats.delete(key);
      else activeStats.add(key);
      persistStats();
      projectStatus();
      const checked = activeStats.has(key);
      item.setAttribute('aria-checked', checked ? 'true' : 'false');
      item.textContent = `${checked ? '✓ ' : '  '}${key.toUpperCase()}`;
    });
    menu.appendChild(item);
  }
  const close = (ev: MouseEvent): void => {
    if (!menu.contains(ev.target as Node)) {
      closeMenu();
    }
  };
  menu.addEventListener('keydown', (event) => {
    handleMenuKeydown(event, menu, { close: closeMenu, restoreFocusTo: opener });
  });
  cleanupMenuListeners = () => document.removeEventListener('mousedown', close, true);
  document.body.appendChild(menu);
  focusMenuItem(menu);
  setTimeout(() => document.addEventListener('mousedown', close, true), 0);
});

const ACTIVE_CLASS = 'demo__rb--active';
const setActive = (id: string, on: boolean): void => {
  const el = document.getElementById(id);
  if (!el) return;
  el.classList.toggle(ACTIVE_CLASS, on);
};
const markCurrentLegacyRibbonBindings = (): void => {
  for (const command of Object.keys(legacyCommandIds)) {
    document
      .querySelector<HTMLButtonElement>(`button[data-ribbon-command="${command}"]`)
      ?.setAttribute('data-legacy-bound', '1');
  }
};
const setRibbonCommandActive = (command: string, on: boolean): void => {
  const el = document.querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`);
  if (!el) return;
  el.classList.toggle(ACTIVE_CLASS, on);
  el.setAttribute('aria-pressed', on ? 'true' : 'false');
};

function projectFormatToolbar(): void {
  if (!inst) return;
  const s = inst.store.getState();
  const a = s.selection.active;
  const key = `${a.sheet}:${a.row}:${a.col}`;
  const f = s.format.formats.get(key);
  setActive('btn-bold', !!f?.bold);
  setActive('btn-italic', !!f?.italic);
  setActive('btn-underline', !!f?.underline);
  setActive('btn-strike', !!f?.strike);
  setActive('btn-align-left', f?.align === 'left');
  setActive('btn-align-center', f?.align === 'center');
  setActive('btn-align-right', f?.align === 'right');
  setActive('btn-currency', f?.numFmt?.kind === 'currency');
  setActive('btn-percent', f?.numFmt?.kind === 'percent');
  setRibbonCommandActive('viewGridlines', s.ui.showGridLines !== false);
  setRibbonCommandActive('viewHeadings', s.ui.showHeaders !== false);
  setRibbonCommandActive('viewFormulas', !!s.ui.showFormulas);
  setRibbonCommandActive('showFormulasFormula', !!s.ui.showFormulas);
  setRibbonCommandActive('viewFormulaBar', formulaBarVisible);
  setRibbonCommandActive('viewR1C1', !!s.ui.r1c1);
  setRibbonCommandActive('viewNormal', s.ui.workbookView === 'normal');
  setRibbonCommandActive('viewPageLayout', s.ui.workbookView === 'pageLayout');
  setRibbonCommandActive('viewPageBreakPreview', s.ui.workbookView === 'pageBreakPreview');
  for (const wrap of document.querySelectorAll<HTMLElement>('[data-ribbon-select]')) {
    const id = wrap.dataset.ribbonSelect;
    if (!id) continue;
    const current = currentRibbonControlValue(id);
    const value = wrap.querySelector<HTMLElement>('.demo__rb-dd__value');
    if (value) value.textContent = ribbonSelectLabel(wrap, current);
    for (const option of wrap.querySelectorAll<HTMLElement>('.demo__rb-dd__opt')) {
      const selected = option.dataset.value === current;
      option.classList.toggle('demo__rb-dd__opt--selected', selected);
      option.setAttribute('aria-selected', selected ? 'true' : 'false');
    }
  }
  const fontColorSwatch = document.querySelector<HTMLElement>(
    '[data-ribbon-command="fontColor"] .demo__rb-color__swatch',
  );
  if (fontColorSwatch) fontColorSwatch.style.background = f?.color ?? '#201f1e';
  const fillColorSwatch = document.querySelector<HTMLElement>(
    '[data-ribbon-command="fillColor"] .demo__rb-color__swatch',
  );
  if (fillColorSwatch) fillColorSwatch.style.background = f?.fill ?? '#ffffff';
}

async function boot(): Promise<void> {
  // Default to the real WASM engine. Pass ?engine=stub to force the JS stub
  // for explicit demos or behavior diffs.
  const params = new URLSearchParams(window.location.search);
  const preferStub = params.get('engine') === 'stub';
  const wb = await WorkbookHandle.createDefault({
    preferStub,
    onFallback: (reason) => {
      // eslint-disable-next-line no-console
      console.info('[formulon-cell]', reason);
    },
  });
  // mount.ts only runs `seed` on workbooks it owns. We construct `wb` here so
  // we can read `isStub` / `version` for the engine pill before mounting,
  // which means we have to seed the workbook ourselves. `?fixture=empty`
  // (used by E2E specs that need a deterministic blank workbook) skips this.
  if (bootParams.get('fixture') !== 'empty') {
    seed(wb);
  }

  inst = await Spreadsheet.mount(sheetEl as HTMLElement, {
    theme: toCore(uiTheme),
    workbook: wb,
    locale: localeParam === 'en' ? 'en' : 'ja',
    features: playgroundFeatureFlags(),
  });
  // Debug-only: expose for browser console / e2e poking. Safe to leave on the
  // playground build; the core package never references this global.
  (window as unknown as { __fcInst?: SpreadsheetInstance }).__fcInst = inst;

  // Visual-regression fixtures. `?fixture=cf|sparkline|selection|frozen`
  // replaces the default seed with a deterministic shape.
  const fixtureParam = bootParams.get('fixture');
  if (fixtureParam && isFixtureName(fixtureParam)) {
    applyFixture(fixtureParam, wb, inst);
  }

  filterDropdown = attachFilterDropdown({ store: inst.store });

  // Read-only badge — chart/drawing/pivot counts and spreadsheet Tables. Hidden
  //  until the loaded workbook actually carries any of these objects.
  inst.host.addEventListener('fc:passthroughs', (ev) => {
    const e = ev as CustomEvent<{ count: number; byCategory: Record<string, number> }>;
    refreshObjectsBadge('passthroughs', e.detail);
  });
  inst.host.addEventListener('fc:tables', (ev) => {
    const e = ev as CustomEvent<{ count: number }>;
    refreshObjectsBadge('tables', e.detail);
  });
  // Header chevron click → mount.ts owns the `fc:openfilter` listener and
  // opens its own dropdown. The playground keeps its `filterDropdown` only
  // for the sort menu's "filter" action.

  const engineLabel = wb.isStub ? 'stub engine' : `formulon ${wb.version}`;
  if (enginePill) enginePill.textContent = `engine · ${engineLabel}`;
  if (statusEngine) statusEngine.textContent = engineLabel;
  if (docState) docState.textContent = shellText.saved;
  if (statusState) statusState.textContent = shellText.ready;

  inst.store.subscribe(() => {
    projectStatus();
    projectFormatToolbar();
    markDirty();
    refreshZoom();
  });
  projectStatus();
  projectFormatToolbar();
  renderSheetTabs();
  refreshZoom();

  // Reflect Format Painter state on the toolbar button (any path can deactivate
  // it — Esc, post-paint, or programmatic).
  inst.formatPainter?.subscribe((active, sticky) => {
    formatPainterBtn?.classList.toggle(ACTIVE_CLASS, active);
    formatPainterBtn?.classList.toggle('app__tool--sticky', active && sticky);
  });
}

document.getElementById('btn-autosum')?.addEventListener('click', () => {
  if (!inst) return;
  const result = autoSum(inst.store.getState(), inst.workbook);
  if (!result) return;
  mutators.replaceCells(inst.store, inst.workbook.cells(result.addr.sheet));
  mutators.setActive(inst.store, result.addr);
  (sheetEl as HTMLElement).focus();
});
document.getElementById('btn-hyperlink')?.addEventListener('click', () => {
  inst?.openHyperlinkDialog();
});
const openCommentDialog = (): void => {
  inst?.openCommentDialog();
};
document.getElementById('btn-comment')?.addEventListener('click', openCommentDialog);
document.getElementById('btn-review-comment')?.addEventListener('click', openCommentDialog);
document.getElementById('btn-help-readme')?.addEventListener('click', () => {
  window.open('https://github.com/libraz/formulon-cell#readme', '_blank', 'noopener,noreferrer');
});

document.getElementById('btn-undo')?.addEventListener('click', () => {
  if (!inst) return;
  if (!inst.undo()) return;
  (sheetEl as HTMLElement).focus();
});

document.getElementById('btn-redo')?.addEventListener('click', () => {
  if (!inst) return;
  if (!inst.redo()) return;
  (sheetEl as HTMLElement).focus();
});

// Format Painter — single click arms one-shot, double click arms sticky mode.
// Re-clicking the active button deactivates.
const formatPainterBtn = document.getElementById('btn-format-painter');
let painterStickyTimer: number | null = null;
formatPainterBtn?.addEventListener('click', () => {
  if (!inst) return;
  // Defer one-shot activation briefly so a follow-up click within the
  // dblclick window can promote it to sticky without painting twice.
  if (painterStickyTimer != null) return;
  painterStickyTimer = window.setTimeout(() => {
    painterStickyTimer = null;
    if (!inst) return;
    const fp = inst.formatPainter;
    if (!fp) return;
    if (fp.isActive()) fp.deactivate();
    else fp.activate(false);
    (sheetEl as HTMLElement).focus();
    formatPainterBtn?.classList.toggle(ACTIVE_CLASS, fp.isActive());
  }, 220);
});
formatPainterBtn?.addEventListener('dblclick', () => {
  if (!inst) return;
  if (painterStickyTimer != null) {
    clearTimeout(painterStickyTimer);
    painterStickyTimer = null;
  }
  const fp = inst.formatPainter;
  if (!fp) return;
  fp.activate(true);
  (sheetEl as HTMLElement).focus();
  formatPainterBtn?.classList.toggle(ACTIVE_CLASS, fp.isActive());
});

const wireFormat = (
  id: string,
  fn: (
    state: ReturnType<SpreadsheetInstance['store']['getState']>,
    store: SpreadsheetInstance['store'],
  ) => void,
): void => {
  document.getElementById(id)?.addEventListener('click', () => {
    const i = inst;
    if (!i) return;
    // Wrap each toolbar mutation so Cmd+Z reverts the format change.
    recordFormatChange(i.history, i.store, () => {
      fn(i.store.getState(), i.store);
    });
    (sheetEl as HTMLElement).focus();
  });
};

wireFormat('btn-bold', toggleBold);
wireFormat('btn-italic', toggleItalic);
wireFormat('btn-underline', toggleUnderline);
wireFormat('btn-strike', toggleStrike);
wireFormat('btn-percent', cyclePercent);
wireFormat('btn-comma', (state, store) =>
  setNumFmt(state, store, { kind: 'fixed', decimals: 2, thousands: true }),
);
wireFormat('btn-font-grow', (state, store) => {
  const a = state.selection.active;
  const f = state.format.formats.get(`${a.sheet}:${a.row}:${a.col}`);
  setFont(state, store, { fontSize: (f?.fontSize ?? 11) + 1 });
});
wireFormat('btn-font-shrink', (state, store) => {
  const a = state.selection.active;
  const f = state.format.formats.get(`${a.sheet}:${a.row}:${a.col}`);
  setFont(state, store, { fontSize: Math.max(1, (f?.fontSize ?? 11) - 1) });
});
wireFormat('btn-align-left', (state, store) => setAlign(state, store, 'left'));
wireFormat('btn-align-center', (state, store) => setAlign(state, store, 'center'));
wireFormat('btn-align-right', (state, store) => setAlign(state, store, 'right'));
wireFormat('btn-top', (state, store) => setVAlign(state, store, 'top'));
wireFormat('btn-middle', (state, store) => setVAlign(state, store, 'middle'));
wireFormat('btn-decimals-up', (state, store) => bumpDecimals(state, store, 1));
wireFormat('btn-decimals-down', (state, store) => bumpDecimals(state, store, -1));

void clearFormat; // Reserved for a "Clear formatting" menu item; keep the import live.

// ── Borders dropdown (Excel-365 parity) ──────────────────────────────────
// Wires the integrated dropdown built by createBordersMenu(): edge / frame
// / combined presets, "More borders..." entry, and the "罫線の作成" block
// which arms the border-draw controller (drives pointer-edge editing on
// the grid) and exposes two submenus for the line color & line style
// brush settings.
const borderBtn = document.getElementById('btn-borders');
const borderMenu = document.getElementById('menu-borders');
const lineStyleSubmenu =
  borderMenu?.querySelector<HTMLElement>('.app__submenu--line-style') ?? null;
const getBorderBtn = (): HTMLButtonElement | null =>
  document.getElementById('btn-borders') as HTMLButtonElement | null;
const getBorderMenu = (): HTMLDivElement | null =>
  document.getElementById('menu-borders') as HTMLDivElement | null;
const getLineColorSubmenu = (): HTMLElement | null =>
  getBorderMenu()?.querySelector<HTMLElement>('.app__submenu--line-color') ?? null;
const getLineStyleSubmenu = (): HTMLElement | null =>
  getBorderMenu()?.querySelector<HTMLElement>('.app__submenu--line-style') ?? null;

const BORDER_DRAW_ACTIVE_CLASS = 'app__menu-item--active';

const closeBorderSubmenus = (): void => {
  const lineColor = getLineColorSubmenu();
  const lineStyle = getLineStyleSubmenu();
  if (lineColor) lineColor.hidden = true;
  if (lineStyle) lineStyle.hidden = true;
  getBorderMenu()
    ?.querySelectorAll<HTMLButtonElement>('[data-border-submenu]')
    .forEach((b) => {
      b.setAttribute('aria-expanded', 'false');
    });
};

const closeBorderMenu = (restoreFocus = false): void => {
  const menu = getBorderMenu();
  const btn = getBorderBtn();
  if (!menu) return;
  menu.hidden = true;
  btn?.setAttribute('aria-expanded', 'false');
  closeBorderSubmenus();
  if (restoreFocus) btn?.focus();
};

const refreshBorderMenuState = (): void => {
  const menu = getBorderMenu();
  if (!menu) return;
  // Reflect currently-armed draw mode in the menu so the user can see
  // (and toggle off) the active brush.
  const mode = inst?.borderDraw?.getMode() ?? null;
  menu.querySelectorAll<HTMLButtonElement>('[data-border-draw]').forEach((btn) => {
    const armed = btn.dataset.borderDraw === mode;
    btn.classList.toggle(BORDER_DRAW_ACTIVE_CLASS, armed);
    btn.setAttribute('aria-checked', armed ? 'true' : 'false');
  });
};

const openBorderMenu = (): void => {
  const menu = getBorderMenu();
  const btn = getBorderBtn();
  if (!menu) return;
  refreshBorderMenuState();
  menu.hidden = false;
  btn?.setAttribute('aria-expanded', 'true');
  focusMenuItem(menu);
};

borderBtn?.addEventListener('click', (e) => {
  e.stopPropagation();
  if (!borderMenu) return;
  if (borderMenu.hidden) openBorderMenu();
  else closeBorderMenu();
});

document.addEventListener('mousedown', (e) => {
  const menu = getBorderMenu();
  const btn = getBorderBtn();
  if (!menu || menu.hidden) return;
  if (menu.contains(e.target as Node)) return;
  if (btn?.contains(e.target as Node)) return;
  closeBorderMenu();
});

document.addEventListener('keydown', (e) => {
  const menu = getBorderMenu();
  if (e.key === 'Escape' && !menu?.hidden) closeBorderMenu(true);
});

borderMenu?.addEventListener('keydown', (e) => {
  handleMenuKeydown(e, borderMenu, { close: closeBorderMenu, restoreFocusTo: borderBtn });
});

document.addEventListener('keydown', (event) => {
  const target = event.target as Element | null;
  const menu = target?.closest<HTMLDivElement>('#menu-borders');
  if (!menu || menu === borderMenu) return;
  handleMenuKeydown(event, menu, { close: closeBorderMenu, restoreFocusTo: getBorderBtn() });
});

type BorderPresetKey =
  | 'none'
  | 'outline'
  | 'thickOutline'
  | 'all'
  | 'top'
  | 'bottom'
  | 'left'
  | 'right'
  | 'doubleBottom'
  | 'thickBottom'
  | 'topAndBottom'
  | 'topAndThickBottom'
  | 'topAndDoubleBottom';

// Map menu key → engine preset. `clear` is the "罫線なし" entry: the
// engine's `'none'` preset wipes every side.
const MENU_TO_PRESET: Record<string, BorderPresetKey> = {
  clear: 'none',
  all: 'all',
  outline: 'outline',
  thickOutline: 'thickOutline',
  top: 'top',
  bottom: 'bottom',
  left: 'left',
  right: 'right',
  doubleBottom: 'doubleBottom',
  thickBottom: 'thickBottom',
  topAndBottom: 'topAndBottom',
  topAndThickBottom: 'topAndThickBottom',
  topAndDoubleBottom: 'topAndDoubleBottom',
};

const applyBorderPresetMenuAction = (key: string): void => {
  const i = inst;
  if (!i) return;
  if (key === 'format') {
    closeBorderMenu();
    i.openFormatDialog();
    return;
  }
  const preset = MENU_TO_PRESET[key];
  if (!preset) return;
  closeBorderMenu();
  i.borderDraw?.deactivate();
  applyRibbonFormat((state, store) => setBorderPreset(state, store, preset, selectedBorderStyle));
};

const applyBorderDrawMenuAction = (action: string | undefined): void => {
  const i = inst;
  if (!i) return;
  if (action !== 'draw' && action !== 'grid' && action !== 'erase') return;
  const draw = i.borderDraw;
  if (!draw) return;
  if (draw.getMode() === action) {
    draw.deactivate();
  } else {
    draw.activate(action, selectedBorderStyle, selectedBorderColor);
  }
  closeBorderMenu();
  refreshBorderMenuState();
  (sheetEl as HTMLElement).focus();
};

borderMenu?.querySelectorAll<HTMLButtonElement>('[data-border-preset]').forEach((btn) => {
  btn.addEventListener('click', () => {
    applyBorderPresetMenuAction(btn.dataset.borderPreset ?? '');
  });
});

borderMenu?.querySelectorAll<HTMLButtonElement>('[data-border-draw]').forEach((btn) => {
  btn.addEventListener('click', () => {
    applyBorderDrawMenuAction(btn.dataset.borderDraw);
  });
});

const openSubmenu = (which: 'lineColor' | 'lineStyle'): void => {
  const menu = getBorderMenu();
  const lineColor = getLineColorSubmenu();
  const lineStyle = getLineStyleSubmenu();
  if (which === 'lineColor') {
    if (lineStyle) lineStyle.hidden = true;
    if (lineColor) lineColor.hidden = false;
  } else {
    if (lineColor) lineColor.hidden = true;
    if (lineStyle) lineStyle.hidden = false;
  }
  menu?.querySelectorAll<HTMLButtonElement>('[data-border-submenu]').forEach((b) => {
    b.setAttribute('aria-expanded', b.dataset.borderSubmenu === which ? 'true' : 'false');
  });
};

borderMenu?.querySelectorAll<HTMLButtonElement>('[data-border-submenu]').forEach((btn) => {
  btn.addEventListener('mouseenter', () => {
    const which = btn.dataset.borderSubmenu as 'lineColor' | 'lineStyle' | undefined;
    if (which) openSubmenu(which);
  });
  btn.addEventListener('click', (e) => {
    e.stopPropagation();
    const which = btn.dataset.borderSubmenu as 'lineColor' | 'lineStyle' | undefined;
    if (which) openSubmenu(which);
  });
});

// Mousing over a non-submenu item dismisses any open submenu — matches
// Excel's single-active-submenu behavior.
borderMenu
  ?.querySelectorAll<HTMLButtonElement>('[data-border-preset], [data-border-draw]')
  .forEach((btn) => {
    btn.addEventListener('mouseenter', closeBorderSubmenus);
  });

// Line-color picks are handled by the shared palette's onPick callback,
// wired in createLineColorSubmenu().

lineStyleSubmenu
  ?.querySelectorAll<HTMLButtonElement>('[data-border-line-style]')
  .forEach((styleBtn) => {
    styleBtn.addEventListener('click', () => {
      const value = styleBtn.dataset.borderLineStyle ?? 'thin';
      if (value !== 'none') {
        selectedBorderStyle = value as CellBorderStyle;
        inst?.borderDraw?.setStyle(selectedBorderStyle);
      }
      lineStyleSubmenu
        .querySelectorAll<HTMLButtonElement>('[data-border-line-style]')
        .forEach((s) => {
          s.setAttribute('aria-checked', s === styleBtn ? 'true' : 'false');
        });
      closeBorderSubmenus();
    });
  });

document.addEventListener('click', (event) => {
  const target = event.target as Element | null;
  const menu = target?.closest<HTMLElement>('#menu-borders');
  if (!menu || menu === borderMenu) return;
  const preset = target?.closest<HTMLButtonElement>('[data-border-preset]');
  if (preset) {
    event.preventDefault();
    applyBorderPresetMenuAction(preset.dataset.borderPreset ?? '');
    return;
  }
  const draw = target?.closest<HTMLButtonElement>('[data-border-draw]');
  if (draw) {
    event.preventDefault();
    applyBorderDrawMenuAction(draw.dataset.borderDraw);
    return;
  }
  const submenu = target?.closest<HTMLButtonElement>('[data-border-submenu]');
  if (submenu) {
    event.preventDefault();
    const which = submenu.dataset.borderSubmenu as 'lineColor' | 'lineStyle' | undefined;
    if (which) openSubmenu(which);
    return;
  }
  const lineStyle = target?.closest<HTMLButtonElement>('[data-border-line-style]');
  if (lineStyle) {
    event.preventDefault();
    const value = lineStyle.dataset.borderLineStyle ?? 'thin';
    const lineStyleMenu = getLineStyleSubmenu();
    if (value !== 'none') {
      selectedBorderStyle = value as CellBorderStyle;
      inst?.borderDraw?.setStyle(selectedBorderStyle);
    }
    lineStyleMenu?.querySelectorAll<HTMLButtonElement>('[data-border-line-style]').forEach((s) => {
      s.setAttribute('aria-checked', s === lineStyle ? 'true' : 'false');
    });
    closeBorderSubmenus();
  }
});

document.addEventListener('mouseover', (event) => {
  const target = event.target as Element | null;
  const menu = target?.closest<HTMLElement>('#menu-borders');
  if (!menu || menu === borderMenu) return;
  const submenu = target?.closest<HTMLButtonElement>('[data-border-submenu]');
  if (submenu) {
    const which = submenu.dataset.borderSubmenu as 'lineColor' | 'lineStyle' | undefined;
    if (which) openSubmenu(which);
    return;
  }
  if (target?.closest('[data-border-preset], [data-border-draw]')) closeBorderSubmenus();
});

// ── Freeze Panes menu ─────────────────────────────────────────────────────
const freezeBtn = document.getElementById('btn-freeze');
const freezeMenu = document.getElementById('menu-freeze');
const getFreezeBtn = (): HTMLButtonElement | null =>
  document.getElementById('btn-freeze') as HTMLButtonElement | null;
const getFreezeMenu = (): HTMLDivElement | null =>
  document.getElementById('menu-freeze') as HTMLDivElement | null;

const closeFreezeMenu = (restoreFocus = false): void => {
  const menu = getFreezeMenu();
  const btn = getFreezeBtn();
  if (!menu) return;
  menu.hidden = true;
  btn?.setAttribute('aria-expanded', 'false');
  if (restoreFocus) btn?.focus();
};
const openFreezeMenu = (): void => {
  const menu = getFreezeMenu();
  const btn = getFreezeBtn();
  if (!menu) return;
  menu.hidden = false;
  btn?.setAttribute('aria-expanded', 'true');
  focusMenuItem(menu);
};

freezeBtn?.addEventListener('click', (e) => {
  e.stopPropagation();
  if (!freezeMenu) return;
  if (freezeMenu.hidden) openFreezeMenu();
  else closeFreezeMenu();
});

document.addEventListener('mousedown', (e) => {
  const menu = getFreezeMenu();
  const btn = getFreezeBtn();
  if (!menu || menu.hidden) return;
  if (menu.contains(e.target as Node)) return;
  if (btn?.contains(e.target as Node)) return;
  closeFreezeMenu();
});

document.addEventListener('keydown', (e) => {
  const menu = getFreezeMenu();
  if (e.key === 'Escape' && !menu?.hidden) closeFreezeMenu(true);
});

freezeMenu?.addEventListener('keydown', (e) => {
  handleMenuKeydown(e, freezeMenu, { close: closeFreezeMenu, restoreFocusTo: freezeBtn });
});

document.addEventListener('keydown', (event) => {
  const target = event.target as Element | null;
  const menu = target?.closest<HTMLDivElement>('#menu-freeze');
  if (!menu || menu === freezeMenu) return;
  handleMenuKeydown(event, menu, { close: closeFreezeMenu, restoreFocusTo: getFreezeBtn() });
});

const applyFreezeMenuAction = (action: string | undefined): void => {
  const i = inst;
  if (!i) return;
  const s = i.store.getState();

  let rows = s.layout.freezeRows;
  let cols = s.layout.freezeCols;
  if (action === 'row') {
    rows = 1;
    cols = 0;
  } else if (action === 'col') {
    rows = 0;
    cols = 1;
  } else if (action === 'selection') {
    // Freeze rows above and columns left of the active cell.
    rows = s.selection.active.row;
    cols = s.selection.active.col;
  } else if (action === 'off') {
    rows = 0;
    cols = 0;
  } else {
    return;
  }

  setFreezePanes(i.store, i.history, rows, cols, i.workbook);
  closeFreezeMenu();
  (sheetEl as HTMLElement).focus();
};

freezeMenu?.querySelectorAll<HTMLButtonElement>('[data-freeze]').forEach((btn) => {
  btn.addEventListener('click', () => {
    applyFreezeMenuAction(btn.dataset.freeze);
  });
});

document.addEventListener('click', (event) => {
  const target = event.target as Element | null;
  const menu = target?.closest<HTMLElement>('#menu-freeze');
  if (!menu || menu === freezeMenu) return;
  const item = target?.closest<HTMLButtonElement>('[data-freeze]');
  if (!item) return;
  event.preventDefault();
  applyFreezeMenuAction(item.dataset.freeze);
});

themeToggle?.addEventListener('click', () => {
  applyUiTheme(uiTheme === 'dark' ? 'light' : 'dark');
});

// ── File menu (New / Open / Save / Save As) ───────────────────────────────
const fileMenuBtn = document.getElementById('menu-file');
const fileMenuDrop = document.getElementById('menu-file-dropdown');
const fileInput = document.getElementById('file-input') as HTMLInputElement | null;

let docName = 'Book1';

const setDocName = (name: string): void => {
  docName = name;
  const el = document.getElementById('doc-name');
  if (el) el.textContent = name;
};

const openFileMenu = (): void => {
  if (!fileMenuDrop) return;
  fileMenuDrop.hidden = false;
  fileMenuBtn?.setAttribute('aria-expanded', 'true');
};
const closeFileMenu = (): void => {
  if (!fileMenuDrop) return;
  fileMenuDrop.hidden = true;
  fileMenuBtn?.setAttribute('aria-expanded', 'false');
};

fileMenuBtn?.addEventListener('click', (e) => {
  e.stopPropagation();
  if (!fileMenuDrop) return;
  if (fileMenuDrop.hidden) openFileMenu();
  else closeFileMenu();
});

document.addEventListener('mousedown', (e) => {
  if (!fileMenuDrop || fileMenuDrop.hidden) return;
  if (fileMenuDrop.contains(e.target as Node)) return;
  if (fileMenuBtn?.contains(e.target as Node)) return;
  closeFileMenu();
});

document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && !fileMenuDrop?.hidden) closeFileMenu();
});

const triggerOpen = (): void => fileInput?.click();

const downloadBytes = (bytes: Uint8Array, filename: string): void => {
  const blob = new Blob([bytes as BlobPart], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1_000);
};

const triggerSave = (filename = `${docName.replace(/\.xlsx$/i, '')}.xlsx`): void => {
  if (!inst) return;
  try {
    const bytes = inst.workbook.save();
    downloadBytes(bytes, filename);
    if (docState) docState.textContent = shellText.saved;
  } catch (err) {
    // eslint-disable-next-line no-console
    console.error('save failed', err);
    if (docState) docState.textContent = shellText.saveFailed;
  }
};

const triggerSaveAs = async (): Promise<void> => {
  const name = await showPrompt({
    title: shellText.saveAs,
    label: shellText.fileName,
    initial: docName,
    okLabel: shellText.save,
    validate: (value) => (value.trim() ? null : shellText.enterFileName),
  });
  if (!name) return;
  const trimmed = name.trim();
  setDocName(trimmed);
  triggerSave(trimmed.endsWith('.xlsx') ? trimmed : `${trimmed}.xlsx`);
};

const inspectWorkbookFromBackstage = (): void => {
  const i = inst;
  if (!i) return;
  const summary = summarizeSpreadsheetCompatibility(i.workbook);
  const objectsText = dictionaries[ribbonLang].workbookObjects;
  const compatibilityLabel = (id: (typeof summary.items)[number]['id']): string => {
    switch (id) {
      case 'cell-formatting':
        return objectsText.compatibilityLabels.cellFormatting;
      case 'conditional-formatting':
        return objectsText.compatibilityLabels.conditionalFormatting;
      case 'data-validation':
        return objectsText.compatibilityLabels.dataValidation;
      case 'hyperlinks':
        return objectsText.compatibilityLabels.hyperlinks;
      case 'comments':
        return objectsText.compatibilityLabels.comments;
      case 'defined-names':
        return objectsText.compatibilityLabels.definedNames;
      case 'sheet-protection':
        return objectsText.compatibilityLabels.sheetProtection;
      case 'sheet-views':
        return objectsText.compatibilityLabels.sheetViews;
      case 'loaded-tables':
        return objectsText.compatibilityLabels.loadedTables;
      case 'format-as-table':
        return objectsText.compatibilityLabels.formatAsTable;
      case 'pivot-layouts':
        return objectsText.compatibilityLabels.pivotLayouts;
      case 'pivot-authoring':
        return objectsText.compatibilityLabels.pivotAuthoring;
      case 'session-charts':
        return objectsText.compatibilityLabels.sessionCharts;
      case 'charts-drawings':
        return objectsText.compatibilityLabels.chartsDrawings;
      case 'chart-authoring':
        return objectsText.compatibilityLabels.chartAuthoring;
      case 'external-links':
        return objectsText.compatibilityLabels.externalLinks;
    }
  };
  const compatibilityDetail = (id: (typeof summary.items)[number]['id']): string => {
    switch (id) {
      case 'cell-formatting':
        return objectsText.compatibilityDetails.cellFormatting;
      case 'conditional-formatting':
        return objectsText.compatibilityDetails.conditionalFormatting;
      case 'data-validation':
        return objectsText.compatibilityDetails.dataValidation;
      case 'hyperlinks':
        return objectsText.compatibilityDetails.hyperlinks;
      case 'comments':
        return objectsText.compatibilityDetails.comments;
      case 'defined-names':
        return objectsText.compatibilityDetails.definedNames;
      case 'sheet-protection':
        return objectsText.compatibilityDetails.sheetProtection;
      case 'sheet-views':
        return objectsText.compatibilityDetails.sheetViews;
      case 'loaded-tables':
        return objectsText.compatibilityDetails.loadedTables;
      case 'format-as-table':
        return objectsText.compatibilityDetails.formatAsTable;
      case 'pivot-layouts':
        return objectsText.compatibilityDetails.pivotLayouts;
      case 'pivot-authoring':
        return objectsText.compatibilityDetails.pivotAuthoring;
      case 'session-charts':
        return objectsText.compatibilityDetails.sessionCharts;
      case 'charts-drawings':
        return objectsText.compatibilityDetails.chartsDrawings;
      case 'chart-authoring':
        return objectsText.compatibilityDetails.chartAuthoring;
      case 'external-links':
        return objectsText.compatibilityDetails.externalLinks;
    }
  };
  const statusLabel = (status: keyof typeof summary.byStatus): string => {
    if (status === 'writable') return objectsText.writable;
    if (status === 'read-only') return objectsText.readOnly;
    if (status === 'session') return objectsText.sessionOnly;
    return objectsText.unsupported;
  };
  showRibbonReport(dictionaries[ribbonLang].backstage.inspect, [
    {
      severity: 'info',
      label: objectsText.compatibility,
      detail: `${objectsText.writable} ${summary.byStatus.writable}, ${objectsText.readOnly} ${summary.byStatus['read-only']}, ${objectsText.sessionOnly} ${summary.byStatus.session}, ${objectsText.unsupported} ${summary.byStatus.unsupported}`,
    },
    ...summary.items.map((item) => ({
      severity:
        item.status === 'unsupported' || item.status === 'read-only'
          ? ('warning' as const)
          : ('info' as const),
      label: `${compatibilityLabel(item.id)} · ${statusLabel(item.status)}`,
      detail: item.count
        ? `${compatibilityDetail(item.id)} (${item.count})`
        : compatibilityDetail(item.id),
    })),
  ]);
};

const loadXlsxFile = async (file: File): Promise<void> => {
  if (!inst) return;
  if (docState) docState.textContent = shellText.loading;
  try {
    const buf = await file.arrayBuffer();
    const next = await WorkbookHandle.loadBytes(new Uint8Array(buf));
    await inst.setWorkbook(next);
    setDocName(file.name);
    if (docState) docState.textContent = shellText.saved;
    renderSheetTabs();
  } catch (err) {
    // eslint-disable-next-line no-console
    console.error('open failed', err);
    if (docState) docState.textContent = shellText.openFailed;
    void showMessage({
      title: shellText.openFailed,
      message: err instanceof Error ? err.message : String(err),
    });
  }
};

fileInput?.addEventListener('change', () => {
  const f = fileInput.files?.[0];
  if (f) void loadXlsxFile(f);
  fileInput.value = ''; // allow same-file re-open
});

fileMenuDrop?.querySelectorAll<HTMLButtonElement>('[data-file]').forEach((btn) => {
  btn.addEventListener('click', () => {
    const action = btn.dataset.file;
    closeFileMenu();
    if (!inst) return;
    if (action === 'new') {
      void (async () => {
        const next = await WorkbookHandle.createDefault();
        await inst?.setWorkbook(next);
        setDocName('Book1');
        if (docState) docState.textContent = shellText.saved;
        renderSheetTabs();
      })();
    } else if (action === 'open') {
      triggerOpen();
    } else if (action === 'save') {
      triggerSave();
    } else if (action === 'save-as') {
      void triggerSaveAs();
    }
  });
});

ribbonRoot?.addEventListener('click', (event) => {
  const button = (event.target as Element | null)?.closest<HTMLButtonElement>(
    '[data-backstage-action]',
  );
  if (!button || button.disabled) return;
  const action = button.dataset.backstageAction;
  if (!action || action === 'info') return;
  event.preventDefault();
  event.stopPropagation();
  if (action === 'back') {
    closeBackstage(true);
  } else if (action === 'new') {
    closeBackstage();
    void (async () => {
      const next = await WorkbookHandle.createDefault();
      await inst?.setWorkbook(next);
      setDocName('Book1');
      if (docState) docState.textContent = shellText.saved;
      renderSheetTabs();
    })();
  } else if (action === 'open') {
    closeBackstage();
    triggerOpen();
  } else if (action === 'save') {
    closeBackstage();
    triggerSave();
  } else if (action === 'save-as') {
    closeBackstage();
    void triggerSaveAs();
  } else if (action === 'print') {
    closeBackstage();
    inst?.print('print');
  } else if (action === 'options') {
    closeBackstage();
    inst?.openIterativeDialog();
  } else if (action === 'protect-workbook') {
    closeBackstage();
    void applyProtectAction(
      inst && isWorkbookStructureProtected(inst.store.getState())
        ? 'unprotect-workbook'
        : 'protect-workbook',
    );
  } else if (action === 'inspect-workbook') {
    closeBackstage();
    inspectWorkbookFromBackstage();
  } else if (action === 'links') {
    closeBackstage();
    inst?.openExternalLinksDialog();
  }
});

// Drag & drop xlsx onto the page.
window.addEventListener('dragover', (e) => {
  if (!e.dataTransfer) return;
  e.preventDefault();
  e.dataTransfer.dropEffect = 'copy';
});
window.addEventListener('drop', (e) => {
  e.preventDefault();
  const f = e.dataTransfer?.files?.[0];
  if (!f) return;
  if (!/\.xlsx?$/i.test(f.name)) return;
  void loadXlsxFile(f);
});

// Ctrl/Cmd-O / Ctrl/Cmd-S / Ctrl/Cmd-N for file actions.
window.addEventListener('keydown', (e) => {
  if (!(e.ctrlKey || e.metaKey)) return;
  const k = e.key.toLowerCase();
  if (k === 'o') {
    e.preventDefault();
    triggerOpen();
  } else if (k === 's') {
    e.preventDefault();
    if (e.shiftKey) void triggerSaveAs();
    else triggerSave();
  } else if (k === 'n' && !e.shiftKey) {
    // Ctrl+N — create a fresh workbook in place.
    e.preventDefault();
    void (async () => {
      const next = await WorkbookHandle.createDefault();
      await inst?.setWorkbook(next);
      setDocName('Book1');
      renderSheetTabs();
    })();
  }
});

// Mark the document dirty whenever any cell change flows through.
let dirtyTimer: number | null = null;
let suppressDirty = false;
const markDirty = (): void => {
  if (suppressDirty) return;
  if (dirtyTimer != null) return;
  dirtyTimer = window.setTimeout(() => {
    dirtyTimer = null;
    if (docState) docState.textContent = shellText.edited;
  }, 200);
};
// Subscribe once boot completes — see end of boot().

const refreshWorkbookCells = (): void => {
  if (!inst) return;
  mutators.replaceCells(inst.store, inst.workbook.cells(inst.store.getState().data.sheetIndex));
};

const focusSheet = (): void => {
  (sheetEl as HTMLElement).focus();
};

const selectedRowCount = (): number => {
  if (!inst) return 1;
  const r = inst.store.getState().selection.range;
  return Math.max(1, r.r1 - r.r0 + 1);
};

const selectedColCount = (): number => {
  if (!inst) return 1;
  const r = inst.store.getState().selection.range;
  return Math.max(1, r.c1 - r.c0 + 1);
};

const openFilterForSelection = (): void => {
  const i = inst;
  if (!i) return;
  const r = inferAutoFilterRange(i.store.getState());
  const active = i.store.getState().ui.filterRange;
  const sameActive =
    active != null &&
    active.sheet === r.sheet &&
    active.r0 === r.r0 &&
    active.c0 === r.c0 &&
    active.r1 === r.r1 &&
    active.c1 === r.c1;
  recordFilterChange(i.history, i.store, () => {
    if (sameActive) clearFilter(i.store.getState(), i.store, r);
    else setAutoFilter(i.store, r);
  });
  if (sameActive) {
    focusSheet();
    return;
  }
  const sheetRect = (sheetEl as HTMLElement).getBoundingClientRect();
  filterDropdown?.open(r, r.c0, { x: sheetRect.left + 80, y: sheetRect.top, h: 32 });
  focusSheet();
};

const sortTargetRange = (state: ReturnType<SpreadsheetInstance['store']['getState']>): Range => {
  const r = state.selection.range;
  if (r.r0 === r.r1 && r.c0 === r.c1) return inferAutoFilterRange(state);
  return r;
};

const sortCellDisplayText = (
  state: ReturnType<SpreadsheetInstance['store']['getState']>,
  row: number,
  col: number,
): string => {
  const value = state.data.cells.get(`${state.selection.active.sheet}:${row}:${col}`)?.value;
  if (!value) return '';
  if (value.kind === 'number') return String(value.value);
  if (value.kind === 'text') return value.value;
  if (value.kind === 'bool') return value.value ? 'TRUE' : 'FALSE';
  if (value.kind === 'error') return value.text;
  return '';
};

const colFromLetters = (letters: string): number => {
  let col = 0;
  const upper = letters.toUpperCase();
  for (let i = 0; i < upper.length; i += 1) {
    const code = upper.charCodeAt(i);
    if (code < 65 || code > 90) return -1;
    col = col * 26 + (code - 64);
  }
  return col - 1;
};

const parseA1Range = (raw: string, sheet: number): Range | null => {
  const normalized = raw.replace(/\$/g, '').trim().toUpperCase();
  const match = normalized.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
  if (!match) return null;
  const c0 = colFromLetters(match[1] ?? '');
  const r0 = Number(match[2]) - 1;
  const c1 = match[3] ? colFromLetters(match[3]) : c0;
  const r1 = match[4] ? Number(match[4]) - 1 : r0;
  if (c0 < 0 || c1 < 0 || r0 < 0 || r1 < 0) return null;
  return {
    sheet,
    r0: Math.min(r0, r1),
    c0: Math.min(c0, c1),
    r1: Math.max(r0, r1),
    c1: Math.max(c0, c1),
  };
};

const rangeRef = (range: Range): string => {
  const start = `${colLetter(range.c0)}${range.r0 + 1}`;
  const end = `${colLetter(range.c1)}${range.r1 + 1}`;
  return start === end ? start : `${start}:${end}`;
};

const syncStoreCellsToWorkbook = (
  i: SpreadsheetInstance,
  sheet: number,
  row: number,
  col: number,
  height: number,
  width: number,
): void => {
  const cells = i.store.getState().data.cells;
  for (let r = row; r < row + height; r += 1) {
    for (let c = col; c < col + width; c += 1) {
      const addr = { sheet, row: r, col: c };
      const cell = cells.get(`${sheet}:${r}:${c}`);
      if (!cell) {
        i.workbook.setBlank(addr);
      } else if (cell.formula) {
        i.workbook.setFormula(addr, cell.formula);
      } else if (cell.value.kind === 'number') {
        i.workbook.setNumber(addr, cell.value.value);
      } else if (cell.value.kind === 'text') {
        i.workbook.setText(addr, cell.value.value);
      } else if (cell.value.kind === 'bool') {
        i.workbook.setBool(addr, cell.value.value);
      } else {
        i.workbook.setBlank(addr);
      }
    }
  }
};

const applyAdvancedFilterAction = async (): Promise<void> => {
  const i = inst;
  if (!i) return;
  const state = i.store.getState();
  const listInitial = rangeRef(state.ui.filterRange ?? inferAutoFilterRange(state));
  const result = await showAdvancedFilterDialog({
    title: ribbonMenuText.advancedFilterDialogTitle,
    listRangeLabel: ribbonMenuText.advancedFilterListRange,
    criteriaRangeLabel: ribbonMenuText.advancedFilterCriteriaRange,
    copyToLabel: ribbonMenuText.advancedFilterCopyTo,
    uniqueOnlyLabel: ribbonMenuText.advancedFilterUniqueOnly,
    initialListRange: listInitial,
    okLabel: 'OK',
    cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
    validateListRange: (value) =>
      parseA1Range(value, state.selection.active.sheet)
        ? null
        : ribbonLang === 'ja'
          ? 'A1:B10 の形式で入力してください。'
          : 'Enter a list range such as A1:B10.',
    validateCriteriaRange: (value) =>
      parseA1Range(value, state.selection.active.sheet)
        ? null
        : ribbonLang === 'ja'
          ? 'A1:B3 の形式で入力してください。'
          : 'Enter a criteria range such as A1:B3.',
    validateCopyTo: (value) => {
      if (!value.trim()) return null;
      return parseA1Range(value, state.selection.active.sheet)
        ? null
        : ribbonLang === 'ja'
          ? 'A1 の形式で入力してください。'
          : 'Enter a cell such as A1.';
    },
  });
  if (result === null) {
    focusSheet();
    return;
  }
  const listRange = parseA1Range(result.listRange, state.selection.active.sheet);
  const criteriaRange = parseA1Range(result.criteriaRange, state.selection.active.sheet);
  if (!listRange || !criteriaRange) return;
  const copyRange = result.copyTo
    ? parseA1Range(result.copyTo, state.selection.active.sheet)
    : null;
  if (copyRange) {
    let copied = 0;
    i.history.begin();
    try {
      copied = copyAdvancedFilterResult(
        i.store.getState(),
        i.store,
        listRange,
        criteriaRange,
        { sheet: copyRange.sheet, row: copyRange.r0, col: copyRange.c0 },
        { uniqueOnly: result.uniqueOnly },
      );
      syncStoreCellsToWorkbook(
        i,
        copyRange.sheet,
        copyRange.r0,
        copyRange.c0,
        copied,
        listRange.c1 - listRange.c0 + 1,
      );
    } finally {
      i.history.end();
    }
    if (statusMetric) {
      statusMetric.textContent = ribbonMenuText.advancedFilterCopiedStatus.replace(
        '{count}',
        String(copied),
      );
    }
  } else {
    recordFilterChange(i.history, i.store, () => {
      applyAdvancedFilter(i.store.getState(), i.store, listRange, criteriaRange);
    });
  }
  focusSheet();
};

const sortSelection = (direction: 'asc' | 'desc'): void => {
  const i = inst;
  if (!i) return;
  const state = i.store.getState();
  const r = sortTargetRange(state);
  if (r.r0 === r.r1) return;
  const activeCol = state.selection.active.col;
  const byCol = activeCol >= r.c0 && activeCol <= r.c1 ? activeCol : r.c0;
  const hasHeader = inferSortHasHeader(state, r);
  let sorted = false;
  i.history.begin();
  try {
    recordFormatChange(i.history, i.store, () => {
      sorted = sortRange(state, i.store, i.workbook, r, { byCol, direction, hasHeader });
    });
  } finally {
    i.history.end();
  }
  if (sorted) refreshWorkbookCells();
  focusSheet();
};

const customSortSelection = async (): Promise<void> => {
  const i = inst;
  if (!i) return;
  const state = i.store.getState();
  const range = sortTargetRange(state);
  if (range.r0 === range.r1) {
    focusSheet();
    return;
  }
  const inferredHeader = inferSortHasHeader(state, range);
  const activeCol =
    state.selection.active.col >= range.c0 && state.selection.active.col <= range.c1
      ? state.selection.active.col
      : range.c0;
  const columns = Array.from({ length: range.c1 - range.c0 + 1 }, (_, offset) => {
    const col = range.c0 + offset;
    const letter = colLetter(col);
    const header = inferredHeader ? sortCellDisplayText(state, range.r0, col).trim() : '';
    return {
      value: String(col),
      label: header ? `${header} (${letter})` : letter,
    };
  });
  const result = await showSortDialog({
    title: ribbonMenuText.sortCustom,
    columnLabel: ribbonMenuText.sortColumn,
    thenByLabel: ribbonMenuText.sortThenBy,
    noThenByLabel: ribbonMenuText.sortNoThenBy,
    orderLabel: ribbonMenuText.sortOrder,
    headerLabel: ribbonMenuText.sortMyDataHasHeaders,
    ascendingLabel: ribbonMenuText.sortAscendingMenu,
    descendingLabel: ribbonMenuText.sortDescendingMenu,
    columns,
    initialColumn: String(activeCol),
    initialDirection: 'asc',
    initialHasHeader: inferredHeader,
    okLabel: 'OK',
    cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
  });
  if (!result) {
    focusSheet();
    return;
  }
  let sorted = false;
  const byCol = Number(result.column);
  i.history.begin();
  try {
    recordFormatChange(i.history, i.store, () => {
      sorted = sortRange(i.store.getState(), i.store, i.workbook, range, {
        byCol,
        direction: result.direction,
        hasHeader: result.hasHeader,
        keys: result.levels.map((level) => ({
          byCol: Number(level.column),
          direction: level.direction,
        })),
      });
    });
  } finally {
    i.history.end();
  }
  if (sorted) {
    refreshWorkbookCells();
    if (statusMetric) {
      const columnLabel = columns.find((column) => column.value === result.column)?.label ?? '';
      statusMetric.textContent = ribbonMenuText.sortStatus.replace('{column}', columnLabel);
    }
  }
  focusSheet();
};

const showZoomDialogFromRibbon = async (): Promise<void> => {
  const i = inst;
  if (!i) return;
  const current = Math.round(i.store.getState().viewport.zoom * 100);
  const value = await showNumberPrompt({
    title: ribbonLang === 'ja' ? 'ズーム' : 'Zoom',
    label: ribbonLang === 'ja' ? '倍率' : 'Magnification',
    initial: current,
    min: 10,
    max: 400,
    step: 1,
    okLabel: 'OK',
    cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
  });
  if (value === null) {
    focusSheet();
    return;
  }
  setSheetZoom(i.store, Math.max(0.1, Math.min(4, value / 100)), i.workbook);
  refreshZoom();
  focusSheet();
};

const removeDuplicateRows = async (): Promise<void> => {
  const i = inst;
  if (!i) return;
  const state = i.store.getState();
  const range = sortTargetRange(state);
  const inferredHeader = inferSortHasHeader(state, range);
  const columns = Array.from({ length: range.c1 - range.c0 + 1 }, (_, offset) => {
    const col = range.c0 + offset;
    const letter = colLetter(col);
    const header = inferredHeader ? sortCellDisplayText(state, range.r0, col).trim() : '';
    return {
      value: String(col),
      label: header ? `${header} (${letter})` : letter,
    };
  });
  const result = await showRemoveDuplicatesDialog({
    title: ribbonMenuText.removeDuplicatesDialogTitle,
    columnsLabel: ribbonMenuText.removeDuplicatesColumns,
    headerLabel: ribbonMenuText.sortMyDataHasHeaders,
    selectAllLabel: ribbonMenuText.removeDuplicatesSelectAll,
    unselectAllLabel: ribbonMenuText.removeDuplicatesUnselectAll,
    noColumnsLabel: ribbonMenuText.removeDuplicatesNoColumns,
    columns,
    initialColumns: columns.map((column) => column.value),
    initialHasHeader: inferredHeader,
    okLabel: 'OK',
    cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
  });
  if (!result) {
    focusSheet();
    return;
  }
  let removed = 0;
  i.history.begin();
  try {
    recordFormatChange(i.history, i.store, () => {
      removed = removeDuplicates(i.store.getState(), i.store, i.workbook, range, {
        columns: result.columns.map(Number),
        hasHeader: result.hasHeader,
      });
    });
  } finally {
    i.history.end();
  }
  if (removed > 0) refreshWorkbookCells();
  if (statusMetric) {
    statusMetric.textContent = ribbonMenuText.removeDuplicatesStatus.replace(
      '{count}',
      String(removed),
    );
  }
  focusSheet();
};

const splitTextToColumns = (delimiter = ','): void => {
  const i = inst;
  if (!i) return;
  const state = i.store.getState();
  let max = 0;
  i.history.begin();
  try {
    recordFormatChange(i.history, i.store, () => {
      max = textToColumns(state, i.store, i.workbook, state.selection.range, delimiter);
    });
  } finally {
    i.history.end();
  }
  if (max > 0) refreshWorkbookCells();
  if (statusMetric)
    statusMetric.textContent =
      max > 0
        ? ribbonMenuText.textToColumnsStatus.replace('{count}', String(max))
        : ribbonMenuText.textToColumnsNoDelimited;
  focusSheet();
};

const splitTextToColumnsCustom = async (): Promise<void> => {
  const delimiter = await showPrompt({
    title: ribbonMenuText.textToColumnsDialogTitle,
    label: ribbonMenuText.textToColumnsDialogDelimiters,
    initial: ',',
  });
  if (delimiter === null) return;
  splitTextToColumns(delimiter || ',');
};

const applyDataValidationAction = (action: string): void => {
  const i = inst;
  if (!i) return;
  const range = i.store.getState().selection.range;
  if (action === 'settings') {
    i.openDataValidationDialog();
    return;
  }
  if (action === 'circle-invalid') {
    const count = recordValidationCirclesChange(i.history, i.store, () =>
      circleInvalidValidationDataInSheet(i.store, range.sheet),
    );
    if (statusMetric)
      statusMetric.textContent = `${ribbonMenuText.validationCircleInvalid} · ${count}`;
    projectFormatToolbar();
    focusSheet();
    return;
  }
  if (action === 'clear-circles') {
    recordValidationCirclesChange(i.history, i.store, () => clearValidationCircles(i.store));
    projectFormatToolbar();
    focusSheet();
    return;
  }
  if (action === 'clear-rules') {
    clearValidationInRangeWithEngine(i.store, i.history, i.workbook, range);
    refreshWorkbookCells();
    projectFormatToolbar();
    focusSheet();
  }
};

const applyFormulaAuditAction = (action: string): void => {
  const i = inst;
  if (!i) return;
  const showTraceEmpty = (message: string): void => {
    void showMessage({ title: ribbonText.formulaAuditing, message });
  };
  if (action === 'clear-all') {
    i.clearTraces();
  } else if (action === 'clear-precedents') {
    clearTraceArrowsByKind(i.store, 'precedent', i.history);
    refreshWorkbookCells();
  } else if (action === 'clear-dependents') {
    clearTraceArrowsByKind(i.store, 'dependent', i.history);
    refreshWorkbookCells();
  } else if (action === 'error-checking') {
    runFormulaErrorChecking();
    return;
  } else if (action === 'trace-error') {
    const found = selectNextFormulaError(i.store);
    if (found) {
      if (i.tracePrecedents() === 0) showTraceEmpty(ribbonMenuText.traceNoPrecedents);
    } else {
      runFormulaErrorChecking();
    }
    return;
  } else if (action === 'ignore-error') {
    const active = i.store.getState().selection.active;
    const key = `${active.sheet}:${active.row}:${active.col}`;
    const cell = i.store.getState().data.cells.get(key);
    if (cell?.formula && cellValueIsFormulaError(cell.value)) {
      recordIgnoredErrorsChange(i.history, i.store, () => {
        ignoreCellError(i.store, active);
      });
      if (statusMetric) statusMetric.textContent = ribbonMenuText.ignoreError;
      projectFormatToolbar();
      focusSheet();
      return;
    }
    runFormulaErrorChecking();
    return;
  }
  projectFormatToolbar();
  focusSheet();
};

const applyWatchAction = (action: string): void => {
  const i = inst;
  if (!i) return;
  const state = i.store.getState();
  if (action === 'open') {
    i.openWatchWindow();
    return;
  }
  if (action === 'add') {
    recordWatchesChange(i.history, i.store, () => {
      watchRange(i.store, state.selection.range);
    });
    i.openWatchWindow();
    return;
  }
  if (action === 'delete') {
    recordWatchesChange(i.history, i.store, () => {
      unwatchCell(i.store, state.selection.active);
    });
    i.openWatchWindow();
    return;
  }
  if (action === 'delete-all') {
    recordWatchesChange(i.history, i.store, () => {
      clearWatchedCells(i.store);
    });
    i.openWatchWindow();
  }
};

const runSheetProtectionFlow = async (): Promise<void> => {
  const i = inst;
  if (!i) return;
  const ja = ribbonLang === 'ja';
  const sheet = i.store.getState().data.sheetIndex;
  if (i.isSheetProtected()) {
    const saved = protectedSheetPassword(i.store.getState(), sheet);
    if (saved) {
      const entered = await showPrompt({
        title: ja ? 'シート保護の解除' : 'Unprotect Sheet',
        label: ja ? 'パスワード' : 'Password',
        initial: '',
      });
      if (entered === null) return;
      if (entered !== saved) {
        void showMessage({
          title: ja ? 'シート保護の解除' : 'Unprotect Sheet',
          message: ja ? 'パスワードが正しくありません。' : 'The password is incorrect.',
        });
        return;
      }
    }
    recordProtectionChange(i.history, i.store, i.workbook, () => {
      i.setSheetProtected(false);
    });
    focusSheet();
    return;
  }
  const password = await showPrompt({
    title: ja ? 'シートの保護' : 'Protect Sheet',
    label: ja ? 'パスワード (省略可)' : 'Password (optional)',
    initial: '',
  });
  if (password === null) return;
  recordProtectionChange(i.history, i.store, i.workbook, () => {
    i.setSheetProtected(true, password || undefined);
  });
  focusSheet();
};

const runWorkbookProtectionFlow = async (protect: boolean): Promise<void> => {
  const i = inst;
  if (!i) return;
  const protectionText = dictionaries[ribbonLang].protection;
  if (protect) {
    if (isWorkbookStructureProtected(i.store.getState())) return;
    const password = await showPrompt({
      title: ribbonMenuText.protectWorkbookCommand.replace(/\.\.\.$/, ''),
      label: `${protectionText.password} (${shellText.optional})`,
      initial: '',
    });
    if (password === null) return;
    recordProtectionChange(i.history, i.store, i.workbook, () => {
      setWorkbookStructureProtected(i.store, true, password ? { password } : undefined);
    });
    if (statusMetric) statusMetric.textContent = ribbonMenuText.workbookProtectedStatus;
    renderSheetTabs();
    focusSheet();
    return;
  }

  if (!isWorkbookStructureProtected(i.store.getState())) return;
  const saved = workbookStructurePassword(i.store.getState());
  if (saved) {
    const entered = await showPrompt({
      title: ribbonMenuText.unprotectWorkbookCommand.replace(/\.\.\.$/, ''),
      label: protectionText.password,
      initial: '',
    });
    if (entered === null) return;
    if (entered !== saved) {
      void showMessage({
        title: ribbonMenuText.unprotectWorkbookCommand.replace(/\.\.\.$/, ''),
        message: ribbonMenuText.workbookIncorrectPassword,
      });
      return;
    }
  }
  recordProtectionChange(i.history, i.store, i.workbook, () => {
    setWorkbookStructureProtected(i.store, false);
  });
  if (statusMetric) statusMetric.textContent = ribbonMenuText.workbookUnprotectedStatus;
  renderSheetTabs();
  focusSheet();
};

const applyProtectAction = async (action: string): Promise<void> => {
  const i = inst;
  if (!i) return;
  if (action === 'protect-sheet') {
    if (i.isSheetProtected()) return;
    await runSheetProtectionFlow();
    return;
  }
  if (action === 'unprotect-sheet') {
    if (!i.isSheetProtected()) return;
    await runSheetProtectionFlow();
    return;
  }
  if (action === 'lock-cell' || action === 'unlock-cell') {
    const locked = action === 'lock-cell';
    recordFormatChange(i.history, i.store, () => {
      setCellLocked(i.store, i.store.getState().selection.range, locked);
    });
    if (statusMetric) {
      statusMetric.textContent = locked
        ? ribbonMenuText.cellsLockedStatus
        : ribbonMenuText.cellsUnlockedStatus;
    }
    projectFormatToolbar();
    focusSheet();
    return;
  }
  if (action === 'protect-workbook') {
    await runWorkbookProtectionFlow(true);
    return;
  }
  if (action === 'unprotect-workbook') {
    await runWorkbookProtectionFlow(false);
    return;
  }
  if (action === 'allow-edit-ranges') {
    const state = i.store.getState();
    const raw = await showPrompt({
      title: ribbonMenuText.allowEditRangesDialogTitle,
      label: ribbonMenuText.allowEditRangesDialogRange,
      initial: rangeRef(state.selection.range),
      okLabel: ribbonLang === 'ja' ? 'OK' : 'OK',
      cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
      validate: (value) =>
        parseA1Range(value, state.selection.active.sheet)
          ? null
          : ribbonMenuText.allowEditRangesDialogInvalid,
    });
    if (raw === null) return;
    const range = parseA1Range(raw, state.selection.active.sheet);
    if (!range) return;
    recordProtectionChange(i.history, i.store, i.workbook, () => {
      addAllowedEditRange(i.store, range, { title: rangeRef(range) });
    });
    if (statusMetric) {
      statusMetric.textContent = ribbonMenuText.allowedEditRangeAddedStatus.replace(
        '{range}',
        rangeRef(range),
      );
    }
    focusSheet();
    return;
  }
  if (action === 'clear-allowed-edit-ranges') {
    recordProtectionChange(i.history, i.store, i.workbook, () => {
      clearAllowedEditRanges(i.store, i.store.getState().data.sheetIndex);
    });
    if (statusMetric) statusMetric.textContent = ribbonMenuText.allowedEditRangesClearedStatus;
    focusSheet();
  }
};

const applyCellStyleFromRibbon = (id: CellStyleId): void => {
  const i = inst;
  if (!i) return;
  const range = i.store.getState().selection.range;
  applyCellStyle(i.store, i.history, range, id);
  refreshWorkbookCells();
  focusSheet();
};

const applyCurrencyPreset = (symbol: string): void => {
  applyRibbonFormat((state, store) =>
    setNumFmt(state, store, { kind: 'currency', decimals: 2, symbol }),
  );
};

const openCurrencyFooterAction = (action: string): void => {
  if (action === 'more') {
    inst?.openFormatDialog();
  }
};

const openCellStyleFooterAction = async (action: string): Promise<void> => {
  const ja = ribbonLang === 'ja';
  if (action === 'new-cell-style') {
    await showMessage({
      title: ja ? '新しいセルのスタイル' : 'New Cell Style',
      message: ja
        ? 'カスタム セル スタイルの作成は今後のリリースで対応予定です。'
        : 'Authoring custom cell styles is coming in a future release.',
    });
    focusSheet();
    return;
  }
  if (action === 'merge-cell-style') {
    await showMessage({
      title: ja ? 'スタイルの結合' : 'Merge Styles',
      message: ja
        ? '他のブックのスタイル結合は今後のリリースで対応予定です。'
        : 'Merging styles from another workbook is coming in a future release.',
    });
    focusSheet();
  }
};

const openTableStyleFooterAction = async (action: string): Promise<void> => {
  const ja = ribbonLang === 'ja';
  if (action === 'new-table-style') {
    await showMessage({
      title: ja ? '新しい表スタイル' : 'New Table Style',
      message: ja
        ? 'カスタム表スタイルの作成は今後のリリースで対応予定です。'
        : 'Authoring custom table styles is coming in a future release.',
    });
    focusSheet();
    return;
  }
  if (action === 'new-pivot-style') {
    await showMessage({
      title: ja ? '新しいピボットテーブル スタイル' : 'New PivotTable Style',
      message: ja
        ? 'カスタム ピボットテーブル スタイルの作成は今後のリリースで対応予定です。'
        : 'Authoring custom PivotTable styles is coming in a future release.',
    });
    focusSheet();
  }
};

const createTableFromSelection = async (
  style: TableStyle = 'medium',
  color?: string,
  variant: TableVariantId = 'banded',
): Promise<void> => {
  const i = inst;
  if (!i) return;
  const state = i.store.getState();
  const result = await showFormatAsTableDialog({
    title: ribbonLang === 'ja' ? 'テーブルとして書式設定' : 'Format as Table',
    rangeLabel:
      ribbonLang === 'ja'
        ? 'テーブルに変換するデータ範囲を指定してください'
        : 'Where is the data for your table?',
    headersLabel:
      ribbonLang === 'ja' ? '先頭行をテーブルの見出しとして使用する' : 'My table has headers',
    initialRange: rangeRef(state.selection.range),
    initialHasHeaders: true,
    okLabel: 'OK',
    cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
    validateRange: (value) =>
      parseA1Range(value, state.selection.active.sheet)
        ? null
        : ribbonLang === 'ja'
          ? 'A1:B10 の形式で入力してください。'
          : 'Enter a range such as A1:B10.',
  });
  if (result === null) {
    focusSheet();
    return;
  }
  const r = parseA1Range(result.range, state.selection.active.sheet);
  if (!r) {
    focusSheet();
    return;
  }
  const variantOptions = tableVariantOptions(variant);
  recordTablesChange(i.history, i.store, () => {
    formatAsTable(i.store, r, {
      showHeader: result.hasHeaders,
      style,
      color,
      banded: variantOptions.banded,
      firstCol: variantOptions.firstCol,
    });
  });
  focusSheet();
};

type RecommendedPivotSpec = {
  rowField: string;
  columnField?: string;
  valueField: string;
  aggregation: PivotAggregation;
  placement: 'existing' | 'new-sheet';
};

const pivotSpecKey = (spec: RecommendedPivotSpec): string =>
  [spec.rowField, spec.columnField ?? '', spec.valueField, spec.aggregation, spec.placement].join(
    '\u0001',
  );

const buildRecommendedPivotSpecs = (
  fields: readonly PivotSourceField[],
  placement: RecommendedPivotSpec['placement'],
): RecommendedPivotSpec[] => {
  const numeric = fields.filter((field) => field.numericCount > 0);
  const values = numeric.length > 0 ? numeric : fields;
  const categories = fields.filter((field) => field.numericCount === 0);
  const rows = categories.length > 0 ? categories : fields;
  const specs: RecommendedPivotSpec[] = [];
  const add = (rowField = rows[0], valueField = values[0], columnField?: PivotSourceField) => {
    if (!rowField || !valueField || rowField.name === valueField.name) return;
    if (columnField && (columnField.name === rowField.name || columnField.name === valueField.name))
      return;
    specs.push({
      rowField: rowField.name,
      columnField: columnField?.name,
      valueField: valueField.name,
      aggregation: valueField.numericCount > 0 ? PivotAggregation.Sum : PivotAggregation.Count,
      placement,
    });
  };
  add(
    rows[0],
    values[0],
    categories.find((field) => field.name !== rows[0]?.name),
  );
  add(rows[0], values[0]);
  add(rows[1], values[0]);
  add(rows[0], values[1] ?? values[0]);
  const seen = new Set<string>();
  return specs.filter((spec) => {
    const key = pivotSpecKey(spec);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
};

const pivotSpecLabel = (spec: RecommendedPivotSpec): string => {
  const valueLabel =
    spec.aggregation === PivotAggregation.Sum
      ? ribbonLang === 'ja'
        ? `合計 / ${spec.valueField}`
        : `Sum of ${spec.valueField}`
      : ribbonLang === 'ja'
        ? `データの個数 / ${spec.valueField}`
        : `Count of ${spec.valueField}`;
  const axisLabel = spec.columnField
    ? ribbonLang === 'ja'
      ? `${spec.rowField} x ${spec.columnField}`
      : `${spec.rowField} by ${spec.columnField}`
    : spec.rowField;
  return `${axisLabel} - ${valueLabel}`;
};

const createRecommendedPivotTable = (
  placement: 'existing' | 'new-sheet',
  sourceOverride?: Range,
  specOverride?: RecommendedPivotSpec,
): void => {
  const i = inst;
  if (!i) return;
  const state = i.store.getState();
  const source = sourceOverride ?? sortTargetRange(state);
  const fields = inferPivotSourceFields(i.workbook, source);
  const spec = specOverride ?? buildRecommendedPivotSpecs(fields, placement)[0];
  if (!spec) {
    void showMessage({
      title: ribbonText.pivotTable,
      message:
        ribbonLang === 'ja'
          ? 'ピボットテーブルを作成できる見出し付きデータ範囲を選択してください。'
          : 'Select a labeled data range that can be used for a PivotTable.',
    });
    return;
  }
  let destinationSheet = source.sheet;
  if (placement === 'new-sheet') {
    const added = addSheet(i.store, i.workbook);
    if (added < 0) {
      if (statusMetric && isWorkbookStructureProtected(i.store.getState())) {
        statusMetric.textContent = ribbonMenuText.workbookStructureProtectedBlocked;
      }
      return;
    }
    destinationSheet = added;
    renderSheetTabs();
  }
  const destination =
    placement === 'new-sheet'
      ? { sheet: destinationSheet, row: 0, col: 0 }
      : { sheet: destinationSheet, row: source.r1 + 3, col: source.c0 };
  const result = createPivotTableFromRange(i.workbook, {
    source,
    destination,
    name: `PivotTable${i.workbook.getPivotTables().length + 1}`,
    rowField: spec.rowField,
    columnField: spec.columnField,
    valueField: spec.valueField,
    aggregation: spec.aggregation,
  });
  if (!result.ok) {
    void showMessage({
      title: ribbonText.pivotTable,
      message:
        ribbonLang === 'ja'
          ? 'ピボットテーブルを作成できませんでした。'
          : 'Could not create a PivotTable from the selected range.',
    });
    return;
  }
  mutators.setActive(i.store, destination);
  if (placement === 'new-sheet') switchSheet(destinationSheet);
  else refreshWorkbookCells();
  if (statusMetric) statusMetric.textContent = ribbonMenuText.pivotTableCreated;
  focusSheet();
};

const openRecommendedPivotTablesDialog = async (): Promise<void> => {
  const i = inst;
  if (!i) return;
  const source = sortTargetRange(i.store.getState());
  const fields = inferPivotSourceFields(i.workbook, source);
  const specs = buildRecommendedPivotSpecs(fields, 'existing');
  if (specs.length === 0) {
    createRecommendedPivotTable('existing', source);
    return;
  }
  const options = specs.map((spec, index) => ({
    value: `pivot-${index}`,
    label: pivotSpecLabel(spec),
  }));
  const choice = await showChoiceDialog<string>({
    title: ribbonMenuText.recommendedPivotTables,
    label: ribbonText.pivotTable,
    initial: options[0]?.value,
    okLabel: 'OK',
    cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
    options,
  });
  if (!choice) {
    focusSheet();
    return;
  }
  const index = Number(choice.replace('pivot-', ''));
  const spec = specs[index];
  if (spec) createRecommendedPivotTable(spec.placement, source, spec);
};

const applyPivotTableAction = (action: string): void => {
  const i = inst;
  if (!i) return;
  if (action === 'recommended') {
    void openRecommendedPivotTablesDialog();
    return;
  }
  if (action === 'new-sheet') {
    createRecommendedPivotTable('new-sheet');
    return;
  }
  i.openPivotTableDialog();
};

const applyDefinedNameAction = (action: string): void => {
  const i = inst;
  if (!i) return;
  if (action === 'define') {
    i.openDefineNameDialog();
    return;
  }
  if (action === 'manager') {
    i.openNamedRangeDialog();
    return;
  }
  if (
    action === 'create-top-row' ||
    action === 'create-bottom-row' ||
    action === 'create-left-column' ||
    action === 'create-right-column'
  ) {
    const source =
      action === 'create-top-row'
        ? 'top-row'
        : action === 'create-bottom-row'
          ? 'bottom-row'
          : action === 'create-left-column'
            ? 'left-column'
            : 'right-column';
    const result = recordDefinedNamesChange(i.history, i.workbook, () =>
      createDefinedNamesFromSelection(i.store.getState(), i.workbook, source),
    );
    if (!result.ok) {
      void showMessage({
        title: ribbonText.definedNames,
        message: ribbonMenuText.definedNamesCreateFailed,
      });
      return;
    }
    if (statusMetric) {
      statusMetric.textContent = ribbonMenuText.definedNamesCreated.replace(
        '{count}',
        String(result.entries.length),
      );
    }
    focusSheet();
    return;
  }
  if (action === 'use-formula') {
    const names = listDefinedNames(i.workbook);
    const firstName = names[0];
    if (firstName) {
      insertDefinedNameFormula(i.store.getState(), i.workbook, firstName.name, i.store);
      refreshWorkbookCells();
      focusSheet();
      return;
    }
    void showMessage({
      title: ribbonText.definedNames,
      message: ribbonMenuText.noDefinedNames,
    });
    return;
  }
  if (action.startsWith('insert:')) {
    const name = action.slice('insert:'.length);
    insertDefinedNameFormula(i.store.getState(), i.workbook, name, i.store);
    refreshWorkbookCells();
    focusSheet();
  }
};

const clearHyperlinksInSelection = (mode: 'clear' | 'remove' = 'clear'): void => {
  const i = inst;
  if (!i) return;
  const range = i.store.getState().selection.range;
  recordFormatChange(i.history, i.store, () => {
    for (let row = range.r0; row <= range.r1; row += 1) {
      for (let col = range.c0; col <= range.c1; col += 1) {
        const addr = { sheet: range.sheet, row, col };
        clearHyperlink(i.store, addr, i.workbook);
        if (mode === 'remove') {
          mutators.setCellFormat(i.store, addr, {
            color: undefined,
            underline: undefined,
          });
        }
      }
    }
  });
  refreshWorkbookCells();
  focusSheet();
};

const applyLinksAction = (action: string): void => {
  const i = inst;
  if (!i) return;
  if (action === 'hyperlink') {
    i.openHyperlinkDialog();
    return;
  }
  if (action === 'external') {
    i.openExternalLinksDialog();
    return;
  }
  if (action === 'clear') {
    clearHyperlinksInSelection('clear');
    return;
  }
  if (action === 'open') {
    const target = hyperlinkAt(i.store.getState(), i.store.getState().selection.active);
    if (!target) {
      void showMessage({
        title: ribbonText.links,
        message: ribbonMenuText.linkNoHyperlink,
      });
      return;
    }
    window.open(target, '_blank', 'noopener,noreferrer');
  }
};

type SessionIllustrationKind = 'image' | 'shape' | 'screenshot';
type SessionShapeKind = 'rectangle' | 'rounded-rectangle' | 'oval' | 'line' | 'arrow';
type SessionIllustration = {
  id: string;
  kind: SessionIllustrationKind;
  shape?: SessionShapeKind;
  url?: string;
  x: number;
  y: number;
  w: number;
  h: number;
};

const sessionIllustrations: SessionIllustration[] = [];
let selectedIllustrationId: string | null = null;

const cloneSessionIllustration = (item: SessionIllustration): SessionIllustration => ({
  ...item,
});

const captureSessionIllustrationsSnapshot = (): {
  items: SessionIllustration[];
  selectedId: string | null;
} => ({
  items: sessionIllustrations.map(cloneSessionIllustration),
  selectedId: selectedIllustrationId,
});

const applySessionIllustrationsSnapshot = (snapshot: {
  items: readonly SessionIllustration[];
  selectedId: string | null;
}): void => {
  sessionIllustrations.splice(
    0,
    sessionIllustrations.length,
    ...snapshot.items.map(cloneSessionIllustration),
  );
  selectedIllustrationId = snapshot.selectedId;
  renderSessionIllustrations();
};

const sameSessionIllustrationsSnapshot = (
  a: { items: readonly SessionIllustration[]; selectedId: string | null },
  b: { items: readonly SessionIllustration[]; selectedId: string | null },
): boolean => a.selectedId === b.selectedId && JSON.stringify(a.items) === JSON.stringify(b.items);

const recordSessionIllustrationsChange = (mutate: () => void): void => {
  const history = inst?.history ?? null;
  if (!history || history.isReplaying()) {
    mutate();
    return;
  }
  const before = captureSessionIllustrationsSnapshot();
  mutate();
  const after = captureSessionIllustrationsSnapshot();
  if (sameSessionIllustrationsSnapshot(before, after)) return;
  history.push({
    undo: () => applySessionIllustrationsSnapshot(before),
    redo: () => applySessionIllustrationsSnapshot(after),
  });
};

type SessionInkStroke = {
  id: string;
  points: Array<{ x: number; y: number }>;
};

const sessionInkStrokes: SessionInkStroke[] = [];
let drawInkMode: 'pen' | 'erase' | null = null;
let inkListenersAttached = false;

const cloneSessionInkStroke = (stroke: SessionInkStroke): SessionInkStroke => ({
  id: stroke.id,
  points: stroke.points.map((point) => ({ ...point })),
});

const captureSessionInkSnapshot = (): SessionInkStroke[] =>
  sessionInkStrokes.map(cloneSessionInkStroke);

const applySessionInkSnapshot = (snapshot: readonly SessionInkStroke[]): void => {
  sessionInkStrokes.splice(0, sessionInkStrokes.length, ...snapshot.map(cloneSessionInkStroke));
  renderSessionInk();
};

const sameSessionInkSnapshot = (
  a: readonly SessionInkStroke[],
  b: readonly SessionInkStroke[],
): boolean => JSON.stringify(a) === JSON.stringify(b);

const pushSessionInkHistory = (
  before: readonly SessionInkStroke[],
  after: readonly SessionInkStroke[],
): void => {
  const history = inst?.history ?? null;
  if (!history || history.isReplaying() || sameSessionInkSnapshot(before, after)) return;
  const undoSnapshot = before.map(cloneSessionInkStroke);
  const redoSnapshot = after.map(cloneSessionInkStroke);
  history.push({
    undo: () => applySessionInkSnapshot(undoSnapshot),
    redo: () => applySessionInkSnapshot(redoSnapshot),
  });
};

const recordSessionInkChange = (mutate: () => void): void => {
  const before = captureSessionInkSnapshot();
  mutate();
  pushSessionInkHistory(before, captureSessionInkSnapshot());
};

const illustrationGrid = (): HTMLElement | null =>
  sheetEl?.querySelector<HTMLElement>('.fc-host__grid') ?? null;

const syncDrawInkButtons = (): void => {
  for (const button of document.querySelectorAll<HTMLButtonElement>(
    '[data-ribbon-command="drawPen"], [data-ribbon-command="drawErase"]',
  )) {
    const active =
      (button.dataset.ribbonCommand === 'drawPen' && drawInkMode === 'pen') ||
      (button.dataset.ribbonCommand === 'drawErase' && drawInkMode === 'erase');
    button.setAttribute('aria-pressed', active ? 'true' : 'false');
  }
};

const inkRoot = (): SVGSVGElement | null => {
  const grid = illustrationGrid();
  if (!grid) return null;
  let root = grid.querySelector<SVGSVGElement>('.app-ink');
  if (!root) {
    root = document.createElementNS(SVG_NS, 'svg');
    root.classList.add('app-ink');
    root.setAttribute('aria-hidden', 'true');
    grid.appendChild(root);
  }
  return root;
};

const renderSessionInk = (): void => {
  const root = inkRoot();
  if (!root) return;
  root.replaceChildren();
  for (const stroke of sessionInkStrokes) {
    const polyline = document.createElementNS(SVG_NS, 'polyline');
    polyline.classList.add('app-ink__stroke');
    polyline.dataset.inkStrokeId = stroke.id;
    polyline.setAttribute('points', stroke.points.map((p) => `${p.x},${p.y}`).join(' '));
    root.appendChild(polyline);
  }
};

const gridPointFromPointer = (event: PointerEvent): { x: number; y: number } | null => {
  const grid = illustrationGrid();
  if (!grid) return null;
  const rect = grid.getBoundingClientRect();
  return {
    x: Math.max(0, event.clientX - rect.left),
    y: Math.max(0, event.clientY - rect.top),
  };
};

const pointToSegmentDistance = (
  point: { x: number; y: number },
  a: { x: number; y: number },
  b: { x: number; y: number },
): number => {
  const dx = b.x - a.x;
  const dy = b.y - a.y;
  const len2 = dx * dx + dy * dy || 1;
  const t = Math.max(0, Math.min(1, ((point.x - a.x) * dx + (point.y - a.y) * dy) / len2));
  const x = a.x + t * dx;
  const y = a.y + t * dy;
  return Math.hypot(point.x - x, point.y - y);
};

const eraseInkAt = (point: { x: number; y: number }): void => {
  const index = sessionInkStrokes.findIndex((stroke) => {
    if (stroke.points.length === 1) {
      return Math.hypot(stroke.points[0]!.x - point.x, stroke.points[0]!.y - point.y) < 12;
    }
    for (let i = 1; i < stroke.points.length; i += 1) {
      if (pointToSegmentDistance(point, stroke.points[i - 1]!, stroke.points[i]!) < 12) {
        return true;
      }
    }
    return false;
  });
  if (index >= 0) {
    recordSessionInkChange(() => {
      sessionInkStrokes.splice(index, 1);
      renderSessionInk();
    });
  }
};

const attachInkPointerListeners = (): void => {
  const grid = illustrationGrid();
  if (!grid || inkListenersAttached) return;
  inkListenersAttached = true;
  grid.addEventListener('pointerdown', (event) => {
    if (!drawInkMode || event.button !== 0) return;
    const point = gridPointFromPointer(event);
    if (!point) return;
    event.preventDefault();
    event.stopPropagation();
    if (drawInkMode === 'erase') {
      eraseInkAt(point);
      return;
    }
    const before = captureSessionInkSnapshot();
    const stroke: SessionInkStroke = {
      id: `ink-${Date.now().toString(36)}-${sessionInkStrokes.length}`,
      points: [point],
    };
    sessionInkStrokes.push(stroke);
    renderSessionInk();
    const onMove = (moveEvent: PointerEvent): void => {
      const next = gridPointFromPointer(moveEvent);
      if (!next) return;
      stroke.points.push(next);
      renderSessionInk();
    };
    const onUp = (): void => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
      window.removeEventListener('pointercancel', onUp);
      renderSessionInk();
      pushSessionInkHistory(before, captureSessionInkSnapshot());
    };
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
    window.addEventListener('pointercancel', onUp);
  });
};

const setDrawInkMode = (mode: 'pen' | 'erase'): void => {
  drawInkMode = drawInkMode === mode ? null : mode;
  attachInkPointerListeners();
  inkRoot();
  sheetEl
    ?.querySelector<HTMLElement>('.fc-host')
    ?.classList.toggle('app-ink--pen', drawInkMode === 'pen');
  sheetEl
    ?.querySelector<HTMLElement>('.fc-host')
    ?.classList.toggle('app-ink--erase', drawInkMode === 'erase');
  syncDrawInkButtons();
  illustrationGrid()?.focus();
};

const illustrationRoot = (): HTMLElement | null => {
  const grid = illustrationGrid();
  if (!grid) return null;
  let root = grid.querySelector<HTMLElement>('.app-illustrations');
  if (!root) {
    root = document.createElement('div');
    root.className = 'app-illustrations';
    grid.appendChild(root);
  }
  return root;
};

const updateSessionIllustration = (
  id: string,
  patch: Partial<Pick<SessionIllustration, 'x' | 'y' | 'w' | 'h'>>,
): void => {
  recordSessionIllustrationsChange(() => {
    const item = sessionIllustrations.find((candidate) => candidate.id === id);
    if (!item) return;
    Object.assign(item, patch);
    renderSessionIllustrations();
  });
};

const removeSessionIllustration = (id: string): void => {
  recordSessionIllustrationsChange(() => {
    const index = sessionIllustrations.findIndex((candidate) => candidate.id === id);
    if (index < 0) return;
    sessionIllustrations.splice(index, 1);
    if (selectedIllustrationId === id) selectedIllustrationId = null;
    renderSessionIllustrations();
  });
};

const selectSessionIllustration = (id: string): void => {
  selectedIllustrationId = id;
  renderSessionIllustrations();
};

const applyIllustrationPointerDrag = (
  event: PointerEvent,
  item: SessionIllustration,
  node: HTMLElement,
  mode: 'move' | 'resize',
): void => {
  if (event.button !== 0) return;
  event.preventDefault();
  event.stopPropagation();
  selectedIllustrationId = item.id;
  illustrationRoot()
    ?.querySelectorAll<HTMLElement>('.app-illustration')
    .forEach((candidate) =>
      candidate.setAttribute(
        'aria-selected',
        candidate.dataset.illustrationId === item.id ? 'true' : 'false',
      ),
    );
  node.focus();
  const startX = event.clientX;
  const startY = event.clientY;
  const start = { x: item.x, y: item.y, w: item.w, h: item.h };
  let next = { ...start };
  const applyLive = (): void => {
    node.style.left = `${next.x}px`;
    node.style.top = `${next.y}px`;
    node.style.width = `${next.w}px`;
    node.style.height = `${next.h}px`;
  };
  const onMove = (moveEvent: PointerEvent): void => {
    const dx = moveEvent.clientX - startX;
    const dy = moveEvent.clientY - startY;
    if (mode === 'resize') {
      next = {
        ...start,
        w: Math.max(24, start.w + dx),
        h: Math.max(12, start.h + dy),
      };
    } else {
      next = {
        ...start,
        x: Math.max(0, start.x + dx),
        y: Math.max(0, start.y + dy),
      };
    }
    applyLive();
  };
  const onUp = (upEvent: PointerEvent): void => {
    upEvent.preventDefault();
    window.removeEventListener('pointermove', onMove);
    window.removeEventListener('pointerup', onUp);
    window.removeEventListener('pointercancel', onUp);
    updateSessionIllustration(item.id, next);
  };
  window.addEventListener('pointermove', onMove);
  window.addEventListener('pointerup', onUp);
  window.addEventListener('pointercancel', onUp);
};

const applyIllustrationKeyboard = (
  event: KeyboardEvent,
  item: SessionIllustration,
  node: HTMLElement,
): void => {
  if (event.key === 'Delete' || event.key === 'Backspace') {
    event.preventDefault();
    removeSessionIllustration(item.id);
    focusSheet();
    return;
  }
  const step = event.shiftKey ? 10 : 1;
  const resize = event.altKey;
  const delta: [number, number] | null =
    event.key === 'ArrowLeft'
      ? [-step, 0]
      : event.key === 'ArrowRight'
        ? [step, 0]
        : event.key === 'ArrowUp'
          ? [0, -step]
          : event.key === 'ArrowDown'
            ? [0, step]
            : null;
  if (!delta) return;
  event.preventDefault();
  if (resize) {
    updateSessionIllustration(item.id, {
      w: Math.max(24, item.w + delta[0]),
      h: Math.max(12, item.h + delta[1]),
    });
  } else {
    updateSessionIllustration(item.id, {
      x: Math.max(0, item.x + delta[0]),
      y: Math.max(0, item.y + delta[1]),
    });
  }
  node.focus();
};

const illustrationLabel = (item: SessionIllustration): string => {
  const t = ribbonMenuText;
  if (item.kind === 'image') return t.pictureOnline;
  if (item.kind === 'screenshot') return t.screenshotCurrentView;
  if (item.shape === 'rounded-rectangle') return t.shapeRoundedRectangle;
  if (item.shape === 'oval') return t.shapeOval;
  if (item.shape === 'line') return t.shapeLine;
  if (item.shape === 'arrow') return t.shapeArrow;
  return t.shapeRectangle;
};

const renderSessionIllustrations = (): void => {
  const root = illustrationRoot();
  if (!root) return;
  root.replaceChildren();
  for (const item of sessionIllustrations) {
    const node = document.createElement('div');
    node.className = `app-illustration app-illustration--${item.kind}`;
    node.setAttribute('role', 'button');
    node.tabIndex = 0;
    if (item.shape) node.classList.add(`app-illustration--${item.shape}`);
    node.dataset.illustrationId = item.id;
    node.dataset.illustrationType = item.kind;
    if (item.shape) node.dataset.shape = item.shape;
    node.setAttribute('aria-label', illustrationLabel(item));
    node.setAttribute('aria-selected', item.id === selectedIllustrationId ? 'true' : 'false');
    node.style.left = `${item.x}px`;
    node.style.top = `${item.y}px`;
    node.style.width = `${item.w}px`;
    node.style.height = `${item.h}px`;
    node.addEventListener('pointerdown', (event) => {
      const rect = node.getBoundingClientRect();
      const nearResizeHandle =
        event.clientX >= rect.right - 14 && event.clientY >= rect.bottom - 14;
      applyIllustrationPointerDrag(event, item, node, nearResizeHandle ? 'resize' : 'move');
    });
    node.addEventListener('keydown', (event) => applyIllustrationKeyboard(event, item, node));
    if (item.kind === 'image') {
      const image = document.createElement('img');
      image.alt = '';
      image.src = item.url ?? '';
      node.appendChild(image);
    } else if (item.kind === 'screenshot') {
      node.appendChild(document.createElement('span'));
    }
    const resize = document.createElement('span');
    resize.className = 'app-illustration__resize';
    resize.setAttribute('aria-hidden', 'true');
    resize.addEventListener('pointerdown', (event) =>
      applyIllustrationPointerDrag(event, item, node, 'resize'),
    );
    node.appendChild(resize);
    root.appendChild(node);
  }
};

const addSessionIllustration = (
  kind: SessionIllustrationKind,
  input: Partial<SessionIllustration> = {},
): void => {
  const count = sessionIllustrations.length;
  const item: SessionIllustration = {
    id: `illustration-${Date.now().toString(36)}-${count}`,
    kind,
    x: 360 + (count % 4) * 28,
    y: 340 + (count % 4) * 24,
    w: kind === 'shape' && (input.shape === 'line' || input.shape === 'arrow') ? 150 : 180,
    h: kind === 'shape' && (input.shape === 'line' || input.shape === 'arrow') ? 28 : 110,
    ...input,
  };
  recordSessionIllustrationsChange(() => {
    sessionIllustrations.push(item);
    selectedIllustrationId = item.id;
    renderSessionIllustrations();
  });
  illustrationGrid()?.focus();
};

const insertPictureFromRibbon = async (action: string): Promise<void> => {
  const title =
    action === 'device' ? ribbonMenuText.pictureThisDevice : ribbonMenuText.pictureOnline;
  const url = await showPrompt({
    title,
    label: ribbonMenuText.pictureUrlPrompt,
    placeholder: 'https://...',
    okLabel: 'OK',
    cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
    validate: (value) =>
      value.trim() ? null : ribbonLang === 'ja' ? 'URLを入力してください。' : 'Enter a URL.',
  });
  if (!url) {
    focusSheet();
    return;
  }
  addSessionIllustration('image', { url: url.trim(), w: 220, h: 140 });
};

const insertShapeFromRibbon = (shape: SessionShapeKind): void => {
  addSessionIllustration('shape', { shape });
};

const insertScreenshotFromRibbon = (): void => {
  addSessionIllustration('screenshot', { w: 230, h: 150 });
};

const createChartFromSelection = (kind: SessionChartKind = 'column'): void => {
  if (!inst) return;
  const r = inst.store.getState().selection.range;
  const count = inst.store.getState().charts.charts.length;
  createSessionChart(
    inst.store,
    r,
    {
      id: `ribbon-chart-${r.sheet}-${r.r0}-${r.c0}-${r.r1}-${r.c1}-${kind}-${count}`,
      kind,
      title: null,
      x: 340 + (count % 3) * 24,
      y: 96 + (count % 3) * 24,
      w: 360,
      h: 220,
    },
    inst.history,
  );
  focusSheet();
};

const recommendedChartKind = (): SessionChartKind => {
  const r = inst?.store.getState().selection.range;
  if (!r) return 'column';
  if (r.r0 === r.r1 && r.c1 - r.c0 >= 2) return 'line';
  if (r.c0 === r.c1 && r.r1 - r.r0 >= 2) return 'bar';
  if (r.c1 - r.c0 === 1 && r.r1 - r.r0 <= 6) return 'pie';
  return 'column';
};

const chartLabel = (kind: SessionChartKind): string => {
  const t = ribbonMenuText;
  if (kind === 'bar') return t.chartBar;
  if (kind === 'line') return t.chartLine;
  if (kind === 'area') return t.chartArea;
  if (kind === 'pie') return t.chartPie;
  if (kind === 'scatter') return t.chartScatter;
  return t.chartColumn;
};

const createRecommendedChartFromSelection = async (): Promise<void> => {
  const t = ribbonMenuText;
  const initial = recommendedChartKind();
  const rawOptions: Array<{ value: SessionChartKind; label: string }> = [
    { value: initial, label: `${t.recommendedCharts}: ${chartLabel(initial)}` },
    { value: 'column', label: t.chartColumn },
    { value: 'bar', label: t.chartBar },
    { value: 'line', label: t.chartLine },
    { value: 'area', label: t.chartArea },
    { value: 'pie', label: t.chartPie },
    { value: 'scatter', label: t.chartScatter },
  ];
  const options = rawOptions.filter(
    (option, index, all) =>
      all.findIndex((candidate) => candidate.value === option.value) === index,
  );
  const choice = await showChoiceDialog<SessionChartKind>({
    title: t.recommendedCharts,
    label: t.chart,
    initial,
    okLabel: 'OK',
    cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
    options,
  });
  if (!choice) {
    focusSheet();
    return;
  }
  createChartFromSelection(choice);
};

const chartKindFromAction = (action: string): SessionChartKind =>
  action === 'bar' ||
  action === 'line' ||
  action === 'area' ||
  action === 'pie' ||
  action === 'scatter'
    ? action
    : action === 'recommended'
      ? recommendedChartKind()
      : 'column';

const copySelectionToClipboard = async (): Promise<void> => {
  if (!inst) return;
  const state = inst.store.getState();
  const result = copy(state);
  if (!result) return;
  ribbonClipboardSnapshot = captureSnapshot(state, result.range);
  ribbonClipboardText = result.tsv;
  await navigator.clipboard?.writeText(result.tsv);
  focusSheet();
};

const cutSelectionToClipboard = async (): Promise<void> => {
  if (!inst) return;
  const state = inst.store.getState();
  ribbonClipboardSnapshot = captureSnapshot(state, state.selection.range);
  inst.history.begin();
  let result: ReturnType<typeof cut> = null;
  try {
    result = cut(state, inst.workbook);
    if (result) {
      const ranges = result.payloadRanges ?? result.ranges ?? [result.range];
      recordFormatChange(inst.history, inst.store, () => {
        inst?.store.setState((s) => {
          const formats = new Map(s.format.formats);
          for (const range of ranges) {
            for (let row = range.r0; row <= range.r1; row += 1) {
              for (let col = range.c0; col <= range.c1; col += 1) {
                formats.delete(`${range.sheet}:${row}:${col}`);
              }
            }
          }
          return { ...s, format: { formats } };
        });
      });
    }
  } finally {
    inst.history.end();
  }
  if (!result) {
    ribbonClipboardSnapshot = null;
    ribbonClipboardText = null;
    return;
  }
  ribbonClipboardText = result.tsv;
  await navigator.clipboard?.writeText(result.tsv);
  refreshWorkbookCells();
  focusSheet();
};

const pasteClipboardIntoSelection = async (): Promise<void> => {
  const i = inst;
  if (!i) return;
  let text = '';
  try {
    text = (await navigator.clipboard?.readText()) ?? '';
  } catch {
    text = '';
  }
  if (!text && ribbonClipboardText) text = ribbonClipboardText;
  if (!text) return;
  if (ribbonClipboardSnapshot && text === ribbonClipboardText) {
    const source = ribbonClipboardSnapshot;
    i.history.begin();
    let result: ReturnType<typeof applyPasteSpecial> = null;
    try {
      recordFormatChange(i.history, i.store, () => {
        result = applyPasteSpecial(i.store.getState(), i.store, i.workbook, source, {
          what: 'all',
          operation: 'none',
          skipBlanks: false,
          transpose: false,
        });
      });
    } finally {
      i.history.end();
    }
    const applied = result as ReturnType<typeof applyPasteSpecial>;
    if (applied) mutators.setRange(i.store, applied.writtenRange);
  } else {
    i.history.begin();
    let result: ReturnType<typeof pasteTSV> = null;
    try {
      result = pasteTSV(i.store.getState(), i.workbook, text);
    } finally {
      i.history.end();
    }
    if (result) mutators.setRange(i.store, result.writtenRange);
  }
  refreshWorkbookCells();
  focusSheet();
};

const pasteOptionsForAction = (action: string): PasteSpecialOptions | null => {
  const base = {
    operation: 'none' as PasteOperation,
    skipBlanks: false,
    transpose: false,
  };
  const whatByAction: Record<string, PasteWhat> = {
    all: 'all',
    formulas: 'formulas',
    'formulas-and-numfmt': 'formulas-and-numfmt',
    values: 'values',
    'values-and-numfmt': 'values-and-numfmt',
    formats: 'formats',
  };
  if (action === 'transpose') return { ...base, what: 'all', transpose: true };
  const what = whatByAction[action];
  return what ? { ...base, what } : null;
};

const ribbonPasteWhatOptions = (): Array<{ value: PasteWhat; label: string }> =>
  ribbonLang === 'ja'
    ? [
        { value: 'all', label: 'すべて' },
        { value: 'formulas', label: '数式' },
        { value: 'values', label: '値' },
        { value: 'formats', label: '書式' },
        { value: 'formulas-and-numfmt', label: '数式と数値の書式' },
        { value: 'values-and-numfmt', label: '値と数値の書式' },
      ]
    : [
        { value: 'all', label: 'All' },
        { value: 'formulas', label: 'Formulas' },
        { value: 'values', label: 'Values' },
        { value: 'formats', label: 'Formats' },
        { value: 'formulas-and-numfmt', label: 'Formulas and number formats' },
        { value: 'values-and-numfmt', label: 'Values and number formats' },
      ];

const ribbonPasteOperationOptions = (): Array<{ value: PasteOperation; label: string }> =>
  ribbonLang === 'ja'
    ? [
        { value: 'none', label: 'しない' },
        { value: 'add', label: '加算' },
        { value: 'subtract', label: '減算' },
        { value: 'multiply', label: '乗算' },
        { value: 'divide', label: '除算' },
      ]
    : [
        { value: 'none', label: 'None' },
        { value: 'add', label: 'Add' },
        { value: 'subtract', label: 'Subtract' },
        { value: 'multiply', label: 'Multiply' },
        { value: 'divide', label: 'Divide' },
      ];

const makeRibbonPasteRadio = <T extends string>(
  name: string,
  value: T,
  label: string,
  checked: boolean,
): HTMLLabelElement => {
  const wrap = document.createElement('label');
  wrap.className = 'fc-fmtdlg__radio';
  const input = document.createElement('input');
  input.type = 'radio';
  input.name = name;
  input.value = value;
  input.checked = checked;
  const span = document.createElement('span');
  span.textContent = label;
  wrap.append(input, span);
  return wrap;
};

const makeRibbonPasteCheck = (
  label: string,
): { input: HTMLInputElement; label: HTMLLabelElement } => {
  const wrap = document.createElement('label');
  wrap.className = 'fc-fmtdlg__check';
  const input = document.createElement('input');
  input.type = 'checkbox';
  const span = document.createElement('span');
  span.textContent = label;
  wrap.append(input, span);
  return { input, label: wrap };
};

const selectedRibbonPasteRadio = <T extends string>(
  root: HTMLElement,
  name: string,
  fallback: T,
): T =>
  (root.querySelector<HTMLInputElement>(`input[name="${name}"]:checked`)?.value as T | undefined) ??
  fallback;

const applyRibbonPasteSpecialSnapshot = (
  source: ClipboardSnapshot,
  opts: PasteSpecialOptions,
): boolean => {
  const i = inst;
  if (!i) return false;
  i.history.begin();
  let result: ReturnType<typeof applyPasteSpecial> = null;
  try {
    recordFormatChange(i.history, i.store, () => {
      result = applyPasteSpecial(i.store.getState(), i.store, i.workbook, source, opts);
    });
  } finally {
    i.history.end();
  }
  const applied = result as ReturnType<typeof applyPasteSpecial>;
  if (!applied) return false;
  mutators.setRange(i.store, applied.writtenRange);
  refreshWorkbookCells();
  focusSheet();
  return true;
};

const openRibbonPasteSpecialDialog = (source: ClipboardSnapshot): void => {
  const ja = ribbonLang === 'ja';
  const title = ja ? '形式を選択して貼り付け' : 'Paste Special';
  const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null;
  const overlay = document.createElement('div');
  overlay.className = 'fc-fmtdlg app__dlg fc-pastesp';
  overlay.setAttribute('role', 'dialog');
  overlay.setAttribute('aria-modal', 'true');
  overlay.setAttribute('aria-label', title);

  const panel = document.createElement('div');
  panel.className = 'fc-fmtdlg__panel app__dlg__panel fc-pastesp__panel';
  overlay.appendChild(panel);

  const header = document.createElement('div');
  header.className = 'fc-fmtdlg__header';
  header.textContent = title;
  panel.appendChild(header);

  const body = document.createElement('div');
  body.className = 'fc-fmtdlg__body fc-pastesp__body';
  panel.appendChild(body);

  const cols = document.createElement('div');
  cols.className = 'fc-pastesp__cols';
  body.appendChild(cols);

  const whatName = `app-ribbon-paste-what-${Math.random().toString(36).slice(2)}`;
  const whatGroup = document.createElement('div');
  whatGroup.className = 'fc-pastesp__group';
  const whatLegend = document.createElement('div');
  whatLegend.className = 'fc-pastesp__legend';
  whatLegend.textContent = ja ? '貼り付け' : 'Paste';
  const whatList = document.createElement('div');
  whatList.className = 'fc-pastesp__list';
  whatList.setAttribute('role', 'radiogroup');
  whatList.setAttribute('aria-label', whatLegend.textContent);
  for (const option of ribbonPasteWhatOptions()) {
    whatList.appendChild(
      makeRibbonPasteRadio(whatName, option.value, option.label, option.value === 'all'),
    );
  }
  whatGroup.append(whatLegend, whatList);
  cols.appendChild(whatGroup);

  const opName = `app-ribbon-paste-op-${Math.random().toString(36).slice(2)}`;
  const opGroup = document.createElement('div');
  opGroup.className = 'fc-pastesp__group';
  const opLegend = document.createElement('div');
  opLegend.className = 'fc-pastesp__legend';
  opLegend.textContent = ja ? '演算' : 'Operation';
  const opList = document.createElement('div');
  opList.className = 'fc-pastesp__list';
  opList.setAttribute('role', 'radiogroup');
  opList.setAttribute('aria-label', opLegend.textContent);
  for (const option of ribbonPasteOperationOptions()) {
    opList.appendChild(
      makeRibbonPasteRadio(opName, option.value, option.label, option.value === 'none'),
    );
  }
  opGroup.append(opLegend, opList);
  cols.appendChild(opGroup);

  const bottomRow = document.createElement('div');
  bottomRow.className = 'fc-pastesp__bottomrow';
  const skipBlanks = makeRibbonPasteCheck(ja ? '空白セルを無視する' : 'Skip blanks');
  const transpose = makeRibbonPasteCheck(ja ? '行/列の入れ替え' : 'Transpose');
  bottomRow.append(skipBlanks.label, transpose.label);
  body.appendChild(bottomRow);

  const footer = document.createElement('div');
  footer.className = 'fc-fmtdlg__footer';
  panel.appendChild(footer);
  const cancelBtn = document.createElement('button');
  cancelBtn.type = 'button';
  cancelBtn.className = 'fc-fmtdlg__btn';
  cancelBtn.textContent = ja ? 'キャンセル' : 'Cancel';
  const okBtn = document.createElement('button');
  okBtn.type = 'button';
  okBtn.className = 'fc-fmtdlg__btn fc-fmtdlg__btn--primary';
  okBtn.textContent = 'OK';
  footer.append(cancelBtn, okBtn);

  const close = (): void => {
    overlay.removeEventListener('keydown', onKey);
    overlay.remove();
    opener?.focus({ preventScroll: true });
  };
  const apply = (): void => {
    const what = selectedRibbonPasteRadio<PasteWhat>(overlay, whatName, 'all');
    const operation = selectedRibbonPasteRadio<PasteOperation>(overlay, opName, 'none');
    applyRibbonPasteSpecialSnapshot(source, {
      what,
      operation,
      skipBlanks: skipBlanks.input.checked,
      transpose: transpose.input.checked,
    });
    close();
  };
  const onKey = (event: KeyboardEvent): void => {
    event.stopPropagation();
    if (event.key === 'Escape') {
      event.preventDefault();
      close();
    } else if (event.key === 'Enter') {
      event.preventDefault();
      apply();
    }
  };
  cancelBtn.addEventListener('click', close);
  okBtn.addEventListener('click', apply);
  overlay.addEventListener('keydown', onKey);
  overlay.addEventListener('click', (event) => {
    if (event.target === overlay) close();
  });
  document.body.appendChild(overlay);
  requestAnimationFrame(() => {
    whatList.querySelector<HTMLInputElement>('input[type="radio"]')?.focus();
  });
};

const applyRibbonPasteAction = async (action: string): Promise<void> => {
  const i = inst;
  if (!i) return;
  if (action === 'dialog') {
    if (ribbonClipboardSnapshot) {
      openRibbonPasteSpecialDialog(ribbonClipboardSnapshot);
      return;
    }
    i.openPasteSpecial();
    return;
  }
  const opts = pasteOptionsForAction(action);
  if (!opts) return;
  let text = '';
  try {
    text = (await navigator.clipboard?.readText()) ?? '';
  } catch {
    text = '';
  }
  if (!text && ribbonClipboardText) text = ribbonClipboardText;
  if (ribbonClipboardSnapshot && text === ribbonClipboardText) {
    const source = ribbonClipboardSnapshot;
    applyRibbonPasteSpecialSnapshot(source, opts);
    return;
  }
  if (action === 'all' || action === 'values') {
    if (!text) return;
    let result: ReturnType<typeof pasteTSV> = null;
    i.history.begin();
    try {
      result = pasteTSV(i.store.getState(), i.workbook, text);
    } finally {
      i.history.end();
    }
    if (result) mutators.setRange(i.store, result.writtenRange);
    refreshWorkbookCells();
    focusSheet();
    return;
  }
  void showMessage({
    title: ribbonLang === 'ja' ? '貼り付け' : 'Paste',
    message:
      ribbonLang === 'ja'
        ? 'この貼り付け形式には、このブック内でコピーしたセルが必要です。'
        : 'This paste option requires cells copied inside this workbook.',
  });
};

const reviewCellsForSheet = (
  sheet: number,
  range?: { sheet: number; r0: number; c0: number; r1: number; c1: number },
): ReviewCell[] => {
  if (!inst) return [];
  return reviewCellsFromState(inst.store.getState(), sheet, range);
};

const showRibbonReport = (title: string, items: readonly RibbonReportItem[]): void => {
  void showReport({
    title,
    items,
    emptyLabel: ribbonReportText.noIssues,
    closeLabel: ribbonLang === 'ja' ? '閉じる' : 'Close',
    infoLabel: ribbonReportText.info,
    warningLabel: ribbonReportText.warning,
  });
};

const selectionRangeLabel = (range: Range): string => {
  const start = `${colLetter(range.c0)}${range.r0 + 1}`;
  const end = `${colLetter(range.c1)}${range.r1 + 1}`;
  return start === end ? start : `${start}:${end}`;
};

const scriptCommandLabel = (command: ScriptCommand): string => {
  switch (command) {
    case 'uppercase':
      return ribbonMenuText.scriptCommandUppercase;
    case 'lowercase':
      return ribbonMenuText.scriptCommandLowercase;
    case 'trim':
      return ribbonMenuText.scriptCommandTrim;
    case 'clear':
      return ribbonMenuText.scriptCommandClear;
  }
};

const runAccessibilityCheck = (): void => {
  if (!inst) return;
  const sheet = inst.store.getState().data.sheetIndex;
  const items = analyzeAccessibilityCells(reviewCellsForSheet(sheet), ribbonLang);
  if (statusMetric)
    statusMetric.textContent = `${ribbonText.accessibility} · ${items.filter((i) => i.severity === 'warning').length} ${ribbonReportText.warning}`;
  showRibbonReport(ribbonText.accessibility, items);
};

const runSpellingReview = (): void => {
  if (!inst) return;
  const sheet = inst.store.getState().data.sheetIndex;
  const items = analyzeSpellingCells(reviewCellsForSheet(sheet), ribbonLang);
  if (statusMetric)
    statusMetric.textContent = `${ribbonText.spelling} · ${items.filter((i) => i.severity === 'warning').length} ${ribbonReportText.warning}`;
  showRibbonReport(ribbonText.spelling, items);
};

const openTranslateReview = (): void => {
  if (!inst) return;
  const state = inst.store.getState();
  const items = buildTranslationReviewItems(
    reviewCellsForSheet(state.data.sheetIndex, state.selection.range),
    ribbonLang,
  );
  showRibbonReport(ribbonText.translate, items);
};

const runPlaygroundScriptCommand = (op: ScriptCommand): void => {
  if (!inst) return;
  const range = inst.store.getState().selection.range;
  inst.history.begin();
  let changed = 0;
  try {
    changed = applyTextScriptToRange(inst.store.getState(), inst.workbook, range, op);
  } finally {
    inst.history.end();
  }
  refreshWorkbookCells();
  automationRuns.unshift({
    label: scriptCommandLabel(op),
    range: selectionRangeLabel(range),
    changed,
  });
  automationRuns.splice(8);
  if (statusMetric)
    statusMetric.textContent = ribbonMenuText.automationRunStatus.replace(
      '{count}',
      String(changed),
    );
  focusSheet();
};

const runPlaygroundScript = async (): Promise<void> => {
  if (!inst) return;
  const command = await showPrompt({
    title: ribbonMenuText.scriptDialogTitle,
    label: ribbonMenuText.scriptDialogCommand,
    placeholder: ribbonMenuText.scriptCommandPrompt,
    okLabel: ribbonMenuText.scriptDialogRun,
    cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
    validate: (value) => (parseScriptCommand(value) ? null : ribbonMenuText.scriptCommandInvalid),
  });
  if (!command || !inst) return;
  const op = parseScriptCommand(command);
  if (!op) return;
  runPlaygroundScriptCommand(op);
};

const applyScriptAction = async (action: string): Promise<void> => {
  if (action === 'custom') {
    await runPlaygroundScript();
    return;
  }
  const op = parseScriptCommand(action);
  if (!op) return;
  runPlaygroundScriptCommand(op);
};

const openAllScripts = (): void => {
  const t = ribbonMenuText;
  const items: RibbonReportItem[] = [
    {
      severity: 'info',
      label: t.automationBuiltInScriptsLabel,
      detail: t.automationBuiltInScriptsDetail,
    },
  ];
  if (automationRuns.length) {
    items.push(
      ...automationRuns.map((run) => ({
        severity: 'info' as const,
        label: `${t.automationRecentRunsLabel}: ${run.label}`,
        detail: t.automationRunDetail
          .replace('{command}', run.label)
          .replace('{range}', run.range)
          .replace('{count}', String(run.changed)),
      })),
    );
  } else {
    items.push({
      severity: 'info',
      label: t.automationRecentRunsLabel,
      detail: t.automationNoRuns,
    });
  }
  showRibbonReport(t.automationScriptsTitle, items);
};

const recordSelectedActions = (): void => {
  if (!inst) return;
  const range = selectionRangeLabel(inst.store.getState().selection.range);
  automationRuns.unshift({
    label: ribbonText.recordActions,
    range,
    changed: 0,
  });
  automationRuns.splice(8);
  if (statusMetric) statusMetric.textContent = `${ribbonText.recordActions} · ${range}`;
  showRibbonReport(ribbonText.recordActions, [
    {
      severity: 'info',
      label: ribbonMenuText.recordActionsStatus,
      detail: ribbonMenuText.recordActionsEmpty,
    },
  ]);
  focusSheet();
};

const openAddInManager = (): void => {
  showRibbonReport(ribbonText.addIn, [
    {
      severity: 'info',
      label: ribbonMenuText.addInBuiltInLabel,
      detail: ribbonMenuText.addInBuiltInDetail,
    },
    {
      severity: 'info',
      label: ribbonMenuText.addInExternalLabel,
      detail: ribbonMenuText.addInExternalDetail,
    },
  ]);
};

const applyAddInAction = (action: string): void => {
  const t = ribbonMenuText;
  if (action === 'get') {
    showRibbonReport(t.addInGet, [
      { severity: 'info', label: t.addInStoreLabel, detail: t.addInStoreDetail },
      { severity: 'info', label: t.addInBuiltInLabel, detail: t.addInBuiltInDetail },
    ]);
    return;
  }
  if (action === 'manage') {
    showRibbonReport(t.addInManage, [
      { severity: 'info', label: t.addInManagedStatus, detail: t.addInExternalDetail },
    ]);
    return;
  }
  if (action === 'my') {
    showRibbonReport(t.addInMy, [
      { severity: 'info', label: t.addInBuiltInLabel, detail: t.addInBuiltInDetail },
      { severity: 'info', label: t.addInExternalLabel, detail: t.addInExternalDetail },
    ]);
    return;
  }
  openAddInManager();
};

const applyPdfAction = (action: string): void => {
  const i = inst;
  if (!i) return;
  if (action === 'preferences') {
    i.openPageSetup();
    return;
  }
  i.print('pdf');
  if (action === 'share') {
    showRibbonReport(ribbonText.pdf, [
      { severity: 'info', label: ribbonMenuText.pdfShare, detail: ribbonMenuText.pdfShareReady },
    ]);
  }
};

const runFormulaErrorChecking = (): void => {
  const i = inst;
  if (!i) return;
  const found = selectNextFormulaError(i.store);
  if (found) {
    projectFormatToolbar();
    focusSheet();
    return;
  }
  void showMessage({
    title: ribbonLang === 'ja' ? 'エラー チェック' : 'Error Checking',
    message:
      ribbonLang === 'ja'
        ? '選択範囲に数式エラーは見つかりませんでした。'
        : 'No formula errors were found in the selected range.',
  });
};

const saveCurrentSheetViewFromRibbon = async (): Promise<void> => {
  const i = inst;
  if (!i) return;
  const count = i.store.getState().sheetViews.views.length + 1;
  const defaultName = `${dictionaries[ribbonLang].viewToolbar.views} ${count}`;
  const name = await showPrompt({
    title: dictionaries[ribbonLang].viewToolbar.saveView,
    label: dictionaries[ribbonLang].viewToolbar.views,
    initial: defaultName,
    okLabel: 'OK',
    cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
    validate: (value) =>
      value.trim() ? null : ribbonLang === 'ja' ? '名前を入力してください。' : 'Enter a name.',
  });
  const trimmed = name?.trim();
  if (!trimmed) {
    focusSheet();
    return;
  }
  const id = `view-${Date.now().toString(36)}-${count}`;
  recordSheetViewsChange(i.history, i.store, () => {
    saveSheetView(i.store, id, trimmed);
    i.store.setState((s) => ({ ...s, sheetViews: { ...s.sheetViews, activeViewId: id } }));
  });
  projectFormatToolbar();
  focusSheet();
};

const deleteActiveSheetViewFromRibbon = (): void => {
  const i = inst;
  if (!i) return;
  const id = i.store.getState().sheetViews.activeViewId;
  if (!id) {
    focusSheet();
    return;
  }
  deleteSheetView(i.store, id, i.history);
  projectFormatToolbar();
  focusSheet();
};

const applyRibbonCommand = (id: string): boolean => {
  const i = inst;
  if (!i) return false;
  const state = i.store.getState();
  const range = state.selection.range;
  switch (id) {
    case 'pageSetup':
    case 'pageSetupAdvanced':
    case 'printTitles':
      i.openPageSetup();
      return true;
    case 'pageBreaks':
      applyPageBreakAction();
      return true;
    case 'pageTheme':
      applyUiTheme(uiTheme === 'dark' ? 'light' : 'dark');
      focusSheet();
      return true;
    case 'sheetBackground':
      void applySheetBackgroundAction('set');
      return true;
    case 'print':
    case 'printPageLayout':
      i.print('print');
      return true;
    case 'pdf':
      applyPdfAction('create');
      return true;
    case 'links':
    case 'linksInsert':
    case 'linksData':
      i.openExternalLinksDialog();
      return true;
    case 'inspect':
      inspectWorkbookFromBackstage();
      return true;
    case 'formatCells':
    case 'formatCellsHome':
      i.openFormatDialog();
      return true;
    case 'gotoSpecial':
    case 'gotoSpecialHome':
      i.openGoToSpecial();
      return true;
    case 'paste':
      void pasteClipboardIntoSelection();
      return true;
    case 'cut':
      void cutSelectionToClipboard();
      return true;
    case 'copy':
      void copySelectionToClipboard();
      return true;
    case 'undoHome':
      if (i.undo()) focusSheet();
      return true;
    case 'redoHome':
      if (i.redo()) focusSheet();
      return true;
    case 'bold':
      applyRibbonFormat((s, store) => toggleBold(s, store));
      return true;
    case 'italic':
      applyRibbonFormat((s, store) => toggleItalic(s, store));
      return true;
    case 'underline':
      applyRibbonFormat((s, store) => toggleUnderline(s, store));
      return true;
    case 'strike':
      applyRibbonFormat((s, store) => toggleStrike(s, store));
      return true;
    case 'currency':
      applyRibbonFormat((s, store) => cycleCurrency(s, store));
      return true;
    case 'percent':
      applyRibbonFormat((s, store) => cyclePercent(s, store));
      return true;
    case 'comma':
      applyRibbonFormat((s, store) =>
        setNumFmt(s, store, { kind: 'fixed', decimals: 2, thousands: true }),
      );
      return true;
    case 'fontGrow':
      applyRibbonFormat((s, store) => {
        const a = s.selection.active;
        const f = s.format.formats.get(`${a.sheet}:${a.row}:${a.col}`);
        setFont(s, store, { fontSize: (f?.fontSize ?? 11) + 1 });
      });
      return true;
    case 'fontShrink':
      applyRibbonFormat((s, store) => {
        const a = s.selection.active;
        const f = s.format.formats.get(`${a.sheet}:${a.row}:${a.col}`);
        setFont(s, store, { fontSize: Math.max(1, (f?.fontSize ?? 11) - 1) });
      });
      return true;
    case 'alignL':
      applyRibbonFormat((s, store) => setAlign(s, store, 'left'));
      return true;
    case 'alignC':
      applyRibbonFormat((s, store) => setAlign(s, store, 'center'));
      return true;
    case 'alignR':
      applyRibbonFormat((s, store) => setAlign(s, store, 'right'));
      return true;
    case 'top':
      applyRibbonFormat((s, store) => setVAlign(s, store, 'top'));
      return true;
    case 'middle':
      applyRibbonFormat((s, store) => setVAlign(s, store, 'middle'));
      return true;
    case 'decUp':
      applyRibbonFormat((s, store) => bumpDecimals(s, store, 1));
      return true;
    case 'decDown':
      applyRibbonFormat((s, store) => bumpDecimals(s, store, -1));
      return true;
    case 'wrap':
      applyRibbonFormat((s, store) => toggleWrap(s, store));
      return true;
    case 'merge': {
      const anchorAt0 = state.merges.byAnchor.get(`${range.sheet}:${range.r0}:${range.c0}`);
      const isExactMerge =
        anchorAt0 &&
        range.r0 === anchorAt0.r0 &&
        range.c0 === anchorAt0.c0 &&
        range.r1 === anchorAt0.r1 &&
        range.c1 === anchorAt0.c1;
      if (isExactMerge) applyUnmerge(i.store, i.workbook, i.history, range);
      else applyMerge(i.store, i.workbook, i.history, range);
      focusSheet();
      return true;
    }
    case 'formatPainter': {
      const fp = i.formatPainter;
      if (!fp) return true;
      if (fp.isActive()) fp.deactivate();
      else fp.activate(false);
      projectFormatToolbar();
      focusSheet();
      return true;
    }
    case 'pivotTableInsert':
      i.openPivotTableDialog();
      return true;
    case 'hyperlinkInsert':
      i.openHyperlinkDialog();
      return true;
    case 'commentInsert':
      i.openCommentDialog();
      return true;
    case 'clearFormat':
      applyRibbonFormat((s, store) => clearVisualFormat(s, store));
      return true;
    case 'general':
      applyRibbonFormat((s, store) => setNumFmt(s, store, { kind: 'general' }));
      return true;
    case 'conditional':
      i.openConditionalDialog();
      return true;
    case 'cellStyles':
      i.openCellStylesGallery();
      return true;
    case 'rules':
      i.openCfRulesDialog();
      return true;
    case 'formatTableHome':
      void createTableFromSelection('medium');
      return true;
    case 'insertRows':
      insertRows(i.store, i.workbook, i.history, range.r0, selectedRowCount());
      refreshWorkbookCells();
      focusSheet();
      return true;
    case 'deleteRows':
      deleteRows(i.store, i.workbook, i.history, range.r0, selectedRowCount());
      refreshWorkbookCells();
      focusSheet();
      return true;
    case 'insertCols':
      insertCols(i.store, i.workbook, i.history, range.c0, selectedColCount());
      refreshWorkbookCells();
      focusSheet();
      return true;
    case 'deleteCols':
      deleteCols(i.store, i.workbook, i.history, range.c0, selectedColCount());
      refreshWorkbookCells();
      focusSheet();
      return true;
    case 'sortAscHome':
    case 'sortAsc':
    case 'sortFilterHome':
      sortSelection('asc');
      return true;
    case 'sortDesc':
      sortSelection('desc');
      return true;
    case 'sortData':
      void customSortSelection();
      return true;
    case 'filterHome':
      openFilterForSelection();
      return true;
    case 'outlineGroup':
      applyOutlineAction('group');
      return true;
    case 'outlineUngroup':
      applyOutlineAction('ungroup');
      return true;
    case 'outlineShowDetail':
      applyOutlineAction('show-detail');
      return true;
    case 'outlineHideDetail':
      applyOutlineAction('hide-detail');
      return true;
    case 'bottomAlign':
      applyRibbonFormat((s, store) => setVAlign(s, store, 'bottom'));
      return true;
    case 'textOrientation':
      applyRibbonFormat((s, store) => setRotation(s, store, 45));
      return true;
    case 'indentDecrease':
      applyRibbonFormat((s, store) => bumpIndent(s, store, -1));
      return true;
    case 'indentIncrease':
      applyRibbonFormat((s, store) => bumpIndent(s, store, 1));
      return true;
    case 'moreBorders':
      i.openFormatDialog();
      return true;
    case 'drawBorder':
      if (i.borderDraw?.getMode() === 'draw') i.borderDraw.deactivate();
      else i.borderDraw?.activate('draw', selectedBorderStyle, selectedBorderColor);
      projectFormatToolbar();
      focusSheet();
      return true;
    case 'drawBorderGrid':
    case 'drawGrid':
      if (i.borderDraw?.getMode() === 'grid') i.borderDraw.deactivate();
      else i.borderDraw?.activate('grid', selectedBorderStyle, selectedBorderColor);
      projectFormatToolbar();
      focusSheet();
      return true;
    case 'eraseBorder':
      if (i.borderDraw?.getMode() === 'erase') i.borderDraw.deactivate();
      else i.borderDraw?.activate('erase', selectedBorderStyle, selectedBorderColor);
      projectFormatToolbar();
      focusSheet();
      return true;
    case 'drawPen':
      i.borderDraw?.deactivate();
      setDrawInkMode('pen');
      return true;
    case 'drawErase':
      i.borderDraw?.deactivate();
      setDrawInkMode('erase');
      return true;
    case 'findHome':
    case 'findReview':
      i.openFindReplace();
      return true;
    case 'spellingReview':
      runSpellingReview();
      return true;
    case 'translateReview':
      openTranslateReview();
      return true;
    case 'accessibility':
      runAccessibilityCheck();
      return true;
    case 'formatTableInsert':
      void createTableFromSelection('medium');
      return true;
    case 'namedRangesInsert':
    case 'namedRanges':
      i.openNamedRangeDialog();
      return true;
    case 'removeDupesInsert':
    case 'removeDupes':
      removeDuplicateRows();
      return true;
    case 'textToColumns':
      splitTextToColumns(',');
      return true;
    case 'dataValidation':
      i.openDataValidationDialog();
      return true;
    case 'chartInsert':
      createChartFromSelection();
      return true;
    case 'pictureInsert':
      void insertPictureFromRibbon('online');
      return true;
    case 'shapesInsert':
      insertShapeFromRibbon('rectangle');
      return true;
    case 'screenshotInsert':
      insertScreenshotFromRibbon();
      return true;
    case 'fxInsert':
    case 'fx':
      i.openFunctionArguments();
      return true;
    case 'autosum':
    case 'autosumFormula': {
      applyAutoSumFormula('SUM');
      return true;
    }
    case 'sum':
      i.openFunctionArguments('SUM');
      return true;
    case 'avg':
      i.openFunctionArguments('AVERAGE');
      return true;
    case 'ifFormula':
      i.openFunctionArguments('IF');
      return true;
    case 'xlookupFormula':
      i.openFunctionArguments('XLOOKUP');
      return true;
    case 'concatFormula':
      i.openFunctionArguments('CONCAT');
      return true;
    case 'todayFormula':
      i.openFunctionArguments('TODAY');
      return true;
    case 'pmtFormula':
      i.openFunctionArguments('PMT');
      return true;
    case 'roundFormula':
      i.openFunctionArguments('ROUND');
      return true;
    case 'precedents':
      if (i.tracePrecedents() === 0) {
        void showMessage({
          title: ribbonText.formulaAuditing,
          message: ribbonMenuText.traceNoPrecedents,
        });
      }
      return true;
    case 'dependents':
      if (i.traceDependents() === 0) {
        void showMessage({
          title: ribbonText.formulaAuditing,
          message: ribbonMenuText.traceNoDependents,
        });
      }
      return true;
    case 'clearArrows':
      i.clearTraces();
      return true;
    case 'errorChecking':
      runFormulaErrorChecking();
      return true;
    case 'showFormulasFormula':
      setShowFormulas(i.store, !i.store.getState().ui.showFormulas);
      projectFormatToolbar();
      focusSheet();
      return true;
    case 'evaluateFormula':
      i.openEvaluateFormulaDialog();
      return true;
    case 'recalcNow':
      i.recalc();
      focusSheet();
      return true;
    case 'calcOptions':
      i.openIterativeDialog();
      return true;
    case 'watch':
    case 'watchView':
      i.toggleWatchWindow();
      return true;
    case 'viewGridlines':
    case 'pageLayoutGridlinesView':
      setGridlinesVisible(i.store, i.store.getState().ui.showGridLines === false);
      projectFormatToolbar();
      focusSheet();
      return true;
    case 'pageLayoutGridlinesPrint': {
      const sheet = i.store.getState().data.sheetIndex;
      const setup = getPageSetup(i.store.getState(), sheet);
      setPrintGridlines(i.store, sheet, !setup.showGridlines, i.history);
      projectFormatToolbar();
      focusSheet();
      return true;
    }
    case 'viewHeadings':
    case 'pageLayoutHeadingsView':
      setHeadingsVisible(i.store, i.store.getState().ui.showHeaders === false);
      projectFormatToolbar();
      focusSheet();
      return true;
    case 'pageLayoutHeadingsPrint': {
      const sheet = i.store.getState().data.sheetIndex;
      const setup = getPageSetup(i.store.getState(), sheet);
      setPrintHeadings(i.store, sheet, !setup.showHeadings, i.history);
      projectFormatToolbar();
      focusSheet();
      return true;
    }
    case 'viewNormal':
      setWorkbookView(i.store, 'normal');
      projectFormatToolbar();
      focusSheet();
      return true;
    case 'viewPageLayout':
      setWorkbookView(i.store, 'pageLayout');
      projectFormatToolbar();
      focusSheet();
      return true;
    case 'viewPageBreakPreview':
      setWorkbookView(i.store, 'pageBreakPreview');
      projectFormatToolbar();
      focusSheet();
      return true;
    case 'sheetViewSave':
      void saveCurrentSheetViewFromRibbon();
      return true;
    case 'sheetViewDelete':
      deleteActiveSheetViewFromRibbon();
      return true;
    case 'workbookObjectsView':
      i.openWorkbookObjects();
      return true;
    case 'viewFormulas':
      setShowFormulas(i.store, !i.store.getState().ui.showFormulas);
      projectFormatToolbar();
      focusSheet();
      return true;
    case 'viewFormulaBar':
      formulaBarVisible = !formulaBarVisible;
      i.setFeatures(playgroundFeatureFlags());
      projectFormatToolbar();
      focusSheet();
      return true;
    case 'viewR1C1':
      setR1C1ReferenceStyle(i.store, !i.store.getState().ui.r1c1);
      projectFormatToolbar();
      focusSheet();
      return true;
    case 'hideRows':
      hideRows(i.store, i.history, range.r0, range.r1, i.workbook);
      focusSheet();
      return true;
    case 'hideCols':
      hideCols(i.store, i.history, range.c0, range.c1, i.workbook);
      focusSheet();
      return true;
    case 'newCommentReview':
      i.openCommentDialog();
      return true;
    case 'deleteCommentReview':
      deleteActiveReviewComment();
      return true;
    case 'previousCommentReview':
      selectReviewComment(-1);
      return true;
    case 'nextCommentReview':
      selectReviewComment(1);
      return true;
    case 'protectReview':
    case 'protect':
      void runSheetProtectionFlow();
      return true;
    case 'protectWorkbookReview':
      void runWorkbookProtectionFlow(!isWorkbookStructureProtected(i.store.getState()));
      return true;
    case 'protectionReview':
      void applyProtectAction('allow-edit-ranges');
      return true;
    case 'script':
      void runPlaygroundScript();
      return true;
    case 'recordActions':
      recordSelectedActions();
      return true;
    case 'allScripts':
      openAllScripts();
      return true;
    case 'addIn':
      openAddInManager();
      return true;
    case 'zoomSelection': {
      const selected = i.store.getState().selection.range;
      const rowCount = Math.max(1, selected.r1 - selected.r0 + 1);
      const colCount = Math.max(1, selected.c1 - selected.c0 + 1);
      const scaleForRows = 20 / rowCount;
      const scaleForCols = 12 / colCount;
      setSheetZoom(i.store, Math.max(0.5, Math.min(4, scaleForRows, scaleForCols)), i.workbook);
      refreshZoom();
      focusSheet();
      return true;
    }
    case 'windowVisibility':
      i.openFormatDialog('more');
      return true;
    case 'zoomDialog':
      void showZoomDialogFromRibbon();
      return true;
    case 'zoom75':
      setSheetZoom(i.store, 0.75, i.workbook);
      refreshZoom();
      focusSheet();
      return true;
    case 'zoom100':
      setSheetZoom(i.store, 1, i.workbook);
      refreshZoom();
      focusSheet();
      return true;
    case 'zoom125':
      setSheetZoom(i.store, 1.25, i.workbook);
      refreshZoom();
      focusSheet();
      return true;
    default:
      return false;
  }
};

// ── Ribbon tab strip ────────────────────────────────────────────────────
const selectRibbonTab = (tabId: RibbonTab, focusTab = false): void => {
  if (!ribbonRoot) return;
  if (tabId === 'file') {
    openBackstage(focusTab ? 'tab' : 'back');
    return;
  }
  backstageOpen = false;
  ribbonDisplayMenuOpen = false;
  activeRibbonTab = tabId;
  for (const item of ribbonRoot.querySelectorAll<HTMLButtonElement>('[data-ribbon-tab]')) {
    const isActive = item.dataset.ribbonTab === activeRibbonTab;
    item.classList.toggle('demo__ribbon-tab--active', isActive);
    item.setAttribute('aria-selected', isActive ? 'true' : 'false');
    item.tabIndex = isActive ? 0 : -1;
    if (focusTab && isActive) item.focus({ preventScroll: true });
  }
  for (const panel of ribbonRoot.querySelectorAll<HTMLElement>('[data-ribbon-panel]')) {
    panel.hidden = panel.dataset.ribbonPanel !== activeRibbonTab;
  }
  ribbonRoot.querySelector('.demo__ribbon-display-menu')?.remove();
  ribbonRoot
    .querySelector<HTMLButtonElement>('[data-ribbon-toggle]')
    ?.setAttribute('aria-expanded', 'false');
};

const setRibbonCollapsed = (next: boolean): void => {
  ribbonCollapsed = next;
  for (const shell of ribbonRoot?.querySelectorAll<HTMLElement>('.demo__ribbon-shell') ?? []) {
    shell.classList.toggle('demo__ribbon-shell--collapsed', ribbonCollapsed);
  }
  for (const tabs of ribbonRoot?.querySelectorAll<HTMLElement>('.demo__ribbon-tabs') ?? []) {
    tabs.dataset.ribbonCollapsed = ribbonCollapsed ? 'true' : 'false';
  }
  for (const item of ribbonRoot?.querySelectorAll<HTMLButtonElement>(
    '[data-ribbon-display-option]',
  ) ?? []) {
    item.setAttribute(
      'aria-checked',
      item.dataset.ribbonDisplayOption === (ribbonCollapsed ? 'collapsed' : 'expanded')
        ? 'true'
        : 'false',
    );
  }
};

const openBackstage = (focus: 'back' | 'tab' = 'back'): void => {
  if (!backstageOpen && activeRibbonTab !== 'file') backstageReturnTab = activeRibbonTab;
  activeRibbonTab = 'file';
  backstageOpen = true;
  ribbonCollapsed = false;
  ribbonDisplayMenuOpen = false;
  renderRibbon();
  const selector = focus === 'tab' ? '[data-ribbon-tab="file"]' : '[data-backstage-action="back"]';
  ribbonRoot?.querySelector<HTMLButtonElement>(selector)?.focus({ preventScroll: true });
};

const closeBackstage = (focusTab = false): void => {
  if (!backstageOpen) return;
  backstageOpen = false;
  activeRibbonTab = backstageReturnTab;
  renderRibbon();
  if (focusTab) {
    ribbonRoot
      ?.querySelector<HTMLButtonElement>(`[data-ribbon-tab="${activeRibbonTab}"]`)
      ?.focus({ preventScroll: true });
  }
};

const titleActionButton = (label: string): HTMLButtonElement | null =>
  document.querySelector<HTMLButtonElement>(`.app__title [data-shell-i18n-label="${label}"]`);

titleActionButton('home')?.addEventListener('click', () => {
  closeBackstage();
  selectRibbonTab('home', true);
});

titleActionButton('save')?.addEventListener('click', () => {
  triggerSave();
});

titleActionButton('saveAs')?.addEventListener('click', () => {
  void triggerSaveAs();
});

titleActionButton('undo')?.addEventListener('click', () => {
  if (inst?.undo()) focusSheet();
});

titleActionButton('redo')?.addEventListener('click', () => {
  if (inst?.redo()) focusSheet();
});

titleActionButton('comments')?.addEventListener('click', () => {
  inst?.openCommentDialog();
});

titleActionButton('share')?.addEventListener('click', () => {
  showRibbonReport(shellText.share, [
    { severity: 'info', label: shellText.share, detail: shellText.shareReady },
  ]);
});

const seedFindDialogQuery = (query: string): void => {
  requestAnimationFrame(() => {
    const input = document.querySelector<HTMLInputElement>('.fc-find input[type="text"]');
    if (!input) return;
    input.value = query;
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.focus();
    input.select();
  });
};

const runTitleSearch = (): void => {
  const query = titleSearchInput?.value.trim() ?? '';
  if (!query || !inst) return;
  const state = inst.store.getState();
  const match = findNext(
    state,
    { query, within: 'sheet', searchBy: 'rows', lookIn: 'values' },
    state.selection.active,
    'next',
  );
  inst.openFindReplace('find');
  seedFindDialogQuery(query);
  if (match) {
    mutators.setActive(inst.store, match.addr);
    projectFormatToolbar();
  } else if (statusMetric) {
    statusMetric.textContent =
      ribbonLang === 'ja' ? `「${query}」は見つかりませんでした` : `No matches for "${query}"`;
  }
};

titleSearchInput?.addEventListener('keydown', (event) => {
  if (event.key !== 'Enter') return;
  event.preventDefault();
  runTitleSearch();
});

document.addEventListener('keydown', (event) => {
  if (event.key.toLowerCase() !== 'u' || !event.metaKey || !event.ctrlKey) return;
  event.preventDefault();
  titleSearchInput?.focus();
  titleSearchInput?.select();
});

const toggleAutosave = (): void => {
  autosaveEnabled = !autosaveEnabled;
  refreshAutosave();
  if (statusMetric)
    statusMetric.textContent = autosaveEnabled ? shellText.autosaveOn : shellText.autosaveOff;
};

autosaveSwitch?.addEventListener('click', toggleAutosave);

const titleMoreButton = titleActionButton('more');
const titleMoreMenu = createMenu('menu-title-more');
titleMoreMenu.classList.add('app__title-more-menu');
titleMoreMenu.append(
  menuButton(shellText.save, 'titleMoreAction', 'save'),
  menuButton(shellText.saveAs, 'titleMoreAction', 'save-as'),
  menuButton(shellText.autosave, 'titleMoreAction', 'autosave'),
  menuSeparator(),
  menuButton(shellText.comments, 'titleMoreAction', 'comments'),
  menuButton(shellText.share, 'titleMoreAction', 'share'),
);
document.body.appendChild(titleMoreMenu);

const closeTitleMoreMenu = (restoreFocus = false): void => {
  titleMoreMenu.hidden = true;
  titleMoreButton?.setAttribute('aria-expanded', 'false');
  if (restoreFocus) titleMoreButton?.focus({ preventScroll: true });
};

const openTitleMoreMenu = (): void => {
  if (!titleMoreButton) return;
  const rect = titleMoreButton.getBoundingClientRect();
  titleMoreMenu.style.left = `${Math.round(rect.left)}px`;
  titleMoreMenu.style.top = `${Math.round(rect.bottom + 4)}px`;
  titleMoreMenu.hidden = false;
  titleMoreButton.setAttribute('aria-haspopup', 'menu');
  titleMoreButton.setAttribute('aria-expanded', 'true');
  focusMenuItem(titleMoreMenu, 'first');
};

titleMoreButton?.addEventListener('click', () => {
  if (titleMoreMenu.hidden) openTitleMoreMenu();
  else closeTitleMoreMenu(true);
});

titleMoreMenu.addEventListener('click', (event) => {
  const item = (event.target as Element | null)?.closest<HTMLButtonElement>(
    '[data-title-more-action]',
  );
  const action = item?.dataset.titleMoreAction;
  if (!action) return;
  closeTitleMoreMenu();
  if (action === 'save') triggerSave();
  else if (action === 'save-as') void triggerSaveAs();
  else if (action === 'autosave') toggleAutosave();
  else if (action === 'comments') inst?.openCommentDialog();
  else if (action === 'share') {
    showRibbonReport(shellText.share, [
      { severity: 'info', label: shellText.share, detail: shellText.shareReady },
    ]);
  }
});

titleMoreMenu.addEventListener('keydown', (event) => {
  handleMenuKeydown(event, titleMoreMenu, {
    close: closeTitleMoreMenu,
    restoreFocusTo: titleMoreButton ?? undefined,
  });
});

document.addEventListener('pointerdown', (event) => {
  if (titleMoreMenu.hidden) return;
  const target = event.target as Element | null;
  if (titleMoreMenu.contains(target)) return;
  if (titleMoreButton?.contains(target)) return;
  closeTitleMoreMenu();
});

ribbonRoot?.addEventListener('click', (event) => {
  const toggle = (event.target as Element | null)?.closest<HTMLButtonElement>(
    '[data-ribbon-toggle]',
  );
  if (toggle) {
    event.preventDefault();
    ribbonDisplayMenuOpen = !ribbonDisplayMenuOpen;
    renderRibbon();
    ribbonRoot
      ?.querySelector<HTMLButtonElement>('[data-ribbon-toggle]')
      ?.focus({ preventScroll: true });
    return;
  }
  const displayOption = (event.target as Element | null)?.closest<HTMLButtonElement>(
    '[data-ribbon-display-option]',
  );
  if (displayOption) {
    event.preventDefault();
    ribbonDisplayMenuOpen = false;
    setRibbonCollapsed(displayOption.dataset.ribbonDisplayOption === 'collapsed');
    renderRibbon();
    return;
  }
  const tab = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-ribbon-tab]');
  if (!tab) return;
  ribbonDisplayMenuOpen = false;
  const nextTab = (tab.dataset.ribbonTab as RibbonTab | undefined) ?? 'home';
  if (nextTab === 'file') openBackstage();
  else {
    closeBackstage();
    selectRibbonTab(nextTab);
  }
});

ribbonRoot?.addEventListener('dblclick', (event) => {
  const tab = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-ribbon-tab]');
  if (!tab) return;
  setRibbonCollapsed(!ribbonCollapsed);
});

document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && backstageOpen) {
    event.preventDefault();
    closeBackstage(true);
    return;
  }
  if (event.key === 'Escape' && ribbonDisplayMenuOpen) {
    event.preventDefault();
    ribbonDisplayMenuOpen = false;
    renderRibbon();
    ribbonRoot
      ?.querySelector<HTMLButtonElement>('[data-ribbon-toggle]')
      ?.focus({ preventScroll: true });
    return;
  }
  if (event.key !== 'F1' || (!event.ctrlKey && !event.metaKey)) return;
  event.preventDefault();
  ribbonDisplayMenuOpen = false;
  setRibbonCollapsed(!ribbonCollapsed);
});

document.addEventListener('pointerdown', (event) => {
  if (!ribbonDisplayMenuOpen) return;
  const target = event.target as Element | null;
  if (target?.closest('.demo__ribbon-display')) return;
  ribbonDisplayMenuOpen = false;
  renderRibbon();
});

ribbonRoot?.addEventListener('keydown', (event) => {
  const display = (event.target as Element | null)?.closest<HTMLElement>('.demo__ribbon-display');
  if (!display) return;
  const options = Array.from(
    display.querySelectorAll<HTMLButtonElement>('[data-ribbon-display-option]'),
  );
  const focusOption = (index: number): void => {
    options[(index + options.length) % options.length]?.focus({ preventScroll: true });
  };
  const activeIndex = Math.max(0, options.indexOf(document.activeElement as HTMLButtonElement));
  if (event.key === 'ArrowDown') {
    event.preventDefault();
    if (!ribbonDisplayMenuOpen) {
      ribbonDisplayMenuOpen = true;
      renderRibbon();
      requestAnimationFrame(() =>
        ribbonRoot
          ?.querySelectorAll<HTMLButtonElement>('[data-ribbon-display-option]')[0]
          ?.focus({ preventScroll: true }),
      );
      return;
    }
    focusOption(activeIndex + 1);
  } else if (event.key === 'ArrowUp') {
    event.preventDefault();
    if (!ribbonDisplayMenuOpen) {
      ribbonDisplayMenuOpen = true;
      renderRibbon();
      requestAnimationFrame(() => {
        const nextOptions = ribbonRoot?.querySelectorAll<HTMLButtonElement>(
          '[data-ribbon-display-option]',
        );
        nextOptions?.[nextOptions.length - 1]?.focus({ preventScroll: true });
      });
      return;
    }
    focusOption(activeIndex - 1);
  } else if (event.key === 'Home' && options.length) {
    event.preventDefault();
    focusOption(0);
  } else if (event.key === 'End' && options.length) {
    event.preventDefault();
    focusOption(options.length - 1);
  }
});

ribbonRoot?.addEventListener('keydown', (event) => {
  const tab = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-ribbon-tab]');
  if (!tab) return;
  const tabs = Array.from(
    ribbonRoot.querySelectorAll<HTMLButtonElement>('[data-ribbon-tab]'),
  ).filter((item) => item.offsetParent !== null);
  const current = Math.max(0, tabs.indexOf(tab));
  let next = current;
  if (event.key === 'ArrowRight') next = (current + 1) % tabs.length;
  else if (event.key === 'ArrowLeft') next = (current - 1 + tabs.length) % tabs.length;
  else if (event.key === 'Home') next = 0;
  else if (event.key === 'End') next = tabs.length - 1;
  else return;
  event.preventDefault();
  const nextTab = tabs[next]?.dataset.ribbonTab as RibbonTab | undefined;
  if (nextTab) selectRibbonTab(nextTab, true);
});

ribbonRoot?.addEventListener('click', (event) => {
  const button = (event.target as Element | null)?.closest<HTMLButtonElement>(
    'button[data-ribbon-command]',
  );
  if (!button || button.disabled) return;
  const id = button.dataset.ribbonCommand;
  if (!id) return;
  if (id === 'printArea') {
    event.preventDefault();
    event.stopPropagation();
    const menu = getPrintAreaMenu();
    if (!menu) return;
    if (menu.hidden) openPrintAreaMenu(button);
    else closePrintAreaMenu(true);
    return;
  }
  if (id === 'symbolInsert') {
    event.preventDefault();
    event.stopPropagation();
    const menu = getSymbolMenu();
    if (!menu) return;
    if (menu.hidden) openSymbolMenu(button);
    else closeSymbolMenu(true);
    return;
  }
  if (id === 'borders') {
    event.preventDefault();
    event.stopPropagation();
    const menu = getBorderMenu();
    if (!menu) return;
    if (menu.hidden) openBorderMenu();
    else closeBorderMenu(true);
    return;
  }
  if (id === 'freeze') {
    event.preventDefault();
    event.stopPropagation();
    const menu = getFreezeMenu();
    if (!menu) return;
    if (menu.hidden) openFreezeMenu();
    else closeFreezeMenu(true);
    return;
  }
  const dynamicSpec = dynamicDropdownSpecForButton(button);
  if (dynamicSpec) {
    event.preventDefault();
    event.stopPropagation();
    const menu = document.getElementById(dynamicSpec.menuId) as HTMLDivElement | null;
    if (!menu) return;
    if (menu.hidden) openDynamicRibbonDropdown(dynamicSpec, button);
    else closeDynamicRibbonDropdown(dynamicSpec, true);
    return;
  }
  if (legacyCommandIds[id] && button.dataset.legacyBound === '1') return;
  if (applyRibbonCommand(id)) {
    event.preventDefault();
    event.stopPropagation();
  }
});

// ── View menu (Show Formulas / R1C1 / Grid / Headers toggles) ────────────
const viewBtn = document.getElementById('menu-view');
const viewDrop = document.getElementById('menu-view-dropdown');
const closeViewMenu = (): void => {
  if (!viewDrop) return;
  viewDrop.hidden = true;
  viewBtn?.setAttribute('aria-expanded', 'false');
};
const refreshViewMenu = (): void => {
  if (!inst || !viewDrop) return;
  const ui = inst.store.getState().ui;
  const update = (action: string, on: boolean): void => {
    const item = viewDrop.querySelector<HTMLElement>(`[data-view="${action}"] [data-fc-check]`);
    if (item) item.textContent = on ? '✓' : '';
  };
  update('show-formulas', !!ui.showFormulas);
  update('r1c1', !!ui.r1c1);
  update('grid', ui.showGridLines !== false);
  update('headers', ui.showHeaders !== false);
};
viewBtn?.addEventListener('click', (e) => {
  e.stopPropagation();
  if (!viewDrop) return;
  refreshViewMenu();
  viewDrop.hidden = !viewDrop.hidden;
  viewBtn.setAttribute('aria-expanded', viewDrop.hidden ? 'false' : 'true');
});
document.addEventListener('mousedown', (e) => {
  if (!viewDrop || viewDrop.hidden) return;
  if (viewDrop.contains(e.target as Node) || viewBtn?.contains(e.target as Node)) return;
  closeViewMenu();
});
viewDrop?.querySelectorAll<HTMLButtonElement>('[data-view]').forEach((btn) => {
  btn.addEventListener('click', () => {
    if (!inst) return;
    const action = btn.dataset.view;
    const ui = inst.store.getState().ui;
    if (action === 'show-formulas') mutators.setShowFormulas(inst.store, !ui.showFormulas);
    else if (action === 'r1c1') mutators.setR1C1(inst.store, !ui.r1c1);
    else if (action === 'grid') mutators.setShowGridLines(inst.store, !ui.showGridLines);
    else if (action === 'headers') mutators.setShowHeaders(inst.store, !ui.showHeaders);
    refreshViewMenu();
  });
});

// ── Tools menu (Iterative / Names / Conditional) ─────────────────────────
const toolsBtn = document.getElementById('menu-tools');
const toolsDrop = document.getElementById('menu-tools-dropdown');
const closeToolsMenu = (): void => {
  if (!toolsDrop) return;
  toolsDrop.hidden = true;
  toolsBtn?.setAttribute('aria-expanded', 'false');
};
toolsBtn?.addEventListener('click', (e) => {
  e.stopPropagation();
  if (!toolsDrop) return;
  toolsDrop.hidden = !toolsDrop.hidden;
  toolsBtn.setAttribute('aria-expanded', toolsDrop.hidden ? 'false' : 'true');
});
document.addEventListener('mousedown', (e) => {
  if (!toolsDrop || toolsDrop.hidden) return;
  if (toolsDrop.contains(e.target as Node) || toolsBtn?.contains(e.target as Node)) return;
  closeToolsMenu();
});
toolsDrop?.querySelectorAll<HTMLButtonElement>('[data-tools]').forEach((btn) => {
  btn.addEventListener('click', () => {
    if (!inst) return;
    const action = btn.dataset.tools;
    closeToolsMenu();
    if (action === 'iterative') inst.openIterativeDialog();
    else if (action === 'named') inst.openNamedRangeDialog();
    else if (action === 'conditional') inst.openConditionalDialog();
  });
});

// ── Sheet tabs ───────────────────────────────────────────────────────────
const tabsList = document.getElementById('sheet-tabs');
const tabAddBtn = document.getElementById('btn-sheet-add');
const tabPrevBtn = document.getElementById('btn-sheet-prev');
const tabNextBtn = document.getElementById('btn-sheet-next');

const renderSheetTabs = (): void => {
  if (!inst || !tabsList) return;
  const wb = inst.workbook;
  const state = inst.store.getState();
  const activeIdx = state.data.sheetIndex;
  const hidden = state.layout.hiddenSheets;
  const n = wb.sheetCount;
  tabsList.replaceChildren();
  for (let i = 0; i < n; i += 1) {
    if (hidden.has(i)) continue;
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'app__tab';
    if (i === activeIdx) btn.classList.add('app__tab--active');
    btn.setAttribute('role', 'tab');
    btn.setAttribute('aria-selected', i === activeIdx ? 'true' : 'false');
    const tabColor = state.layout.sheetTabColors.get(i);
    if (tabColor) {
      btn.dataset.sheetTabColor = 'true';
      btn.style.setProperty('--app-sheet-tab-color', tabColor);
    }
    const label = document.createElement('span');
    label.className = 'app__tab-label';
    label.textContent = wb.sheetName(i);
    btn.appendChild(label);
    btn.addEventListener('click', () => switchSheet(i));
    btn.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      openTabMenu(i, e.clientX, e.clientY);
    });
    tabsList.appendChild(btn);
  }
  // "Unhide…" affordance — surfaced as an extra tab pill when at least one
  // sheet is hidden. Click opens a list of hidden sheets to restore.
  if (hidden.size > 0) {
    const unhide = document.createElement('button');
    unhide.type = 'button';
    unhide.className = 'app__tab app__tab--unhide';
    unhide.textContent = `Unhide… (${hidden.size})`;
    unhide.addEventListener('click', (e) => {
      const r = (e.currentTarget as HTMLElement).getBoundingClientRect();
      openUnhideMenu(r.left, r.bottom);
    });
    tabsList.appendChild(unhide);
  }
};

const openUnhideMenu = (x: number, y: number): void => {
  if (!inst) return;
  closeTabMenu();
  const wb = inst.workbook;
  const store = inst.store;
  const hidden = store.getState().layout.hiddenSheets;
  if (hidden.size === 0) return;

  const menu = document.createElement('div');
  menu.className = 'app__menu';
  prepareMenu(menu, 'Unhide sheet');
  menu.style.position = 'fixed';
  menu.style.left = `${x}px`;
  menu.style.top = `${y}px`;
  menu.style.zIndex = '90';
  let cleanupMenuListeners = (): void => {};

  for (const i of Array.from(hidden).sort((a, b) => a - b)) {
    const it = document.createElement('button');
    it.type = 'button';
    it.className = 'app__menu-item';
    it.setAttribute('role', 'menuitem');
    it.tabIndex = -1;
    it.textContent = wb.sheetName(i);
    it.addEventListener('click', () => {
      closeTabMenu();
      cleanupMenuListeners();
      if (setSheetHidden(store, wb, inst?.history ?? null, i, false)) {
        renderSheetTabs();
      }
    });
    menu.appendChild(it);
  }

  document.body.appendChild(menu);
  tabMenuEl = menu;
  focusMenuItem(menu);

  const rect = menu.getBoundingClientRect();
  if (rect.right > window.innerWidth) {
    menu.style.left = `${Math.max(0, window.innerWidth - rect.width - 4)}px`;
  }
  if (rect.bottom > window.innerHeight) {
    menu.style.top = `${Math.max(0, window.innerHeight - rect.height - 4)}px`;
  }

  const onDocDown = (ev: MouseEvent): void => {
    if (!tabMenuEl) return;
    if (ev.target instanceof Node && tabMenuEl.contains(ev.target)) return;
    closeTabMenu();
    cleanupMenuListeners();
  };
  const onDocKey = (ev: KeyboardEvent): void => {
    handleMenuKeydown(ev, menu, {
      close: (restoreFocus) => {
        closeTabMenu();
        cleanupMenuListeners();
        if (restoreFocus) {
          document.querySelector<HTMLButtonElement>('.app__tab--unhide')?.focus();
        }
      },
    });
  };
  cleanupMenuListeners = () => {
    document.removeEventListener('mousedown', onDocDown, true);
    document.removeEventListener('keydown', onDocKey, true);
  };
  document.addEventListener('mousedown', onDocDown, true);
  document.addEventListener('keydown', onDocKey, true);
};

let tabMenuEl: HTMLDivElement | null = null;
const closeTabMenu = (): void => {
  if (!tabMenuEl) return;
  tabMenuEl.remove();
  tabMenuEl = null;
};

const openTabMenu = (idx: number, x: number, y: number): void => {
  if (!inst) return;
  openSheetTabMenu({
    closeTabMenu,
    idx,
    inst,
    renderSheetTabs,
    setTabMenuEl: (el) => {
      tabMenuEl = el;
    },
    x,
    y,
  });
};

const switchSheet = (idx: number): void => {
  if (!inst) return;
  const n = inst.workbook.sheetCount;
  if (idx < 0 || idx >= n) return;
  if (inst.store.getState().data.sheetIndex === idx) return;
  mutators.setSheetIndex(inst.store, idx);
  mutators.replaceCells(inst.store, inst.workbook.cells(idx));
  renderSheetTabs();
  (sheetEl as HTMLElement).focus();
};

tabAddBtn?.addEventListener('click', () => {
  if (!inst) return;
  const idx = addSheet(inst.store, inst.workbook);
  if (idx < 0) {
    if (statusMetric && isWorkbookStructureProtected(inst.store.getState())) {
      statusMetric.textContent = ribbonMenuText.workbookStructureProtectedBlocked;
    }
    return;
  }
  // The wb.subscribe handler in mount.ts will pick up sheet-add as a no-op for cells,
  // but we re-render tabs and switch to the new sheet here.
  renderSheetTabs();
  switchSheet(idx);
});

const { refreshZoom } = setupZoomControls(() => inst);

tabPrevBtn?.addEventListener('click', () => {
  if (!inst) return;
  switchSheet(inst.store.getState().data.sheetIndex - 1);
});
tabNextBtn?.addEventListener('click', () => {
  if (!inst) return;
  switchSheet(inst.store.getState().data.sheetIndex + 1);
});

// ── Merge / Wrap / Sort buttons ───────────────────────────────────────────
document.getElementById('btn-merge')?.addEventListener('click', () => {
  if (!inst) return;
  const s = inst.store.getState();
  const r = s.selection.range;
  const anchorAt0 = s.merges.byAnchor.get(`${r.sheet}:${r.r0}:${r.c0}`);
  const isExactMerge =
    anchorAt0 &&
    r.r0 === anchorAt0.r0 &&
    r.c0 === anchorAt0.c0 &&
    r.r1 === anchorAt0.r1 &&
    r.c1 === anchorAt0.c1;
  if (isExactMerge) applyUnmerge(inst.store, inst.workbook, inst.history, r);
  else applyMerge(inst.store, inst.workbook, inst.history, r);
  (sheetEl as HTMLElement).focus();
});

document.getElementById('btn-wrap')?.addEventListener('click', () => {
  if (!inst) return;
  const current = inst;
  recordFormatChange(inst.history, inst.store, () => {
    toggleWrap(current.store.getState(), current.store);
  });
  (sheetEl as HTMLElement).focus();
});
markCurrentLegacyRibbonBindings();

let filterDropdown: ReturnType<typeof attachFilterDropdown> | null = null;

boot().catch((err) => {
  // eslint-disable-next-line no-console
  console.error('formulon-cell boot failed', err);
  if (sheetEl) {
    sheetEl.innerHTML = `<pre style="padding:24px;color:#d24545;font-family:'IBM Plex Mono',monospace;white-space:pre-wrap">${
      err instanceof Error ? (err.stack ?? err.message) : String(err)
    }</pre>`;
  }
});
