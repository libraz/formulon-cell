import {
  type AutoSumFormulaName,
  activateSheetView,
  addAllowedEditRange,
  addSheet,
  aggregateSelection,
  analyzeAccessibilityCells,
  analyzeSpellingCells,
  applyAdvancedFilter,
  applyCellFormatAction as applyCellFormatActionImpl,
  applyCellStyle,
  applyConditionalMenuAction as applyConditionalMenuActionImpl,
  applyConditionalPresetAction,
  applyMerge,
  applyPasteSpecial,
  applyTextScriptToRange,
  applyUnmerge,
  type attachFilterDropdown,
  autofitColWidth,
  autofitRowHeight,
  autoSum,
  BORDER_PRESETS,
  type BorderPreviewSpec,
  backstageMenuText,
  boundingRange,
  buildCfMenuText,
  buildRibbonModel,
  buildSpreadsheetCompatibilityReport,
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
  copy,
  copyAdvancedFilterResult,
  createBackstageFactories,
  createBackstageTitle,
  createBorderMenu,
  createBorderPreview,
  createBordersMenu as createBordersMenuImpl,
  createColorPalette,
  createConditionalMenu as createConditionalMenuImpl,
  createControlDispatch,
  createDefinedNamesFromSelection,
  createDynamicDropdowns,
  createFormulasMenuFactories,
  createHomeMenuFactories,
  createInsertMenuFactories,
  createLineSamplePreview,
  createMenu,
  createPageLayoutMenuFactories,
  createPasteMenu as createPasteMenuImpl,
  createPivotTableFromRange,
  createReviewMenuFactories,
  createRibbonChartFromSelection,
  createSelectColorRibbon,
  createStylesMenuFactories,
  createTextOrientationMenu as createTextOrientationMenuImpl,
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
  fillSeriesSourceRange,
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
  inferFillSeriesDirection,
  inferPivotSourceFields,
  inferRecommendedChartKind,
  inferSortHasHeader,
  insertCells,
  insertCols,
  insertDefinedNameFormula,
  insertManualPageBreak,
  insertRows,
  isCellWritable,
  isJapaneseFontName,
  isWorkbookStructureProtected,
  LEGACY_COMMAND_IDS,
  LINE_STYLES_ALL,
  listComments,
  listDefinedNames,
  type MarginPreset,
  marginPresetOf,
  menuButton,
  menuSectionHeader,
  menuSeparator,
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
  RIBBON_BORDER_DRAW_MODES,
  RIBBON_DIALOG_OPENERS,
  RIBBON_FORMAT_MUTATORS,
  RIBBON_FUNCTION_ARG_OPENERS,
  RIBBON_KEYSHORTCUTS,
  RIBBON_VIEW_MODES,
  RIBBON_ZOOM_PRESETS,
  type RibbonCommand,
  type RibbonFillDirection,
  type RibbonFillSeriesMode,
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
  SPLIT_BUTTON_COMMANDS,
  Spreadsheet,
  type SpreadsheetInstance,
  SVG_NS,
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
  shouldShowFontOption,
  showCols,
  showFillSeriesDialog,
  showRows,
  sortRange,
  TABLE_STYLE_COLORS,
  type TableStyle,
  type TableVariantId,
  tableStyleSwatch,
  tableVariantOptions,
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
import { createBootWiring } from './boot-wiring.js';
import { createClipboard } from './clipboard.js';
import { createCommandPalette } from './command-palette.js';
import { createDataMenuWirings } from './data-menu-wirings.js';
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
import { createHomeMenuWirings } from './home-menu-wirings.js';
import { createIllustrations, type SessionShapeKind } from './illustrations.js';
import { focusMenuItem, handleMenuKeydown, prepareMenu } from './menu-a11y.js';
import { createProtectionFlows } from './protection-flows.js';
import { createRangeUtils } from './range-utils.js';
import { createRibbonActions, type PrintTitlesAction } from './ribbon-actions.js';
import { createScriptAddinActions } from './script-addin-actions.js';
import { seedWorkbook } from './seed.js';
import { openSheetTabMenu } from './sheet-tab-menu.js';
import { createSheetTabs } from './sheet-tabs-runtime.js';
import { createShellLocale } from './shell-locale.js';
import { createShellMenus } from './shell-menus.js';
import { createSortFilter } from './sort-filter.js';
import { ACTIVE_CLASS, createStatusProjection } from './status-projection.js';
import { createWorkbookActions } from './workbook-actions.js';
import { createXlsxIo } from './xlsx-io.js';
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
const {
  createSymbolMenu,
  createPivotTableMenu,
  createDefinedNamesMenu,
  createLinksMenu,
  createDataValidationMenu,
  createChartInsertMenu,
  createPictureInsertMenu,
  createShapesInsertMenu,
  createScreenshotInsertMenu,
  createScriptMenu,
  createAddInMenu,
  createPdfMenu,
} = createInsertMenuFactories(ribbonMenuText);
const {
  createPrintAreaMenu,
  createPageBreaksMenu,
  createSheetBackgroundMenu,
  createPrintTitlesMenu,
  createPageThemeMenu,
} = createPageLayoutMenuFactories(ribbonMenuText);
const { createAutoSumMenu, createCalcOptionsMenu, createClearArrowsMenu, createErrorCheckingMenu } =
  createFormulasMenuFactories(ribbonMenuText, ribbonLang);
const { createWatchMenu, createReviewCommentsMenu, createProtectMenu } =
  createReviewMenuFactories(ribbonMenuText);
const pageScaleText = pageScaleMenuText(ribbonLang);
const {
  createFreezeMenu,
  createFillMenu,
  createClearMenu,
  createInsertCellsMenu,
  createDeleteCellsMenu,
  createFormatCellsMenu,
  createSortMenu,
  createTextToColumnsMenu,
  createFindSelectMenu,
} = createHomeMenuFactories({
  ribbonLang,
  ribbonMenuText,
  ribbonText,
  sheetTabs: dictionaries[ribbonLang].sheetTabs,
});
const createTextOrientationMenu = (): HTMLDivElement =>
  createTextOrientationMenuImpl(ribbonMenuText);
const { createTableStyleMenu, createCellStylesMenu, createCurrencyMenu } =
  createStylesMenuFactories({ ribbonLang, ribbonMenuText, ribbonText });
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
        searchCommands: 'コマンドの検索',
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
        searchCommands: 'Search commands',
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

const xlsxIo = createXlsxIo({
  getInst: () => inst,
  setInst: (next) => {
    inst = next;
  },
  ribbonLang,
  markDirty: () => markDirty(),
  refreshWorkbookCells: () => refreshWorkbookCells(),
  shellText,
  docState,
  getRenderSheetTabs: () => renderSheetTabs,
  showPrompt,
  showMessage,
  getShowRibbonReport: () => showRibbonReport,
});
const {
  openFileMenu,
  closeFileMenu,
  triggerOpen,
  triggerSave,
  triggerSaveAs,
  loadXlsxFile,
  inspectWorkbookFromBackstage,
  setDocName,
} = xlsxIo;

let autosaveEnabled = false;
const { setShellLabel, refreshAutosave, applyShellLocale } = createShellLocale({
  autosaveSwitch,
  shellText,
  html,
  ribbonLang,
  getAutosaveEnabled: () => autosaveEnabled,
});
void setShellLabel;

let activeRibbonTab: RibbonTab = 'home';
let ribbonCollapsed = false;
let ribbonDisplayMenuOpen = false;
let backstageOpen = false;
let backstageReturnTab: RibbonTab = 'home';
let selectedBorderStyle: CellBorderStyle = 'thin';
let selectedBorderColor = '#000000';
let formulaBarVisible = true;
let filterDropdown: ReturnType<typeof attachFilterDropdown> | null = null;

const { createBackstageView } = createBackstageFactories({
  backstageText,
  ribbonText,
  shellSavedText: shellText.saved,
  docName: () => xlsxIo.getDocName(),
  docState,
});

const controlDispatch = createControlDispatch({
  getInst: () => inst,
  ribbonLang,
  ribbonText,
  pageScaleText,
  sheetEl: sheetEl as HTMLElement,
  focusSheet: () => focusSheet(),
  refreshWorkbookCells: () => refreshWorkbookCells(),
  projectFormatToolbar: () => projectFormatToolbar(),
});
const {
  createRibbonIcon,
  currentRibbonControlValue,
  applyRibbonFormat,
  applyRibbonControl,
  applyMergeControl,
} = controlDispatch;

const selectColorRibbon = createSelectColorRibbon({
  ribbonLang,
  ribbonText,
  pageScaleText,
  getInst: () => inst,
  applyRibbonControl,
  currentRibbonControlValue,
  shouldShowFontOption,
  createRibbonIcon,
});
const {
  makeSvg,
  createRibbonSelect,
  createRibbonColor,
  closeOpenRibbonDropdowns,
  updateRibbonSelectDisplay,
  ribbonSelectLabel,
  RIBBON_CHEVRON_PATH,
} = selectColorRibbon;

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

// ── Borders dropdown (Excel-365 parity) ─────────────────────────────────
// Renders a small SVG cell-outline icon for each border preset. Sides are
// drawn solid in the foreground color (thin/thick/double); the unset sides
// show as a faint dashed cell-outline base so the user can still see the
// ── Borders dropdown lives in ribbon/menus/borders.ts ──

const createPasteMenu = (): HTMLDivElement => createPasteMenuImpl(ribbonLang);

const createConditionalMenu = (): HTMLDivElement => createConditionalMenuImpl(ribbonLang);

const createBordersMenu = (): HTMLDivElement =>
  createBordersMenuImpl({
    ribbonText,
    getBorderColor: () => selectedBorderColor,
    onPickColor: (color) => {
      selectedBorderColor = color;
      inst?.borderDraw?.setColor(color);
      closeBorderSubmenus();
    },
  });

// Forward declaration so the mountToolbar closure below can resolve it at
// call-time even though `projectFormatToolbar` (the const this points at)
// is created later by statusProjection. Without this indirection the
// renderRibbon() that mountToolbar fires at the end of its setup would
// TDZ-fault on the still-unborn const.
let deferredProjectFormatToolbar: (() => void) | undefined;
// Same TDZ deferral for the dynamic-dropdowns interceptor — created later
// (line ~1276 after action handlers exist).
let deferredInterceptCommand:
  | ((id: string, button: HTMLButtonElement, event: MouseEvent) => boolean)
  | undefined;
// Hooks the toolbar dispatcher fans into for clipboard / etc. They live in
// playground modules created after mountToolbar (clipboard.ts, sortFilter,
// illustrations, …) — the thunks let those modules attach late without
// breaking TDZ.
let deferredCopyHook: (() => unknown) | undefined;
let deferredCutHook: (() => unknown) | undefined;
let deferredPasteHook: (() => unknown) | undefined;
const deferredHooksBag: {
  sortFilter?: {
    sort: (dir: 'asc' | 'desc') => void;
    customSort: () => Promise<unknown> | unknown;
    openFilter: () => void;
    removeDuplicates: () => void;
    splitTextToColumns: (sep: string) => void;
  };
  insert?: {
    createTable: (variant: 'medium') => Promise<unknown> | unknown;
    createChart: () => void;
    insertPicture: (source: 'online') => Promise<unknown> | unknown;
    insertShape: (shape: 'rectangle') => void;
    insertScreenshot: () => void;
  };
  page?: {
    pageBreak: () => void;
    sheetBackground: (action: 'set') => Promise<unknown> | unknown;
    pdf: (action: 'create') => void;
    inspect: () => void;
    outline: (action: 'group' | 'ungroup' | 'show-detail' | 'hide-detail') => void;
  };
  drawing?: { setInkMode: (mode: 'pen' | 'erase') => void };
  review?: {
    spelling: () => void;
    translate: () => void;
    accessibility: () => void;
    deleteComment: () => void;
    selectComment: (direction: 1 | -1) => void;
  };
  protection?: {
    runSheet: () => Promise<unknown> | unknown;
    runWorkbook: (protect: boolean) => Promise<unknown> | unknown;
    allowEditRanges: () => Promise<unknown> | unknown;
  };
  automation?: {
    runScript: () => Promise<unknown> | unknown;
    recordActions: () => void;
    allScripts: () => void;
    addInManager: () => void;
  };
  formula?: {
    autoSum: (fn: AutoSumFormulaName) => void;
    errorChecking: () => void;
  };
  sheetView?: {
    save: () => Promise<unknown> | unknown;
    deleteActive: () => void;
    zoomDialog: () => Promise<unknown> | unknown;
  };
} = {};

// Toolbar driven by the core `Spreadsheet.mountToolbar` API. State (active
// tab, collapsed, backstage, theme, border style/color, formula-bar) lives
// inside the toolbar instance — the `let activeRibbonTab/…` mirrors above
// stay synced via the on*Change callbacks so existing module-scope readers
// keep working. `commandDelegation: false` keeps the toolbar from
// double-firing the legacy `wireFormat('btn-…')` listeners and the
// `home-menu-wirings.ts` / `data-menu-wirings.ts` per-command click
// handlers; a later phase will collapse those into the toolbar's `hooks`.
const tb = Spreadsheet.mountToolbar(ribbonRoot as HTMLElement, () => inst, {
  lang: ribbonLang,
  text: ribbonText,
  menuText: ribbonMenuText,
  helpers: {
    createSelect: createRibbonSelect,
    createColor: createRibbonColor,
    createIcon: createRibbonIcon,
    makeSvg,
    chevronPath: RIBBON_CHEVRON_PATH,
  },
  menus: {
    paste: createPasteMenu,
    pivotTable: createPivotTableMenu,
    definedNames: createDefinedNamesMenu,
    links: createLinksMenu,
    borders: createBordersMenu,
    textOrientation: createTextOrientationMenu,
    conditional: createConditionalMenu,
    fill: createFillMenu,
    insertCells: createInsertCellsMenu,
    deleteCells: createDeleteCellsMenu,
    formatCells: createFormatCellsMenu,
    autoSum: createAutoSumMenu,
    freeze: createFreezeMenu,
    clearArrows: createClearArrowsMenu,
    errorChecking: createErrorCheckingMenu,
    watch: createWatchMenu,
    reviewComments: createReviewCommentsMenu,
    protect: createProtectMenu,
    calcOptions: createCalcOptionsMenu,
    sort: createSortMenu,
    textToColumns: createTextToColumnsMenu,
    dataValidation: createDataValidationMenu,
    findSelect: createFindSelectMenu,
    pictureInsert: createPictureInsertMenu,
    shapesInsert: createShapesInsertMenu,
    screenshotInsert: createScreenshotInsertMenu,
    chartInsert: createChartInsertMenu,
    tableStyle: createTableStyleMenu,
    cellStyles: createCellStylesMenu,
    currency: createCurrencyMenu,
    pageTheme: createPageThemeMenu,
    printArea: createPrintAreaMenu,
    pageBreaks: createPageBreaksMenu,
    sheetBackground: createSheetBackgroundMenu,
    printTitles: createPrintTitlesMenu,
    symbol: createSymbolMenu,
    script: createScriptMenu,
    addIn: createAddInMenu,
    pdf: createPdfMenu,
    clear: createClearMenu,
  },
  createBackstageView,
  activeTab: activeRibbonTab,
  collapsed: ribbonCollapsed,
  formulaBarVisible,
  theme: uiTheme,
  borderStyle: selectedBorderStyle,
  borderColor: selectedBorderColor,
  commandDelegation: true,
  interceptCommand: (id, button, event) => deferredInterceptCommand?.(id, button, event) ?? false,
  hooks: {
    clipboard: {
      copy: () => deferredCopyHook?.(),
      cut: () => deferredCutHook?.(),
      paste: () => deferredPasteHook?.(),
    },
    sortFilter: {
      sort: (dir) => deferredHooksBag.sortFilter?.sort(dir),
      customSort: () => deferredHooksBag.sortFilter?.customSort(),
      openFilter: () => deferredHooksBag.sortFilter?.openFilter(),
      removeDuplicates: () => deferredHooksBag.sortFilter?.removeDuplicates(),
      splitTextToColumns: (sep) => deferredHooksBag.sortFilter?.splitTextToColumns(sep),
    },
    insert: {
      createTable: (v) => deferredHooksBag.insert?.createTable(v),
      createChart: () => deferredHooksBag.insert?.createChart(),
      insertPicture: (s) => deferredHooksBag.insert?.insertPicture(s),
      insertShape: (s) => deferredHooksBag.insert?.insertShape(s),
      insertScreenshot: () => deferredHooksBag.insert?.insertScreenshot(),
    },
    page: {
      pageBreak: () => deferredHooksBag.page?.pageBreak(),
      sheetBackground: (a) => deferredHooksBag.page?.sheetBackground(a),
      pdf: (a) => deferredHooksBag.page?.pdf(a),
      inspect: () => deferredHooksBag.page?.inspect(),
      outline: (a) => deferredHooksBag.page?.outline(a),
    },
    drawing: { setInkMode: (m) => deferredHooksBag.drawing?.setInkMode(m) },
    review: {
      spelling: () => deferredHooksBag.review?.spelling(),
      translate: () => deferredHooksBag.review?.translate(),
      accessibility: () => deferredHooksBag.review?.accessibility(),
      deleteComment: () => deferredHooksBag.review?.deleteComment(),
      selectComment: (d) => deferredHooksBag.review?.selectComment(d),
    },
    protection: {
      runSheet: () => deferredHooksBag.protection?.runSheet(),
      runWorkbook: (p) => deferredHooksBag.protection?.runWorkbook(p),
      allowEditRanges: () => deferredHooksBag.protection?.allowEditRanges(),
    },
    automation: {
      runScript: () => deferredHooksBag.automation?.runScript(),
      recordActions: () => deferredHooksBag.automation?.recordActions(),
      allScripts: () => deferredHooksBag.automation?.allScripts(),
      addInManager: () => deferredHooksBag.automation?.addInManager(),
    },
    formula: {
      autoSum: (fn) => deferredHooksBag.formula?.autoSum(fn),
      errorChecking: () => deferredHooksBag.formula?.errorChecking(),
    },
    sheetView: {
      save: () => deferredHooksBag.sheetView?.save(),
      deleteActive: () => deferredHooksBag.sheetView?.deleteActive(),
      zoomDialog: () => deferredHooksBag.sheetView?.zoomDialog(),
    },
  },
  applyRibbonFormat,
  focusSheet: () => focusSheet(),
  refreshCells: () => refreshWorkbookCells(),
  // projectFormatToolbar is declared later — defer the call so TDZ doesn't
  // fire during the initial mount's renderRibbon(). statusProjection's
  // store.subscribe takes over once boot completes.
  projectFormatToolbar: () => deferredProjectFormatToolbar?.(),
  onTabChange: (tab) => {
    activeRibbonTab = tab;
  },
  onCollapsedChange: (next) => {
    ribbonCollapsed = next;
  },
  onBackstageOpenChange: (next) => {
    backstageOpen = next;
  },
  onThemeChange: (next) => {
    applyUiTheme(next);
  },
  onFormulaBarChange: (next) => {
    formulaBarVisible = next;
  },
});
const renderRibbon = (): void => tb.rerender();
const playgroundFeatureFlags = (): FeatureFlags => ({
  viewToolbar: false,
  watchWindow: true,
  workbookObjects: true,
  formulaBar: formulaBarVisible,
});
const legacyCommandIds = LEGACY_COMMAND_IDS;
const RIBBON_SPLIT_BUTTON_COMMANDS = SPLIT_BUTTON_COMMANDS;
void RIBBON_SPLIT_BUTTON_COMMANDS;
void tb;

const applyCellFormatAction = (action: string): Promise<void> =>
  applyCellFormatActionImpl(action, {
    inst,
    ribbonLang,
    range: normalizedSelectionRange(),
    statusMetric,
    ribbonMenuText,
    renameSheetLabel: dictionaries[ribbonLang].sheetTabs.rename,
    runSheetProtectionFlow,
    showPrompt,
    promptDimension,
    renderSheetTabs,
    switchSheet,
    refreshWorkbookCells,
    sheetTabColorByAction,
    projectFormatToolbar,
    focusSheet,
  });

const cfFill = { fill: '#ffc7ce', color: '#9c0006' } as const;
const cfTopFill = { fill: '#c6efce', color: '#006100' } as const;

const applyConditionalMenuAction = (action: string, panel?: string): Promise<void> =>
  applyConditionalMenuActionImpl(
    {
      inst,
      ribbonLang,
      range: cfSelectionRange(),
      cfFill,
      cfTopFill,
      promptCfNumber,
      promptCfText,
      showChoiceDialog,
      showMessage,
      refreshWorkbookCells,
      addConditionalRuleFromRibbon,
    },
    action,
    panel,
  );

applyShellLocale();
renderRibbon();

const statusProjection = createStatusProjection({
  getInst: () => inst,
  ribbonLang,
  statusSelection,
  statusMetric,
  statusObjects,
  legacyCommandIds,
  getFormulaBarVisible: () => formulaBarVisible,
  currentRibbonControlValue,
  ribbonSelectLabel,
});
const {
  projectStatus,
  projectFormatToolbar,
  refreshObjectsBadge,
  setRibbonCommandActive,
  markCurrentLegacyRibbonBindings,
  persistStats,
} = statusProjection;
// Late-bind the toolbar's deferred projection hook now that the function
// exists. See the `let deferredProjectFormatToolbar` declaration above the
// mountToolbar call for the initial-mount TDZ this works around.
deferredProjectFormatToolbar = projectFormatToolbar;
void persistStats;
void setRibbonCommandActive;

const { boot, openCommentDialog, wireFormat, formatPainterBtn } = createBootWiring({
  getInst: () => inst,
  setInst: (next) => {
    inst = next;
  },
  ribbonLang,
  localeParam,
  getUiTheme: () => uiTheme,
  toCore,
  sheetEl: sheetEl as HTMLElement,
  enginePill,
  statusEngine,
  docState,
  statusState,
  projectStatus,
  projectFormatToolbar,
  refreshObjectsBadge,
  markDirty: () => markDirty(),
  refreshZoom: () => refreshZoom(),
  renderSheetTabs: () => renderSheetTabs(),
  shellText,
  bootParams,
  seed,
  playgroundFeatureFlags,
  setFilterDropdown: (drop) => {
    filterDropdown = drop;
  },
  activeClass: ACTIVE_CLASS,
});
void formatPainterBtn;

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

// Format-button click handlers (bold/italic/align/font-grow/…) are dispatched
// by mountToolbar via commandDelegation against `[data-ribbon-command]`. No
// per-button wireFormat needed — see RIBBON_FORMAT_MUTATORS and the fontGrow
// /fontShrink cases in core's applyRibbonCommand.

void clearFormat; // Reserved for a "Clear formatting" menu item; keep the import live.
void wireFormat;
void toggleBold;
void toggleItalic;
void toggleUnderline;
void toggleStrike;
void cyclePercent;
void setNumFmt;
void setFont;
void setAlign;
void setVAlign;
void bumpDecimals;

// ── Borders dropdown (Excel-365 parity) ──────────────────────────────────

// ── Freeze Panes menu ─────────────────────────────────────────────────────

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

const illustrations = createIllustrations({
  getInst: () => inst,
  getSheetEl: () => sheetEl,
  getLabels: () => ribbonMenuText,
  focusSheet,
});
const { addSessionIllustration, setDrawInkMode } = illustrations;

const clipboard = createClipboard({
  getInst: () => inst,
  refreshWorkbookCells,
  focusSheet,
  ribbonLang,
});
const {
  copySelectionToClipboard,
  cutSelectionToClipboard,
  pasteClipboardIntoSelection,
  applyRibbonPasteAction,
} = clipboard;

// Late-bind clipboard hooks so the toolbar's applyRibbonCommand path can
// fan copy / cut / paste clicks into the playground's clipboard helpers.
deferredCopyHook = () => copySelectionToClipboard();
deferredCutHook = () => cutSelectionToClipboard();
deferredPasteHook = () => pasteClipboardIntoSelection();

const sortFilter = createSortFilter({
  getInst: () => inst,
  ribbonLang,
  ribbonMenuText,
  sheetEl: sheetEl as HTMLElement,
  statusMetric,
  getFilterDropdown: () => filterDropdown,
  focusSheet,
  refreshWorkbookCells,
});
const {
  openFilterForSelection,
  applyAdvancedFilterAction,
  sortSelection,
  customSortSelection,
  removeDuplicateRows,
  splitTextToColumns,
  splitTextToColumnsCustom,
} = sortFilter;

const borderMenuApi = createBorderMenu({
  getInst: () => inst,
  sheetEl: sheetEl as HTMLElement,
  getSelectedBorderStyle: () => selectedBorderStyle,
  setSelectedBorderStyle: (v) => {
    selectedBorderStyle = v;
  },
  getSelectedBorderColor: () => selectedBorderColor,
  applyRibbonFormat,
});
const {
  openBorderMenu,
  closeBorderMenu,
  closeBorderSubmenus,
  refreshBorderMenuState,
  applyBorderPresetMenuAction,
  applyBorderDrawMenuAction,
} = borderMenuApi;

const sheetTabsApi = createSheetTabs({
  getInst: () => inst,
  focusSheet,
  statusMetric,
  workbookStructureProtectedBlockedText: ribbonMenuText.workbookStructureProtectedBlocked,
});
const { renderSheetTabs, switchSheet, openTabMenu, openUnhideMenu, closeTabMenu } = sheetTabsApi;

const {
  selectedRowCount,
  selectedColCount,
  sortTargetRange,
  sortCellDisplayText,
  colFromLetters,
  parseA1Range,
  rangeRef,
  syncStoreCellsToWorkbook,
  showZoomDialogFromRibbon,
} = createRangeUtils({
  getInst: () => inst,
  ribbonLang,
  refreshZoom: () => refreshZoom(),
  focusSheet,
});
void sortCellDisplayText;
void colFromLetters;
void syncStoreCellsToWorkbook;

const { runSheetProtectionFlow, runWorkbookProtectionFlow, applyProtectAction } =
  createProtectionFlows({
    getInst: () => inst,
    ribbonLang,
    ribbonMenuText,
    shellText,
    protectionText: dictionaries[ribbonLang].protection,
    statusMetric,
    parseA1Range,
    rangeRef,
    renderSheetTabs,
    projectFormatToolbar,
    focusSheet,
  });

const workbookActions = createWorkbookActions({
  getInst: () => inst,
  ribbonLang,
  ribbonText,
  ribbonMenuText,
  refreshWorkbookCells,
  focusSheet,
  renderSheetTabs,
  switchSheet,
  applyRibbonFormat,
  sortTargetRange,
  rangeRef,
  parseA1Range,
  getStatusMetric: () => statusMetric,
});
const {
  applyCellStyleFromRibbon,
  applyCurrencyPreset,
  openCurrencyFooterAction,
  openCellStyleFooterAction,
  openTableStyleFooterAction,
  createTableFromSelection,
  applyPivotTableAction,
  applyDefinedNameAction,
  clearHyperlinksInSelection,
  applyLinksAction,
} = workbookActions;

const scriptAddinActions = createScriptAddinActions({
  getInst: () => inst,
  ribbonLang,
  ribbonText,
  ribbonMenuText,
  ribbonReportText,
  viewToolbarText: dictionaries[ribbonLang].viewToolbar,
  automationRuns,
  statusMetric,
  refreshWorkbookCells,
  projectFormatToolbar,
  focusSheet,
});
const {
  showRibbonReport,
  runAccessibilityCheck,
  runSpellingReview,
  openTranslateReview,
  runPlaygroundScript,
  applyScriptAction,
  openAllScripts,
  recordSelectedActions,
  openAddInManager,
  applyAddInAction,
  applyPdfAction,
  runFormulaErrorChecking,
  saveCurrentSheetViewFromRibbon,
  deleteActiveSheetViewFromRibbon,
} = scriptAddinActions;

const ribbonActions = createRibbonActions({
  getInst: () => inst,
  ribbonLang,
  ribbonText,
  ribbonMenuText,
  sheetEl: sheetEl as HTMLElement,
  getStatusMetric: () => statusMetric,
  refreshWorkbookCells,
  focusSheet,
  projectFormatToolbar,
  applyRibbonFormat,
  renderSheetTabs,
  switchSheet,
  selectedRowCount,
  selectedColCount,
  clearHyperlinksInSelection,
  addSessionIllustration,
  runFormulaErrorChecking,
});
const {
  selectMatchingAddresses,
  applyFindSelectAction,
  applyAutoSumFormula,
  cfSelectionRange,
  normalizedSelectionRange,
  clearSelectionContents,
  applyFillDirection,
  runFillSeries,
  applyFillSeries,
  applyClearAction,
  promptDimension,
  applyCellInsertAction,
  applyCellDeleteAction,
  sheetTabColorByAction,
  applyTextOrientationAction,
  addConditionalRuleFromRibbon,
  promptCfNumber,
  promptCfText,
  selectionToA1Range,
  applyPrintAreaAction,
  applyPageBreakAction,
  applySheetBackgroundAction,
  applyPrintTitlesAction,
  selectionOutlineAxis,
  selectionDetailOutlineAxis,
  selectedRowOutlineRange,
  selectedColOutlineRange,
  applyOutlineAction,
  selectReviewComment,
  deleteActiveReviewComment,
  applyReviewCommentAction,
  insertSymbolIntoActiveCell,
  insertCustomSymbolIntoActiveCell,
  applyDataValidationAction,
  applyFormulaAuditAction,
  applyWatchAction,
  insertPictureFromRibbon,
  insertShapeFromRibbon,
  insertScreenshotFromRibbon,
  createChartFromSelection,
  recommendedChartKind,
  chartLabel,
  createRecommendedChartFromSelection,
  chartKindFromAction,
} = ribbonActions;

// Late-bind toolbar hooks now that every action helper exists. Mirrors the
// old per-call `applyRibbonCommandImpl` deps struct but as a single object
// the toolbar can call into on every dispatch.
deferredHooksBag.sortFilter = {
  sort: (dir) => sortSelection(dir),
  customSort: () => customSortSelection(),
  openFilter: () => openFilterForSelection(),
  removeDuplicates: () => removeDuplicateRows(),
  splitTextToColumns: (sep) => splitTextToColumns(sep),
};
deferredHooksBag.insert = {
  createTable: () => createTableFromSelection(),
  createChart: () => createChartFromSelection(recommendedChartKind()),
  insertPicture: (s) => insertPictureFromRibbon(s),
  insertShape: (s) => insertShapeFromRibbon(s),
  insertScreenshot: () => insertScreenshotFromRibbon(),
};
deferredHooksBag.page = {
  pageBreak: () => applyPageBreakAction('insert-auto'),
  sheetBackground: (a) => applySheetBackgroundAction(a),
  pdf: (a) => applyPdfAction(a),
  inspect: () => inspectWorkbookFromBackstage(),
  outline: (a) => applyOutlineAction(a),
};
deferredHooksBag.drawing = { setInkMode: (m) => setDrawInkMode(m) };
deferredHooksBag.review = {
  spelling: () => runSpellingReview(),
  translate: () => openTranslateReview(),
  accessibility: () => runAccessibilityCheck(),
  deleteComment: () => deleteActiveReviewComment(),
  selectComment: (d) => selectReviewComment(d),
};
deferredHooksBag.protection = {
  runSheet: () => runSheetProtectionFlow(),
  runWorkbook: (p) => runWorkbookProtectionFlow(p),
  allowEditRanges: () => applyProtectAction('allow-edit-ranges'),
};
deferredHooksBag.automation = {
  runScript: () => runPlaygroundScript(),
  recordActions: () => recordSelectedActions(),
  allScripts: () => openAllScripts(),
  addInManager: () => openAddInManager(),
};
deferredHooksBag.formula = {
  autoSum: (fn) => applyAutoSumFormula(fn),
  errorChecking: () => runFormulaErrorChecking(),
};
deferredHooksBag.sheetView = {
  save: () => saveCurrentSheetViewFromRibbon(),
  deleteActive: () => deleteActiveSheetViewFromRibbon(),
  zoomDialog: () => showZoomDialogFromRibbon(),
};

const dataMenus = createDataMenuWirings({
  getInst: () => inst,
  ribbonLang,
  getSheetEl: () => sheetEl as HTMLElement,
  focusSheet,
  refreshWorkbookCells,
  closeBorderMenu,
  closeConditionalMenu: (rf?: boolean) => closeConditionalMenu(rf),
  closeFillMenu: (rf?: boolean) => closeFillMenu(rf),
  closeClearMenu: (rf?: boolean) => closeClearMenu(rf),
  closeCellsMenus: (rf?: HTMLElement | null) => closeCellsMenus(rf),
  closeTextOrientationMenu: (rf?: boolean) => closeTextOrientationMenu(rf),
  closePasteMenu: (rf?: boolean) => closePasteMenu(rf),
  sortSelection,
  customSortSelection,
  openFilterForSelection,
  applyAdvancedFilterAction,
  removeDuplicateRows,
  applyFindSelectAction,
  applyAutoSumFormula,
  createChartFromSelection,
  createRecommendedChartFromSelection,
  chartKindFromAction,
});
const {
  applySortMenuAction,
  applyCalcOptionAction,
  applyFreezeMenuAction,
  updateCalcOptionsMenu,
  openSortFilterHomeMenu,
  closeSortFilterHomeMenu,
  openDataSortMenu,
  closeDataSortMenu,
  openFindSelectMenu,
  closeFindSelectMenu,
  openCalcOptionsMenu,
  closeCalcOptionsMenu,
  openChartInsertMenu,
  closeChartInsertMenu,
  openFreezeMenu,
  closeFreezeMenu,
  getFreezeBtn,
  getFreezeMenu,
} = dataMenus;

const homeMenus = createHomeMenuWirings({
  getInst: () => inst,
  closeBorderMenu,
  closeFreezeMenu,
  closeFindSelectMenu,
  closeSortFilterHomeMenu,
  pasteClipboardIntoSelection,
  applyRibbonPasteAction,
  applyConditionalMenuAction,
  applyFillSeries,
  applyFillDirection,
  applyClearAction,
  applyTextOrientationAction,
  applyCellInsertAction,
  applyCellDeleteAction,
  applyCellFormatAction,
  applyPrintAreaAction,
  insertSymbolIntoActiveCell,
  insertCustomSymbolIntoActiveCell,
});
const {
  closePasteMenu,
  closeConditionalMenu,
  closeFillMenu,
  closeClearMenu,
  closePrintAreaMenu,
  closeSymbolMenu,
  closeTextOrientationMenu,
  closeCellsMenus,
  openPrintAreaMenu,
  openSymbolMenu,
  getPrintAreaMenu,
  getSymbolMenu,
} = homeMenus;

const dynamicDropdowns = createDynamicDropdowns({
  getInst: () => inst,
  updateCalcOptionsMenu,
  updateDefinedNamesMenu,
  closeBorderMenu,
  closeFreezeMenu,
  closePrintAreaMenu,
  closeSymbolMenu,
  getConditionalMenu: () => document.getElementById('menu-conditional') as HTMLDivElement | null,
  applyRibbonPasteAction,
  applyPivotTableAction,
  applyDefinedNameAction,
  applyLinksAction,
  applyFillSeries,
  applyFillDirection,
  applyClearAction,
  applyTextOrientationAction,
  applyCellInsertAction,
  applyCellDeleteAction,
  applyCellFormatAction,
  applyPageBreakAction,
  applySheetBackgroundAction,
  applyPrintTitlesAction,
  applyUiTheme,
  focusSheet,
  applySortMenuAction,
  applyFindSelectAction,
  applyAutoSumFormula,
  applyFormulaAuditAction,
  applyWatchAction,
  applyReviewCommentAction,
  applyProtectAction,
  applyCalcOptionAction,
  createRecommendedChartFromSelection,
  createChartFromSelection,
  chartKindFromAction,
  insertPictureFromRibbon,
  insertShapeFromRibbon,
  insertScreenshotFromRibbon,
  applyScriptAction,
  applyPdfAction,
  createTableFromSelection,
  openTableStyleFooterAction,
  applyCellStyleFromRibbon,
  openCellStyleFooterAction,
  applyCurrencyPreset,
  openCurrencyFooterAction,
  splitTextToColumns,
  splitTextToColumnsCustom,
  applyDataValidationAction,
  applyAddInAction,
  applyConditionalMenuAction,
});
const {
  DYNAMIC_RIBBON_DROPDOWN_IDS,
  dynamicDropdownSpecForButton,
  dynamicDropdownSpecForMenu,
  dynamicDropdownButtonForSpec,
  openDynamicRibbonDropdown,
  closeDynamicRibbonDropdown,
  closeAllDynamicRibbonDropdowns,
  closeDynamicConditionalSubmenus,
  openDynamicConditionalSubmenu,
  dynamicRibbonDropdownClick,
} = dynamicDropdowns;

// Late-bind the toolbar's interceptCommand hook so dropdown-owning buttons
// (calcOptions, dataValidation, sort, etc.) open / close their menu instead
// of falling through to applyRibbonCommand (which would otherwise open the
// dialog associated with that command id). Survives ribbon re-renders since
// the dropdown spec is looked up per-click via the current DOM.
//
// Split-button commands (paste, autosum, …) are skipped so the primary
// action still fires through applyRibbonCommand → hooks; the chevron-open
// path lives in home-menu-wirings and only works pre-rerender, which
// matches the legacy behaviour.
// Split buttons whose primary click runs an action (paste / autosum /
// currency) rather than opening the dropdown — the chevron-vs-main click
// split lives in home-menu-wirings / data-menu-wirings. The remaining split
// buttons (pageTheme / addIn / script) are menu-only: a plain click opens
// the dropdown the same as a non-split menu owner.
const PRIMARY_ACTION_SPLIT_COMMANDS = new Set<string>([
  'paste',
  'autosum',
  'autosumFormula',
  'currency',
]);
// Menus that aren't owned by dynamic-dropdowns (border / freeze / printArea
// / symbol live in border-menu / data-menu-wirings / home-menu-wirings),
// but the toolbar still needs to open them on a fresh click so the wiring
// survives ribbon re-renders.
const LEGACY_MENU_OPENERS: Record<string, () => void> = {
  printArea: () => openPrintAreaMenu(),
  symbolInsert: () => openSymbolMenu(),
  freeze: () => openFreezeMenu(),
  borders: () => openBorderMenu(),
};
deferredInterceptCommand = (id, button, event) => {
  const legacyOpener = LEGACY_MENU_OPENERS[id];
  if (legacyOpener) {
    legacyOpener();
    return true;
  }
  // Split buttons: chevron click opens dropdown, primary click runs the
  // action via applyRibbonCommand. event.target lets us tell the two apart
  // even when the legacy data-menu-wirings handler is stale (post-rerender).
  const target = event.target as Element | null;
  const isChevronClick = !!target?.closest('.demo__rb-split-chevron');
  if (PRIMARY_ACTION_SPLIT_COMMANDS.has(id) && !isChevronClick) return false;
  const spec = dynamicDropdownSpecForButton(button);
  if (!spec) return false;
  const menu = document.getElementById(spec.menuId) as HTMLDivElement | null;
  if (!menu) return false;
  if (menu.hidden) openDynamicRibbonDropdown(spec, button);
  else closeDynamicRibbonDropdown(spec, true);
  return true;
};

document.addEventListener('click', (event) => {
  dynamicRibbonDropdownClick(event);
});

document.addEventListener('mouseover', (event) => {
  const target = event.target as Element | null;
  const menu = target?.closest<HTMLElement>('#menu-conditional');
  if (!menu || menu === document.getElementById('menu-conditional')) return;
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

const backstageTitle = createBackstageTitle({
  getInst: () => inst,
  ribbonLang,
  shellText,
  ribbonRoot,
  titleSearchInput,
  autosaveSwitch,
  statusMetric,
  focusSheet: () => focusSheet(),
  triggerSave: () => triggerSave(),
  triggerSaveAs: () => triggerSaveAs(),
  renderRibbon: () => renderRibbon(),
  refreshAutosave: () => refreshAutosave(),
  projectFormatToolbar: () => projectFormatToolbar(),
  showRibbonReport: (t, items) => showRibbonReport(t, items),
  setRibbonDisplayMenuOpen: (v) => {
    ribbonDisplayMenuOpen = v;
  },
  getActiveRibbonTab: () => activeRibbonTab,
  setActiveRibbonTab: (tab) => {
    activeRibbonTab = tab;
  },
  getRibbonCollapsed: () => ribbonCollapsed,
  setRibbonCollapsed: (next) => {
    ribbonCollapsed = next;
  },
  getBackstageOpen: () => backstageOpen,
  setBackstageOpen: (next) => {
    backstageOpen = next;
  },
  getBackstageReturnTab: () => backstageReturnTab,
  setBackstageReturnTab: (tab) => {
    backstageReturnTab = tab;
  },
  getAutosaveEnabled: () => autosaveEnabled,
  setAutosaveEnabled: (next) => {
    autosaveEnabled = next;
  },
});
const {
  selectRibbonTab,
  setRibbonCollapsedExternal: setRibbonCollapsed,
  openBackstage,
  closeBackstage,
} = backstageTitle;
void selectRibbonTab;
void setRibbonCollapsed;

const { refreshZoom } = setupZoomControls(() => inst);

if (titleSearchInput) {
  const searchContainer =
    titleSearchInput.closest<HTMLElement>('.app__search') ?? titleSearchInput.parentElement;
  if (searchContainer) {
    createCommandPalette({
      input: titleSearchInput,
      container: searchContainer,
      ribbonText,
      ribbonLang,
      applyCommand: (id) => tb.applyCommand(id),
    });
  }
}

createShellMenus({
  getInst: () => inst,
  ribbonLang,
  ribbonRoot,
  sheetEl: sheetEl as HTMLElement,
  docState,
  shellText,
  openFileMenu,
  closeFileMenu,
  triggerOpen,
  triggerSave,
  triggerSaveAs,
  loadXlsxFile,
  inspectWorkbookFromBackstage,
  setDocName,
  renderSheetTabs,
  applyUiTheme,
  getUiTheme: () => uiTheme,
  applyProtectAction,
  closeBackstage,
});

markCurrentLegacyRibbonBindings();

boot().catch((err) => {
  // eslint-disable-next-line no-console
  console.error('formulon-cell boot failed', err);
  if (sheetEl) {
    sheetEl.innerHTML = `<pre style="padding:24px;color:#d24545;font-family:'IBM Plex Mono',monospace;white-space:pre-wrap">${
      err instanceof Error ? (err.stack ?? err.message) : String(err)
    }</pre>`;
  }
});
