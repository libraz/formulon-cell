<script setup lang="ts">
import {
  activateSheetView,
  addSheet,
  applyAdvancedFilter,
  applyMerge,
  applyUnmerge,
  applyConditionalPresetAction,
  applyCellStyle,
  analyzeAccessibilityCells,
  analyzeSpellingCells,
  autoSum,
  autofitColsWidth,
  autofitRowsHeight,
  applyTextScriptToRange,
  buildTranslationReviewItems,
  CELL_STYLE_GROUPS,
  CELL_STYLES,
  type CellStyleId,
  bumpDecimals,
  bumpIndent,
  type CellBorderStyle,
  type ConditionalPresetAction,
  clearPrintArea,
  clearPrintTitles,
  clearSheetBackgroundImage,
  clearTraceArrows,
  clearTraceArrowsByKind,
  clearValidationCircles,
  clearValidationInRangeWithEngine,
  clearWatchedCells,
  circleInvalidValidationDataInSheet,
  collapseColGroup,
  collapseRowGroup,
  copyAdvancedFilterResult,
  createDefinedNamesFromSelection,
  buildRibbonAddInReport,
  createColorPalette,
  createRibbonChartFromSelection,
  executeRibbonClearAction,
  executeRibbonCommentAction,
  executeRibbonFilterDataAction,
  executeRibbonFormulaAuditingAction,
  executeRibbonHyperlinkAction,
  executeRibbonPivotTableAction,
  executeRibbonProtectionAction,
  resolveRibbonPdfAction,
  deleteCells,
  deleteSheetView,
  cycleBorders,
  cycleCurrency,
  cyclePercent,
  deleteCols,
  deleteRows,
  executeRibbonFillAction,
  executeRibbonFindAction,
  formatAsTable,
  type FeatureFlags,
  groupCols,
  groupRows,
  hiddenInSelection,
  insertCells,
  hideCols,
  hideRows,
  insertCols,
  insertRows,
  insertManualPageBreak,
  listComments,
  listDefinedNames,
  makeRangeResolver,
  type MarginPreset,
  moveSheet,
  mutators,
  type NumberFormatAction,
  numberFormatForAction,
  type PageOrientation,
  type PaperSize,
  reviewCellsFromState,
  recordConditionalRulesChange,
  recordDefinedNamesChange,
  recordFilterChange,
  recordFormatChange,
  recordLayoutChange,
  recordMergesChangeWithEngine,
  recordPageSetupChange,
  recordTablesChange,
  recordValidationCirclesChange,
  recordWatchesChange,
  renameSheet,
  removeSheet,
  removeDuplicates,
  removeManualPageBreak,
  resetManualPageBreaks,
  saveSheetView,
  setAlign,
  setBorderPreset,
  setColsWidth,
  setFillColor,
  setFitToPages,
  setFreezePanes,
  setFont,
  setFontColor,
  setGridlinesVisible,
  setHeadingsVisible,
  setMarginPreset,
  setNumFmt,
  setPageOrientation,
  setPageScale,
  setPaperSize,
  setPrintArea,
  setPrintGridlines,
  setPrintHeadings,
  setPrintTitleCols,
  setPrintTitleRows,
  setRowsHeight,
  setSheetHidden,
  setR1C1ReferenceStyle,
  setRotation,
  setSheetBackgroundImage,
  setSheetZoom,
  setWorkbookStructureProtected,
  setShowFormulas,
  setWorkbookView,
  toggleAutoFilterFromSelection,
  type ScriptCommand,
  showCols,
  showColsAroundSelection,
  showRows,
  showRowsAroundSelection,
  inferSortHasHeader,
  sortActiveColumnAuto,
  sortRangeWithHistory,
  type SpreadsheetInstance,
  insertDefinedNameFormula,
  isCellWritable,
  isWorkbookStructureProtected,
  buildSpreadsheetCompatibilityReport,
  textToColumns,
  setVAlign,
  toggleBold,
  toggleItalic,
  toggleStrike,
  toggleUnderline,
  toggleWrap,
  unwatchCell,
  ungroupCols,
  ungroupRows,
  warnProtected,
  watchRange,
  type CellDeleteAction,
  type CellInsertAction,
  type ConditionalMenuAction,
  type FreezeAction,
  type MergeAction,
  type PasteAction,
  type WindowAction,
  deleteSelectedCols,
  deleteSelectedRows,
  dispatchHostClipboard,
  handleAutoSum,
  handleAutoSumAction,
  handleConditionalAction,
  handleDeleteCellsAction,
  handleFreezeAction,
  handleInsertCellsAction,
  handleMergeAction,
  handlePasteAction,
  handleWindowAction,
  insertSelectedCols,
  insertSelectedRows,
  toggleSelectedColsHidden,
  toggleSelectedRowsHidden,
  type AddInAction,
  type AdvancedFilterDialogDraft,
  type AutomationRunDraft,
  type AutoSumAction,
  type CalculationAction,
  CELL_STYLE_SECTION_ACTION_PREFIX,
  type CellFormatAction,
  type CellStyleAction,
  type ChartAction,
  type ClearAction,
  type ClearArrowsAction,
  type CommentAction,
  cellLabel,
  colLetter,
  type DataValidationAction,
  type DefinedNameAction,
  type DimensionDialogDraft,
  type FillAction,
  type FilterDataAction,
  type FindAction,
  type FormatTableAction,
  type FormulaAuditingAction,
  type FunctionAction,
  formatA1Range,
  type HyperlinkAction,
  MORE_SYMBOL_ACTION,
  type OutlineAxisAction,
  type PageBreakAction,
  type PdfAction,
  type PictureAction,
  type PivotTableAction,
  type PrintAreaAction,
  type PrintTitleAction,
  type ProtectionAction,
  parseA1Range,
  type RemoveDuplicatesDialogDraft,
  type RibbonReportDialogDraft,
  type ConditionalIconSetAction,
  conditionalColorScaleLabel,
  conditionalColorScaleSwatchColors,
  conditionalDataBarLabel,
  conditionalDataBarSwatchColor,
  conditionalIconSetLabel,
  type ScreenshotAction,
  type ScriptDialogDraft,
  SHEET_TAB_COLOR_ACTIONS,
  type ShapeAction,
  type SheetBackgroundAction,
  type SheetCell,
  type SheetRange,
  type SheetRenameDialogDraft,
  type SortAction,
  type SortDialogDraft,
  type SymbolAction,
  TEXT_TO_COLUMNS_DIALOG_KEYS,
  type TextOrientationAction,
  type TextToColumnsAction,
  type TextToColumnsDialogDraft,
  type ThemeAction,
  type WatchAction,
} from '@libraz/formulon-cell';
import { RibbonIcon } from './toolbar/icons.js';
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue';
import { useToolbarActive } from './toolbar/active.js';
import { useToolbarDropdown } from './toolbar/dropdown.js';
import {
  type BorderPreset,
  FONT_SIZES,
  localizeBorderPresets,
  localizeBorderStyles,
  projectActiveState,
  RIBBON_KEYSHORTCUTS,
  type RibbonTab,
} from './toolbar/model.js';
import { useI18n } from './composables.js';
import { toolbarTabs } from './toolbar/tabs.js';
import { dictionaries, dictionaryLocaleFor } from './toolbar/translations.js';

interface Props {
  instance: SpreadsheetInstance | null;
  features?: FeatureFlags;
  activeTab: RibbonTab;
  locale: string;
  onSpellingReview?: () => void;
  onAccessibilityCheck?: () => void;
  onRunScript?: () => void;
  onDrawPen?: () => void;
  onDrawEraser?: () => void;
  onTranslate?: () => void;
  onAddIn?: () => void;
  onNewWorkbook?: () => void;
  onOpenWorkbook?: () => void;
  onSaveWorkbook?: () => void;
  onSaveWorkbookAs?: () => void;
}


const props = defineProps<Props>();
const emit = defineEmits<{
  tabChange: [tab: RibbonTab];
}>();
const sortDialog = ref<SortDialogDraft | null>(null);
const removeDuplicatesDialog = ref<RemoveDuplicatesDialogDraft | null>(null);
const dimensionDialog = ref<DimensionDialogDraft | null>(null);
const sheetRenameDialog = ref<SheetRenameDialogDraft | null>(null);
const scriptDialog = ref<ScriptDialogDraft | null>(null);
const automationRunCount = ref(0);
const lastAutomationRun = ref<AutomationRunDraft | null>(null);
const advancedFilterDialog = ref<AdvancedFilterDialogDraft | null>(null);
const zoomDialog = ref<string | null>(null);
const ribbonReportDialog = ref<RibbonReportDialogDraft | null>(null);
const textToColumnsDialog = ref<TextToColumnsDialogDraft | null>(null);

const instanceRef = computed(() => props.instance);
const liveI18n = useI18n(instanceRef);
const lang = computed(() => dictionaryLocaleFor(props.locale));
const liveLang = computed(() => dictionaryLocaleFor(liveI18n.locale.value));
const strings = computed(() =>
  (liveI18n.locale.value === props.locale || liveLang.value === lang.value) &&
  'ribbon' in liveI18n.strings.value
    ? liveI18n.strings.value
    : dictionaries[lang.value],
);
const tabs = computed(() => toolbarTabs(strings.value));
const tr = computed(() => strings.value.ribbon);
const cfText = computed(() => strings.value.conditionalMenu);
const cellText = computed(() => strings.value.ribbonMenu);
const sheetTabsText = computed(() => strings.value.sheetTabs);
const sheetTabColorLabel = (action: CellFormatAction): string => {
  const sheetTabs = sheetTabsText.value;
  switch (action) {
    case 'tabColorRed':
      return sheetTabs.tabColorRed;
    case 'tabColorOrange':
      return sheetTabs.tabColorOrange;
    case 'tabColorYellow':
      return sheetTabs.tabColorYellow;
    case 'tabColorGreen':
      return sheetTabs.tabColorGreen;
    case 'tabColorBlue':
      return sheetTabs.tabColorBlue;
    case 'tabColorPurple':
      return sheetTabs.tabColorPurple;
    case 'tabColorGray':
      return sheetTabs.tabColorGray;
    default:
      return sheetTabs.tabColor;
  }
};
const pageScaleText = computed(() => strings.value.pageScale);
const viewText = computed(() => strings.value.viewToggle);
const viewToolbarText = computed(() => strings.value.viewToolbar);
const insertSymbols = [
  '±',
  '×',
  '÷',
  '≤',
  '≥',
  '≠',
  '≈',
  '∞',
  '√',
  '∑',
  '∫',
  'π',
  'Α',
  'Β',
  'Γ',
  'Δ',
  'Θ',
  'Λ',
  'Ξ',
  'Π',
  'Σ',
  'Φ',
  'Ψ',
  'Ω',
  '$',
  '€',
  '¥',
  '£',
  '¢',
  '₩',
  '₹',
  '₽',
  '©',
  '®',
  '™',
  '§',
  '¶',
  '†',
  '‡',
  '•',
] as const;

const cfDataBarColor = (action: string): string => conditionalDataBarSwatchColor(action);
const cfScaleColors = (action: string): readonly string[] =>
  conditionalColorScaleSwatchColors(action);

const cfDataBarLabel = (action: ConditionalPresetAction | string): string =>
  conditionalDataBarLabel(action, cfText.value);

const cfScaleLabel = (action: ConditionalPresetAction | string): string =>
  conditionalColorScaleLabel(action, cfText.value);

const cfIconSetLabel = (action: ConditionalIconSetAction | string): string =>
  conditionalIconSetLabel(action, cfText.value);

const cfIconSetGroups = computed(() => [
  {
    title: cfText.value.direction,
    items: [
      { action: 'icons-arrows3', family: 'arrow', slots: ['up-green', 'right-yellow', 'down-red'] },
      {
        action: 'icons-arrows5',
        family: 'arrow',
        slots: ['up-green', 'up-right-gray', 'right-gray', 'down-right-gray', 'down-gray'],
      },
      {
        action: 'icons-triangles3',
        family: 'triangle',
        slots: ['up-green', 'flat-yellow', 'down-red'],
      },
    ],
  },
  {
    title: cfText.value.shapes,
    items: [
      { action: 'icons-traffic3', family: 'circle', slots: ['green', 'yellow', 'red'] },
      { action: 'icons-trafficRim3', family: 'rim', slots: ['green', 'yellow', 'red'] },
      { action: 'icons-symbols3', family: 'symbol', slots: ['check-green', 'bang-yellow', 'x-red'] },
      { action: 'icons-flags3', family: 'flag', slots: ['green', 'yellow', 'red'] },
    ],
  },
  {
    title: cfText.value.ratings,
    items: [
      { action: 'icons-stars3', family: 'star', slots: ['gold', 'half', 'empty'] },
      { action: 'icons-quarters5', family: 'quarter', slots: ['q4', 'q3', 'q2', 'q1', 'q0'] },
      { action: 'icons-ratings5', family: 'rating', slots: ['r4', 'r3', 'r2', 'r1', 'r0'] },
      { action: 'icons-bars5', family: 'bars', slots: ['b4', 'b3', 'b2', 'b1', 'b0'] },
      { action: 'icons-boxes5', family: 'boxes', slots: ['b4', 'b3', 'b2', 'b1', 'b0'] },
    ],
  },
] satisfies {
  title: string;
  items: { action: ConditionalIconSetAction; family: string; slots: string[] }[];
}[]);

const cellStyleLabel = (id: CellStyleId): string => {
  const labels = cellText.value;
  switch (id) {
    case 'normal':
      return labels.cellStyleNormal;
    case 'title':
      return labels.cellStyleTitle;
    case 'heading1':
      return labels.cellStyleHeading1;
    case 'heading2':
      return labels.cellStyleHeading2;
    case 'heading3':
      return labels.cellStyleHeading3;
    case 'heading4':
      return labels.cellStyleHeading4;
    case 'good':
      return labels.cellStyleGood;
    case 'bad':
      return labels.cellStyleBad;
    case 'neutral':
      return labels.cellStyleNeutral;
    case 'note':
      return labels.cellStyleNote;
    case 'warning':
      return labels.cellStyleWarning;
    case 'inputCell':
      return labels.cellStyleInputCell;
    case 'outputCell':
      return labels.cellStyleOutputCell;
    case 'calculation':
      return labels.cellStyleCalculation;
    case 'linkedCell':
      return labels.cellStyleLinkedCell;
    case 'totalCell':
      return labels.cellStyleTotalCell;
    case 'currency':
      return labels.cellStyleCurrency;
    case 'currency0':
      return labels.cellStyleCurrency0;
    case 'percent':
      return labels.cellStylePercent;
    case 'comma':
      return labels.cellStyleComma;
    case 'comma0':
      return labels.cellStyleComma0;
    default:
      return id;
  }
};
const cellStyleGroups = computed(() =>
  CELL_STYLE_GROUPS.map((group) => ({
    ...group,
    label: strings.value.cellStylesGallery.groups[group.id],
  })),
);
const definedNameEntries = computed(() =>
  props.instance ? listDefinedNames(props.instance.workbook) : [],
);
const tablistRef = ref<HTMLDivElement | null>(null);
const ribbonDisplayRef = ref<HTMLDivElement | null>(null);
const sheetBackgroundInput = ref<HTMLInputElement | null>(null);
const ribbonCollapsed = ref(false);
const ribbonDisplayMenuOpen = ref(false);
const formulaBarVisible = ref(props.features?.formulaBar !== false);
const previousNonFileTab = ref<RibbonTab>('home');
const fileLabel = computed(() => strings.value.ribbon.tabs.file);
const ribbonDisplayCopy = computed(() => strings.value.ribbonDisplay);
const ribbonDisplayLabel = computed(() => ribbonDisplayCopy.value.label);
const ribbonDisplayOptions = computed(() =>
  [
    { id: 'expanded', label: ribbonDisplayCopy.value.expanded },
    { id: 'collapsed', label: ribbonDisplayCopy.value.collapsed },
  ],
);
const backstageText = computed(() => strings.value.backstage);
const workbookStructureProtected = computed(() => {
  // Depend on active so store-only protection toggles trigger a template refresh.
  active.value;
  const inst = props.instance;
  return !!inst && isWorkbookStructureProtected(inst.store.getState());
});
const keyShortcuts = (id: string): string | undefined => RIBBON_KEYSHORTCUTS[id];
const borderPresets = computed(() =>
  localizeBorderPresets(tr.value),
);
const borderStyles = computed(() =>
  localizeBorderStyles(tr.value),
);
const sortColumnOptions = computed(() => {
  const inst = props.instance;
  if (!inst) return [];
  const state = inst.store.getState();
  const range = state.selection.range;
  const options: { value: number; label: string }[] = [];
  for (let col = range.c0; col <= range.c1; col += 1) {
    const header = cellLabel(
      state.data.cells.get(`${state.data.sheetIndex}:${range.r0}:${col}`) as SheetCell | undefined,
    );
    options.push({
      value: col,
      label: sortDialog.value?.hasHeader && header ? header : colLetter(col),
    });
  }
  return options;
});
const removeDuplicateColumnOptions = computed(() => {
  const inst = props.instance;
  if (!inst) return [];
  const state = inst.store.getState();
  const range = state.selection.range;
  const options: { value: number; label: string }[] = [];
  for (let col = range.c0; col <= range.c1; col += 1) {
    const header = cellLabel(
      state.data.cells.get(`${state.data.sheetIndex}:${range.r0}:${col}`) as SheetCell | undefined,
    );
    options.push({
      value: col,
      label: removeDuplicatesDialog.value?.hasHeader && header ? header : colLetter(col),
    });
  }
  return options;
});
const scriptOptions = computed<{ value: ScriptCommand; label: string }[]>(() => [
  { value: 'uppercase', label: cellText.value.scriptCommandUppercase },
  { value: 'lowercase', label: cellText.value.scriptCommandLowercase },
  { value: 'trim', label: cellText.value.scriptCommandTrim },
  { value: 'clear', label: cellText.value.scriptCommandClear },
]);

const active = useToolbarActive(() => props.instance);
const currentTheme = computed(() => (props.instance?.store.getState().ui.theme ?? 'paper') as ThemeAction);
const sheetViewOptions = computed(() => {
  const state = props.instance?.store.getState();
  const current = { value: 'current', label: viewToolbarText.value.currentView };
  if (!state) return [current];
  return [
    current,
    ...state.sheetViews.views
      .filter((view) => view.sheet === state.data.sheetIndex)
      .map((view) => ({ value: view.id, label: view.name })),
  ];
});
const activeSheetViewId = computed(
  () => props.instance?.store.getState().sheetViews.activeViewId ?? 'current',
);
const activeCalcAction = computed<CalculationAction | null>(() => {
  if (active.value.calcMode == null) return null;
  if (active.value.calcMode === 0) return 'auto';
  if (active.value.calcMode === 1) return 'manual';
  return 'autoNoTable';
});
const activeFreezeAction = computed<FreezeAction>(() => {
  const layout = props.instance?.store.getState().layout;
  if (!layout || (layout.freezeRows === 0 && layout.freezeCols === 0)) return 'none';
  if (layout.freezeRows === 1 && layout.freezeCols === 0) return 'topRow';
  if (layout.freezeRows === 0 && layout.freezeCols === 1) return 'firstColumn';
  return 'panes';
});

const disabled = computed(() => !props.instance);
const showBuiltInReview = (
  title: string,
  items: ReturnType<typeof analyzeSpellingCells>,
): void => {
  ribbonReportDialog.value = { title, items };
};
const onSpellingReview = (): void => {
  if (props.onSpellingReview) {
    props.onSpellingReview();
    return;
  }
  const inst = props.instance;
  if (!inst) return;
  showBuiltInReview(
    tr.value.spelling,
    analyzeSpellingCells(reviewCellsFromState(inst.store.getState()), lang.value),
  );
};
const onAccessibilityReview = (): void => {
  if (props.onAccessibilityCheck) {
    props.onAccessibilityCheck();
    return;
  }
  const inst = props.instance;
  if (!inst) return;
  showBuiltInReview(
    tr.value.accessibility,
    analyzeAccessibilityCells(reviewCellsFromState(inst.store.getState()), lang.value),
  );
};
const onTranslateReview = (): void => {
  if (props.onTranslate) {
    props.onTranslate();
    return;
  }
  const inst = props.instance;
  if (!inst) return;
  const state = inst.store.getState();
  showBuiltInReview(
    tr.value.translate,
    buildTranslationReviewItems(
      reviewCellsFromState(state, state.data.sheetIndex, state.selection.range),
      lang.value,
    ),
  );
};
const onRunScript = (): void => {
  if (props.onRunScript) {
    props.onRunScript();
    return;
  }
  const inst = props.instance;
  if (!inst) return;
  scriptDialog.value = { command: 'uppercase' };
};

const applyScriptDialog = (): void => {
  const inst = props.instance;
  const draft = scriptDialog.value;
  if (!inst || !draft) return;
  const state = inst.store.getState();
  const range = state.selection.range;
  inst.history.begin();
  try {
    const changed = applyTextScriptToRange(
      state,
      inst.workbook,
      range,
      draft.command,
    );
    if (changed > 0) mutators.replaceCells(inst.store, inst.workbook.cells(state.data.sheetIndex));
    automationRunCount.value += 1;
    lastAutomationRun.value = { command: draft.command, range: formatA1Range(range), changed };
  } finally {
    inst.history.end();
  }
  scriptDialog.value = null;
};
const onRecordActions = (): void => {
  if (!props.instance) return;
  const recordedDetail = lastAutomationRun.value
    ? cellText.value.automationRunDetail
        .replace('{command}', automationCommandLabel(lastAutomationRun.value.command))
        .replace('{range}', lastAutomationRun.value.range)
        .replace('{count}', String(lastAutomationRun.value.changed))
    : cellText.value.recordActionsEmpty;
  ribbonReportDialog.value = {
    title: tr.value.recordActions,
    items: [
      {
        severity: 'info',
        label: cellText.value.recordActionsStatus,
        detail: recordedDetail,
      },
    ],
  };
};
const automationCommandLabel = (command: ScriptCommand): string => {
  switch (command) {
    case 'uppercase':
      return cellText.value.scriptCommandUppercase;
    case 'lowercase':
      return cellText.value.scriptCommandLowercase;
    case 'trim':
      return cellText.value.scriptCommandTrim;
    case 'clear':
      return cellText.value.scriptCommandClear;
  }
};
const onAllScripts = (): void => {
  const runStatus =
    automationRunCount.value > 0
      ? cellText.value.automationRunStatus.replace('{count}', String(automationRunCount.value))
      : cellText.value.automationNoRuns;
  const runDetail = lastAutomationRun.value
    ? cellText.value.automationRunDetail
        .replace('{command}', automationCommandLabel(lastAutomationRun.value.command))
        .replace('{range}', lastAutomationRun.value.range)
        .replace('{count}', String(lastAutomationRun.value.changed))
    : null;
  ribbonReportDialog.value = {
    title: cellText.value.automationScriptsTitle,
    items: [
      {
        severity: 'info',
        label: cellText.value.automationBuiltInScriptsLabel,
        detail: cellText.value.automationBuiltInScriptsDetail,
      },
      {
        severity: 'info',
        label: cellText.value.automationRecentRunsLabel,
        detail: runDetail ? `${runStatus}\n${runDetail}` : runStatus,
      },
    ],
  };
};
const onAddInAction = (action: 'get' | 'my' | 'manage'): void => {
  const report = buildRibbonAddInReport(action, {
    cellMenu: cellText.value,
    addInDefaultTitle: tr.value.addIn,
  });
  if (report) ribbonReportDialog.value = report;
};
const onPdfAction = (action: 'create' | 'share' | 'preferences'): void => {
  const inst = props.instance;
  if (!inst) return;
  const result = resolveRibbonPdfAction(action, {
    cellMenu: cellText.value,
    pdfTitle: tr.value.pdf,
  });
  if (result.kind === 'open-page-setup') {
    inst.openPageSetup();
    return;
  }
  inst.print('pdf');
  if (result.report) ribbonReportDialog.value = result.report;
};
const onPivotTableAction = (action: PivotTableAction): void => {
  const inst = props.instance;
  if (!inst) return;
  const result = executeRibbonPivotTableAction({
    store: inst.store,
    workbook: inst.workbook,
    history: inst.history,
    action,
    strings: {
      pivotTable: tr.value.pivotTable,
      pivotTableNewSheet: cellText.value.pivotTableNewSheet,
      recommendedPivotTables: cellText.value.recommendedPivotTables,
      pivotAuthoringDetail: strings.value.workbookObjects.compatibilityDetails.pivotAuthoring,
      workbookStructureProtectedBlocked: cellText.value.workbookStructureProtectedBlocked,
    },
  });
  if (result.kind === 'open-dialog') inst.openPivotTableDialog();
  else if (result.kind === 'report') ribbonReportDialog.value = result.report;
  else active.value = projectActiveState(inst);
};
const onIllustrationAction = (label: string): void => {
  if (!props.instance) return;
  ribbonReportDialog.value = {
    title: tr.value.illustrations,
    items: [
      {
        severity: 'info',
        label,
        detail: strings.value.workbookObjects.compatibilityDetails.chartsDrawings,
      },
    ],
  };
};
const onBackstageProtectWorkbook = (): void => {
  const inst = props.instance;
  if (!inst) return;
  setWorkbookStructureProtected(inst.store, !isWorkbookStructureProtected(inst.store.getState()));
  active.value = projectActiveState(inst);
};
const onBackstageInspectWorkbook = (): void => {
  const inst = props.instance;
  if (!inst) return;
  ribbonReportDialog.value = {
    title: backstageText.value.inspect,
    items: buildSpreadsheetCompatibilityReport(inst.workbook, strings.value.workbookObjects),
  };
};
const onDrawPen = (): void => {
  if (props.onDrawPen) {
    props.onDrawPen();
    return;
  }
  props.instance?.borderDraw?.activate('draw', borderStyle.value, borderColor.value);
};
const onDrawGrid = (): void => {
  props.instance?.borderDraw?.activate('grid', borderStyle.value, borderColor.value);
};
const onDrawEraser = (): void => {
  if (props.onDrawEraser) {
    props.onDrawEraser();
    return;
  }
  props.instance?.borderDraw?.activate('erase');
};
const setActiveTab = (tab: RibbonTab): void => {
  ribbonDisplayMenuOpen.value = false;
  if (tab !== 'file') previousNonFileTab.value = tab;
  emit('tabChange', tab);
};

const focusRibbonTab = async (tab: RibbonTab): Promise<void> => {
  await nextTick();
  tablistRef.value
    ?.querySelector<HTMLButtonElement>(`[data-ribbon-tab="${tab}"]`)
    ?.focus({ preventScroll: true });
};

const onGlobalKeydown = (event: KeyboardEvent): void => {
  if (event.key === 'Escape' && props.activeTab === 'file') {
    event.preventDefault();
    emit('tabChange', previousNonFileTab.value);
    void focusRibbonTab(previousNonFileTab.value);
    return;
  }
  if (event.key === 'Escape' && ribbonDisplayMenuOpen.value) {
    event.preventDefault();
    ribbonDisplayMenuOpen.value = false;
    return;
  }
  if (event.key !== 'F1' || (!event.ctrlKey && !event.metaKey)) return;
  event.preventDefault();
  ribbonDisplayMenuOpen.value = false;
  ribbonCollapsed.value = !ribbonCollapsed.value;
};

const onGlobalPointerdown = (event: PointerEvent): void => {
  if (!ribbonDisplayMenuOpen.value) return;
  const target = event.target as Node | null;
  if (target && ribbonDisplayRef.value?.contains(target)) return;
  ribbonDisplayMenuOpen.value = false;
};

onMounted(() => {
  window.addEventListener('keydown', onGlobalKeydown);
  document.addEventListener('pointerdown', onGlobalPointerdown, true);
});
onBeforeUnmount(() => {
  window.removeEventListener('keydown', onGlobalKeydown);
  document.removeEventListener('pointerdown', onGlobalPointerdown, true);
});

const onRibbonTabKeydown = (event: KeyboardEvent): void => {
  const target = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-ribbon-tab]');
  if (!target) return;
  const list = tabs.value;
  const currentId = (target.dataset.ribbonTab as RibbonTab | undefined) ?? props.activeTab;
  const current = Math.max(
    0,
    list.findIndex((tab) => tab.id === currentId),
  );
  let next = current;
  if (event.key === 'ArrowRight') next = (current + 1) % list.length;
  else if (event.key === 'ArrowLeft') next = (current - 1 + list.length) % list.length;
  else if (event.key === 'Home') next = 0;
  else if (event.key === 'End') next = list.length - 1;
  else return;
  event.preventDefault();
  const nextTab = list[next]?.id;
  if (!nextTab) return;
  setActiveTab(nextTab);
  void focusRibbonTab(nextTab);
};

const focusRibbonDisplayOption = async (index: number): Promise<void> => {
  await nextTick();
  const options = Array.from(
    ribbonDisplayRef.value?.querySelectorAll<HTMLButtonElement>('.demo__ribbon-display-option') ??
      [],
  );
  options[(index + options.length) % options.length]?.focus({ preventScroll: true });
};

const onRibbonDisplayKeydown = (event: KeyboardEvent): void => {
  const options = Array.from(
    ribbonDisplayRef.value?.querySelectorAll<HTMLButtonElement>('.demo__ribbon-display-option') ??
      [],
  );
  const activeIndex = Math.max(0, options.indexOf(document.activeElement as HTMLButtonElement));
  if (event.key === 'ArrowDown') {
    event.preventDefault();
    if (!ribbonDisplayMenuOpen.value) {
      ribbonDisplayMenuOpen.value = true;
      void focusRibbonDisplayOption(0);
      return;
    }
    void focusRibbonDisplayOption(activeIndex + 1);
  } else if (event.key === 'ArrowUp') {
    event.preventDefault();
    if (!ribbonDisplayMenuOpen.value) {
      ribbonDisplayMenuOpen.value = true;
      void focusRibbonDisplayOption(-1);
      return;
    }
    void focusRibbonDisplayOption(activeIndex - 1);
  } else if (event.key === 'Home' && options.length) {
    event.preventDefault();
    void focusRibbonDisplayOption(0);
  } else if (event.key === 'End' && options.length) {
    event.preventDefault();
    void focusRibbonDisplayOption(options.length - 1);
  }
};

const wrapFormat = (
  fn: (
    state: ReturnType<SpreadsheetInstance['store']['getState']>,
    store: SpreadsheetInstance['store'],
  ) => void,
): void => {
  const inst = props.instance;
  if (!inst) return;
  recordFormatChange(inst.history, inst.store, () => fn(inst.store.getState(), inst.store));
};

const onUndo = (): void => {
  props.instance?.undo();
};
const onRedo = (): void => {
  props.instance?.redo();
};

const onPasteAction = (action: PasteAction): void => {
  handlePasteAction(props.instance, action);
};

const onFormatPainter = (): void => {
  props.instance?.formatPainter?.activate(false);
};

const onAutoSum = (functionName: AutoSumAction = 'SUM'): void => {
  handleAutoSumAction(props.instance, functionName);
  closeDropdown();
};

const onFunctionAction = (action: FunctionAction): void => {
  props.instance?.openFunctionArguments(action);
  closeDropdown();
};

const onMergeAction = (action: MergeAction): void => {
  handleMergeAction(props.instance, action);
};

const onConditionalAction = (action: ConditionalMenuAction): void => {
  handleConditionalAction(props.instance, action);
  closeDropdown();
};

const onInsertCellsAction = (action: CellInsertAction): void => {
  handleInsertCellsAction(props.instance, action);
  closeDropdown();
};

const onDeleteCellsAction = (action: CellDeleteAction): void => {
  handleDeleteCellsAction(props.instance, action);
  closeDropdown();
};

const onCellFormatAction = (action: CellFormatAction): void => {
  const inst = props.instance;
  if (!inst) return;
  const state = inst.store.getState();
  const r = state.selection.range;
  if (action === 'dialog') inst.openFormatDialog();
  else if (action === 'rowHeight') {
    const current = state.layout.rowHeights.get(r.r0) ?? state.layout.defaultRowHeight;
    dimensionDialog.value = { kind: 'rowHeight', value: String(current) };
  } else if (action === 'colWidth') {
    const current = state.layout.colWidths.get(r.c0) ?? state.layout.defaultColWidth;
    dimensionDialog.value = { kind: 'colWidth', value: String(current) };
  } else if (action === 'autoFitRowHeight') {
    autofitRowsHeight(inst.store, inst.history, r.r0, r.r1, inst.workbook);
  } else if (action === 'autoFitColWidth') {
    autofitColsWidth(inst.store, inst.history, r.c0, r.c1, inst.workbook);
  } else if (action === 'protectSheet') inst.toggleSheetProtection();
  else if (action === 'hideRows') hideRows(inst.store, inst.history, r.r0, r.r1, inst.workbook);
  else if (action === 'showRows')
    showRowsAroundSelection(inst.store, inst.history, r.r0, r.r1, inst.workbook);
  else if (action === 'hideCols') hideCols(inst.store, inst.history, r.c0, r.c1, inst.workbook);
  else if (action === 'showCols')
    showColsAroundSelection(inst.store, inst.history, r.c0, r.c1, inst.workbook);
  else if (action === 'renameSheet')
    sheetRenameDialog.value = { value: inst.workbook.sheetName(state.data.sheetIndex) };
  else if (action === 'hideSheet')
    setSheetHidden(inst.store, inst.workbook, inst.history, state.data.sheetIndex, true);
  else if (action === 'unhideSheet') {
    const firstHidden = [...state.layout.hiddenSheets].sort((a, b) => a - b)[0];
    if (firstHidden != null) setSheetHidden(inst.store, inst.workbook, inst.history, firstHidden, false);
  } else if (action === 'moveSheetLeft') {
    const sheet = state.data.sheetIndex;
    if (sheet > 0) moveSheet(inst.store, inst.workbook, sheet, sheet - 1, inst.history);
  } else if (action === 'moveSheetRight') {
    const sheet = state.data.sheetIndex;
    if (sheet < inst.workbook.sheetCount - 1)
      moveSheet(inst.store, inst.workbook, sheet, sheet + 1, inst.history);
  } else if (action === 'tabColorNone') {
    recordLayoutChange(inst.history, inst.store, () => {
      mutators.setSheetTabColor(inst.store, state.data.sheetIndex, null);
    });
  } else if (action.startsWith('tabColor')) {
    const entry = SHEET_TAB_COLOR_ACTIONS.find((item) => item.action === action);
    if (entry) {
      recordLayoutChange(inst.history, inst.store, () => {
        mutators.setSheetTabColor(inst.store, state.data.sheetIndex, entry.color);
      });
    }
  }
  closeDropdown();
};

const applyDimensionDialog = (): void => {
  const inst = props.instance;
  const draft = dimensionDialog.value;
  if (!inst || !draft) return;
  const px = Number.parseFloat(draft.value);
  if (!Number.isFinite(px) || px <= 0) return;
  const r = inst.store.getState().selection.range;
  if (draft.kind === 'rowHeight')
    setRowsHeight(inst.store, inst.history, r.r0, r.r1, px, inst.workbook);
  else setColsWidth(inst.store, inst.history, r.c0, r.c1, px, inst.workbook);
  dimensionDialog.value = null;
};

const applySheetRenameDialog = (): void => {
  const inst = props.instance;
  const draft = sheetRenameDialog.value;
  if (!inst || !draft) return;
  const name = draft.value.trim();
  if (!name) return;
  renameSheet(inst.workbook, inst.store.getState().data.sheetIndex, name, inst.store, inst.history);
  sheetRenameDialog.value = null;
};

const onCellStyleAction = (action: CellStyleAction): void => {
  const inst = props.instance;
  if (!inst) return;
  const r = inst.store.getState().selection.range;
  applyCellStyle(inst.store, inst.history, r, action);
  closeDropdown();
};

const onFreezeAction = (action: FreezeAction): void => {
  handleFreezeAction(props.instance, action);
  closeDropdown();
};

const onAlign = (kind: 'left' | 'center' | 'right'): void => {
  wrapFormat((s, st) => setAlign(s, st, kind));
};
const onBumpDecimals = (delta: 1 | -1): void => {
  wrapFormat((s, st) => bumpDecimals(s, st, delta));
};
const onFontFamily = (value: string): void => {
  wrapFormat((s, st) => setFont(s, st, { fontFamily: value }));
};
const onFontSize = (value: string | number): void => {
  wrapFormat((s, st) => setFont(s, st, { fontSize: Number(value) }));
};
const borderColor = ref('#000000');
const onBorderPreset = (preset: BorderPreset): void => {
  wrapFormat((s, st) => setBorderPreset(s, st, preset, borderStyle.value, borderColor.value));
};
const onBorderStyle = (next: CellBorderStyle): void => {
  props.instance?.borderDraw?.setStyle(next);
};
const onPageOrientation = (next: PageOrientation): void => {
  const inst = props.instance;
  if (!inst) return;
  const sheet = inst.store.getState().data.sheetIndex;
  recordPageSetupChange(inst.history, inst.store, () => setPageOrientation(inst.store, sheet, next));
};
const onPaperSize = (next: PaperSize): void => {
  const inst = props.instance;
  if (!inst) return;
  const sheet = inst.store.getState().data.sheetIndex;
  recordPageSetupChange(inst.history, inst.store, () => setPaperSize(inst.store, sheet, next));
};
const onMarginPreset = (next: MarginPreset): void => {
  const inst = props.instance;
  if (!inst) return;
  const sheet = inst.store.getState().data.sheetIndex;
  recordPageSetupChange(inst.history, inst.store, () => setMarginPreset(inst.store, sheet, next));
};
const onScaleFit = (axis: 'width' | 'height', pages: string | number): void => {
  const inst = props.instance;
  if (!inst) return;
  const sheet = inst.store.getState().data.sheetIndex;
  const n = Number.parseInt(String(pages), 10);
  setFitToPages(inst.store, sheet, axis, n > 0 ? n : undefined, inst.history);
  closeDropdown();
};
const onScalePercent = (percent: string | number): void => {
  const inst = props.instance;
  if (!inst) return;
  const sheet = inst.store.getState().data.sheetIndex;
  const scale = Number.parseInt(String(percent), 10) / 100;
  setPageScale(inst.store, sheet, Number.isFinite(scale) ? scale : 1, inst.history);
  closeDropdown();
};
const onNumberFormat = (next: string): void => {
  const inst = props.instance;
  if (!inst) return;
  const action = next as NumberFormatAction;
  if (action === 'more') {
    inst.openFormatDialog('number');
    return;
  }
  const fmt = numberFormatForAction(action, lang.value);
  if (!fmt) return;
  wrapFormat((s, st) => setNumFmt(s, st, fmt));
};

const { borderStyle, closeDropdown, onDropdownKeydown, onDropdownPick, openDropdown, toggleDropdown } =
  useToolbarDropdown({
    onBorderPreset,
    onBorderStyle,
    onFontFamily,
    onFontSize,
    onMarginPreset,
    onNumberFormat,
    onOpenPageSetup: () => props.instance?.openPageSetup(),
    onPageOrientation,
    onPaperSize,
  });
const onFontColor = (value: string): void => {
  wrapFormat((s, st) => setFontColor(s, st, value));
};
const onFillColor = (value: string): void => {
  wrapFormat((s, st) => setFillColor(s, st, value));
};
const onPaletteColor = (kind: 'fontColor' | 'fillColor' | 'borderColor', value: string): void => {
  if (kind === 'fontColor') onFontColor(value);
  else if (kind === 'fillColor') onFillColor(value);
  else {
    borderColor.value = value;
    props.instance?.borderDraw?.setColor(value);
  }
  closeDropdown();
};

// Color flyouts host the shared vanilla color palette. The watcher mounts it
// when a flyout opens and tears it down on close.
const fontColorHost = ref<HTMLDivElement | null>(null);
const fillColorHost = ref<HTMLDivElement | null>(null);
const borderColorHost = ref<HTMLDivElement | null>(null);
const fontColorNative = ref<HTMLInputElement | null>(null);
const fillColorNative = ref<HTMLInputElement | null>(null);
const borderColorNative = ref<HTMLInputElement | null>(null);
let activePalette: HTMLElement | null = null;

const mountColorPalette = (kind: 'fontColor' | 'fillColor' | 'borderColor'): void => {
  const host =
    kind === 'fontColor'
      ? fontColorHost.value
      : kind === 'fillColor'
        ? fillColorHost.value
        : borderColorHost.value;
  if (!host) return;
  const nativeRef =
    kind === 'fontColor' ? fontColorNative : kind === 'fillColor' ? fillColorNative : borderColorNative;
  const ariaLabel =
    kind === 'fontColor' ? tr.value.fontColor : kind === 'fillColor' ? tr.value.fillColor : tr.value.lineColor;
  const value =
    kind === 'fontColor'
      ? active.value.fontColor
      : kind === 'fillColor'
        ? active.value.fillColor
        : borderColor.value;
  const palette = createColorPalette({
    themeLabel: tr.value.themeColors,
    standardLabel: tr.value.standardColors,
    moreColorsLabel: tr.value.moreColors,
    ariaLabel,
    value,
    automatic:
      kind === 'fontColor' || kind === 'borderColor'
        ? { label: tr.value.automatic, color: '#000000' }
        : null,
    onPick: (color) => onPaletteColor(kind, color),
    onMoreColors: () => {
      closeDropdown();
      nativeRef.value?.click();
    },
  });
  host.appendChild(palette.el);
  palette.focus();
  activePalette = palette.el;
};

watch(openDropdown, async (value) => {
  activePalette?.remove();
  activePalette = null;
  if (value === 'fontColor' || value === 'fillColor' || value === 'borderColor') {
    await nextTick();
    mountColorPalette(value);
  }
});

watch(
  () => props.features,
  (features) => {
    formulaBarVisible.value = features?.formulaBar !== false;
  },
);

const onFormatAsTable = (style: FormatTableAction = 'medium'): void => {
  const inst = props.instance;
  if (!inst) return;
  const r = inst.store.getState().selection.range;
  recordTablesChange(inst.history, inst.store, () => {
    formatAsTable(inst.store, r, { style });
  });
  closeDropdown();
};

const onInsertRows = (): void => insertSelectedRows(props.instance);
const onDeleteRows = (): void => deleteSelectedRows(props.instance);
const onInsertCols = (): void => insertSelectedCols(props.instance);
const onDeleteCols = (): void => deleteSelectedCols(props.instance);

const onToggleRowsHidden = (): void => toggleSelectedRowsHidden(props.instance);
const onToggleColsHidden = (): void => toggleSelectedColsHidden(props.instance);

const onWindowAction = (action: WindowAction): void => {
  handleWindowAction(props.instance, action);
  closeDropdown();
};

const onFilterToggle = (): void => {
  const inst = props.instance;
  if (!inst) return;
  toggleAutoFilterFromSelection(inst.store, inst.history);
};

const onSort = (direction: 'asc' | 'desc'): void => {
  const inst = props.instance;
  if (!inst) return;
  sortActiveColumnAuto({
    store: inst.store,
    workbook: inst.workbook,
    history: inst.history,
    direction,
  });
};

const onRemoveDuplicates = (): void => {
  const inst = props.instance;
  if (!inst) return;
  const s = inst.store.getState();
  const range = s.selection.range;
  removeDuplicatesDialog.value = {
    columns: Array.from({ length: range.c1 - range.c0 + 1 }, (_, i) => range.c0 + i),
    hasHeader: inferSortHasHeader(s, range),
  };
};

const applyRemoveDuplicatesDialog = (): void => {
  const inst = props.instance;
  const draft = removeDuplicatesDialog.value;
  if (!inst || !draft) return;
  if (draft.columns.length === 0) {
    ribbonReportDialog.value = {
      title: cellText.value.removeDuplicatesDialogTitle,
      items: [
        {
          severity: 'warning',
          label: cellText.value.removeDuplicatesNoColumns,
          detail: '',
        },
      ],
    };
    return;
  }
  const s = inst.store.getState();
  inst.history.begin();
  let removed = 0;
  try {
    removed = removeDuplicates(s, inst.store, inst.workbook, s.selection.range, {
      columns: draft.columns,
      hasHeader: draft.hasHeader,
    });
  } finally {
    inst.history.end();
  }
  if (removed > 0) mutators.replaceCells(inst.store, inst.workbook.cells(s.data.sheetIndex));
  removeDuplicatesDialog.value = null;
};

const setRemoveDuplicatesColumns = (columns: number[]): void => {
  if (!removeDuplicatesDialog.value) return;
  removeDuplicatesDialog.value = { ...removeDuplicatesDialog.value, columns };
};

const toggleRemoveDuplicatesColumn = (col: number, checked: boolean): void => {
  const draft = removeDuplicatesDialog.value;
  if (!draft) return;
  const columns = checked
    ? [...draft.columns, col].sort((a, b) => a - b)
    : draft.columns.filter((item) => item !== col);
  removeDuplicatesDialog.value = { ...draft, columns };
};

const onCustomSort = (): void => {
  const inst = props.instance;
  if (!inst) return;
  const s = inst.store.getState();
  const range = s.selection.range;
  sortDialog.value = {
    byCol:
      s.selection.active.col >= range.c0 && s.selection.active.col <= range.c1
        ? s.selection.active.col
        : range.c0,
    direction: 'asc',
    hasHeader: range.r0 < range.r1,
  };
};

const applyCustomSort = (): void => {
  const inst = props.instance;
  const draft = sortDialog.value;
  if (!inst || !draft) return;
  sortRangeWithHistory({
    store: inst.store,
    workbook: inst.workbook,
    history: inst.history,
    range: inst.store.getState().selection.range,
    options: {
      byCol: draft.byCol,
      direction: draft.direction,
      hasHeader: draft.hasHeader,
    },
  });
  sortDialog.value = null;
};

const onSortMenuAction = (action: SortAction): void => {
  const inst = props.instance;
  if (!inst) return;
  const s = inst.store.getState();
  if (action === 'asc' || action === 'desc') onSort(action);
  else if (action === 'custom') onCustomSort();
  else if (action === 'filter') onFilterToggle();
  else if (action === 'filter-clear' && s.ui.filterRange)
    executeRibbonFilterDataAction({ store: inst.store, history: inst.history, action: 'clear' });
  else if (action === 'filter-reapply')
    executeRibbonFilterDataAction({ store: inst.store, history: inst.history, action: 'reapply' });
  else if (action === 'filter-by-selected')
    executeRibbonFilterDataAction({
      store: inst.store,
      history: inst.history,
      action: 'filter-by-selected',
    });
  else if (action === 'filter-advanced')
    advancedFilterDialog.value = {
      listRange: formatA1Range(s.selection.range),
      criteriaRange: '',
      copyTo: '',
      uniqueOnly: false,
    };
  else if (action === 'dedupe') onRemoveDuplicates();
  else if (action === 'conditional') inst.openCfRulesDialog();
  else if (action === 'named') inst.openNamedRangeDialog();
  closeDropdown();
};

const onFilterDataAction = (action: FilterDataAction): void => {
  const inst = props.instance;
  if (!inst) return;
  const result = executeRibbonFilterDataAction({
    store: inst.store,
    history: inst.history,
    action,
  });
  if (result.kind === 'open-advanced') {
    advancedFilterDialog.value = {
      listRange: formatA1Range(result.range),
      criteriaRange: '',
      copyTo: '',
      uniqueOnly: false,
    };
  } else if (result.kind === 'open-filter-dropdown') {
    inst.openFilterDropdown(result.range, result.column);
  }
  closeDropdown();
};

const applyAdvancedFilterDialog = (): void => {
  const inst = props.instance;
  const draft = advancedFilterDialog.value;
  if (!inst || !draft) return;
  const state = inst.store.getState();
  const sheet = state.data.sheetIndex;
  const sheetName = inst.workbook.sheetName(sheet);
  const listRange = parseA1Range(draft.listRange, sheet, sheetName);
  const criteriaRange = parseA1Range(draft.criteriaRange, sheet, sheetName);
  if (!listRange || !criteriaRange) return;
  const copyToRange = draft.copyTo.trim() ? parseA1Range(draft.copyTo, sheet, sheetName) : null;
  if (draft.copyTo.trim()) {
    if (!copyToRange) return;
    inst.history.begin();
    let copied = 0;
    try {
      copied = copyAdvancedFilterResult(
        inst.store.getState(),
        inst.store,
        listRange,
        criteriaRange,
        { sheet, row: copyToRange.r0, col: copyToRange.c0 },
        { uniqueOnly: draft.uniqueOnly },
        inst.workbook,
      );
    } finally {
      inst.history.end();
    }
    ribbonReportDialog.value = {
      title: cellText.value.advancedFilterDialogTitle,
      items: [
        {
          severity: 'info',
          label: cellText.value.filterAdvanced,
          detail: cellText.value.advancedFilterCopiedStatus.replace('{count}', String(copied)),
        },
      ],
    };
  } else {
    recordFilterChange(inst.history, inst.store, () =>
      applyAdvancedFilter(inst.store.getState(), inst.store, listRange, criteriaRange),
    );
  }
  advancedFilterDialog.value = null;
};

const onTextOrientationAction = (action: TextOrientationAction): void => {
  if (action === 'formatAlignment') {
    props.instance?.openFormatDialog('align');
    closeDropdown();
    return;
  }
  const rotation =
    action === 'angleCounterclockwise'
      ? 45
      : action === 'angleClockwise'
        ? -45
        : action === 'rotateTextUp' || action === 'verticalText'
          ? 90
          : action === 'rotateTextDown'
            ? -90
            : 0;
  wrapFormat((s, st) => setRotation(s, st, rotation));
  closeDropdown();
};

const onTextToColumnsAction = (action: TextToColumnsAction): void => {
  const inst = props.instance;
  if (!inst) return;
  if (action === 'custom') {
    textToColumnsDialog.value = {
      comma: true,
      tab: false,
      semicolon: false,
      space: false,
      collapseConsecutive: false,
    };
    closeDropdown();
    return;
  }
  const delimiter =
    action === 'tab' ? '\t' : action === 'semicolon' ? ';' : action === 'space' ? ' ' : ',';
  applyTextToColumns([delimiter]);
  closeDropdown();
};

const applyTextToColumns = (delimiters: readonly string[], collapseConsecutive = false): void => {
  const inst = props.instance;
  if (!inst || delimiters.length === 0) return;
  const state = inst.store.getState();
  inst.history.begin();
  let max = 0;
  try {
    recordFormatChange(inst.history, inst.store, () => {
      max = textToColumns(state, inst.store, inst.workbook, state.selection.range, delimiters, {
        collapseConsecutiveDelimiters: collapseConsecutive,
      });
    });
  } finally {
    inst.history.end();
  }
  if (max > 0) mutators.replaceCells(inst.store, inst.workbook.cells(state.data.sheetIndex));
};

const applyTextToColumnsDialog = (): void => {
  const draft = textToColumnsDialog.value;
  if (!draft) return;
  const delimiters = [
    draft.comma ? ',' : '',
    draft.tab ? '\t' : '',
    draft.semicolon ? ';' : '',
    draft.space ? ' ' : '',
  ].filter(Boolean);
  applyTextToColumns(delimiters, draft.collapseConsecutive);
  textToColumnsDialog.value = null;
};

const onDataValidationAction = (action: DataValidationAction): void => {
  const inst = props.instance;
  if (!inst) return;
  if (action === 'settings') inst.openDataValidationDialog();
  else if (action === 'clearValidation') {
    const state = inst.store.getState();
    clearValidationInRangeWithEngine(inst.store, inst.history, inst.workbook, state.selection.range);
  } else if (action === 'clearCircles')
    recordValidationCirclesChange(inst.history, inst.store, () => {
      clearValidationCircles(inst.store);
    });
  else {
    const state = inst.store.getState();
    recordValidationCirclesChange(inst.history, inst.store, () => {
      circleInvalidValidationDataInSheet(
        inst.store,
        state.selection.range.sheet,
        makeRangeResolver(inst.workbook, state.data.sheetIndex),
      );
    });
  }
  closeDropdown();
};

const onFormulaAuditingAction = (action: FormulaAuditingAction): void => {
  const inst = props.instance;
  if (!inst) return;
  const result = executeRibbonFormulaAuditingAction({
    store: inst.store,
    workbook: inst.workbook,
    history: inst.history,
    action,
    strings: { errorChecking: cellText.value.errorChecking },
  });
  if (result.kind === 'trace-precedents') inst.tracePrecedents();
  else if (result.kind === 'report') ribbonReportDialog.value = result.report;
  closeDropdown();
};

const onClearArrowsAction = (action: ClearArrowsAction): void => {
  const inst = props.instance;
  if (!inst) return;
  if (action === 'clear-precedents') clearTraceArrowsByKind(inst.store, 'precedent', inst.history);
  else if (action === 'clear-dependents')
    clearTraceArrowsByKind(inst.store, 'dependent', inst.history);
  else clearTraceArrows(inst.store, inst.history);
  closeDropdown();
};

const onFindAction = (action: FindAction): void => {
  const inst = props.instance;
  if (!inst) return;
  const result = executeRibbonFindAction({
    store: inst.store,
    workbook: inst.workbook,
    action,
    strings: {
      findSelect: cellText.value.findSelect,
      findNoMatches: cellText.value.findNoMatches,
      commentNone: cellText.value.commentNone,
    },
  });
  if (result.kind === 'open-find') inst.openFindReplace(result.mode);
  else if (result.kind === 'open-go-to') inst.openGoTo();
  else if (result.kind === 'open-go-to-special') inst.openGoToSpecial();
  else if (result.kind === 'report') ribbonReportDialog.value = result.report;
  closeDropdown();
};

const onCommentAction = (action: CommentAction): void => {
  const inst = props.instance;
  if (!inst) return;
  executeRibbonCommentAction({
    store: inst.store,
    workbook: inst.workbook,
    history: inst.history,
    action,
  });
  closeDropdown();
};

const onProtectionAction = (action: ProtectionAction): void => {
  const inst = props.instance;
  if (!inst) return;
  ribbonReportDialog.value = executeRibbonProtectionAction({
    store: inst.store,
    action,
    strings: cellText.value,
  });
  closeDropdown();
};

const onHyperlinkAction = (action: HyperlinkAction): void => {
  const inst = props.instance;
  if (!inst) return;
  const result = executeRibbonHyperlinkAction({
    store: inst.store,
    workbook: inst.workbook,
    history: inst.history,
    action,
    strings: cellText.value,
  });
  if (result.kind === 'open-hyperlink-dialog') inst.openHyperlinkDialog();
  else if (result.kind === 'open-external-dialog') inst.openExternalLinksDialog();
  else if (result.kind === 'open-url') window.open(result.url, '_blank', 'noopener,noreferrer');
  else if (result.kind === 'report') ribbonReportDialog.value = result.report;
  closeDropdown();
};

const onSelectComment = (direction: 1 | -1): void => {
  const inst = props.instance;
  if (!inst) return;
  const state = inst.store.getState();
  const comments = listComments(state);
  if (comments.length === 0) return;
  const activeAddr = state.selection.active;
  const current = comments.findIndex(
    (entry) => entry.addr.row === activeAddr.row && entry.addr.col === activeAddr.col,
  );
  const nextIndex =
    current >= 0
      ? (current + direction + comments.length) % comments.length
      : direction > 0
        ? 0
        : comments.length - 1;
  const next = comments[nextIndex]?.addr;
  if (next) mutators.setActive(inst.store, next);
};

const selectionOutlineAxis = (): 'row' | 'col' => {
  const inst = props.instance;
  if (!inst) return 'row';
  const range = inst.store.getState().selection.range;
  return range.r1 - range.r0 >= range.c1 - range.c0 ? 'row' : 'col';
};

const onOutlineAction = (
  action: 'group' | 'ungroup' | 'show-detail' | 'hide-detail',
  axis?: OutlineAxisAction,
): void => {
  const inst = props.instance;
  if (!inst) return;
  const range = inst.store.getState().selection.range;
  const targetAxis = axis === 'rows' ? 'row' : axis === 'cols' ? 'col' : selectionOutlineAxis();
  if (targetAxis === 'row') {
    if (action === 'group') groupRows(inst.store, inst.history, range.r0, range.r1, inst.workbook);
    else if (action === 'ungroup')
      ungroupRows(inst.store, inst.history, range.r0, range.r1, inst.workbook);
    else if (action === 'show-detail') showRows(inst.store, inst.history, range.r0, range.r1, inst.workbook);
    else collapseRowGroup(inst.store, inst.history, range.r0, range.r1, inst.workbook);
  } else {
    if (action === 'group') groupCols(inst.store, inst.history, range.c0, range.c1, inst.workbook);
    else if (action === 'ungroup')
      ungroupCols(inst.store, inst.history, range.c0, range.c1, inst.workbook);
    else if (action === 'show-detail') showCols(inst.store, inst.history, range.c0, range.c1, inst.workbook);
    else collapseColGroup(inst.store, inst.history, range.c0, range.c1, inst.workbook);
  }
};

const onViewFlag = (flag: 'gridlines' | 'headings' | 'formulas' | 'r1c1'): void => {
  const inst = props.instance;
  if (!inst) return;
  const ui = inst.store.getState().ui;
  if (flag === 'gridlines') setGridlinesVisible(inst.store, ui.showGridLines === false);
  else if (flag === 'headings') setHeadingsVisible(inst.store, ui.showHeaders === false);
  else if (flag === 'formulas') setShowFormulas(inst.store, !ui.showFormulas);
  else setR1C1ReferenceStyle(inst.store, !ui.r1c1);
};

const onPrintSheetOption = (option: 'gridlines' | 'headings'): void => {
  const inst = props.instance;
  if (!inst) return;
  const sheet = inst.store.getState().data.sheetIndex;
  if (option === 'gridlines')
    setPrintGridlines(inst.store, sheet, !active.value.printGridlines, inst.history);
  else setPrintHeadings(inst.store, sheet, !active.value.printHeadings, inst.history);
};

const onWorkbookView = (mode: 'normal' | 'pageLayout' | 'pageBreakPreview'): void => {
  const inst = props.instance;
  if (!inst) return;
  setWorkbookView(inst.store, mode);
};

const onSheetViewSelect = (value: string): void => {
  const inst = props.instance;
  if (!inst) return;
  if (value === 'current') {
    inst.store.setState((state) => ({
      ...state,
      sheetViews: { ...state.sheetViews, activeViewId: null },
    }));
    return;
  }
  activateSheetView(inst.store, value);
};

const onSheetViewSelectEvent = (event: Event): void => {
  const target = event.target;
  if (target instanceof HTMLSelectElement) onSheetViewSelect(target.value);
};

const onSheetViewSave = (): void => {
  const inst = props.instance;
  if (!inst) return;
  const count = inst.store.getState().sheetViews.views.length + 1;
  const id = `view-${Date.now().toString(36)}-${count}`;
  saveSheetView(inst.store, id, `${viewToolbarText.value.views} ${count}`);
  inst.store.setState((state) => ({
    ...state,
    sheetViews: { ...state.sheetViews, activeViewId: id },
  }));
};

const onSheetViewDelete = (): void => {
  const inst = props.instance;
  if (!inst) return;
  const id = inst.store.getState().sheetViews.activeViewId;
  if (id) deleteSheetView(inst.store, id);
};

const onToggleFormulaBar = (): void => {
  const inst = props.instance;
  if (!inst) return;
  const next = !formulaBarVisible.value;
  formulaBarVisible.value = next;
  inst.setFeatures({ ...(props.features ?? {}), formulaBar: next });
};

const onSymbolAction = (symbol: SymbolAction): void => {
  const inst = props.instance;
  if (!inst) return;
  const addr = inst.store.getState().selection.active;
  if (inst.workbook.cellFormula(addr)) return;
  if (!isCellWritable(inst.store.getState(), addr)) {
    warnProtected(addr);
    closeDropdown();
    return;
  }
  const text =
    symbol === MORE_SYMBOL_ACTION
      ? typeof window.prompt === 'function'
        ? window.prompt(cellText.value.symbolPrompt, '')?.trim() ?? ''
        : ''
      : symbol;
  if (text.length === 0) {
    if (symbol === MORE_SYMBOL_ACTION) {
      ribbonReportDialog.value = {
        title: cellText.value.symbol,
        items: [
          {
            severity: 'warning',
            label: cellText.value.symbolMore,
            detail: cellText.value.symbolInvalid,
          },
        ],
      };
    }
    closeDropdown();
    return;
  }
  const value = inst.workbook.getValue(addr);
  const current = value.kind === 'text' ? value.value : '';
  inst.history.begin();
  try {
    inst.workbook.setText(addr, `${current}${text}`);
  } finally {
    inst.history.end();
  }
  mutators.replaceCells(inst.store, inst.workbook.cells(addr.sheet));
  closeDropdown();
};

const onFillDown = (): void => {
  const inst = props.instance;
  if (!inst) return;
  executeRibbonFillAction({
    store: inst.store,
    workbook: inst.workbook,
    history: inst.history,
    action: 'down',
  });
};

const onFillAction = (action: FillAction): void => {
  const inst = props.instance;
  if (!inst) return;
  executeRibbonFillAction({
    store: inst.store,
    workbook: inst.workbook,
    history: inst.history,
    action,
  });
  closeDropdown();
};

const onClearAction = (action: ClearAction): void => {
  const inst = props.instance;
  if (!inst) return;
  executeRibbonClearAction({
    store: inst.store,
    workbook: inst.workbook,
    history: inst.history,
    action,
  });
  closeDropdown();
};

const onCreateChart = (action: ChartAction = 'column'): void => {
  const inst = props.instance;
  if (!inst) return;
  createRibbonChartFromSelection({
    store: inst.store,
    history: inst.history,
    range: inst.store.getState().selection.range,
    action,
    idPrefix: 'vue-ribbon-chart',
  });
  closeDropdown();
};

const onPrintAreaAction = (action: PrintAreaAction): void => {
  const inst = props.instance;
  if (!inst) return;
  const state = inst.store.getState();
  const sheet = state.data.sheetIndex;
  const range = state.selection.range;
  recordPageSetupChange(inst.history, inst.store, () => {
    if (action === 'clear') {
      clearPrintArea(inst.store, sheet);
      return;
    }
    const start = `${colLetter(range.c0)}${range.r0 + 1}`;
    const end = `${colLetter(range.c1)}${range.r1 + 1}`;
    setPrintArea(inst.store, sheet, start === end ? start : `${start}:${end}`);
  });
  closeDropdown();
};

const onPrintTitleAction = (action: PrintTitleAction): void => {
  const inst = props.instance;
  if (!inst) return;
  const state = inst.store.getState();
  const sheet = state.data.sheetIndex;
  const range = state.selection.range;
  recordPageSetupChange(inst.history, inst.store, () => {
    if (action === 'clear') {
      clearPrintTitles(inst.store, sheet);
    } else if (action === 'rows') {
      const rows = range.r0 === range.r1 ? `${range.r0 + 1}` : `${range.r0 + 1}:${range.r1 + 1}`;
      setPrintTitleRows(inst.store, sheet, rows);
    } else {
      const cols =
        range.c0 === range.c1
          ? colLetter(range.c0)
          : `${colLetter(range.c0)}:${colLetter(range.c1)}`;
      setPrintTitleCols(inst.store, sheet, cols);
    }
  });
  closeDropdown();
};

const onPageBreakAction = (action: PageBreakAction): void => {
  const inst = props.instance;
  if (!inst) return;
  const state = inst.store.getState();
  const sheet = state.data.sheetIndex;
  const activeCell = state.selection.active;
  recordPageSetupChange(inst.history, inst.store, () => {
    if (action === 'insert-row') insertManualPageBreak(inst.store, sheet, 'row', activeCell.row);
    else if (action === 'insert-col') insertManualPageBreak(inst.store, sheet, 'col', activeCell.col);
    else if (action === 'remove-row') removeManualPageBreak(inst.store, sheet, 'row', activeCell.row);
    else if (action === 'remove-col') removeManualPageBreak(inst.store, sheet, 'col', activeCell.col);
    else resetManualPageBreaks(inst.store, sheet);
  });
  closeDropdown();
};

const onSheetBackgroundAction = (action: SheetBackgroundAction): void => {
  const inst = props.instance;
  if (!inst) return;
  const sheet = inst.store.getState().data.sheetIndex;
  if (action === 'clear') {
    clearSheetBackgroundImage(inst.store, sheet, inst.history);
    closeDropdown();
    return;
  }
  sheetBackgroundInput.value?.click();
  closeDropdown();
};

const onSheetBackgroundFileChange = (event: Event): void => {
  const inst = props.instance;
  if (!inst) return;
  const input = event.currentTarget as HTMLInputElement;
  const file = input.files?.[0];
  input.value = '';
  if (!file || !file.type.startsWith('image/')) return;
  const sheet = inst.store.getState().data.sheetIndex;
  const reader = new FileReader();
  reader.onload = () => {
    if (typeof reader.result === 'string')
      setSheetBackgroundImage(inst.store, sheet, reader.result, inst.history);
  };
  reader.readAsDataURL(file);
};

const onThemeAction = (action: ThemeAction): void => {
  const inst = props.instance;
  if (!inst) return;
  inst.setTheme(action);
  active.value = projectActiveState(inst);
  closeDropdown();
};

const onDefinedNameAction = (action: DefinedNameAction): void => {
  const inst = props.instance;
  if (!inst) return;
  if (action === 'manager' || action === 'define') {
    inst.openNamedRangeDialog();
    closeDropdown();
    return;
  }
  if (
    action === 'createTopRow' ||
    action === 'createBottomRow' ||
    action === 'createLeftColumn' ||
    action === 'createRightColumn'
  ) {
    const result = recordDefinedNamesChange(inst.history, inst.workbook, () =>
      createDefinedNamesFromSelection(
        inst.store.getState(),
        inst.workbook,
        action === 'createTopRow'
          ? 'top-row'
          : action === 'createBottomRow'
            ? 'bottom-row'
            : action === 'createLeftColumn'
              ? 'left-column'
              : 'right-column',
      ),
    );
    const sheet = inst.store.getState().data.sheetIndex;
    if (result.ok) mutators.replaceCells(inst.store, inst.workbook.cells(sheet));
    closeDropdown();
    return;
  }
  if (action.startsWith('use:')) {
    const result = insertDefinedNameFormula(
      inst.store.getState(),
      inst.workbook,
      action.slice('use:'.length),
      inst.store,
    );
    if (result) {
      mutators.replaceCells(inst.store, inst.workbook.cells(result.addr.sheet));
      mutators.setActive(inst.store, result.addr);
    }
    closeDropdown();
  }
};

const onCalculationAction = (action: CalculationAction): void => {
  const inst = props.instance;
  if (!inst) return;
  if (action === 'iterative') {
    inst.openIterativeDialog();
    closeDropdown();
    return;
  }
  const mode = action === 'auto' ? 0 : action === 'manual' ? 1 : 2;
  inst.workbook.setCalcMode(mode);
  active.value = projectActiveState(inst);
  closeDropdown();
};

const onWatchAction = (action: WatchAction): void => {
  const inst = props.instance;
  if (!inst) return;
  const state = inst.store.getState();
  if (action === 'add') {
    recordWatchesChange(inst.history, inst.store, () => {
      watchRange(inst.store, state.selection.range);
    });
  } else if (action === 'delete') {
    recordWatchesChange(inst.history, inst.store, () => {
      unwatchCell(inst.store, state.selection.active);
    });
  } else if (action === 'delete-all') {
    recordWatchesChange(inst.history, inst.store, () => {
      clearWatchedCells(inst.store);
    });
  }
  inst.openWatchWindow();
  active.value = projectActiveState(inst);
  closeDropdown();
};

const onZoom = (zoom: number): void => {
  const inst = props.instance;
  if (!inst) return;
  setSheetZoom(inst.store, zoom, inst.workbook);
};
const openZoomDialog = (): void => {
  const inst = props.instance;
  if (!inst) return;
  zoomDialog.value = String(Math.round(inst.store.getState().viewport.zoom * 100));
};
const applyZoomDialog = (): void => {
  const inst = props.instance;
  if (!inst || zoomDialog.value == null) return;
  const percent = Number.parseFloat(zoomDialog.value);
  if (!Number.isFinite(percent)) return;
  const clamped = Math.max(10, Math.min(400, percent));
  setSheetZoom(inst.store, clamped / 100, inst.workbook);
  zoomDialog.value = null;
};
const onZoomSelection = (): void => {
  const inst = props.instance;
  if (!inst) return;
  const state = inst.store.getState();
  const range = state.selection.range;
  const selectedRows = Math.max(1, range.r1 - range.r0 + 1);
  const selectedCols = Math.max(1, range.c1 - range.c0 + 1);
  const rowFit = state.viewport.rowCount / selectedRows;
  const colFit = state.viewport.colCount / selectedCols;
  setSheetZoom(inst.store, state.viewport.zoom * Math.min(rowFit, colFit), inst.workbook);
};
</script>

<template>
  <div class="demo__ribbon-shell" :class="{ 'demo__ribbon-shell--collapsed': ribbonCollapsed }" @keydown="onDropdownKeydown">
    <div
      ref="tablistRef"
      class="demo__ribbon-tabs"
      role="tablist"
      :aria-label="tr.ribbonTabs"
      :data-ribbon-collapsed="ribbonCollapsed ? 'true' : 'false'"
      @keydown="onRibbonTabKeydown"
    >
      <button
        v-for="tab in tabs"
        :key="tab.id"
        :class="[
          'demo__ribbon-tab',
          {
            'demo__ribbon-tab--file': tab.id === 'file',
            'demo__ribbon-tab--active': props.activeTab === tab.id,
          },
        ]"
        type="button"
        role="tab"
        :data-ribbon-tab="tab.id"
        :aria-selected="props.activeTab === tab.id"
        :tabindex="props.activeTab === tab.id ? 0 : -1"
        @click="setActiveTab(tab.id)"
        @dblclick="ribbonCollapsed = !ribbonCollapsed"
      >
        {{ tab.label }}
      </button>
    </div>
    <div
      v-if="props.activeTab !== 'file'"
      ref="ribbonDisplayRef"
      class="demo__ribbon-display"
      @keydown="onRibbonDisplayKeydown"
    >
      <button
        class="demo__ribbon-toggle"
        type="button"
        :aria-label="ribbonDisplayLabel"
        aria-haspopup="menu"
        :aria-expanded="ribbonDisplayMenuOpen ? 'true' : 'false'"
        :title="ribbonDisplayLabel"
        @click="ribbonDisplayMenuOpen = !ribbonDisplayMenuOpen"
      />
      <div v-if="ribbonDisplayMenuOpen" class="demo__ribbon-display-menu" role="menu">
        <button
          v-for="option in ribbonDisplayOptions"
          :key="option.id"
          class="demo__ribbon-display-option"
          type="button"
          role="menuitemradio"
          :aria-checked="(option.id === 'collapsed') === ribbonCollapsed ? 'true' : 'false'"
          @click="ribbonCollapsed = option.id === 'collapsed'; ribbonDisplayMenuOpen = false"
        >
          {{ option.label }}
        </button>
      </div>
    </div>
    <div v-if="props.activeTab === 'file'" class="demo__backstage" role="dialog" aria-modal="true" :aria-label="fileLabel">
      <nav class="demo__backstage-nav" :aria-label="fileLabel">
        <button class="demo__backstage-navitem" type="button" :aria-label="backstageText.back" @click="setActiveTab(previousNonFileTab)">←</button>
        <strong>{{ fileLabel }}</strong>
        <button class="demo__backstage-navitem demo__backstage-navitem--active" type="button">{{ backstageText.info }}</button>
        <button class="demo__backstage-navitem" type="button" :disabled="!props.onNewWorkbook" @click="props.onNewWorkbook?.()">{{ backstageText.newLabel }}</button>
        <button class="demo__backstage-navitem" type="button" :disabled="!props.onOpenWorkbook" @click="props.onOpenWorkbook?.()">{{ backstageText.open }}</button>
        <button class="demo__backstage-navitem" type="button" :disabled="!props.onSaveWorkbook" @click="props.onSaveWorkbook?.()">{{ backstageText.save }}</button>
        <button class="demo__backstage-navitem" type="button" :disabled="!props.onSaveWorkbookAs" @click="props.onSaveWorkbookAs?.()">{{ backstageText.saveAs }}</button>
        <button class="demo__backstage-navitem" data-ribbon-command="print" type="button" :disabled="disabled" @click="props.instance?.print('print')">{{ tr.print }}</button>
        <button class="demo__backstage-navitem" data-ribbon-command="pageSetup" type="button" :disabled="disabled" @click="props.instance?.openPageSetup()">{{ backstageText.options }}</button>
      </nav>
      <main class="demo__backstage-main">
        <div class="demo__backstage-title">
          <span class="demo__backstage-xl">X</span>
          <div>
            <h1>{{ backstageText.title }}</h1>
            <p>{{ backstageText.subtitle }}</p>
          </div>
        </div>
        <section class="demo__backstage-info">
          <div>
            <h2 class="demo__backstage-section-title">{{ backstageText.workbookInfo }}</h2>
            <div class="demo__backstage-command-list">
              <button class="demo__backstage-command" :class="{ 'demo__backstage-command--active': workbookStructureProtected }" data-ribbon-command="protect" type="button" :aria-pressed="workbookStructureProtected ? 'true' : undefined" :disabled="disabled" @click="onBackstageProtectWorkbook"><span class="demo__backstage-command-icon">P</span><span><strong>{{ backstageText.protect }}</strong><span>{{ backstageText.protectBody }}</span></span></button>
              <button class="demo__backstage-command" data-ribbon-command="inspect" type="button" :disabled="disabled" @click="onBackstageInspectWorkbook"><span class="demo__backstage-command-icon">!</span><span><strong>{{ backstageText.inspect }}</strong><span>{{ backstageText.inspectBody }}</span></span></button>
              <button class="demo__backstage-command" type="button" :disabled="!props.onSaveWorkbookAs" @click="props.onSaveWorkbookAs?.()"><span class="demo__backstage-command-icon">S</span><span><strong>{{ backstageText.manage }}</strong><span>{{ backstageText.manageBody }}</span></span></button>
            </div>
          </div>
          <aside class="demo__backstage-properties">
            <h2 class="demo__backstage-section-title">{{ backstageText.properties }}</h2>
            <div class="demo__backstage-preview">X</div>
            <dl class="demo__backstage-prop-list">
              <dt>{{ backstageText.name }}</dt><dd>{{ backstageText.title }}</dd>
              <dt>{{ backstageText.type }}</dt><dd>{{ backstageText.typeValue }}</dd>
              <dt>{{ backstageText.status }}</dt><dd>{{ backstageText.statusValue }}</dd>
              <dt>{{ backstageText.location }}</dt><dd>{{ backstageText.locationValue }}</dd>
            </dl>
          </aside>
        </section>
        <div class="demo__backstage-grid">
          <button class="demo__backstage-card" type="button" :disabled="!props.onNewWorkbook" @click="props.onNewWorkbook?.()"><strong>{{ backstageText.newLabel }}</strong><span>{{ backstageText.newBody }}</span></button>
          <button class="demo__backstage-card" type="button" :disabled="!props.onOpenWorkbook" @click="props.onOpenWorkbook?.()"><strong>{{ backstageText.open }}</strong><span>{{ backstageText.openBody }}</span></button>
          <button class="demo__backstage-card" type="button" :disabled="!props.onSaveWorkbook" @click="props.onSaveWorkbook?.()"><strong>{{ backstageText.save }}</strong><span>{{ backstageText.saveBody }}</span></button>
          <button class="demo__backstage-card" type="button" :disabled="!props.onSaveWorkbookAs" @click="props.onSaveWorkbookAs?.()"><strong>{{ backstageText.saveAs }}</strong><span>{{ backstageText.saveAsBody }}</span></button>
          <button class="demo__backstage-card" data-ribbon-command="print" type="button" :disabled="disabled" @click="props.instance?.print('print')"><strong>{{ tr.print }}</strong><span>{{ backstageText.printBody }}</span></button>
          <button class="demo__backstage-card" data-ribbon-command="pageSetup" type="button" :disabled="disabled" @click="props.instance?.openPageSetup()"><strong>{{ backstageText.options }}</strong><span>{{ backstageText.optionsBody }}</span></button>
        </div>
      </main>
    </div>

    <div
      v-else
      :class="['demo__ribbon', { 'demo__ribbon--office365-home': props.activeTab === 'home' }]"
      role="toolbar"
      :aria-label="`${strings.ribbon.tabs[props.activeTab]} ${tr.ribbon}`"
    >
      <template v-if="props.activeTab === 'home'">
      <section class="demo__ribbon-group demo__ribbon-group--clipboard" :aria-label="tr.clipboard">
        <div class="demo__ribbon-tools">
    <div class="demo__rb-menu" data-ribbon-command="paste" :class="{ 'demo__rb-menu--open': openDropdown === 'paste' }" data-dropdown-name="paste">
      <button class="demo__rb demo__rb-menu__btn demo__rb--large" type="button" :disabled="disabled" :title="tr.paste" :aria-label="tr.paste" :aria-keyshortcuts="keyShortcuts('paste')" aria-haspopup="menu" :aria-expanded="openDropdown === 'paste'" @click="toggleDropdown('paste')">
        <RibbonIcon name="paste" />
        <span>{{ tr.paste }}</span>
        <svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
      </button>
      <div v-if="openDropdown === 'paste'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.paste">
        <button class="demo__merge-menu__item" type="button" role="menuitem" data-paste-action="paste" @click="onPasteAction('paste'); closeDropdown()"><RibbonIcon name="paste" /><span>{{ tr.paste }}</span></button>
        <div class="demo__cf-menu__sep" role="presentation" />
        <button class="demo__merge-menu__item" type="button" role="menuitem" data-paste-action="pasteFormulas" @click="onPasteAction('pasteFormulas'); closeDropdown()"><RibbonIcon name="paste" /><span>{{ strings.contextMenu.pasteFormulas }}</span></button>
        <button class="demo__merge-menu__item" type="button" role="menuitem" data-paste-action="pasteFormulasNumFmt" @click="onPasteAction('pasteFormulasNumFmt'); closeDropdown()"><RibbonIcon name="paste" /><span>{{ strings.contextMenu.pasteFormulasNumFmt }}</span></button>
        <div class="demo__cf-menu__sep" role="presentation" />
        <button class="demo__merge-menu__item" type="button" role="menuitem" data-paste-action="pasteValues" @click="onPasteAction('pasteValues'); closeDropdown()"><RibbonIcon name="paste" /><span>{{ strings.contextMenu.pasteValues }}</span></button>
        <button class="demo__merge-menu__item" type="button" role="menuitem" data-paste-action="pasteValuesNumFmt" @click="onPasteAction('pasteValuesNumFmt'); closeDropdown()"><RibbonIcon name="paste" /><span>{{ strings.contextMenu.pasteValuesNumFmt }}</span></button>
        <button class="demo__merge-menu__item" type="button" role="menuitem" data-paste-action="pasteFormatsOnly" @click="onPasteAction('pasteFormatsOnly'); closeDropdown()"><RibbonIcon name="paste" /><span>{{ strings.contextMenu.pasteFormatsOnly }}</span></button>
        <button class="demo__merge-menu__item" type="button" role="menuitem" data-paste-action="pasteTranspose" @click="onPasteAction('pasteTranspose'); closeDropdown()"><RibbonIcon name="paste" /><span>{{ strings.contextMenu.pasteTranspose }}</span></button>
        <div class="demo__cf-menu__sep" role="presentation" />
        <button class="demo__merge-menu__item" type="button" role="menuitem" data-paste-action="insertCopiedCells" @click="onPasteAction('insertCopiedCells'); closeDropdown()"><RibbonIcon name="paste" /><span>{{ strings.contextMenu.insertCopiedCells }}</span></button>
        <div class="demo__cf-menu__sep" role="presentation" />
        <button class="demo__merge-menu__item" type="button" role="menuitem" data-paste-action="pasteSpecial" @click="onPasteAction('pasteSpecial'); closeDropdown()"><RibbonIcon name="paste" /><span>{{ strings.contextMenu.pasteSpecial }}</span></button>
      </div>
    </div>
    <button class="demo__rb" data-ribbon-command="cut" type="button" :disabled="disabled" :title="tr.cut" :aria-label="tr.cut" :aria-keyshortcuts="keyShortcuts('cut')" @click="dispatchHostClipboard(props.instance, 'cut')">
      <RibbonIcon name="cut" />
    </button>
    <button class="demo__rb" data-ribbon-command="copy" type="button" :disabled="disabled" :title="tr.copy" :aria-label="tr.copy" :aria-keyshortcuts="keyShortcuts('copy')" @click="dispatchHostClipboard(props.instance, 'copy')">
      <RibbonIcon name="copy" />
    </button>
    <button class="demo__rb" data-ribbon-command="formatPainter" :class="{ 'demo__rb--active': active.formatPainterArmed }" type="button" :disabled="disabled" :title="tr.formatPainter" :aria-label="tr.formatPainter" @click="onFormatPainter">
      <RibbonIcon name="paint" />
    </button>
        </div>
        <div class="demo__ribbon-label">{{ tr.clipboard }}</div>
      </section>

      <section class="demo__ribbon-group demo__ribbon-group--font" :aria-label="tr.font">
        <div class="demo__ribbon-tools">
    <div
      class="demo__rb-dd demo__rb-select--font"
      data-dropdown-name="fontFamily"
      data-ribbon-command="fontFamily"
      :class="{ 'demo__rb-dd--open': openDropdown === 'fontFamily' }"
    >
      <button
        type="button"
        class="demo__rb-dd__btn"
        :disabled="disabled"
        :title="tr.font"
        :aria-label="tr.font"
        aria-haspopup="listbox"
        :aria-expanded="openDropdown === 'fontFamily'"
        @click="toggleDropdown('fontFamily')"
      >
        <span class="demo__rb-dd__value">{{ active.fontFamily }}</span>
        <svg class="demo__rb-dd__chev" viewBox="0 0 12 12" aria-hidden="true">
          <path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" />
        </svg>
      </button>
      <div v-if="openDropdown === 'fontFamily'" class="demo__rb-dd__list" role="listbox" :aria-label="tr.font" tabindex="-1">
        <button
          v-for="font in strings.ribbon.fontFamilies"
          :key="font"
          type="button"
          role="option"
          :aria-selected="active.fontFamily === font"
          class="demo__rb-dd__opt"
          :class="{ 'demo__rb-dd__opt--selected': active.fontFamily === font }"
          @click="onDropdownPick('fontFamily', font)"
        >
          <span class="demo__rb-dd__check" aria-hidden="true">
            <svg v-if="active.fontFamily === font" viewBox="0 0 16 16">
              <path d="M3.5 8.5l3 3 6-6.5" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" />
            </svg>
          </span>
          <span class="demo__rb-dd__label">{{ font }}</span>
        </button>
      </div>
    </div>
    <div
      class="demo__rb-dd"
      data-dropdown-name="fontSize"
      data-ribbon-command="fontSize"
      :class="{ 'demo__rb-dd--open': openDropdown === 'fontSize' }"
    >
      <button
        type="button"
        class="demo__rb-dd__btn"
        :disabled="disabled"
        :title="tr.fontSize"
        :aria-label="tr.fontSize"
        aria-haspopup="listbox"
        :aria-expanded="openDropdown === 'fontSize'"
        @click="toggleDropdown('fontSize')"
      >
        <span class="demo__rb-dd__value">{{ active.fontSize }}</span>
        <svg class="demo__rb-dd__chev" viewBox="0 0 12 12" aria-hidden="true">
          <path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" />
        </svg>
      </button>
      <div v-if="openDropdown === 'fontSize'" class="demo__rb-dd__list" role="listbox" :aria-label="tr.fontSize" tabindex="-1">
        <button
          v-for="size in FONT_SIZES"
          :key="size"
          type="button"
          role="option"
          :aria-selected="active.fontSize === size"
          class="demo__rb-dd__opt"
          :class="{ 'demo__rb-dd__opt--selected': active.fontSize === size }"
          @click="onDropdownPick('fontSize', size)"
        >
          <span class="demo__rb-dd__check" aria-hidden="true">
            <svg v-if="active.fontSize === size" viewBox="0 0 16 16">
              <path d="M3.5 8.5l3 3 6-6.5" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" />
            </svg>
          </span>
          <span class="demo__rb-dd__label">{{ size }}</span>
        </button>
      </div>
    </div>
    <button class="demo__rb" data-ribbon-command="fontGrow" type="button" :disabled="disabled" :title="tr.increaseFontSize" :aria-label="tr.increaseFontSize" @click="wrapFormat((s, st) => setFont(s, st, { fontSize: active.fontSize + 1 }))">
      <RibbonIcon name="fontGrow" />
    </button>
    <button class="demo__rb" data-ribbon-command="fontShrink" type="button" :disabled="disabled" :title="tr.decreaseFontSize" :aria-label="tr.decreaseFontSize" @click="wrapFormat((s, st) => setFont(s, st, { fontSize: Math.max(1, active.fontSize - 1) }))">
      <RibbonIcon name="fontShrink" />
    </button>
    <span class="demo__rb-break" data-ribbon-command="font-row-2" aria-hidden="true" />
    <button class="demo__rb" data-ribbon-command="bold" :class="{ 'demo__rb--active': active.bold }" type="button" :disabled="disabled" :title="`${tr.bold} (⌘B)`" :aria-label="`${tr.bold} (⌘B)`" @click="wrapFormat(toggleBold)">
      <RibbonIcon name="bold" />
    </button>
    <button class="demo__rb" data-ribbon-command="italic" :class="{ 'demo__rb--active': active.italic }" type="button" :disabled="disabled" :title="`${tr.italic} (⌘I)`" :aria-label="`${tr.italic} (⌘I)`" @click="wrapFormat(toggleItalic)">
      <RibbonIcon name="italic" />
    </button>
    <button class="demo__rb" data-ribbon-command="underline" :class="{ 'demo__rb--active': active.underline }" type="button" :disabled="disabled" :title="`${tr.underline} (⌘U)`" :aria-label="`${tr.underline} (⌘U)`" @click="wrapFormat(toggleUnderline)">
      <RibbonIcon name="underline" />
    </button>
    <button class="demo__rb" data-ribbon-command="strike" :class="{ 'demo__rb--active': active.strike }" type="button" :disabled="disabled" :title="tr.strikethrough" :aria-label="tr.strikethrough" @click="wrapFormat(toggleStrike)">
      <RibbonIcon name="strike" />
    </button>
    <button class="demo__rb" data-ribbon-command="borders" type="button" :disabled="disabled" :title="tr.borders" :aria-label="tr.borders" @click="wrapFormat(cycleBorders)">
      <RibbonIcon name="borders" />
    </button>
    <div
      class="demo__rb-dd demo__rb-select--border"
      data-ribbon-command="borderPreset"
      data-dropdown-name="borderPreset"
      :class="{ 'demo__rb-dd--open': openDropdown === 'borderPreset' }"
    >
      <button
        type="button"
        class="demo__rb-dd__btn"
        :disabled="disabled"
        :title="tr.borderPattern"
        :aria-label="tr.borderPattern"
        aria-haspopup="listbox"
        :aria-expanded="openDropdown === 'borderPreset'"
        @click="toggleDropdown('borderPreset')"
      >
        <span class="demo__rb-dd__value">{{ tr.outsideBorders }}</span>
        <svg class="demo__rb-dd__chev" viewBox="0 0 12 12" aria-hidden="true">
          <path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" />
        </svg>
      </button>
      <div v-if="openDropdown === 'borderPreset'" class="demo__rb-dd__list" role="listbox" :aria-label="tr.borderPattern" tabindex="-1">
        <button
          v-for="preset in borderPresets"
          :key="preset.value"
          type="button"
          role="option"
          :aria-selected="false"
          class="demo__rb-dd__opt"
          @click="onDropdownPick('borderPreset', preset.value)"
        >
          <span class="demo__rb-dd__check" aria-hidden="true" />
          <span class="demo__rb-dd__label">{{ preset.label }}</span>
        </button>
      </div>
    </div>
    <div
      class="demo__rb-dd demo__rb-select--border-style"
      data-ribbon-command="borderStyle"
      data-dropdown-name="borderStyle"
      :class="{ 'demo__rb-dd--open': openDropdown === 'borderStyle' }"
    >
      <button
        type="button"
        class="demo__rb-dd__btn"
        :disabled="disabled"
        :title="tr.borderLineStyle"
        :aria-label="tr.borderLineStyle"
        aria-haspopup="listbox"
        :aria-expanded="openDropdown === 'borderStyle'"
        @click="toggleDropdown('borderStyle')"
      >
        <span class="demo__rb-dd__value">{{ borderStyles.find((style) => style.value === borderStyle)?.label ?? borderStyle }}</span>
        <svg class="demo__rb-dd__chev" viewBox="0 0 12 12" aria-hidden="true">
          <path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" />
        </svg>
      </button>
      <div v-if="openDropdown === 'borderStyle'" class="demo__rb-dd__list" role="listbox" :aria-label="tr.borderLineStyle" tabindex="-1">
        <button
          v-for="style in borderStyles"
          :key="style.value"
          type="button"
          role="option"
          :aria-selected="borderStyle === style.value"
          :class="['demo__rb-dd__opt', { 'demo__rb-dd__opt--selected': borderStyle === style.value }]"
          @click="onDropdownPick('borderStyle', style.value)"
        >
          <span class="demo__rb-dd__check" aria-hidden="true" />
          <span class="demo__rb-dd__label">{{ style.label }}</span>
        </button>
      </div>
    </div>
    <div class="demo__rb-color" data-ribbon-command="borderColor" :class="{ 'demo__rb-color--open': openDropdown === 'borderColor' }" data-dropdown-name="borderColor" :title="tr.lineColor" :aria-label="tr.lineColor">
      <button type="button" class="demo__rb-color__btn" :disabled="disabled" :aria-label="tr.lineColor" aria-haspopup="menu" :aria-expanded="openDropdown === 'borderColor'" @click="toggleDropdown('borderColor')">
        <span class="demo__rb-color__icon"><RibbonIcon name="fontColor" /></span><span class="demo__rb-color__swatch" :style="{ backgroundColor: borderColor }" /><svg class="demo__rb-color__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
      </button>
      <div v-if="openDropdown === 'borderColor'" ref="borderColorHost" class="demo__color-flyout" />
      <input ref="borderColorNative" class="demo__color-flyout__native" type="color" :value="borderColor" aria-hidden="true" tabindex="-1" @change="onPaletteColor('borderColor', ($event.target as HTMLInputElement).value)" />
    </div>
    <button class="demo__rb" data-ribbon-command="moreBorders" type="button" :disabled="disabled" :title="tr.moreBorders" :aria-label="tr.moreBorders" @click="props.instance?.openFormatDialog('border')">
      <RibbonIcon name="formatCells" />
    </button>
    <button class="demo__rb" data-ribbon-command="drawBorder" type="button" :disabled="!props.onDrawPen && !props.instance?.borderDraw" :title="tr.drawBorder" :aria-label="tr.drawBorder" @click="onDrawPen">
      <RibbonIcon name="pen" />
    </button>
    <button class="demo__rb" data-ribbon-command="drawBorderGrid" type="button" :disabled="!props.instance?.borderDraw" :title="tr.drawBorderGrid" :aria-label="tr.drawBorderGrid" @click="onDrawGrid">
      <RibbonIcon name="borders" />
    </button>
    <button class="demo__rb" data-ribbon-command="eraseBorder" type="button" :disabled="!props.onDrawEraser && !props.instance?.borderDraw" :title="tr.eraseBorder" :aria-label="tr.eraseBorder" @click="onDrawEraser">
      <RibbonIcon name="eraser" />
    </button>
    <div class="demo__rb-color" data-ribbon-command="fontColor" :class="{ 'demo__rb-color--open': openDropdown === 'fontColor' }" data-dropdown-name="fontColor" :title="tr.fontColor" :aria-label="tr.fontColor">
      <button type="button" class="demo__rb-color__btn" :disabled="disabled" :aria-label="tr.fontColor" aria-haspopup="menu" :aria-expanded="openDropdown === 'fontColor'" @click="toggleDropdown('fontColor')">
        <span class="demo__rb-color__icon"><RibbonIcon name="fontColor" /></span><span class="demo__rb-color__swatch" :style="{ backgroundColor: active.fontColor }" /><svg class="demo__rb-color__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
      </button>
      <div v-if="openDropdown === 'fontColor'" ref="fontColorHost" class="demo__color-flyout" />
      <input ref="fontColorNative" class="demo__color-flyout__native" type="color" :value="active.fontColor" aria-hidden="true" tabindex="-1" @change="onPaletteColor('fontColor', ($event.target as HTMLInputElement).value)" />
    </div>
    <div class="demo__rb-color" data-ribbon-command="fillColor" :class="{ 'demo__rb-color--open': openDropdown === 'fillColor' }" data-dropdown-name="fillColor" :title="tr.fillColor" :aria-label="tr.fillColor">
      <button type="button" class="demo__rb-color__btn" :disabled="disabled" :aria-label="tr.fillColor" aria-haspopup="menu" :aria-expanded="openDropdown === 'fillColor'" @click="toggleDropdown('fillColor')">
        <span class="demo__rb-color__icon"><RibbonIcon name="fillColor" /></span><span class="demo__rb-color__swatch" :style="{ backgroundColor: active.fillColor }" /><svg class="demo__rb-color__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
      </button>
      <div v-if="openDropdown === 'fillColor'" ref="fillColorHost" class="demo__color-flyout" />
      <input ref="fillColorNative" class="demo__color-flyout__native" type="color" :value="active.fillColor" aria-hidden="true" tabindex="-1" @change="onPaletteColor('fillColor', ($event.target as HTMLInputElement).value)" />
    </div>
        </div>
        <div class="demo__ribbon-label">{{ tr.font }}</div>
      </section>

      <section class="demo__ribbon-group demo__ribbon-group--alignment" :aria-label="tr.alignment">
        <div class="demo__ribbon-tools">
    <button class="demo__rb" data-ribbon-command="top" :class="{ 'demo__rb--active': active.vAlignTop }" type="button" :disabled="disabled" :title="tr.topAlign" :aria-label="tr.topAlign" @click="wrapFormat((s, st) => setVAlign(s, st, 'top'))">
      <RibbonIcon name="top" />
    </button>
    <button class="demo__rb" data-ribbon-command="middle" :class="{ 'demo__rb--active': active.vAlignMiddle }" type="button" :disabled="disabled" :title="tr.middleAlign" :aria-label="tr.middleAlign" @click="wrapFormat((s, st) => setVAlign(s, st, 'middle'))">
      <RibbonIcon name="middle" />
    </button>
    <button class="demo__rb" data-ribbon-command="bottomAlign" :class="{ 'demo__rb--active': active.vAlignBottom }" type="button" :disabled="disabled" :title="tr.bottomAlign" :aria-label="tr.bottomAlign" @click="wrapFormat((s, st) => setVAlign(s, st, 'bottom'))">
      <RibbonIcon name="bottomAlign" />
    </button>
    <div class="demo__rb-menu" data-ribbon-command="textOrientation" :class="{ 'demo__rb-menu--open': openDropdown === 'textOrientation' }" data-dropdown-name="textOrientation">
      <button class="demo__rb demo__rb-menu__btn" :class="{ 'demo__rb--active': active.textOrientation !== 'horizontalText' }" type="button" :disabled="disabled" :title="tr.textOrientation" :aria-label="tr.textOrientation" aria-haspopup="menu" :aria-expanded="openDropdown === 'textOrientation'" @click="toggleDropdown('textOrientation')">
        <RibbonIcon name="textOrientation" /><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
      </button>
      <div v-if="openDropdown === 'textOrientation'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.textOrientation">
        <button class="demo__merge-menu__item" :class="{ 'demo__rb--active': active.textOrientation === 'angleCounterclockwise' }" :aria-checked="active.textOrientation === 'angleCounterclockwise'" data-cell-action="angleCounterclockwise" type="button" role="menuitemradio" @click="onTextOrientationAction('angleCounterclockwise')"><RibbonIcon name="orientation" /><span>{{ cellText.orientationAngleCounterclockwise }}</span></button>
        <button class="demo__merge-menu__item" :class="{ 'demo__rb--active': active.textOrientation === 'angleClockwise' }" :aria-checked="active.textOrientation === 'angleClockwise'" data-cell-action="angleClockwise" type="button" role="menuitemradio" @click="onTextOrientationAction('angleClockwise')"><RibbonIcon name="orientation" /><span>{{ cellText.orientationAngleClockwise }}</span></button>
        <button class="demo__merge-menu__item" data-cell-action="verticalText" type="button" role="menuitem" @click="onTextOrientationAction('verticalText')"><RibbonIcon name="orientation" /><span>{{ cellText.orientationVerticalText }}</span></button>
        <button class="demo__merge-menu__item" :class="{ 'demo__rb--active': active.textOrientation === 'rotateTextUp' }" :aria-checked="active.textOrientation === 'rotateTextUp'" data-cell-action="rotateTextUp" type="button" role="menuitemradio" @click="onTextOrientationAction('rotateTextUp')"><RibbonIcon name="orientation" /><span>{{ cellText.orientationRotateTextUp }}</span></button>
        <button class="demo__merge-menu__item" :class="{ 'demo__rb--active': active.textOrientation === 'rotateTextDown' }" :aria-checked="active.textOrientation === 'rotateTextDown'" data-cell-action="rotateTextDown" type="button" role="menuitemradio" @click="onTextOrientationAction('rotateTextDown')"><RibbonIcon name="orientation" /><span>{{ cellText.orientationRotateTextDown }}</span></button>
        <div class="demo__cf-menu__sep" role="presentation" />
        <button class="demo__merge-menu__item" :class="{ 'demo__rb--active': active.textOrientation === 'horizontalText' }" :aria-checked="active.textOrientation === 'horizontalText'" data-cell-action="horizontalText" type="button" role="menuitemradio" @click="onTextOrientationAction('horizontalText')"><RibbonIcon name="orientation" /><span>{{ cellText.orientationHorizontalText }}</span></button>
        <div class="demo__cf-menu__sep" role="presentation" />
        <button class="demo__merge-menu__item" data-cell-action="formatAlignment" type="button" role="menuitem" @click="onTextOrientationAction('formatAlignment')"><RibbonIcon name="formatCells" /><span>{{ cellText.orientationFormatAlignment }}</span></button>
      </div>
    </div>
    <button class="demo__rb demo__rb--wide" data-ribbon-command="wrap" :class="{ 'demo__rb--active': active.wrapText }" type="button" :disabled="disabled" :title="tr.wrapText" :aria-label="tr.wrapText" @click="wrapFormat(toggleWrap)">
      <RibbonIcon name="wrap" /><span>{{ tr.wrapText }}</span>
    </button>
    <span class="demo__rb-break" data-ribbon-command="alignment-row-2" aria-hidden="true" />
    <button class="demo__rb" data-ribbon-command="alignL" :class="{ 'demo__rb--active': active.alignLeft }" type="button" :disabled="disabled" :title="tr.alignLeft" :aria-label="tr.alignLeft" @click="onAlign('left')">
      <RibbonIcon name="alignLeft" />
    </button>
    <button class="demo__rb" data-ribbon-command="alignC" :class="{ 'demo__rb--active': active.alignCenter }" type="button" :disabled="disabled" :title="tr.alignCenter" :aria-label="tr.alignCenter" @click="onAlign('center')">
      <RibbonIcon name="alignCenter" />
    </button>
    <button class="demo__rb" data-ribbon-command="alignR" :class="{ 'demo__rb--active': active.alignRight }" type="button" :disabled="disabled" :title="tr.alignRight" :aria-label="tr.alignRight" @click="onAlign('right')">
      <RibbonIcon name="alignRight" />
    </button>
    <button class="demo__rb" data-ribbon-command="indentDecrease" type="button" :disabled="disabled" :title="tr.decreaseIndent" :aria-label="tr.decreaseIndent" @click="wrapFormat((s, st) => bumpIndent(s, st, -1))">
      <RibbonIcon name="indentDecrease" />
    </button>
    <button class="demo__rb" data-ribbon-command="indentIncrease" type="button" :disabled="disabled" :title="tr.increaseIndent" :aria-label="tr.increaseIndent" @click="wrapFormat((s, st) => bumpIndent(s, st, 1))">
      <RibbonIcon name="indentIncrease" />
    </button>
    <div class="demo__rb-menu" data-ribbon-command="merge" :class="{ 'demo__rb-menu--open': openDropdown === 'merge' }" data-dropdown-name="merge">
      <button class="demo__rb demo__rb-menu__btn" :class="{ 'demo__rb--active': active.merged }" type="button" :disabled="disabled" :title="tr.mergeCells" :aria-label="tr.mergeCells" aria-haspopup="menu" :aria-expanded="openDropdown === 'merge'" @click="toggleDropdown('merge')">
        <RibbonIcon name="merge" /><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
      </button>
      <div v-if="openDropdown === 'merge'" class="demo__merge-menu" role="menu" :aria-label="tr.mergeCells">
        <button class="demo__merge-menu__item" :class="{ 'demo__rb--active': active.mergeCenter }" :aria-checked="active.mergeCenter" type="button" role="menuitemradio" @click="onMergeAction('mergeCenter'); closeDropdown()"><RibbonIcon name="merge" /><span>{{ tr.mergeAndCenter }}</span></button>
        <button class="demo__merge-menu__item" type="button" role="menuitem" @click="onMergeAction('mergeAcross'); closeDropdown()"><RibbonIcon name="merge" /><span>{{ tr.mergeAcross }}</span></button>
        <button class="demo__merge-menu__item" :class="{ 'demo__rb--active': active.merged && !active.mergeCenter }" :aria-checked="active.merged && !active.mergeCenter" type="button" role="menuitemradio" @click="onMergeAction('mergeCells'); closeDropdown()"><RibbonIcon name="merge" /><span>{{ tr.mergeCells }}</span></button>
        <button class="demo__merge-menu__item" type="button" role="menuitem" @click="onMergeAction('unmergeCells'); closeDropdown()"><RibbonIcon name="merge" /><span>{{ tr.unmergeCells }}</span></button>
      </div>
    </div>
        </div>
        <div class="demo__ribbon-label">{{ tr.alignment }}</div>
      </section>

      <section class="demo__ribbon-group demo__ribbon-group--number" :aria-label="tr.number">
        <div class="demo__ribbon-tools">
    <div class="demo__rb-dd demo__rb-select--number-format" data-ribbon-command="numberFormat" data-dropdown-name="numberFormat" :class="{ 'demo__rb-dd--open': openDropdown === 'numberFormat' }">
      <button type="button" class="demo__rb-dd__btn" :disabled="disabled" :title="tr.number" :aria-label="tr.number" aria-haspopup="listbox" :aria-expanded="openDropdown === 'numberFormat'" @click="toggleDropdown('numberFormat')">
        <span class="demo__rb-dd__value">{{ [
          { value: 'general', label: tr.general },
          { value: 'fixed', label: tr.fixedNumber },
          { value: 'currency', label: tr.currency },
          { value: 'accounting', label: tr.accounting },
          { value: 'shortDate', label: tr.shortDate },
          { value: 'longDate', label: tr.longDate },
          { value: 'time', label: tr.timeFormat },
          { value: 'percent', label: tr.percent },
          { value: 'fraction', label: tr.fraction },
          { value: 'scientific', label: tr.scientific },
          { value: 'text', label: tr.textFormat },
        ].find((option) => option.value === active.numberFormat)?.label ?? tr.general }}</span>
        <svg class="demo__rb-dd__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
      </button>
      <div v-if="openDropdown === 'numberFormat'" class="demo__rb-dd__list" role="listbox" :aria-label="tr.number" tabindex="-1">
        <button v-for="option in [
          { value: 'general', label: tr.general },
          { value: 'fixed', label: tr.fixedNumber },
          { value: 'currency', label: tr.currency },
          { value: 'accounting', label: tr.accounting },
          { value: 'shortDate', label: tr.shortDate },
          { value: 'longDate', label: tr.longDate },
          { value: 'time', label: tr.timeFormat },
          { value: 'percent', label: tr.percent },
          { value: 'fraction', label: tr.fraction },
          { value: 'scientific', label: tr.scientific },
          { value: 'text', label: tr.textFormat },
          { value: 'more', label: tr.moreNumberFormats },
        ]" :key="option.value" class="demo__rb-dd__opt" :data-fc-value="option.value" type="button" role="option" :aria-selected="active.numberFormat === option.value" @click="onDropdownPick('numberFormat', option.value)">
          <span class="demo__rb-dd__check" aria-hidden="true" />
          <span class="demo__rb-dd__label">{{ option.label }}</span>
        </button>
      </div>
    </div>
    <span class="demo__rb-break" data-ribbon-command="number-row-2" aria-hidden="true" />
    <button class="demo__rb" data-ribbon-command="currency" :class="{ 'demo__rb--active': active.currency }" type="button" :disabled="disabled" :title="tr.currency" :aria-label="tr.currency" @click="wrapFormat((s, st) => cycleCurrency(s, st, lang))">
      <RibbonIcon name="currency" />
    </button>
    <button class="demo__rb" data-ribbon-command="percent" :class="{ 'demo__rb--active': active.percent }" type="button" :disabled="disabled" :title="tr.percent" :aria-label="tr.percent" @click="wrapFormat(cyclePercent)">
      <RibbonIcon name="percent" />
    </button>
    <button class="demo__rb" data-ribbon-command="comma" :class="{ 'demo__rb--active': active.commaStyle }" type="button" :disabled="disabled" :title="tr.commaStyle" :aria-label="tr.commaStyle" @click="wrapFormat((s, st) => setNumFmt(s, st, { kind: 'fixed', decimals: 2, thousands: true }))">
      <RibbonIcon name="comma" />
    </button>
    <button class="demo__rb" data-ribbon-command="decDown" type="button" :disabled="disabled" :title="tr.decreaseDecimals" :aria-label="tr.decreaseDecimals" @click="onBumpDecimals(-1)">
      <RibbonIcon name="decDown" />
    </button>
    <button class="demo__rb" data-ribbon-command="decUp" type="button" :disabled="disabled" :title="tr.increaseDecimals" :aria-label="tr.increaseDecimals" @click="onBumpDecimals(1)">
      <RibbonIcon name="decUp" />
    </button>
        </div>
        <div class="demo__ribbon-label">{{ tr.number }}</div>
      </section>

      <section class="demo__ribbon-group demo__ribbon-group--styles" :aria-label="tr.styles">
        <div class="demo__ribbon-tools">
          <div class="demo__rb-menu demo__cf-menu-wrap" data-ribbon-command="conditional" :class="{ 'demo__rb-menu--open': openDropdown === 'conditional' }" data-dropdown-name="conditional">
            <button class="demo__rb demo__rb-menu__btn demo__rb--wide" :class="{ 'demo__rb--active': active.conditionalFormatting }" type="button" :disabled="disabled" :title="cfText.title" :aria-label="cfText.title" aria-haspopup="menu" :aria-expanded="openDropdown === 'conditional'" @click="toggleDropdown('conditional')">
              <RibbonIcon name="conditional" /><span>{{ cfText.title }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
            </button>
            <div v-if="openDropdown === 'conditional'" class="demo__merge-menu demo__cf-menu" role="menu" :aria-label="cfText.title">
              <div class="demo__cf-menu__submenu" role="none">
                <button type="button" class="demo__merge-menu__item demo__cf-menu__item" role="menuitem"><RibbonIcon name="conditional" /><span>{{ cfText.highlight }}</span><span class="demo__cf-menu__arrow">›</span></button>
                <div class="demo__cf-menu__panel" role="menu">
                  <button class="demo__merge-menu__item demo__cf-menu__item" data-cf-action="cell-greater" type="button" role="menuitem" @click="onConditionalAction('cell-greater')"><RibbonIcon name="conditional" /><span>{{ cfText.greater }}</span></button>
                  <button class="demo__merge-menu__item demo__cf-menu__item" data-cf-action="cell-less" type="button" role="menuitem" @click="onConditionalAction('cell-less')"><RibbonIcon name="conditional" /><span>{{ cfText.less }}</span></button>
                  <button class="demo__merge-menu__item demo__cf-menu__item" data-cf-action="cell-between" type="button" role="menuitem" @click="onConditionalAction('cell-between')"><RibbonIcon name="conditional" /><span>{{ cfText.between }}</span></button>
                  <button class="demo__merge-menu__item demo__cf-menu__item" data-cf-action="cell-equal" type="button" role="menuitem" @click="onConditionalAction('cell-equal')"><RibbonIcon name="conditional" /><span>{{ cfText.equal }}</span></button>
                  <button class="demo__merge-menu__item demo__cf-menu__item" data-cf-action="text-contains" type="button" role="menuitem" @click="onConditionalAction('text-contains')"><RibbonIcon name="conditional" /><span>{{ cfText.textContains }}</span></button>
                  <button class="demo__merge-menu__item demo__cf-menu__item" data-cf-action="date-occurring" type="button" role="menuitem" @click="onConditionalAction('date-occurring')"><RibbonIcon name="conditional" /><span>{{ cfText.dateOccurring }}</span></button>
                  <button class="demo__merge-menu__item demo__cf-menu__item" data-cf-action="duplicates" type="button" role="menuitem" @click="onConditionalAction('duplicates')"><RibbonIcon name="conditional" /><span>{{ cfText.duplicates }}</span></button>
                  <button class="demo__merge-menu__item demo__cf-menu__item" data-cf-action="unique" type="button" role="menuitem" @click="onConditionalAction('unique')"><RibbonIcon name="conditional" /><span>{{ cfText.unique }}</span></button>
                  <button class="demo__merge-menu__item demo__cf-menu__item" data-cf-action="highlight-more" type="button" role="menuitem" @click="onConditionalAction('highlight-more')"><RibbonIcon name="conditional" /><span>{{ cfText.otherRules }}</span></button>
                </div>
              </div>
              <div class="demo__cf-menu__submenu" role="none">
                <button type="button" class="demo__merge-menu__item demo__cf-menu__item" role="menuitem"><RibbonIcon name="conditional" /><span>{{ cfText.topBottom }}</span><span class="demo__cf-menu__arrow">›</span></button>
                <div class="demo__cf-menu__panel" role="menu">
                  <button class="demo__merge-menu__item demo__cf-menu__item" data-cf-action="top10" type="button" role="menuitem" @click="onConditionalAction('top10')"><RibbonIcon name="conditional" /><span>{{ cfText.top10 }}</span></button>
                  <button class="demo__merge-menu__item demo__cf-menu__item" data-cf-action="bottom10" type="button" role="menuitem" @click="onConditionalAction('bottom10')"><RibbonIcon name="conditional" /><span>{{ cfText.bottom10 }}</span></button>
                  <button class="demo__merge-menu__item demo__cf-menu__item" data-cf-action="top10-percent" type="button" role="menuitem" @click="onConditionalAction('top10-percent')"><RibbonIcon name="conditional" /><span>{{ cfText.top10Percent }}</span></button>
                  <button class="demo__merge-menu__item demo__cf-menu__item" data-cf-action="bottom10-percent" type="button" role="menuitem" @click="onConditionalAction('bottom10-percent')"><RibbonIcon name="conditional" /><span>{{ cfText.bottom10Percent }}</span></button>
                  <button class="demo__merge-menu__item demo__cf-menu__item" data-cf-action="above-avg" type="button" role="menuitem" @click="onConditionalAction('above-avg')"><RibbonIcon name="conditional" /><span>{{ cfText.aboveAvg }}</span></button>
                  <button class="demo__merge-menu__item demo__cf-menu__item" data-cf-action="below-avg" type="button" role="menuitem" @click="onConditionalAction('below-avg')"><RibbonIcon name="conditional" /><span>{{ cfText.belowAvg }}</span></button>
                  <button class="demo__merge-menu__item demo__cf-menu__item" data-cf-action="top-bottom-more" type="button" role="menuitem" @click="onConditionalAction('top-bottom-more')"><RibbonIcon name="conditional" /><span>{{ cfText.otherRules }}</span></button>
                </div>
              </div>
              <div class="demo__cf-menu__submenu" role="none">
                <button type="button" class="demo__merge-menu__item demo__cf-menu__item" role="menuitem"><RibbonIcon name="conditional" /><span>{{ cfText.dataBars }}</span><span class="demo__cf-menu__arrow">›</span></button>
                <div class="demo__cf-menu__panel" role="menu">
                  <button v-for="action in ['data-blue', 'data-green', 'data-red', 'data-orange', 'data-purple', 'data-teal', 'data-solid-blue', 'data-solid-green', 'data-solid-red', 'data-solid-orange', 'data-solid-purple', 'data-solid-gray']" :key="action" class="demo__cf-menu__swatch" :data-cf-action="action" type="button" role="menuitem" :title="cfDataBarLabel(action)" :aria-label="cfDataBarLabel(action)" @click="onConditionalAction(action as ConditionalMenuAction)"><span style="background:#fff" /><span :style="{ background: cfDataBarColor(action) }" /><span :style="{ background: action.includes('solid') ? cfDataBarColor(action) : '#fff' }" /></button>
                  <button class="demo__merge-menu__item demo__cf-menu__item" data-cf-action="data-bars-more" type="button" role="menuitem" @click="onConditionalAction('data-bars-more')"><RibbonIcon name="conditional" /><span>{{ cfText.otherRules }}</span></button>
                </div>
              </div>
              <div class="demo__cf-menu__submenu" role="none">
                <button type="button" class="demo__merge-menu__item demo__cf-menu__item" role="menuitem"><RibbonIcon name="conditional" /><span>{{ cfText.colorScales }}</span><span class="demo__cf-menu__arrow">›</span></button>
                <div class="demo__cf-menu__panel" role="menu">
                  <button v-for="action in ['scale-gyr', 'scale-ryg', 'scale-gw', 'scale-rw', 'scale-bwr', 'scale-rwb', 'scale-gwg', 'scale-ywg', 'scale-rwr', 'scale-bwb', 'scale-yry', 'scale-gyg']" :key="action" class="demo__cf-menu__swatch" :data-cf-action="action" type="button" role="menuitem" :title="cfScaleLabel(action)" :aria-label="cfScaleLabel(action)" @click="onConditionalAction(action as ConditionalMenuAction)"><span v-for="color in cfScaleColors(action)" :key="`${action}-${color}`" :style="{ background: color }" /></button>
                  <button class="demo__merge-menu__item demo__cf-menu__item" data-cf-action="color-scales-more" type="button" role="menuitem" @click="onConditionalAction('color-scales-more')"><RibbonIcon name="conditional" /><span>{{ cfText.otherRules }}</span></button>
                </div>
              </div>
              <div class="demo__cf-menu__submenu" role="none">
                <button type="button" class="demo__merge-menu__item demo__cf-menu__item" role="menuitem"><RibbonIcon name="conditional" /><span>{{ cfText.iconSets }}</span><span class="demo__cf-menu__arrow">›</span></button>
                <div class="demo__cf-menu__panel demo__cf-menu__panel--icons" role="menu">
                  <template v-for="group in cfIconSetGroups" :key="group.title">
                    <div class="demo__cf-menu__panel-title" role="presentation">{{ group.title }}</div>
                    <button v-for="item in group.items" :key="item.action" class="demo__cf-menu__iconset" :data-cf-action="item.action" type="button" role="menuitem" :title="cfIconSetLabel(item.action)" :aria-label="cfIconSetLabel(item.action)" @click="onConditionalAction(item.action)">
                      <span v-for="(slot, slotIndex) in item.slots" :key="`${item.action}-${slot}-${slotIndex}`" class="demo__cf-icon" :class="[`demo__cf-icon--${item.family}`, `demo__cf-icon--${slot}`]" />
                    </button>
                  </template>
                  <button class="demo__merge-menu__item demo__cf-menu__item" data-cf-action="icon-sets-more" type="button" role="menuitem" @click="onConditionalAction('icon-sets-more')"><RibbonIcon name="conditional" /><span>{{ cfText.otherRules }}</span></button>
                </div>
              </div>
              <div class="demo__cf-menu__sep" role="separator" />
              <button class="demo__merge-menu__item demo__cf-menu__item" data-cf-action="new-rule" type="button" role="menuitem" @click="onConditionalAction('new-rule')"><RibbonIcon name="conditional" /><span>{{ cfText.newRule }}</span></button>
              <div class="demo__cf-menu__submenu" role="none">
                <button type="button" class="demo__merge-menu__item demo__cf-menu__item" role="menuitem"><RibbonIcon name="conditional" /><span>{{ cfText.clear }}</span><span class="demo__cf-menu__arrow">›</span></button>
                <div class="demo__cf-menu__panel" role="menu">
                  <button class="demo__merge-menu__item demo__cf-menu__item" data-cf-action="clear-selection" type="button" role="menuitem" @click="onConditionalAction('clear-selection')"><RibbonIcon name="conditional" /><span>{{ cfText.clearSelection }}</span></button>
                  <button class="demo__merge-menu__item demo__cf-menu__item" data-cf-action="clear-sheet" type="button" role="menuitem" @click="onConditionalAction('clear-sheet')"><RibbonIcon name="conditional" /><span>{{ cfText.clearSheet }}</span></button>
                </div>
              </div>
              <button class="demo__merge-menu__item demo__cf-menu__item" data-cf-action="manage" type="button" role="menuitem" @click="onConditionalAction('manage')"><RibbonIcon name="conditional" /><span>{{ cfText.manage }}</span></button>
            </div>
          </div>
          <div class="demo__rb-menu" data-ribbon-command="formatTableHome" :class="{ 'demo__rb-menu--open': openDropdown === 'formatTableHome' }" data-dropdown-name="formatTableHome">
            <button class="demo__rb demo__rb-menu__btn demo__rb--wide" :class="{ 'demo__rb--active': active.formatAsTable }" type="button" :disabled="disabled" :title="tr.formatTable" :aria-label="tr.formatTable" aria-haspopup="menu" :aria-expanded="openDropdown === 'formatTableHome'" @click="toggleDropdown('formatTableHome')">
              <RibbonIcon name="tableStyle" /><span>{{ tr.formatTable }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
            </button>
            <div v-if="openDropdown === 'formatTableHome'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.formatTable">
              <button class="demo__merge-menu__item" data-table-style="light" type="button" role="menuitem" @click="onFormatAsTable('light')"><RibbonIcon name="tableStyle" /><span>{{ cellText.tableStyleLight }}</span></button>
              <button class="demo__merge-menu__item" data-table-style="medium" type="button" role="menuitem" @click="onFormatAsTable('medium')"><RibbonIcon name="tableStyle" /><span>{{ cellText.tableStyleMedium }}</span></button>
              <button class="demo__merge-menu__item" data-table-style="dark" type="button" role="menuitem" @click="onFormatAsTable('dark')"><RibbonIcon name="tableStyle" /><span>{{ cellText.tableStyleDark }}</span></button>
            </div>
          </div>
          <div class="demo__rb-menu" data-ribbon-command="cellStyles" :class="{ 'demo__rb-menu--open': openDropdown === 'cellStyles' }" data-dropdown-name="cellStyles">
            <button class="demo__rb demo__rb-menu__btn demo__rb--wide" :class="{ 'demo__rb--active': active.cellStyle !== null }" type="button" :disabled="disabled" :title="tr.cellStyles" :aria-label="tr.cellStyles" aria-haspopup="menu" :aria-expanded="openDropdown === 'cellStyles'" @click="toggleDropdown('cellStyles')">
              <RibbonIcon name="tableStyle" /><span>{{ tr.cellStyles }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
            </button>
            <div v-if="openDropdown === 'cellStyles'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.cellStyles">
              <template v-for="group in cellStyleGroups" :key="group.id">
                <div class="demo__cf-menu__panel-title demo__cell-menu__section" role="presentation">{{ group.label }}</div>
                <button v-for="styleId in group.styleIds" :key="styleId" class="demo__merge-menu__item" :class="{ 'demo__rb--active': active.cellStyle === styleId }" :aria-checked="active.cellStyle === styleId" :data-cell-style="styleId" type="button" role="menuitemradio" @click="onCellStyleAction(styleId as CellStyleAction)">
                  <RibbonIcon name="tableStyle" /><span>{{ cellStyleLabel(styleId as CellStyleId) }}</span>
                </button>
              </template>
            </div>
          </div>
        </div>
        <div class="demo__ribbon-label">{{ tr.styles }}</div>
      </section>

      <section class="demo__ribbon-group demo__ribbon-group--cells" :aria-label="tr.cells">
        <div class="demo__ribbon-tools">
    <div class="demo__rb-menu" data-ribbon-command="insertRows" :class="{ 'demo__rb-menu--open': openDropdown === 'cellInsert' }" data-dropdown-name="cellInsert">
      <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="cellText.insert" :aria-label="cellText.insert" aria-haspopup="menu" :aria-expanded="openDropdown === 'cellInsert'" @click="toggleDropdown('cellInsert')">
        <RibbonIcon name="insertRows" /><span>{{ cellText.insert }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
      </button>
      <div v-if="openDropdown === 'cellInsert'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="cellText.insert">
        <button class="demo__merge-menu__item" data-cell-action="shiftDown" type="button" role="menuitem" @click="onInsertCellsAction('shiftDown')"><RibbonIcon name="insertRows" /><span>{{ cellText.insertShiftDown }}</span></button>
        <button class="demo__merge-menu__item" data-cell-action="shiftRight" type="button" role="menuitem" @click="onInsertCellsAction('shiftRight')"><RibbonIcon name="insertRows" /><span>{{ cellText.insertShiftRight }}</span></button>
        <div class="demo__cf-menu__sep" role="separator" />
        <button class="demo__merge-menu__item" data-cell-action="rows" type="button" role="menuitem" @click="onInsertCellsAction('rows')"><RibbonIcon name="insertRows" /><span>{{ cellText.insertRows }}</span></button>
        <button class="demo__merge-menu__item" data-cell-action="cols" type="button" role="menuitem" @click="onInsertCellsAction('cols')"><RibbonIcon name="insertRows" /><span>{{ cellText.insertCols }}</span></button>
        <div class="demo__cf-menu__sep" role="separator" />
        <button class="demo__merge-menu__item" data-cell-action="sheet" type="button" role="menuitem" @click="onInsertCellsAction('sheet')"><RibbonIcon name="insertRows" /><span>{{ sheetTabsText.insertSheet }}</span></button>
      </div>
    </div>
    <div class="demo__rb-menu" data-ribbon-command="deleteRows" :class="{ 'demo__rb-menu--open': openDropdown === 'cellDelete' }" data-dropdown-name="cellDelete">
      <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="cellText.delete" :aria-label="cellText.delete" aria-haspopup="menu" :aria-expanded="openDropdown === 'cellDelete'" @click="toggleDropdown('cellDelete')">
        <RibbonIcon name="deleteRows" /><span>{{ cellText.delete }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
      </button>
      <div v-if="openDropdown === 'cellDelete'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="cellText.delete">
        <button class="demo__merge-menu__item" data-cell-action="shiftUp" type="button" role="menuitem" @click="onDeleteCellsAction('shiftUp')"><RibbonIcon name="deleteRows" /><span>{{ cellText.deleteShiftUp }}</span></button>
        <button class="demo__merge-menu__item" data-cell-action="shiftLeft" type="button" role="menuitem" @click="onDeleteCellsAction('shiftLeft')"><RibbonIcon name="deleteRows" /><span>{{ cellText.deleteShiftLeft }}</span></button>
        <div class="demo__cf-menu__sep" role="separator" />
        <button class="demo__merge-menu__item" data-cell-action="rows" type="button" role="menuitem" @click="onDeleteCellsAction('rows')"><RibbonIcon name="deleteRows" /><span>{{ cellText.deleteRows }}</span></button>
        <button class="demo__merge-menu__item" data-cell-action="cols" type="button" role="menuitem" @click="onDeleteCellsAction('cols')"><RibbonIcon name="deleteRows" /><span>{{ cellText.deleteCols }}</span></button>
        <div class="demo__cf-menu__sep" role="separator" />
        <button class="demo__merge-menu__item" data-cell-action="sheet" type="button" role="menuitem" @click="onDeleteCellsAction('sheet')"><RibbonIcon name="deleteRows" /><span>{{ sheetTabsText.deleteSheet }}</span></button>
      </div>
    </div>
    <div class="demo__rb-menu" data-ribbon-command="formatCellsHome" :class="{ 'demo__rb-menu--open': openDropdown === 'cellFormat' }" data-dropdown-name="cellFormat">
      <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="cellText.format" :aria-label="cellText.format" :aria-keyshortcuts="keyShortcuts('formatCellsHome')" aria-haspopup="menu" :aria-expanded="openDropdown === 'cellFormat'" @click="toggleDropdown('cellFormat')">
        <RibbonIcon name="formatCells" /><span>{{ cellText.format }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
      </button>
      <div v-if="openDropdown === 'cellFormat'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="cellText.format">
        <button class="demo__merge-menu__item" data-cell-action="dialog" type="button" role="menuitem" @click="onCellFormatAction('dialog')"><RibbonIcon name="formatCells" /><span>{{ cellText.formatCells }}</span></button>
        <div class="demo__cf-menu__sep" role="separator" />
        <button class="demo__merge-menu__item" data-cell-action="rowHeight" type="button" role="menuitem" @click="onCellFormatAction('rowHeight')"><RibbonIcon name="formatCells" /><span>{{ cellText.rowHeight }}</span></button>
        <button class="demo__merge-menu__item" data-cell-action="autoFitRowHeight" type="button" role="menuitem" @click="onCellFormatAction('autoFitRowHeight')"><RibbonIcon name="formatCells" /><span>{{ cellText.autoFitRowHeight }}</span></button>
        <button class="demo__merge-menu__item" data-cell-action="colWidth" type="button" role="menuitem" @click="onCellFormatAction('colWidth')"><RibbonIcon name="formatCells" /><span>{{ cellText.colWidth }}</span></button>
        <button class="demo__merge-menu__item" data-cell-action="autoFitColWidth" type="button" role="menuitem" @click="onCellFormatAction('autoFitColWidth')"><RibbonIcon name="formatCells" /><span>{{ cellText.autoFitColWidth }}</span></button>
        <div class="demo__cf-menu__sep" role="separator" />
        <button class="demo__merge-menu__item" data-cell-action="hideRows" type="button" role="menuitem" @click="onCellFormatAction('hideRows')"><RibbonIcon name="formatCells" /><span>{{ cellText.hideRows }}</span></button>
        <button class="demo__merge-menu__item" data-cell-action="showRows" type="button" role="menuitem" @click="onCellFormatAction('showRows')"><RibbonIcon name="formatCells" /><span>{{ cellText.showRows }}</span></button>
        <button class="demo__merge-menu__item" data-cell-action="hideCols" type="button" role="menuitem" @click="onCellFormatAction('hideCols')"><RibbonIcon name="formatCells" /><span>{{ cellText.hideCols }}</span></button>
        <button class="demo__merge-menu__item" data-cell-action="showCols" type="button" role="menuitem" @click="onCellFormatAction('showCols')"><RibbonIcon name="formatCells" /><span>{{ cellText.showCols }}</span></button>
        <div class="demo__cf-menu__sep" role="separator" />
        <button class="demo__merge-menu__item" data-cell-action="renameSheet" type="button" role="menuitem" @click="onCellFormatAction('renameSheet')"><RibbonIcon name="formatCells" /><span>{{ sheetTabsText.rename }}</span></button>
        <button class="demo__merge-menu__item" data-cell-action="moveSheetLeft" type="button" role="menuitem" @click="onCellFormatAction('moveSheetLeft')"><RibbonIcon name="formatCells" /><span>{{ sheetTabsText.moveLeft }}</span></button>
        <button class="demo__merge-menu__item" data-cell-action="moveSheetRight" type="button" role="menuitem" @click="onCellFormatAction('moveSheetRight')"><RibbonIcon name="formatCells" /><span>{{ sheetTabsText.moveRight }}</span></button>
        <button class="demo__merge-menu__item" data-cell-action="hideSheet" type="button" role="menuitem" @click="onCellFormatAction('hideSheet')"><RibbonIcon name="formatCells" /><span>{{ sheetTabsText.hideSheet }}</span></button>
        <button class="demo__merge-menu__item" data-cell-action="unhideSheet" type="button" role="menuitem" @click="onCellFormatAction('unhideSheet')"><RibbonIcon name="formatCells" /><span>{{ sheetTabsText.unhideSheet }}</span></button>
        <div class="demo__cf-menu__sep" role="separator" />
        <button class="demo__merge-menu__item" data-cell-action="tabColorNone" type="button" role="menuitem" @click="onCellFormatAction('tabColorNone')"><RibbonIcon name="formatCells" /><span>{{ sheetTabsText.tabColor }}: {{ sheetTabsText.noColor }}</span></button>
        <button v-for="entry in SHEET_TAB_COLOR_ACTIONS" :key="entry.action" class="demo__merge-menu__item" :data-cell-action="entry.action" type="button" role="menuitem" @click="onCellFormatAction(entry.action)"><RibbonIcon name="formatCells" /><span>{{ sheetTabsText.tabColor }}: {{ sheetTabColorLabel(entry.action) }}</span></button>
        <div class="demo__cf-menu__sep" role="separator" />
        <button class="demo__merge-menu__item" :class="{ 'demo__rb--active': active.protected }" data-cell-action="protectSheet" type="button" role="menuitemradio" :aria-checked="active.protected" @click="onCellFormatAction('protectSheet')"><RibbonIcon name="protect" /><span>{{ cellText.protectSheet }}</span></button>
      </div>
    </div>
        </div>
        <div class="demo__ribbon-label">{{ tr.cells }}</div>
      </section>

      <section class="demo__ribbon-group demo__ribbon-group--editing" :aria-label="tr.editing">
        <div class="demo__ribbon-tools">
    <div class="demo__rb-menu" data-ribbon-command="autosum" :class="{ 'demo__rb-menu--open': openDropdown === 'autosum' }" data-dropdown-name="autosum">
      <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="`${tr.autoSum} (Σ)`" :aria-label="`${tr.autoSum} (Σ)`" aria-haspopup="menu" :aria-expanded="openDropdown === 'autosum'" @click="toggleDropdown('autosum')">
        <RibbonIcon name="autosum" /><span>{{ tr.autoSum }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
      </button>
      <div v-if="openDropdown === 'autosum'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.autoSum">
        <button class="demo__merge-menu__item" data-autosum-action="SUM" type="button" role="menuitem" @click="onAutoSum('SUM')"><RibbonIcon name="autosum" /><span>{{ cellText.autosumSum }}</span></button>
        <button class="demo__merge-menu__item" data-autosum-action="AVERAGE" type="button" role="menuitem" @click="onAutoSum('AVERAGE')"><RibbonIcon name="autosum" /><span>{{ cellText.autosumAverage }}</span></button>
        <button class="demo__merge-menu__item" data-autosum-action="COUNT" type="button" role="menuitem" @click="onAutoSum('COUNT')"><RibbonIcon name="autosum" /><span>{{ cellText.autosumCount }}</span></button>
        <button class="demo__merge-menu__item" data-autosum-action="MAX" type="button" role="menuitem" @click="onAutoSum('MAX')"><RibbonIcon name="autosum" /><span>{{ cellText.autosumMax }}</span></button>
        <button class="demo__merge-menu__item" data-autosum-action="MIN" type="button" role="menuitem" @click="onAutoSum('MIN')"><RibbonIcon name="autosum" /><span>{{ cellText.autosumMin }}</span></button>
        <div class="demo__cf-menu__sep" role="separator" />
        <button class="demo__merge-menu__item" data-autosum-action="MORE" type="button" role="menuitem" @click="onAutoSum('MORE')"><RibbonIcon name="function" /><span>{{ cellText.autosumMoreFunctions }}</span></button>
      </div>
    </div>
    <div class="demo__rb-menu" data-ribbon-command="fillHome" :class="{ 'demo__rb-menu--open': openDropdown === 'fillHome' }" data-dropdown-name="fillHome">
      <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="cellText.fill" :aria-label="cellText.fill" aria-haspopup="menu" :aria-expanded="openDropdown === 'fillHome'" @click="toggleDropdown('fillHome')">
        <RibbonIcon name="fillColor" /><span>{{ cellText.fill }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
      </button>
      <div v-if="openDropdown === 'fillHome'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="cellText.fill">
        <button class="demo__merge-menu__item" data-editing-action="down" type="button" role="menuitem" @click="onFillAction('down')"><RibbonIcon name="fillColor" /><span>{{ cellText.fillDown }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="right" type="button" role="menuitem" @click="onFillAction('right')"><RibbonIcon name="fillColor" /><span>{{ cellText.fillRight }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="up" type="button" role="menuitem" @click="onFillAction('up')"><RibbonIcon name="fillColor" /><span>{{ cellText.fillUp }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="left" type="button" role="menuitem" @click="onFillAction('left')"><RibbonIcon name="fillColor" /><span>{{ cellText.fillLeft }}</span></button>
        <div class="demo__cf-menu__sep" role="separator" />
        <button class="demo__merge-menu__item" data-editing-action="flash" type="button" role="menuitem" @click="onFillAction('flash')"><RibbonIcon name="fillColor" /><span>{{ cellText.flashFill }}</span></button>
        <div class="demo__cf-menu__sep" role="separator" />
        <button class="demo__merge-menu__item" data-editing-action="series" type="button" role="menuitem" @click="onFillAction('series')"><RibbonIcon name="fillColor" /><span>{{ cellText.series }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="days" type="button" role="menuitem" @click="onFillAction('days')"><RibbonIcon name="fillColor" /><span>{{ cellText.fillDays }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="weekdays" type="button" role="menuitem" @click="onFillAction('weekdays')"><RibbonIcon name="fillColor" /><span>{{ cellText.fillWeekdays }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="months" type="button" role="menuitem" @click="onFillAction('months')"><RibbonIcon name="fillColor" /><span>{{ cellText.fillMonths }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="years" type="button" role="menuitem" @click="onFillAction('years')"><RibbonIcon name="fillColor" /><span>{{ cellText.fillYears }}</span></button>
      </div>
    </div>
    <div class="demo__rb-menu" data-ribbon-command="clearFormat" :class="{ 'demo__rb-menu--open': openDropdown === 'clearHome' }" data-dropdown-name="clearHome">
      <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="cellText.clear" :aria-label="cellText.clear" aria-haspopup="menu" :aria-expanded="openDropdown === 'clearHome'" @click="toggleDropdown('clearHome')">
        <RibbonIcon name="clear" /><span>{{ cellText.clear }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
      </button>
      <div v-if="openDropdown === 'clearHome'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="cellText.clear">
        <button class="demo__merge-menu__item" data-editing-action="all" type="button" role="menuitem" @click="onClearAction('all')"><RibbonIcon name="clear" /><span>{{ cellText.clearAll }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="formats" type="button" role="menuitem" @click="onClearAction('formats')"><RibbonIcon name="clear" /><span>{{ cellText.clearFormats }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="contents" type="button" role="menuitem" @click="onClearAction('contents')"><RibbonIcon name="clear" /><span>{{ cellText.clearContents }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="comments" type="button" role="menuitem" @click="onClearAction('comments')"><RibbonIcon name="clear" /><span>{{ cellText.clearComments }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="hyperlinks" type="button" role="menuitem" @click="onClearAction('hyperlinks')"><RibbonIcon name="clear" /><span>{{ cellText.clearHyperlinks }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="conditional" type="button" role="menuitem" @click="onClearAction('conditional')"><RibbonIcon name="clear" /><span>{{ cellText.clearConditional }}</span></button>
      </div>
    </div>
    <div class="demo__rb-menu" data-ribbon-command="sortFilterHome" :class="{ 'demo__rb-menu--open': openDropdown === 'sortHome' }" data-dropdown-name="sortHome">
      <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="cellText.sortFilter" :aria-label="cellText.sortFilter" aria-haspopup="menu" :aria-expanded="openDropdown === 'sortHome'" @click="toggleDropdown('sortHome')">
        <RibbonIcon name="sortAsc" /><span>{{ cellText.sortFilter }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
      </button>
      <div v-if="openDropdown === 'sortHome'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="cellText.sortFilter">
        <button class="demo__merge-menu__item" data-editing-action="asc" type="button" role="menuitem" @click="onSortMenuAction('asc')"><RibbonIcon name="sortAsc" /><span>{{ cellText.sortAsc }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="desc" type="button" role="menuitem" @click="onSortMenuAction('desc')"><RibbonIcon name="sortAsc" /><span>{{ cellText.sortDesc }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="custom" type="button" role="menuitem" @click="onSortMenuAction('custom')"><RibbonIcon name="sortAsc" /><span>{{ cellText.sortCustom }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="filter" type="button" role="menuitem" @click="onSortMenuAction('filter')"><RibbonIcon name="sortAsc" /><span>{{ cellText.filter }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="filter-clear" type="button" role="menuitem" @click="onSortMenuAction('filter-clear')"><RibbonIcon name="sortAsc" /><span>{{ cellText.clearFilter }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="filter-reapply" type="button" role="menuitem" @click="onSortMenuAction('filter-reapply')"><RibbonIcon name="filter" /><span>{{ cellText.filterReapply }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="filter-by-selected" type="button" role="menuitem" @click="onSortMenuAction('filter-by-selected')"><RibbonIcon name="filter" /><span>{{ cellText.filterBySelectedCellValue }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="filter-advanced" type="button" role="menuitem" @click="onSortMenuAction('filter-advanced')"><RibbonIcon name="filter" /><span>{{ cellText.filterAdvanced }}</span></button>
        <div class="demo__cf-menu__sep" role="separator" />
        <button class="demo__merge-menu__item" data-editing-action="dedupe" type="button" role="menuitem" @click="onSortMenuAction('dedupe')"><RibbonIcon name="sortAsc" /><span>{{ cellText.removeDuplicates }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="conditional" type="button" role="menuitem" @click="onSortMenuAction('conditional')"><RibbonIcon name="sortAsc" /><span>{{ cellText.conditional }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="named" type="button" role="menuitem" @click="onSortMenuAction('named')"><RibbonIcon name="sortAsc" /><span>{{ cellText.namedRanges }}</span></button>
      </div>
    </div>
    <div class="demo__rb-menu" data-ribbon-command="findHome" :class="{ 'demo__rb-menu--open': openDropdown === 'findHome' }" data-dropdown-name="findHome">
      <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="`${cellText.findSelect} (⌘F)`" :aria-label="`${cellText.findSelect} (⌘F)`" :aria-keyshortcuts="keyShortcuts('findHome')" aria-haspopup="menu" :aria-expanded="openDropdown === 'findHome'" @click="toggleDropdown('findHome')">
        <RibbonIcon name="find" /><span>{{ cellText.findSelect }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
      </button>
      <div v-if="openDropdown === 'findHome'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="cellText.findSelect">
        <button class="demo__merge-menu__item" data-editing-action="find" type="button" role="menuitem" @click="onFindAction('find')"><RibbonIcon name="find" /><span>{{ cellText.find }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="replace" type="button" role="menuitem" @click="onFindAction('replace')"><RibbonIcon name="find" /><span>{{ cellText.replace }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="go-to" type="button" role="menuitem" @click="onFindAction('go-to')"><RibbonIcon name="find" /><span>{{ cellText.goTo }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="go-to-special" type="button" role="menuitem" @click="onFindAction('go-to-special')"><RibbonIcon name="find" /><span>{{ cellText.goToSpecial }}</span></button>
        <div class="demo__cf-menu__sep" role="separator" />
        <button class="demo__merge-menu__item" data-editing-action="formulas" type="button" role="menuitem" @click="onFindAction('formulas')"><RibbonIcon name="find" /><span>{{ cellText.findFormulas }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="constants" type="button" role="menuitem" @click="onFindAction('constants')"><RibbonIcon name="find" /><span>{{ cellText.findConstants }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="numbers" type="button" role="menuitem" @click="onFindAction('numbers')"><RibbonIcon name="find" /><span>{{ strings.goToDialog.kindNumbers }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="text" type="button" role="menuitem" @click="onFindAction('text')"><RibbonIcon name="find" /><span>{{ strings.goToDialog.kindText }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="errors" type="button" role="menuitem" @click="onFindAction('errors')"><RibbonIcon name="find" /><span>{{ strings.goToDialog.kindErrors }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="conditional-format" type="button" role="menuitem" @click="onFindAction('conditional-format')"><RibbonIcon name="find" /><span>{{ cellText.findConditionalFormatting }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="data-validation" type="button" role="menuitem" @click="onFindAction('data-validation')"><RibbonIcon name="find" /><span>{{ cellText.findDataValidation }}</span></button>
        <button class="demo__merge-menu__item" data-editing-action="comments" type="button" role="menuitem" @click="onFindAction('comments')"><RibbonIcon name="find" /><span>{{ cellText.comments }}</span></button>
      </div>
    </div>
        </div>
        <div class="demo__ribbon-label">{{ tr.editing }}</div>
      </section>
      </template>

      <template v-else-if="props.activeTab === 'insert'">
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.tables">
          <div class="demo__ribbon-tools">
            <div class="demo__rb-menu" data-ribbon-command="pivotTableInsert" :class="{ 'demo__rb-menu--open': openDropdown === 'pivotTableInsert' }" data-dropdown-name="pivotTableInsert">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.pivotTable" :aria-label="tr.pivotTable" aria-haspopup="menu" :aria-expanded="openDropdown === 'pivotTableInsert'" @click="toggleDropdown('pivotTableInsert')">
                <RibbonIcon name="table" /><span>{{ tr.pivotTable }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'pivotTableInsert'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.pivotTable">
                <button class="demo__merge-menu__item" data-cell-action="dialog" type="button" role="menuitem" @click="onPivotTableAction('dialog'); closeDropdown()"><RibbonIcon name="table" /><span>{{ cellText.pivotTableFromRange }}</span></button>
                <button class="demo__merge-menu__item" data-cell-action="recommended" type="button" role="menuitem" @click="onPivotTableAction('recommended'); closeDropdown()"><RibbonIcon name="table" /><span>{{ cellText.recommendedPivotTables }}</span></button>
                <div class="demo__cf-menu__sep" role="separator" />
                <button class="demo__merge-menu__item" data-cell-action="new-sheet" type="button" role="menuitem" @click="onPivotTableAction('new-sheet'); closeDropdown()"><RibbonIcon name="table" /><span>{{ cellText.pivotTableNewSheet }}</span></button>
                <button class="demo__merge-menu__item" data-cell-action="existing-sheet" type="button" role="menuitem" @click="onPivotTableAction('existing-sheet'); closeDropdown()"><RibbonIcon name="table" /><span>{{ cellText.pivotTableExistingSheet }}</span></button>
              </div>
            </div>
            <div class="demo__rb-menu" data-ribbon-command="formatTableInsert" :class="{ 'demo__rb-menu--open': openDropdown === 'formatTableInsert' }" data-dropdown-name="formatTableInsert">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" :class="{ 'demo__rb--active': active.formatAsTable }" type="button" :disabled="disabled" :title="tr.formatTable" :aria-label="tr.formatTable" aria-haspopup="menu" :aria-expanded="openDropdown === 'formatTableInsert'" @click="toggleDropdown('formatTableInsert')">
                <RibbonIcon name="tableStyle" /><span>{{ tr.formatTable }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'formatTableInsert'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.formatTable">
                <button class="demo__merge-menu__item" data-table-style="light" type="button" role="menuitem" @click="onFormatAsTable('light')"><RibbonIcon name="tableStyle" /><span>{{ cellText.tableStyleLight }}</span></button>
                <button class="demo__merge-menu__item" data-table-style="medium" type="button" role="menuitem" @click="onFormatAsTable('medium')"><RibbonIcon name="tableStyle" /><span>{{ cellText.tableStyleMedium }}</span></button>
                <button class="demo__merge-menu__item" data-table-style="dark" type="button" role="menuitem" @click="onFormatAsTable('dark')"><RibbonIcon name="tableStyle" /><span>{{ cellText.tableStyleDark }}</span></button>
              </div>
            </div>
            <div class="demo__rb-menu" data-ribbon-command="namedRangesInsert" :class="{ 'demo__rb-menu--open': openDropdown === 'definedNamesInsert' }" data-dropdown-name="definedNamesInsert">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.names" :aria-label="tr.names" :aria-keyshortcuts="keyShortcuts('namedRangesInsert')" aria-haspopup="menu" :aria-expanded="openDropdown === 'definedNamesInsert'" @click="toggleDropdown('definedNamesInsert')">
                <RibbonIcon name="names" /><span>{{ tr.names }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'definedNamesInsert'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.names">
                <button class="demo__merge-menu__item" data-defined-name-action="define" type="button" role="menuitem" @click="onDefinedNameAction('define')"><RibbonIcon name="names" /><span>{{ cellText.defineName }}</span></button>
                <button v-for="entry in definedNameEntries" :key="`insert-${entry.name}`" class="demo__merge-menu__item" :data-defined-name-action="`use:${entry.name}`" type="button" role="menuitem" @click="onDefinedNameAction(`use:${entry.name}` as DefinedNameAction)"><RibbonIcon name="names" /><span>{{ cellText.useInFormula }}: {{ entry.name }}</span></button>
                <div class="demo__cf-menu__sep" role="separator" />
                <button class="demo__merge-menu__item" data-defined-name-action="createTopRow" type="button" role="menuitem" @click="onDefinedNameAction('createTopRow')"><RibbonIcon name="names" /><span>{{ cellText.createFromSelectionTop }}</span></button>
                <button class="demo__merge-menu__item" data-defined-name-action="createBottomRow" type="button" role="menuitem" @click="onDefinedNameAction('createBottomRow')"><RibbonIcon name="names" /><span>{{ cellText.createFromSelectionBottom }}</span></button>
                <button class="demo__merge-menu__item" data-defined-name-action="createLeftColumn" type="button" role="menuitem" @click="onDefinedNameAction('createLeftColumn')"><RibbonIcon name="names" /><span>{{ cellText.createFromSelectionLeft }}</span></button>
                <button class="demo__merge-menu__item" data-defined-name-action="createRightColumn" type="button" role="menuitem" @click="onDefinedNameAction('createRightColumn')"><RibbonIcon name="names" /><span>{{ cellText.createFromSelectionRight }}</span></button>
                <div class="demo__cf-menu__sep" role="separator" />
                <button class="demo__merge-menu__item" data-defined-name-action="manager" type="button" role="menuitem" @click="onDefinedNameAction('manager')"><RibbonIcon name="names" /><span>{{ cellText.nameManager }}</span></button>
              </div>
            </div>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="removeDupesInsert" type="button" :disabled="disabled" @click="onRemoveDuplicates">
              <RibbonIcon name="removeDuplicates" /><span>{{ tr.removeDuplicates }}</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ tr.tables }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.illustrations">
          <div class="demo__ribbon-tools">
            <div class="demo__rb-menu" data-ribbon-command="pictureInsert" :class="{ 'demo__rb-menu--open': openDropdown === 'pictureInsert' }" data-dropdown-name="pictureInsert">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.pictures" :aria-label="tr.pictures" aria-haspopup="menu" :aria-expanded="openDropdown === 'pictureInsert'" @click="toggleDropdown('pictureInsert')">
                <RibbonIcon name="page" /><span>{{ tr.pictures }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'pictureInsert'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.pictures">
                <button class="demo__merge-menu__item" data-cell-action="device" type="button" role="menuitem" @click="onIllustrationAction(cellText.pictureThisDevice); closeDropdown()"><RibbonIcon name="page" /><span>{{ cellText.pictureThisDevice }}</span></button>
                <button class="demo__merge-menu__item" data-cell-action="online" type="button" role="menuitem" @click="onIllustrationAction(cellText.pictureOnline); closeDropdown()"><RibbonIcon name="page" /><span>{{ cellText.pictureOnline }}</span></button>
              </div>
            </div>
            <div class="demo__rb-menu" data-ribbon-command="shapesInsert" :class="{ 'demo__rb-menu--open': openDropdown === 'shapesInsert' }" data-dropdown-name="shapesInsert">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.shapes" :aria-label="tr.shapes" aria-haspopup="menu" :aria-expanded="openDropdown === 'shapesInsert'" @click="toggleDropdown('shapesInsert')">
                <RibbonIcon name="options" /><span>{{ tr.shapes }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'shapesInsert'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.shapes">
                <button class="demo__merge-menu__item" data-cell-action="rectangle" type="button" role="menuitem" @click="onIllustrationAction(cellText.shapeRectangle); closeDropdown()"><RibbonIcon name="options" /><span>{{ cellText.shapeRectangle }}</span></button>
                <button class="demo__merge-menu__item" data-cell-action="rounded-rectangle" type="button" role="menuitem" @click="onIllustrationAction(cellText.shapeRoundedRectangle); closeDropdown()"><RibbonIcon name="options" /><span>{{ cellText.shapeRoundedRectangle }}</span></button>
                <button class="demo__merge-menu__item" data-cell-action="oval" type="button" role="menuitem" @click="onIllustrationAction(cellText.shapeOval); closeDropdown()"><RibbonIcon name="options" /><span>{{ cellText.shapeOval }}</span></button>
                <div class="demo__cf-menu__sep" role="separator" />
                <button class="demo__merge-menu__item" data-cell-action="line" type="button" role="menuitem" @click="onIllustrationAction(cellText.shapeLine); closeDropdown()"><RibbonIcon name="options" /><span>{{ cellText.shapeLine }}</span></button>
                <button class="demo__merge-menu__item" data-cell-action="arrow" type="button" role="menuitem" @click="onIllustrationAction(cellText.shapeArrow); closeDropdown()"><RibbonIcon name="options" /><span>{{ cellText.shapeArrow }}</span></button>
              </div>
            </div>
            <div class="demo__rb-menu" data-ribbon-command="screenshotInsert" :class="{ 'demo__rb-menu--open': openDropdown === 'screenshotInsert' }" data-dropdown-name="screenshotInsert">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.screenshot" :aria-label="tr.screenshot" aria-haspopup="menu" :aria-expanded="openDropdown === 'screenshotInsert'" @click="toggleDropdown('screenshotInsert')">
                <RibbonIcon name="page" /><span>{{ tr.screenshot }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'screenshotInsert'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.screenshot">
                <button class="demo__merge-menu__item" data-cell-action="current-view" type="button" role="menuitem" @click="onIllustrationAction(cellText.screenshotCurrentView); closeDropdown()"><RibbonIcon name="page" /><span>{{ cellText.screenshotCurrentView }}</span></button>
                <button class="demo__merge-menu__item" data-cell-action="screen-clipping" type="button" role="menuitem" @click="onIllustrationAction(cellText.screenshotScreenClipping); closeDropdown()"><RibbonIcon name="page" /><span>{{ cellText.screenshotScreenClipping }}</span></button>
              </div>
            </div>
          </div>
          <div class="demo__ribbon-label">{{ tr.illustrations }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.charts">
          <div class="demo__ribbon-tools">
            <div class="demo__rb-menu" data-ribbon-command="chartInsert" :class="{ 'demo__rb-menu--open': openDropdown === 'chartInsert' }" data-dropdown-name="chartInsert">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="cellText.chart" :aria-label="cellText.chart" aria-haspopup="menu" :aria-expanded="openDropdown === 'chartInsert'" @click="toggleDropdown('chartInsert')">
                <RibbonIcon name="chart" /><span>{{ cellText.chart }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'chartInsert'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="cellText.chart">
                <button class="demo__merge-menu__item" data-insert-action="column" type="button" role="menuitem" @click="onCreateChart('column')"><RibbonIcon name="chart" /><span>{{ cellText.chartColumn }}</span></button>
                <button class="demo__merge-menu__item" data-insert-action="bar" type="button" role="menuitem" @click="onCreateChart('bar')"><RibbonIcon name="chart" /><span>{{ cellText.chartBar }}</span></button>
                <button class="demo__merge-menu__item" data-insert-action="line" type="button" role="menuitem" @click="onCreateChart('line')"><RibbonIcon name="chart" /><span>{{ cellText.chartLine }}</span></button>
                <button class="demo__merge-menu__item" data-insert-action="area" type="button" role="menuitem" @click="onCreateChart('area')"><RibbonIcon name="chart" /><span>{{ cellText.chartArea }}</span></button>
                <button class="demo__merge-menu__item" data-insert-action="pie" type="button" role="menuitem" @click="onCreateChart('pie')"><RibbonIcon name="chart" /><span>{{ cellText.chartPie }}</span></button>
                <button class="demo__merge-menu__item" data-insert-action="scatter" type="button" role="menuitem" @click="onCreateChart('scatter')"><RibbonIcon name="chart" /><span>{{ cellText.chartScatter }}</span></button>
                <div class="demo__cf-menu__sep" role="separator" />
                <button class="demo__merge-menu__item" data-insert-action="recommended" type="button" role="menuitem" @click="onCreateChart('recommended')"><RibbonIcon name="chart" /><span>{{ cellText.recommendedCharts }}</span></button>
              </div>
            </div>
          </div>
          <div class="demo__ribbon-label">{{ tr.charts }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.links">
          <div class="demo__ribbon-tools">
            <div class="demo__rb-menu" data-ribbon-command="hyperlinkInsert" :class="{ 'demo__rb-menu--open': openDropdown === 'hyperlinkInsert' }" data-dropdown-name="hyperlinkInsert">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.hyperlink" :aria-label="tr.hyperlink" :aria-keyshortcuts="keyShortcuts('hyperlinkInsert')" aria-haspopup="menu" :aria-expanded="openDropdown === 'hyperlinkInsert'" @click="toggleDropdown('hyperlinkInsert')">
                <RibbonIcon name="link" /><span>{{ tr.hyperlink }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'hyperlinkInsert'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.hyperlink">
                <button class="demo__merge-menu__item" data-link-action="edit" type="button" role="menuitem" @click="onHyperlinkAction('edit')"><RibbonIcon name="link" /><span>{{ cellText.linkInsertOrEdit }}</span></button>
                <button class="demo__merge-menu__item" data-link-action="open" type="button" role="menuitem" @click="onHyperlinkAction('open')"><RibbonIcon name="link" /><span>{{ cellText.linkOpen }}</span></button>
                <button class="demo__merge-menu__item" data-link-action="clear" type="button" role="menuitem" @click="onHyperlinkAction('clear')"><RibbonIcon name="clear" /><span>{{ cellText.linkClear }}</span></button>
                <div class="demo__cf-menu__sep" role="separator" />
                <button class="demo__merge-menu__item" data-link-action="external" type="button" role="menuitem" @click="onHyperlinkAction('external')"><RibbonIcon name="link" /><span>{{ cellText.linkExternalLinks }}</span></button>
              </div>
            </div>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="linksInsert" type="button" :disabled="disabled" @click="props.instance?.openExternalLinksDialog()">
              <RibbonIcon name="link" /><span>{{ tr.links }}</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ tr.links }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.comments">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" data-ribbon-command="commentInsert" :class="{ 'demo__rb--active': active.hasComment }" type="button" :disabled="disabled" @click="props.instance?.openCommentDialog()">
              <RibbonIcon name="comment" /><span>{{ active.hasComment ? tr.editComment : tr.newComment }}</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ tr.comments }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.symbols">
          <div class="demo__ribbon-tools">
            <div class="demo__rb-menu" data-ribbon-command="symbolInsert" :class="{ 'demo__rb-menu--open': openDropdown === 'symbolInsert' }" data-dropdown-name="symbolInsert">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="cellText.symbol" :aria-label="cellText.symbol" aria-haspopup="menu" :aria-expanded="openDropdown === 'symbolInsert'" @click="toggleDropdown('symbolInsert')">
                <RibbonIcon name="function" /><span>{{ cellText.symbol }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'symbolInsert'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="cellText.symbol">
                <template v-for="(symbol, index) in insertSymbols" :key="symbol">
                  <div v-if="[12, 24, 32].includes(index)" class="demo__cf-menu__sep" role="separator" />
                  <button class="demo__merge-menu__item" :data-insert-action="symbol" type="button" role="menuitem" @click="onSymbolAction(symbol)"><RibbonIcon name="function" /><span>{{ symbol }}</span></button>
                </template>
                <div class="demo__cf-menu__sep" role="separator" />
                <button class="demo__merge-menu__item" :data-insert-action="MORE_SYMBOL_ACTION" type="button" role="menuitem" @click="onSymbolAction(MORE_SYMBOL_ACTION)"><RibbonIcon name="function" /><span>{{ cellText.symbolMore }}</span></button>
              </div>
            </div>
          </div>
          <div class="demo__ribbon-label">{{ tr.symbols }}</div>
        </section>
      </template>

      <template v-else-if="props.activeTab === 'draw'">
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="strings.ribbon.tabs.draw">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" data-ribbon-command="drawPen" type="button" :disabled="!props.onDrawPen && !props.instance?.borderDraw" @click="onDrawPen">
              <RibbonIcon name="pen" /><span>{{ tr.pen }}</span>
            </button>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="drawGrid" type="button" :disabled="!props.instance?.borderDraw" @click="onDrawGrid">
              <RibbonIcon name="borders" /><span>{{ tr.drawBorderGrid }}</span>
            </button>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="drawErase" type="button" :disabled="!props.onDrawEraser && !props.instance?.borderDraw" @click="onDrawEraser">
              <RibbonIcon name="eraser" /><span>{{ tr.eraser }}</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ strings.ribbon.tabs.draw }}</div>
        </section>
      </template>

      <template v-else-if="props.activeTab === 'pageLayout'">
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="cellText.theme">
          <div class="demo__ribbon-tools">
            <div class="demo__rb-menu" data-ribbon-command="pageTheme" :class="{ 'demo__rb-menu--open': openDropdown === 'pageTheme' }" data-dropdown-name="pageTheme">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="cellText.theme" :aria-label="cellText.theme" aria-haspopup="menu" :aria-expanded="openDropdown === 'pageTheme'" @click="toggleDropdown('pageTheme')">
                <RibbonIcon name="options" /><span>{{ cellText.theme }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'pageTheme'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="cellText.theme">
                <button class="demo__merge-menu__item" :class="{ 'demo__rb--active': currentTheme === 'paper' }" data-theme-action="paper" type="button" role="menuitemradio" :aria-checked="currentTheme === 'paper'" @click="onThemeAction('paper')"><RibbonIcon name="options" /><span>{{ cellText.themePaper }}</span></button>
                <button class="demo__merge-menu__item" :class="{ 'demo__rb--active': currentTheme === 'ink' }" data-theme-action="ink" type="button" role="menuitemradio" :aria-checked="currentTheme === 'ink'" @click="onThemeAction('ink')"><RibbonIcon name="options" /><span>{{ cellText.themeInk }}</span></button>
                <button class="demo__merge-menu__item" :class="{ 'demo__rb--active': currentTheme === 'contrast' }" data-theme-action="contrast" type="button" role="menuitemradio" :aria-checked="currentTheme === 'contrast'" @click="onThemeAction('contrast')"><RibbonIcon name="options" /><span>{{ cellText.themeContrast }}</span></button>
              </div>
            </div>
          </div>
          <div class="demo__ribbon-label">{{ cellText.theme }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.pageSetup">
          <div class="demo__ribbon-tools">
            <div class="demo__rb-dd demo__rb-select--border" data-ribbon-command="marginsPreset" data-dropdown-name="margins" :class="{ 'demo__rb-dd--open': openDropdown === 'margins' }">
              <button type="button" class="demo__rb-dd__btn" :disabled="disabled" :title="tr.margins" :aria-label="tr.margins" aria-haspopup="listbox" :aria-expanded="openDropdown === 'margins'" @click="toggleDropdown('margins')"><span class="demo__rb-dd__value">{{ active.marginPreset === 'wide' ? tr.marginsWide : active.marginPreset === 'narrow' ? tr.marginsNarrow : active.marginPreset === 'normal' ? tr.marginsNormal : tr.marginsCustom }}</span><svg class="demo__rb-dd__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg></button>
              <div v-if="openDropdown === 'margins'" class="demo__rb-dd__list" role="listbox" :aria-label="tr.margins" tabindex="-1">
                <button class="demo__rb-dd__opt" type="button" role="option" :aria-selected="active.marginPreset === 'normal'" @click="onDropdownPick('margins', 'normal')"><span class="demo__rb-dd__check" aria-hidden="true" /><span class="demo__rb-dd__label">{{ tr.marginsNormal }}</span></button>
                <button class="demo__rb-dd__opt" type="button" role="option" :aria-selected="active.marginPreset === 'wide'" @click="onDropdownPick('margins', 'wide')"><span class="demo__rb-dd__check" aria-hidden="true" /><span class="demo__rb-dd__label">{{ tr.marginsWide }}</span></button>
                <button class="demo__rb-dd__opt" type="button" role="option" :aria-selected="active.marginPreset === 'narrow'" @click="onDropdownPick('margins', 'narrow')"><span class="demo__rb-dd__check" aria-hidden="true" /><span class="demo__rb-dd__label">{{ tr.marginsNarrow }}</span></button>
                <button class="demo__rb-dd__opt" type="button" role="option" :aria-selected="false" @click="onDropdownPick('margins', 'custom')"><span class="demo__rb-dd__check" aria-hidden="true" /><span class="demo__rb-dd__label">{{ tr.marginsCustom }}</span></button>
              </div>
            </div>
            <div class="demo__rb-dd demo__rb-select--border" data-ribbon-command="orientationPreset" data-dropdown-name="orientation" :class="{ 'demo__rb-dd--open': openDropdown === 'orientation' }">
              <button type="button" class="demo__rb-dd__btn" :disabled="disabled" :title="tr.orientation" :aria-label="tr.orientation" aria-haspopup="listbox" :aria-expanded="openDropdown === 'orientation'" @click="toggleDropdown('orientation')"><span class="demo__rb-dd__value">{{ active.pageOrientation === 'landscape' ? tr.landscape : tr.portrait }}</span><svg class="demo__rb-dd__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg></button>
              <div v-if="openDropdown === 'orientation'" class="demo__rb-dd__list" role="listbox" :aria-label="tr.orientation" tabindex="-1">
                <button class="demo__rb-dd__opt" type="button" role="option" :aria-selected="active.pageOrientation === 'portrait'" @click="onDropdownPick('orientation', 'portrait')"><span class="demo__rb-dd__check" aria-hidden="true" /><span class="demo__rb-dd__label">{{ tr.portrait }}</span></button>
                <button class="demo__rb-dd__opt" type="button" role="option" :aria-selected="active.pageOrientation === 'landscape'" @click="onDropdownPick('orientation', 'landscape')"><span class="demo__rb-dd__check" aria-hidden="true" /><span class="demo__rb-dd__label">{{ tr.landscape }}</span></button>
              </div>
            </div>
            <div class="demo__rb-dd demo__rb-select--border" data-ribbon-command="paperSizePreset" data-dropdown-name="paperSize" :class="{ 'demo__rb-dd--open': openDropdown === 'paperSize' }">
              <button type="button" class="demo__rb-dd__btn" :disabled="disabled" :title="tr.paperSize" :aria-label="tr.paperSize" aria-haspopup="listbox" :aria-expanded="openDropdown === 'paperSize'" @click="toggleDropdown('paperSize')"><span class="demo__rb-dd__value">{{ active.paperSize }}</span><svg class="demo__rb-dd__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg></button>
              <div v-if="openDropdown === 'paperSize'" class="demo__rb-dd__list" role="listbox" :aria-label="tr.paperSize" tabindex="-1">
                <button v-for="paper in ['A4', 'A3', 'A5', 'letter', 'legal', 'tabloid']" :key="paper" class="demo__rb-dd__opt" type="button" role="option" :aria-selected="active.paperSize === paper" @click="onDropdownPick('paperSize', paper)"><span class="demo__rb-dd__check" aria-hidden="true" /><span class="demo__rb-dd__label">{{ paper === 'letter' ? tr.paperLetter : paper === 'legal' ? tr.paperLegal : paper === 'tabloid' ? tr.paperTabloid : paper }}</span></button>
              </div>
            </div>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="pageSetupAdvanced" type="button" :disabled="disabled" @click="props.instance?.openPageSetup()"><RibbonIcon name="options" /><span>{{ tr.pageSetup }}</span></button>
            <div class="demo__rb-menu" data-ribbon-command="printArea" :class="{ 'demo__rb-menu--open': openDropdown === 'printArea' }" data-dropdown-name="printArea">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.printArea" :aria-label="tr.printArea" aria-haspopup="menu" :aria-expanded="openDropdown === 'printArea'" @click="toggleDropdown('printArea')">
                <RibbonIcon name="table" /><span>{{ tr.printArea }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'printArea'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.printArea">
                <button class="demo__merge-menu__item" data-page-layout-action="print-area-set" type="button" role="menuitem" @click="onPrintAreaAction('set')"><RibbonIcon name="table" /><span>{{ cellText.printAreaSet }}</span></button>
                <button class="demo__merge-menu__item" data-page-layout-action="print-area-clear" type="button" role="menuitem" @click="onPrintAreaAction('clear')"><RibbonIcon name="table" /><span>{{ cellText.printAreaClear }}</span></button>
              </div>
            </div>
            <div class="demo__rb-menu" data-ribbon-command="pageBreaks" :class="{ 'demo__rb-menu--open': openDropdown === 'pageBreaks' }" data-dropdown-name="pageBreaks">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.breaks" :aria-label="tr.breaks" aria-haspopup="menu" :aria-expanded="openDropdown === 'pageBreaks'" @click="toggleDropdown('pageBreaks')">
                <RibbonIcon name="page" /><span>{{ tr.breaks }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'pageBreaks'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.breaks">
                <button class="demo__merge-menu__item" data-page-break-action="insert-row" type="button" role="menuitem" @click="onPageBreakAction('insert-row')"><RibbonIcon name="page" /><span>{{ cellText.pageBreakInsertRow }}</span></button>
                <button class="demo__merge-menu__item" data-page-break-action="insert-col" type="button" role="menuitem" @click="onPageBreakAction('insert-col')"><RibbonIcon name="page" /><span>{{ cellText.pageBreakInsertCol }}</span></button>
                <div class="demo__cf-menu__sep" role="separator" />
                <button class="demo__merge-menu__item" data-page-break-action="remove-row" type="button" role="menuitem" @click="onPageBreakAction('remove-row')"><RibbonIcon name="page" /><span>{{ cellText.pageBreakRemoveRow }}</span></button>
                <button class="demo__merge-menu__item" data-page-break-action="remove-col" type="button" role="menuitem" @click="onPageBreakAction('remove-col')"><RibbonIcon name="page" /><span>{{ cellText.pageBreakRemoveCol }}</span></button>
                <div class="demo__cf-menu__sep" role="separator" />
                <button class="demo__merge-menu__item" data-page-break-action="reset" type="button" role="menuitem" @click="onPageBreakAction('reset')"><RibbonIcon name="page" /><span>{{ cellText.pageBreakResetAll }}</span></button>
              </div>
            </div>
            <div class="demo__rb-menu" data-ribbon-command="sheetBackground" :class="{ 'demo__rb-menu--open': openDropdown === 'sheetBackground' }" data-dropdown-name="sheetBackground">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.background" :aria-label="tr.background" aria-haspopup="menu" :aria-expanded="openDropdown === 'sheetBackground'" @click="toggleDropdown('sheetBackground')">
                <RibbonIcon name="page" /><span>{{ tr.background }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'sheetBackground'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.background">
                <button class="demo__merge-menu__item" data-sheet-background-action="set" type="button" role="menuitem" @click="onSheetBackgroundAction('set')"><RibbonIcon name="page" /><span>{{ cellText.sheetBackgroundSet }}</span></button>
                <button class="demo__merge-menu__item" data-sheet-background-action="clear" type="button" role="menuitem" @click="onSheetBackgroundAction('clear')"><RibbonIcon name="page" /><span>{{ cellText.sheetBackgroundClear }}</span></button>
              </div>
            </div>
            <div class="demo__rb-menu" data-ribbon-command="printTitles" :class="{ 'demo__rb-menu--open': openDropdown === 'printTitles' }" data-dropdown-name="printTitles">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.printTitles" :aria-label="tr.printTitles" aria-haspopup="menu" :aria-expanded="openDropdown === 'printTitles'" @click="toggleDropdown('printTitles')">
                <RibbonIcon name="table" /><span>{{ tr.printTitles }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'printTitles'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.printTitles">
                <button class="demo__merge-menu__item" data-page-layout-action="print-title-rows" type="button" role="menuitem" @click="onPrintTitleAction('rows')"><RibbonIcon name="table" /><span>{{ cellText.printTitleRowsSet }}</span></button>
                <button class="demo__merge-menu__item" data-page-layout-action="print-title-cols" type="button" role="menuitem" @click="onPrintTitleAction('cols')"><RibbonIcon name="table" /><span>{{ cellText.printTitleColsSet }}</span></button>
                <div class="demo__cf-menu__sep" role="separator" />
                <button class="demo__merge-menu__item" data-page-layout-action="print-titles-clear" type="button" role="menuitem" @click="onPrintTitleAction('clear')"><RibbonIcon name="table" /><span>{{ cellText.printTitlesClear }}</span></button>
              </div>
            </div>
          </div>
          <div class="demo__ribbon-label">{{ tr.pageSetup }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.scale">
          <div class="demo__ribbon-tools">
            <div class="demo__rb-dd demo__rb-select--border" data-ribbon-command="scaleWidth" data-dropdown-name="scaleWidth" :class="{ 'demo__rb-dd--open': openDropdown === 'scaleWidth' }">
              <button type="button" class="demo__rb-dd__btn" :disabled="disabled" :title="pageScaleText.width" :aria-label="pageScaleText.width" aria-haspopup="listbox" :aria-expanded="openDropdown === 'scaleWidth'" @click="toggleDropdown('scaleWidth')"><span class="demo__rb-dd__value">{{ active.fitWidth == null ? pageScaleText.automatic : `${active.fitWidth}` }}</span><svg class="demo__rb-dd__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg></button>
              <div v-if="openDropdown === 'scaleWidth'" class="demo__rb-dd__list" role="listbox" :aria-label="pageScaleText.width" tabindex="-1">
                <button v-for="value in ['0', '1', '2', '3']" :key="`w-${value}`" class="demo__rb-dd__opt" type="button" role="option" :aria-selected="(active.fitWidth == null ? '0' : String(active.fitWidth)) === value" @click="onScaleFit('width', value)"><span class="demo__rb-dd__check" aria-hidden="true" /><span class="demo__rb-dd__label">{{ value === '0' ? pageScaleText.automatic : `${value} ${value === '1' ? pageScaleText.page : pageScaleText.pages}` }}</span></button>
              </div>
            </div>
            <div class="demo__rb-dd demo__rb-select--border" data-ribbon-command="scaleHeight" data-dropdown-name="scaleHeight" :class="{ 'demo__rb-dd--open': openDropdown === 'scaleHeight' }">
              <button type="button" class="demo__rb-dd__btn" :disabled="disabled" :title="pageScaleText.height" :aria-label="pageScaleText.height" aria-haspopup="listbox" :aria-expanded="openDropdown === 'scaleHeight'" @click="toggleDropdown('scaleHeight')"><span class="demo__rb-dd__value">{{ active.fitHeight == null ? pageScaleText.automatic : `${active.fitHeight}` }}</span><svg class="demo__rb-dd__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg></button>
              <div v-if="openDropdown === 'scaleHeight'" class="demo__rb-dd__list" role="listbox" :aria-label="pageScaleText.height" tabindex="-1">
                <button v-for="value in ['0', '1', '2', '3']" :key="`h-${value}`" class="demo__rb-dd__opt" type="button" role="option" :aria-selected="(active.fitHeight == null ? '0' : String(active.fitHeight)) === value" @click="onScaleFit('height', value)"><span class="demo__rb-dd__check" aria-hidden="true" /><span class="demo__rb-dd__label">{{ value === '0' ? pageScaleText.automatic : `${value} ${value === '1' ? pageScaleText.page : pageScaleText.pages}` }}</span></button>
              </div>
            </div>
            <div class="demo__rb-dd demo__rb-select--border" data-ribbon-command="scalePercent" data-dropdown-name="scalePercent" :class="{ 'demo__rb-dd--open': openDropdown === 'scalePercent' }">
              <button type="button" class="demo__rb-dd__btn" :disabled="disabled" :title="pageScaleText.scale" :aria-label="pageScaleText.scale" aria-haspopup="listbox" :aria-expanded="openDropdown === 'scalePercent'" @click="toggleDropdown('scalePercent')"><span class="demo__rb-dd__value">{{ Math.round(active.pageScale * 100) }}%</span><svg class="demo__rb-dd__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg></button>
              <div v-if="openDropdown === 'scalePercent'" class="demo__rb-dd__list" role="listbox" :aria-label="pageScaleText.scale" tabindex="-1">
                <button v-for="value in ['25', '50', '75', '100', '125', '150', '200', '400']" :key="`s-${value}`" class="demo__rb-dd__opt" type="button" role="option" :aria-selected="String(Math.round(active.pageScale * 100)) === value" @click="onScalePercent(value)"><span class="demo__rb-dd__check" aria-hidden="true" /><span class="demo__rb-dd__label">{{ value }}%</span></button>
              </div>
            </div>
          </div>
          <div class="demo__ribbon-label">{{ tr.scale }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.sheetOptions">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" data-ribbon-command="pageLayoutGridlinesView" :class="{ 'demo__rb--active': active.gridlinesVisible }" type="button" :disabled="disabled" @click="onViewFlag('gridlines')"><RibbonIcon name="table" /><span>{{ viewText.gridlines }}</span></button>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="pageLayoutGridlinesPrint" :class="{ 'demo__rb--active': active.printGridlines }" type="button" :disabled="disabled" @click="onPrintSheetOption('gridlines')"><RibbonIcon name="print" /><span>{{ strings.pageSetup.showGridlines }}</span></button>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="pageLayoutHeadingsView" :class="{ 'demo__rb--active': active.headingsVisible }" type="button" :disabled="disabled" @click="onViewFlag('headings')"><RibbonIcon name="table" /><span>{{ viewText.headings }}</span></button>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="pageLayoutHeadingsPrint" :class="{ 'demo__rb--active': active.printHeadings }" type="button" :disabled="disabled" @click="onPrintSheetOption('headings')"><RibbonIcon name="print" /><span>{{ strings.pageSetup.showHeadings }}</span></button>
          </div>
          <div class="demo__ribbon-label">{{ tr.sheetOptions }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.print">
          <div class="demo__ribbon-tools"><button class="demo__rb demo__rb--wide" data-ribbon-command="printPageLayout" type="button" :disabled="disabled" @click="props.instance?.print('print')"><RibbonIcon name="print" /><span>{{ tr.print }}</span></button></div>
          <div class="demo__ribbon-label">{{ tr.print }}</div>
        </section>
      </template>
      <template v-else-if="props.activeTab === 'formulas'">
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.functionLibrary">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" data-ribbon-command="fx" type="button" :disabled="disabled" :aria-keyshortcuts="keyShortcuts('fx')" @click="props.instance?.openFunctionArguments()">
              <RibbonIcon name="function" />
            </button>
            <div class="demo__rb-menu" data-ribbon-command="autosumFormula" :class="{ 'demo__rb-menu--open': openDropdown === 'autosumFormula' }" data-dropdown-name="autosumFormula">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="`${tr.autoSum} (Σ)`" :aria-label="`${tr.autoSum} (Σ)`" aria-haspopup="menu" :aria-expanded="openDropdown === 'autosumFormula'" @click="toggleDropdown('autosumFormula')">
                <RibbonIcon name="autosum" /><span>{{ tr.autoSum }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'autosumFormula'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.autoSum">
                <button class="demo__merge-menu__item" data-autosum-action="SUM" type="button" role="menuitem" @click="onAutoSum('SUM')"><RibbonIcon name="autosum" /><span>{{ cellText.autosumSum }}</span></button>
                <button class="demo__merge-menu__item" data-autosum-action="AVERAGE" type="button" role="menuitem" @click="onAutoSum('AVERAGE')"><RibbonIcon name="autosum" /><span>{{ cellText.autosumAverage }}</span></button>
                <button class="demo__merge-menu__item" data-autosum-action="COUNT" type="button" role="menuitem" @click="onAutoSum('COUNT')"><RibbonIcon name="autosum" /><span>{{ cellText.autosumCount }}</span></button>
                <button class="demo__merge-menu__item" data-autosum-action="MAX" type="button" role="menuitem" @click="onAutoSum('MAX')"><RibbonIcon name="autosum" /><span>{{ cellText.autosumMax }}</span></button>
                <button class="demo__merge-menu__item" data-autosum-action="MIN" type="button" role="menuitem" @click="onAutoSum('MIN')"><RibbonIcon name="autosum" /><span>{{ cellText.autosumMin }}</span></button>
                <div class="demo__cf-menu__sep" role="separator" />
                <button class="demo__merge-menu__item" data-autosum-action="MORE" type="button" role="menuitem" @click="onAutoSum('MORE')"><RibbonIcon name="function" /><span>{{ cellText.autosumMoreFunctions }}</span></button>
              </div>
            </div>
            <button class="demo__rb demo__rb--mono" data-ribbon-command="sum" type="button" :disabled="disabled" @click="props.instance?.openFunctionArguments('SUM')">
              <RibbonIcon name="function" /><span>SUM</span>
            </button>
            <button class="demo__rb demo__rb--mono" data-ribbon-command="avg" type="button" :disabled="disabled" @click="props.instance?.openFunctionArguments('AVERAGE')">
              <RibbonIcon name="function" /><span>AVG</span>
            </button>
            <div class="demo__rb-menu" data-ribbon-command="ifFormula" :class="{ 'demo__rb-menu--open': openDropdown === 'ifFormula' }" data-dropdown-name="ifFormula">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.functionLogical" :aria-label="tr.functionLogical" aria-haspopup="menu" :aria-expanded="openDropdown === 'ifFormula'" @click="toggleDropdown('ifFormula')"><RibbonIcon name="function" /><span>{{ tr.functionLogical }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg></button>
              <div v-if="openDropdown === 'ifFormula'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.functionLogical">
                <button v-for="name in ['IF', 'IFS', 'AND', 'OR']" :key="name" class="demo__merge-menu__item" :data-function-action="name" type="button" role="menuitem" @click="onFunctionAction(name as FunctionAction)"><RibbonIcon name="function" /><span>{{ name }}</span></button>
              </div>
            </div>
            <div class="demo__rb-menu" data-ribbon-command="xlookupFormula" :class="{ 'demo__rb-menu--open': openDropdown === 'xlookupFormula' }" data-dropdown-name="xlookupFormula">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.functionLookupReference" :aria-label="tr.functionLookupReference" aria-haspopup="menu" :aria-expanded="openDropdown === 'xlookupFormula'" @click="toggleDropdown('xlookupFormula')"><RibbonIcon name="function" /><span>{{ tr.functionLookupReference }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg></button>
              <div v-if="openDropdown === 'xlookupFormula'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.functionLookupReference">
                <button v-for="name in ['XLOOKUP', 'VLOOKUP', 'INDEX', 'MATCH']" :key="name" class="demo__merge-menu__item" :data-function-action="name" type="button" role="menuitem" @click="onFunctionAction(name as FunctionAction)"><RibbonIcon name="function" /><span>{{ name }}</span></button>
              </div>
            </div>
            <div class="demo__rb-menu" data-ribbon-command="concatFormula" :class="{ 'demo__rb-menu--open': openDropdown === 'concatFormula' }" data-dropdown-name="concatFormula">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.functionText" :aria-label="tr.functionText" aria-haspopup="menu" :aria-expanded="openDropdown === 'concatFormula'" @click="toggleDropdown('concatFormula')"><RibbonIcon name="function" /><span>{{ tr.functionText }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg></button>
              <div v-if="openDropdown === 'concatFormula'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.functionText">
                <button v-for="name in ['CONCAT', 'TEXT', 'LEFT', 'RIGHT']" :key="name" class="demo__merge-menu__item" :data-function-action="name" type="button" role="menuitem" @click="onFunctionAction(name as FunctionAction)"><RibbonIcon name="function" /><span>{{ name }}</span></button>
              </div>
            </div>
            <div class="demo__rb-menu" data-ribbon-command="todayFormula" :class="{ 'demo__rb-menu--open': openDropdown === 'todayFormula' }" data-dropdown-name="todayFormula">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.functionDateTime" :aria-label="tr.functionDateTime" aria-haspopup="menu" :aria-expanded="openDropdown === 'todayFormula'" @click="toggleDropdown('todayFormula')"><RibbonIcon name="function" /><span>{{ tr.functionDateTime }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg></button>
              <div v-if="openDropdown === 'todayFormula'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.functionDateTime">
                <button v-for="name in ['TODAY', 'NOW', 'DATE', 'YEAR']" :key="name" class="demo__merge-menu__item" :data-function-action="name" type="button" role="menuitem" @click="onFunctionAction(name as FunctionAction)"><RibbonIcon name="function" /><span>{{ name }}</span></button>
              </div>
            </div>
            <div class="demo__rb-menu" data-ribbon-command="pmtFormula" :class="{ 'demo__rb-menu--open': openDropdown === 'pmtFormula' }" data-dropdown-name="pmtFormula">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.functionFinancial" :aria-label="tr.functionFinancial" aria-haspopup="menu" :aria-expanded="openDropdown === 'pmtFormula'" @click="toggleDropdown('pmtFormula')"><RibbonIcon name="function" /><span>{{ tr.functionFinancial }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg></button>
              <div v-if="openDropdown === 'pmtFormula'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.functionFinancial">
                <button v-for="name in ['PMT', 'NPV', 'IRR', 'RATE']" :key="name" class="demo__merge-menu__item" :data-function-action="name" type="button" role="menuitem" @click="onFunctionAction(name as FunctionAction)"><RibbonIcon name="function" /><span>{{ name }}</span></button>
              </div>
            </div>
            <div class="demo__rb-menu" data-ribbon-command="roundFormula" :class="{ 'demo__rb-menu--open': openDropdown === 'roundFormula' }" data-dropdown-name="roundFormula">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.functionMathTrig" :aria-label="tr.functionMathTrig" aria-haspopup="menu" :aria-expanded="openDropdown === 'roundFormula'" @click="toggleDropdown('roundFormula')"><RibbonIcon name="function" /><span>{{ tr.functionMathTrig }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg></button>
              <div v-if="openDropdown === 'roundFormula'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.functionMathTrig">
                <button v-for="name in ['ROUND', 'SUMIF', 'COUNTIF', 'ABS']" :key="name" class="demo__merge-menu__item" :data-function-action="name" type="button" role="menuitem" @click="onFunctionAction(name as FunctionAction)"><RibbonIcon name="function" /><span>{{ name }}</span></button>
              </div>
            </div>
          </div>
          <div class="demo__ribbon-label">{{ tr.functionLibrary }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.definedNames">
          <div class="demo__ribbon-tools">
            <div class="demo__rb-menu" data-ribbon-command="namedRanges" :class="{ 'demo__rb-menu--open': openDropdown === 'definedNames' }" data-dropdown-name="definedNames">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :aria-keyshortcuts="keyShortcuts('namedRanges')" :title="tr.names" :aria-label="tr.names" aria-haspopup="menu" :aria-expanded="openDropdown === 'definedNames'" @click="toggleDropdown('definedNames')">
                <RibbonIcon name="names" /><span>{{ tr.names }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'definedNames'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.names">
                <button class="demo__merge-menu__item" data-defined-name-action="define" type="button" role="menuitem" @click="onDefinedNameAction('define')"><RibbonIcon name="names" /><span>{{ cellText.defineName }}</span></button>
                <button v-for="entry in definedNameEntries" :key="entry.name" class="demo__merge-menu__item" :data-defined-name-action="`use:${entry.name}`" type="button" role="menuitem" @click="onDefinedNameAction(`use:${entry.name}` as DefinedNameAction)"><RibbonIcon name="names" /><span>{{ cellText.useInFormula }}: {{ entry.name }}</span></button>
                <div class="demo__cf-menu__sep" role="separator" />
                <button class="demo__merge-menu__item" data-defined-name-action="createTopRow" type="button" role="menuitem" @click="onDefinedNameAction('createTopRow')"><RibbonIcon name="names" /><span>{{ cellText.createFromSelectionTop }}</span></button>
                <button class="demo__merge-menu__item" data-defined-name-action="createBottomRow" type="button" role="menuitem" @click="onDefinedNameAction('createBottomRow')"><RibbonIcon name="names" /><span>{{ cellText.createFromSelectionBottom }}</span></button>
                <button class="demo__merge-menu__item" data-defined-name-action="createLeftColumn" type="button" role="menuitem" @click="onDefinedNameAction('createLeftColumn')"><RibbonIcon name="names" /><span>{{ cellText.createFromSelectionLeft }}</span></button>
                <button class="demo__merge-menu__item" data-defined-name-action="createRightColumn" type="button" role="menuitem" @click="onDefinedNameAction('createRightColumn')"><RibbonIcon name="names" /><span>{{ cellText.createFromSelectionRight }}</span></button>
                <div class="demo__cf-menu__sep" role="separator" />
                <button class="demo__merge-menu__item" data-defined-name-action="manager" type="button" role="menuitem" @click="onDefinedNameAction('manager')"><RibbonIcon name="names" /><span>{{ cellText.nameManager }}</span></button>
              </div>
            </div>
          </div>
          <div class="demo__ribbon-label">{{ tr.definedNames }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.formulaAuditing">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" data-ribbon-command="precedents" type="button" :disabled="disabled" @click="props.instance?.tracePrecedents()">
              <RibbonIcon name="trace" /><span>{{ tr.tracePrecedents }}</span>
            </button>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="dependents" type="button" :disabled="disabled" @click="props.instance?.traceDependents()">
              <RibbonIcon name="dependents" /><span>{{ tr.traceDependents }}</span>
            </button>
            <div class="demo__rb-menu" data-ribbon-command="clearArrows" :class="{ 'demo__rb-menu--open': openDropdown === 'clearArrows' }" data-dropdown-name="clearArrows">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.removeArrows" :aria-label="tr.removeArrows" aria-haspopup="menu" :aria-expanded="openDropdown === 'clearArrows'" @click="toggleDropdown('clearArrows')">
                <RibbonIcon name="clearArrows" /><span>{{ tr.removeArrows }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'clearArrows'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.removeArrows">
                <button class="demo__merge-menu__item" data-clear-arrows-action="clear-all" type="button" role="menuitem" @click="onClearArrowsAction('clear-all')"><RibbonIcon name="clearArrows" /><span>{{ cellText.removeArrowsAll }}</span></button>
                <button class="demo__merge-menu__item" data-clear-arrows-action="clear-precedents" type="button" role="menuitem" @click="onClearArrowsAction('clear-precedents')"><RibbonIcon name="clearArrows" /><span>{{ cellText.removePrecedentArrows }}</span></button>
                <button class="demo__merge-menu__item" data-clear-arrows-action="clear-dependents" type="button" role="menuitem" @click="onClearArrowsAction('clear-dependents')"><RibbonIcon name="clearArrows" /><span>{{ cellText.removeDependentArrows }}</span></button>
              </div>
            </div>
            <div class="demo__rb-menu" data-ribbon-command="errorChecking" :class="{ 'demo__rb-menu--open': openDropdown === 'errorChecking' }" data-dropdown-name="errorChecking">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.errorChecking" :aria-label="tr.errorChecking" aria-haspopup="menu" :aria-expanded="openDropdown === 'errorChecking'" @click="toggleDropdown('errorChecking')">
                <RibbonIcon name="options" /><span>{{ tr.errorChecking }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'errorChecking'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.errorChecking">
                <button class="demo__merge-menu__item" data-formula-auditing-action="errorChecking" type="button" role="menuitem" @click="onFormulaAuditingAction('errorChecking')"><RibbonIcon name="options" /><span>{{ cellText.errorChecking }}</span></button>
                <button class="demo__merge-menu__item" data-formula-auditing-action="traceError" type="button" role="menuitem" @click="onFormulaAuditingAction('traceError')"><RibbonIcon name="trace" /><span>{{ cellText.traceError }}</span></button>
                <div class="demo__cf-menu__sep" role="separator" />
                <button class="demo__merge-menu__item" data-formula-auditing-action="ignoreError" type="button" role="menuitem" @click="onFormulaAuditingAction('ignoreError')"><RibbonIcon name="options" /><span>{{ cellText.ignoreError }}</span></button>
                <button class="demo__merge-menu__item" data-formula-auditing-action="circleInvalid" type="button" role="menuitem" @click="onFormulaAuditingAction('circleInvalid')"><RibbonIcon name="options" /><span>{{ cellText.validationCircleInvalid }}</span></button>
                <button class="demo__merge-menu__item" data-formula-auditing-action="clearCircles" type="button" role="menuitem" @click="onFormulaAuditingAction('clearCircles')"><RibbonIcon name="options" /><span>{{ cellText.validationClearCircles }}</span></button>
              </div>
            </div>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="showFormulasFormula" :class="{ 'demo__rb--active': active.formulasVisible }" type="button" :disabled="disabled" @click="onViewFlag('formulas')">
              <RibbonIcon name="function" /><span>{{ viewText.formulas }}</span>
            </button>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="evaluateFormula" type="button" :disabled="disabled" @click="props.instance?.openEvaluateFormulaDialog()">
              <RibbonIcon name="function" /><span>{{ tr.evaluateFormula }}</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ tr.formulaAuditing }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.calculation">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" data-ribbon-command="recalcNow" type="button" :disabled="disabled" :aria-keyshortcuts="keyShortcuts('recalcNow')" @click="props.instance?.recalc()">
              <RibbonIcon name="autosum" /><span>{{ tr.recalc }}</span>
            </button>
            <div class="demo__rb-menu" data-ribbon-command="calcOptions" :class="{ 'demo__rb-menu--open': openDropdown === 'calcOptions' }" data-dropdown-name="calcOptions">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.options" :aria-label="tr.options" aria-haspopup="menu" :aria-expanded="openDropdown === 'calcOptions'" @click="toggleDropdown('calcOptions')">
                <RibbonIcon name="options" /><span>{{ tr.options }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'calcOptions'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.options">
                <button class="demo__merge-menu__item" :class="{ 'demo__rb--active': activeCalcAction === 'auto' }" data-calculation-action="auto" type="button" role="menuitemradio" :aria-checked="activeCalcAction === 'auto'" @click="onCalculationAction('auto')"><RibbonIcon name="options" /><span>{{ cellText.calcAutomatic }}</span></button>
                <button class="demo__merge-menu__item" :class="{ 'demo__rb--active': activeCalcAction === 'autoNoTable' }" data-calculation-action="autoNoTable" type="button" role="menuitemradio" :aria-checked="activeCalcAction === 'autoNoTable'" @click="onCalculationAction('autoNoTable')"><RibbonIcon name="options" /><span>{{ cellText.calcAutoNoTable }}</span></button>
                <button class="demo__merge-menu__item" :class="{ 'demo__rb--active': activeCalcAction === 'manual' }" data-calculation-action="manual" type="button" role="menuitemradio" :aria-checked="activeCalcAction === 'manual'" @click="onCalculationAction('manual')"><RibbonIcon name="options" /><span>{{ cellText.calcManual }}</span></button>
                <div class="demo__cf-menu__sep" role="separator" />
                <button class="demo__merge-menu__item" data-calculation-action="iterative" type="button" role="menuitem" @click="onCalculationAction('iterative')"><RibbonIcon name="options" /><span>{{ cellText.calcIterative }}</span></button>
              </div>
            </div>
            <div class="demo__rb-menu" data-ribbon-command="watch" :class="{ 'demo__rb-menu--open': openDropdown === 'watch' }" data-dropdown-name="watch">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.watch" :aria-label="tr.watch" aria-haspopup="menu" :aria-expanded="openDropdown === 'watch'" @click="toggleDropdown('watch')">
                <RibbonIcon name="watch" /><span>{{ tr.watch }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'watch'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.watch">
                <button class="demo__merge-menu__item" data-watch-action="open" type="button" role="menuitem" @click="onWatchAction('open')"><RibbonIcon name="watch" /><span>{{ cellText.watchWindow }}</span></button>
                <button class="demo__merge-menu__item" data-watch-action="add" type="button" role="menuitem" @click="onWatchAction('add')"><RibbonIcon name="watch" /><span>{{ cellText.watchAdd }}</span></button>
                <button class="demo__merge-menu__item" data-watch-action="delete" type="button" role="menuitem" @click="onWatchAction('delete')"><RibbonIcon name="watch" /><span>{{ cellText.watchDelete }}</span></button>
                <div class="demo__cf-menu__sep" role="separator" />
                <button class="demo__merge-menu__item" data-watch-action="delete-all" type="button" role="menuitem" @click="onWatchAction('delete-all')"><RibbonIcon name="watch" /><span>{{ cellText.watchDeleteAll }}</span></button>
              </div>
            </div>
          </div>
          <div class="demo__ribbon-label">{{ tr.calculation }}</div>
        </section>
      </template>

      <template v-else-if="props.activeTab === 'data'">
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.sortFilter">
          <div class="demo__ribbon-tools">
            <div class="demo__rb-menu" data-ribbon-command="filter" :class="{ 'demo__rb-menu--open': openDropdown === 'dataFilter' }" data-dropdown-name="dataFilter">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" :class="{ 'demo__rb--active': active.filterOn }" type="button" :disabled="disabled" :title="tr.filter" :aria-label="tr.filter" aria-haspopup="menu" :aria-expanded="openDropdown === 'dataFilter'" @click="toggleDropdown('dataFilter')">
                <RibbonIcon name="filter" /><span>{{ tr.filter }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'dataFilter'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.filter">
                <button class="demo__merge-menu__item" data-filter-action="toggle" type="button" role="menuitem" @click="onFilterDataAction('toggle')"><RibbonIcon name="filter" /><span>{{ cellText.filterToggle }}</span></button>
                <button class="demo__merge-menu__item" data-filter-action="clear" type="button" role="menuitem" @click="onFilterDataAction('clear')"><RibbonIcon name="filter" /><span>{{ cellText.filterClearAll }}</span></button>
                <button class="demo__merge-menu__item" data-filter-action="reapply" type="button" role="menuitem" @click="onFilterDataAction('reapply')"><RibbonIcon name="filter" /><span>{{ cellText.filterReapply }}</span></button>
                <button class="demo__merge-menu__item" data-filter-action="filter-by-selected" type="button" role="menuitem" @click="onFilterDataAction('filter-by-selected')"><RibbonIcon name="filter" /><span>{{ cellText.filterBySelectedCellValue }}</span></button>
                <div class="demo__cf-menu__sep" role="separator" />
                <button class="demo__merge-menu__item" data-filter-action="advanced" type="button" role="menuitem" @click="onFilterDataAction('advanced')"><RibbonIcon name="filter" /><span>{{ cellText.filterAdvanced }}</span></button>
              </div>
            </div>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="sortAsc" type="button" :disabled="disabled" @click="onSort('asc')">
              <RibbonIcon name="sortAsc" /><span>A-Z</span>
            </button>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="sortDesc" type="button" :disabled="disabled" @click="onSort('desc')">
              <RibbonIcon name="sortDesc" /><span>Z-A</span>
            </button>
            <div class="demo__rb-menu" data-ribbon-command="sortData" :class="{ 'demo__rb-menu--open': openDropdown === 'sortData' }" data-dropdown-name="sortData">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="cellText.sortCustom" :aria-label="cellText.sortCustom" aria-haspopup="menu" :aria-expanded="openDropdown === 'sortData'" @click="toggleDropdown('sortData')">
                <RibbonIcon name="sortAsc" /><span>{{ cellText.sortCustom }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'sortData'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="cellText.sortCustom">
                <button class="demo__merge-menu__item" data-sort-action="custom" type="button" role="menuitem" @click="onSortMenuAction('custom')"><RibbonIcon name="sortAsc" /><span>{{ cellText.sortCustom }}</span></button>
                <div class="demo__cf-menu__sep" role="separator" />
                <button class="demo__merge-menu__item" data-sort-action="asc" type="button" role="menuitem" @click="onSortMenuAction('asc')"><RibbonIcon name="sortAsc" /><span>{{ cellText.sortAsc }}</span></button>
                <button class="demo__merge-menu__item" data-sort-action="desc" type="button" role="menuitem" @click="onSortMenuAction('desc')"><RibbonIcon name="sortDesc" /><span>{{ cellText.sortDesc }}</span></button>
              </div>
            </div>
          </div>
          <div class="demo__ribbon-label">{{ tr.sortFilter }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.dataTools">
          <div class="demo__ribbon-tools">
            <div class="demo__rb-menu" data-ribbon-command="textToColumns" :class="{ 'demo__rb-menu--open': openDropdown === 'textToColumns' }" data-dropdown-name="textToColumns">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="cellText.textToColumns" :aria-label="cellText.textToColumns" aria-haspopup="menu" :aria-expanded="openDropdown === 'textToColumns'" @click="toggleDropdown('textToColumns')">
                <RibbonIcon name="table" /><span>{{ cellText.textToColumns }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'textToColumns'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="cellText.textToColumns">
                <button class="demo__merge-menu__item" data-text-to-columns-action="comma" type="button" role="menuitem" @click="onTextToColumnsAction('comma')"><RibbonIcon name="table" /><span>{{ cellText.textToColumnsComma }}</span></button>
                <button class="demo__merge-menu__item" data-text-to-columns-action="tab" type="button" role="menuitem" @click="onTextToColumnsAction('tab')"><RibbonIcon name="table" /><span>{{ cellText.textToColumnsTab }}</span></button>
                <button class="demo__merge-menu__item" data-text-to-columns-action="semicolon" type="button" role="menuitem" @click="onTextToColumnsAction('semicolon')"><RibbonIcon name="table" /><span>{{ cellText.textToColumnsSemicolon }}</span></button>
                <button class="demo__merge-menu__item" data-text-to-columns-action="space" type="button" role="menuitem" @click="onTextToColumnsAction('space')"><RibbonIcon name="table" /><span>{{ cellText.textToColumnsSpace }}</span></button>
                <div class="demo__cf-menu__sep" role="separator" />
                <button class="demo__merge-menu__item" data-text-to-columns-action="custom" type="button" role="menuitem" @click="onTextToColumnsAction('custom')"><RibbonIcon name="table" /><span>{{ cellText.textToColumnsCustom }}</span></button>
              </div>
            </div>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="removeDupes" type="button" :disabled="disabled" @click="onRemoveDuplicates">
              <RibbonIcon name="removeDuplicates" /><span>{{ tr.removeDuplicates }}</span>
            </button>
            <div class="demo__rb-menu" data-ribbon-command="dataValidation" :class="{ 'demo__rb-menu--open': openDropdown === 'dataValidation' }" data-dropdown-name="dataValidation">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.dataValidation" :aria-label="tr.dataValidation" aria-haspopup="menu" :aria-expanded="openDropdown === 'dataValidation'" @click="toggleDropdown('dataValidation')">
                <RibbonIcon name="options" /><span>{{ tr.dataValidation }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'dataValidation'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.dataValidation">
                <button class="demo__merge-menu__item" data-validation-action="settings" type="button" role="menuitem" @click="onDataValidationAction('settings')"><RibbonIcon name="options" /><span>{{ cellText.validationSettings }}</span></button>
                <button class="demo__merge-menu__item" data-validation-action="circleInvalid" type="button" role="menuitem" @click="onDataValidationAction('circleInvalid')"><RibbonIcon name="options" /><span>{{ cellText.validationCircleInvalid }}</span></button>
                <button class="demo__merge-menu__item" data-validation-action="clearCircles" type="button" role="menuitem" @click="onDataValidationAction('clearCircles')"><RibbonIcon name="options" /><span>{{ cellText.validationClearCircles }}</span></button>
                <div class="demo__cf-menu__sep" role="separator" />
                <button class="demo__merge-menu__item" data-validation-action="clearValidation" type="button" role="menuitem" @click="onDataValidationAction('clearValidation')"><RibbonIcon name="options" /><span>{{ cellText.validationClearRules }}</span></button>
              </div>
            </div>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="linksData" type="button" :disabled="disabled" @click="props.instance?.openExternalLinksDialog()">
              <RibbonIcon name="link" /><span>{{ tr.links }}</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ tr.dataTools }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.outline">
          <div class="demo__ribbon-tools">
            <div class="demo__rb-menu" data-ribbon-command="outlineGroup" :class="{ 'demo__rb-menu--open': openDropdown === 'outlineGroup' }" data-dropdown-name="outlineGroup">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.groupOutline" :aria-label="tr.groupOutline" aria-haspopup="menu" :aria-expanded="openDropdown === 'outlineGroup'" @click="toggleDropdown('outlineGroup')">
                <RibbonIcon name="table" /><span>{{ tr.groupOutline }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'outlineGroup'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.groupOutline">
                <button class="demo__merge-menu__item" data-outline-action="group-rows" type="button" role="menuitem" @click="onOutlineAction('group', 'rows'); closeDropdown()"><RibbonIcon name="table" /><span>{{ strings.contextMenu.rowGroup }}</span></button>
                <button class="demo__merge-menu__item" data-outline-action="group-cols" type="button" role="menuitem" @click="onOutlineAction('group', 'cols'); closeDropdown()"><RibbonIcon name="table" /><span>{{ strings.contextMenu.colGroup }}</span></button>
              </div>
            </div>
            <div class="demo__rb-menu" data-ribbon-command="outlineUngroup" :class="{ 'demo__rb-menu--open': openDropdown === 'outlineUngroup' }" data-dropdown-name="outlineUngroup">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.ungroupOutline" :aria-label="tr.ungroupOutline" aria-haspopup="menu" :aria-expanded="openDropdown === 'outlineUngroup'" @click="toggleDropdown('outlineUngroup')">
                <RibbonIcon name="table" /><span>{{ tr.ungroupOutline }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'outlineUngroup'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.ungroupOutline">
                <button class="demo__merge-menu__item" data-outline-action="ungroup-rows" type="button" role="menuitem" @click="onOutlineAction('ungroup', 'rows'); closeDropdown()"><RibbonIcon name="table" /><span>{{ strings.contextMenu.rowUngroup }}</span></button>
                <button class="demo__merge-menu__item" data-outline-action="ungroup-cols" type="button" role="menuitem" @click="onOutlineAction('ungroup', 'cols'); closeDropdown()"><RibbonIcon name="table" /><span>{{ strings.contextMenu.colUngroup }}</span></button>
              </div>
            </div>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="outlineShowDetail" type="button" :disabled="disabled" @click="onOutlineAction('show-detail')">
              <RibbonIcon name="table" /><span>{{ tr.showDetail }}</span>
            </button>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="outlineHideDetail" type="button" :disabled="disabled" @click="onOutlineAction('hide-detail')">
              <RibbonIcon name="table" /><span>{{ tr.hideDetail }}</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ tr.outline }}</div>
        </section>
      </template>

      <template v-else-if="props.activeTab === 'review'">
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.proofing">
          <div class="demo__ribbon-tools"><button class="demo__rb demo__rb--wide" data-ribbon-command="spellingReview" type="button" :disabled="disabled && !props.onSpellingReview" @click="onSpellingReview"><RibbonIcon name="spelling" /><span>{{ tr.spelling }}</span></button></div>
          <div class="demo__ribbon-label">{{ tr.proofing }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.language">
          <div class="demo__ribbon-tools"><button class="demo__rb demo__rb--wide" data-ribbon-command="translateReview" type="button" :disabled="disabled && !props.onTranslate" @click="onTranslateReview"><RibbonIcon name="translate" /><span>{{ tr.translate }}</span></button></div>
          <div class="demo__ribbon-label">{{ tr.language }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.comments">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" data-ribbon-command="newCommentReview" :class="{ 'demo__rb--active': active.hasComment }" type="button" :disabled="disabled" @click="props.instance?.openCommentDialog()"><RibbonIcon :name="active.hasComment ? 'commentMultiple' : 'commentAdd'" /><span>{{ active.hasComment ? tr.editComment : tr.newComment }}</span></button>
            <div class="demo__rb-menu" data-ribbon-command="deleteCommentReview" :class="{ 'demo__rb-menu--open': openDropdown === 'deleteCommentReview' }" data-dropdown-name="deleteCommentReview">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.deleteComment" :aria-label="tr.deleteComment" aria-haspopup="menu" :aria-expanded="openDropdown === 'deleteCommentReview'" @click="toggleDropdown('deleteCommentReview')"><RibbonIcon name="clear" /><span>{{ tr.deleteComment }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg></button>
              <div v-if="openDropdown === 'deleteCommentReview'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.deleteComment">
                <button class="demo__merge-menu__item" data-comment-action="delete-active" type="button" role="menuitem" @click="onCommentAction('delete-active')"><RibbonIcon name="clear" /><span>{{ cellText.commentDelete }}</span></button>
                <button class="demo__merge-menu__item" data-comment-action="delete-all" type="button" role="menuitem" @click="onCommentAction('delete-all')"><RibbonIcon name="clear" /><span>{{ cellText.commentDeleteAll }}</span></button>
              </div>
            </div>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="previousCommentReview" type="button" :disabled="disabled" @click="onSelectComment(-1)"><RibbonIcon name="goTo" /><span>{{ tr.previousComment }}</span></button>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="nextCommentReview" type="button" :disabled="disabled" @click="onSelectComment(1)"><RibbonIcon name="goTo" /><span>{{ tr.nextComment }}</span></button>
          </div>
          <div class="demo__ribbon-label">{{ tr.comments }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.find">
          <div class="demo__ribbon-tools"><button class="demo__rb demo__rb--wide" data-ribbon-command="findReview" type="button" :disabled="disabled" :aria-keyshortcuts="keyShortcuts('findReview')" @click="props.instance?.openFindReplace('find')"><RibbonIcon name="find" /><span>{{ tr.find }}</span></button></div>
          <div class="demo__ribbon-label">{{ tr.find }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.protection">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" data-ribbon-command="protectReview" :class="{ 'demo__rb--active': active.protected }" type="button" :disabled="disabled" @click="props.instance?.toggleSheetProtection()"><RibbonIcon name="protect" /><span>{{ active.protected ? tr.unprotect : tr.protect }}</span></button>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="protectWorkbookReview" :class="{ 'demo__rb--active': workbookStructureProtected }" type="button" :disabled="disabled" @click="onBackstageProtectWorkbook"><RibbonIcon name="protect" /><span>{{ workbookStructureProtected ? cellText.unprotectWorkbookCommand : cellText.protectWorkbookCommand }}</span></button>
            <div class="demo__rb-menu" data-ribbon-command="protectionReview" :class="{ 'demo__rb-menu--open': openDropdown === 'protectionReview' }" data-dropdown-name="protectionReview">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="cellText.allowEditRangesCommand" :aria-label="cellText.allowEditRangesCommand" aria-haspopup="menu" :aria-expanded="openDropdown === 'protectionReview'" @click="toggleDropdown('protectionReview')"><RibbonIcon name="protect" /><span>{{ cellText.allowEditRangesCommand }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg></button>
              <div v-if="openDropdown === 'protectionReview'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="cellText.allowEditRangesCommand">
                <button class="demo__merge-menu__item" data-protection-action="allow-edit-range" type="button" role="menuitem" @click="onProtectionAction('allow-edit-range')"><RibbonIcon name="protect" /><span>{{ cellText.allowEditRangesCommand }}</span></button>
                <button class="demo__merge-menu__item" data-protection-action="clear-allowed-edit-ranges" type="button" role="menuitem" @click="onProtectionAction('clear-allowed-edit-ranges')"><RibbonIcon name="clear" /><span>{{ cellText.allowEditRangesClearCommand }}</span></button>
              </div>
            </div>
          </div>
          <div class="demo__ribbon-label">{{ tr.protection }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.accessibility">
          <div class="demo__ribbon-tools"><button class="demo__rb demo__rb--wide" data-ribbon-command="accessibility" type="button" :disabled="disabled && !props.onAccessibilityCheck" @click="onAccessibilityReview"><RibbonIcon name="accessibility" /><span>{{ tr.accessibility }}</span></button></div>
          <div class="demo__ribbon-label">{{ tr.accessibility }}</div>
        </section>
      </template>
      <template v-else-if="props.activeTab === 'view'">
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.workbookViews">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" data-ribbon-command="viewNormal" :class="{ 'demo__rb--active': active.workbookView === 'normal' }" type="button" :disabled="disabled" @click="onWorkbookView('normal')"><RibbonIcon name="table" /><span>{{ tr.normalView }}</span></button>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="viewPageLayout" :class="{ 'demo__rb--active': active.workbookView === 'pageLayout' }" type="button" :disabled="disabled" @click="onWorkbookView('pageLayout')"><RibbonIcon name="page" /><span>{{ tr.pageLayoutView }}</span></button>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="viewPageBreakPreview" :class="{ 'demo__rb--active': active.workbookView === 'pageBreakPreview' }" type="button" :disabled="disabled" @click="onWorkbookView('pageBreakPreview')"><RibbonIcon name="table" /><span>{{ tr.pageBreakPreview }}</span></button>
            <div class="demo__rb-menu" data-ribbon-command="watchView" :class="{ 'demo__rb-menu--open': openDropdown === 'watchView' }" data-dropdown-name="watchView">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.watch" :aria-label="tr.watch" aria-haspopup="menu" :aria-expanded="openDropdown === 'watchView'" @click="toggleDropdown('watchView')">
                <RibbonIcon name="watch" /><span>{{ tr.watch }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'watchView'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.watch">
                <button class="demo__merge-menu__item" data-watch-action="open" type="button" role="menuitem" @click="onWatchAction('open')"><RibbonIcon name="watch" /><span>{{ cellText.watchWindow }}</span></button>
                <button class="demo__merge-menu__item" data-watch-action="add" type="button" role="menuitem" @click="onWatchAction('add')"><RibbonIcon name="watch" /><span>{{ cellText.watchAdd }}</span></button>
                <button class="demo__merge-menu__item" data-watch-action="delete" type="button" role="menuitem" @click="onWatchAction('delete')"><RibbonIcon name="watch" /><span>{{ cellText.watchDelete }}</span></button>
                <div class="demo__cf-menu__sep" role="separator" />
                <button class="demo__merge-menu__item" data-watch-action="delete-all" type="button" role="menuitem" @click="onWatchAction('delete-all')"><RibbonIcon name="watch" /><span>{{ cellText.watchDeleteAll }}</span></button>
              </div>
            </div>
          </div>
          <div class="demo__ribbon-label">{{ tr.workbookViews }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="viewToolbarText.views">
          <div class="demo__ribbon-tools">
            <select class="demo__rb-select--border" data-ribbon-command="sheetViewSelect" :disabled="disabled" :value="activeSheetViewId" @change="onSheetViewSelectEvent">
              <option v-for="option in sheetViewOptions" :key="option.value" :value="option.value">{{ option.label }}</option>
            </select>
            <button class="demo__rb" data-ribbon-command="sheetViewSave" type="button" :disabled="disabled" @click="onSheetViewSave"><RibbonIcon name="options" /><span>{{ viewToolbarText.saveView }}</span></button>
            <button class="demo__rb" data-ribbon-command="sheetViewDelete" type="button" :disabled="disabled || activeSheetViewId === 'current'" @click="onSheetViewDelete"><RibbonIcon name="clear" /><span>{{ viewToolbarText.deleteView }}</span></button>
            <button class="demo__rb" data-ribbon-command="workbookObjectsView" type="button" :disabled="disabled" @click="props.instance?.openWorkbookObjects()"><RibbonIcon name="options" /><span>{{ viewToolbarText.objects }}</span></button>
          </div>
          <div class="demo__ribbon-label">{{ viewToolbarText.views }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.show">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" data-ribbon-command="viewGridlines" :class="{ 'demo__rb--active': active.gridlinesVisible }" type="button" :disabled="disabled" @click="onViewFlag('gridlines')"><RibbonIcon name="table" /><span>{{ viewText.gridlines }}</span></button>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="viewHeadings" :class="{ 'demo__rb--active': active.headingsVisible }" type="button" :disabled="disabled" @click="onViewFlag('headings')"><RibbonIcon name="table" /><span>{{ viewText.headings }}</span></button>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="viewFormulas" :class="{ 'demo__rb--active': active.formulasVisible }" type="button" :disabled="disabled" @click="onViewFlag('formulas')"><RibbonIcon name="function" /><span>{{ viewText.formulas }}</span></button>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="viewFormulaBar" :class="{ 'demo__rb--active': formulaBarVisible }" type="button" :disabled="disabled" @click="onToggleFormulaBar"><RibbonIcon name="function" /><span>{{ viewText.formulaBar }}</span></button>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="viewR1C1" :class="{ 'demo__rb--active': active.r1c1 }" type="button" :disabled="disabled" @click="onViewFlag('r1c1')"><RibbonIcon name="options" /><span>R1C1</span></button>
          </div>
          <div class="demo__ribbon-label">{{ tr.show }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.window">
          <div class="demo__ribbon-tools">
            <div class="demo__rb-menu" data-ribbon-command="freeze" :class="{ 'demo__rb-menu--open': openDropdown === 'freeze' }" data-dropdown-name="freeze">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" :class="{ 'demo__rb--active': active.frozen }" type="button" :disabled="disabled" :title="tr.freeze" :aria-label="tr.freeze" aria-haspopup="menu" :aria-expanded="openDropdown === 'freeze'" @click="toggleDropdown('freeze')">
                <RibbonIcon name="freeze" /><span>{{ tr.freeze }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'freeze'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.freeze">
                <button class="demo__merge-menu__item" :class="{ 'demo__rb--active': activeFreezeAction === 'none' }" data-freeze-action="none" type="button" role="menuitemradio" :aria-checked="activeFreezeAction === 'none'" @click="onFreezeAction('none')"><RibbonIcon name="freeze" /><span>{{ viewToolbarText.freezeNone }}</span></button>
                <button class="demo__merge-menu__item" :class="{ 'demo__rb--active': activeFreezeAction === 'panes' }" data-freeze-action="panes" type="button" role="menuitemradio" :aria-checked="activeFreezeAction === 'panes'" @click="onFreezeAction('panes')"><RibbonIcon name="freeze" /><span>{{ viewToolbarText.freezePanes }}</span></button>
                <button class="demo__merge-menu__item" :class="{ 'demo__rb--active': activeFreezeAction === 'topRow' }" data-freeze-action="topRow" type="button" role="menuitemradio" :aria-checked="activeFreezeAction === 'topRow'" @click="onFreezeAction('topRow')"><RibbonIcon name="freeze" /><span>{{ viewToolbarText.freezeTopRow }}</span></button>
                <button class="demo__merge-menu__item" :class="{ 'demo__rb--active': activeFreezeAction === 'firstColumn' }" data-freeze-action="firstColumn" type="button" role="menuitemradio" :aria-checked="activeFreezeAction === 'firstColumn'" @click="onFreezeAction('firstColumn')"><RibbonIcon name="freeze" /><span>{{ viewToolbarText.freezeFirstColumn }}</span></button>
              </div>
            </div>
            <div class="demo__rb-menu" data-ribbon-command="windowVisibility" :class="{ 'demo__rb-menu--open': openDropdown === 'windowVisibility' }" data-dropdown-name="windowVisibility">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="cellText.format" :aria-label="cellText.format" aria-haspopup="menu" :aria-expanded="openDropdown === 'windowVisibility'" @click="toggleDropdown('windowVisibility')">
                <RibbonIcon name="table" /><span>{{ cellText.format }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'windowVisibility'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="cellText.format">
                <button class="demo__merge-menu__item" data-window-action="hideRows" type="button" role="menuitem" @click="onWindowAction('hideRows')"><RibbonIcon name="table" /><span>{{ cellText.hideRows }}</span></button>
                <button class="demo__merge-menu__item" data-window-action="showRows" type="button" role="menuitem" @click="onWindowAction('showRows')"><RibbonIcon name="table" /><span>{{ cellText.showRows }}</span></button>
                <div class="demo__cf-menu__sep" role="separator" />
                <button class="demo__merge-menu__item" data-window-action="hideCols" type="button" role="menuitem" @click="onWindowAction('hideCols')"><RibbonIcon name="table" /><span>{{ cellText.hideCols }}</span></button>
                <button class="demo__merge-menu__item" data-window-action="showCols" type="button" role="menuitem" @click="onWindowAction('showCols')"><RibbonIcon name="table" /><span>{{ cellText.showCols }}</span></button>
              </div>
            </div>
          </div>
          <div class="demo__ribbon-label">{{ tr.window }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.zoom">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" data-ribbon-command="zoomDialog" type="button" :disabled="disabled" @click="openZoomDialog"><RibbonIcon name="zoom" /><span>{{ tr.zoom }}</span></button>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="zoomSelection" type="button" :disabled="disabled" @click="onZoomSelection"><RibbonIcon name="zoom" /><span>{{ tr.zoomSelection }}</span></button>
            <button class="demo__rb demo__rb--mono" data-ribbon-command="zoom75" :class="{ 'demo__rb--active': active.zoom === 0.75 }" type="button" :disabled="disabled" @click="onZoom(0.75)"><RibbonIcon name="zoom" /><span>75%</span></button>
            <button class="demo__rb demo__rb--mono" data-ribbon-command="zoom100" :class="{ 'demo__rb--active': active.zoom === 1 }" type="button" :disabled="disabled" @click="onZoom(1)"><RibbonIcon name="zoom" /><span>100%</span></button>
            <button class="demo__rb demo__rb--mono" data-ribbon-command="zoom125" :class="{ 'demo__rb--active': active.zoom === 1.25 }" type="button" :disabled="disabled" @click="onZoom(1.25)"><RibbonIcon name="zoom" /><span>125%</span></button>
          </div>
          <div class="demo__ribbon-label">{{ tr.zoom }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.protection">
          <div class="demo__ribbon-tools"><button class="demo__rb demo__rb--wide" data-ribbon-command="protect" :class="{ 'demo__rb--active': active.protected }" type="button" :disabled="disabled" @click="props.instance?.toggleSheetProtection()"><RibbonIcon name="protect" /><span>{{ active.protected ? tr.unprotect : tr.protect }}</span></button></div>
          <div class="demo__ribbon-label">{{ tr.protection }}</div>
        </section>
      </template>

      <template v-else-if="props.activeTab === 'automate'">
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="strings.ribbon.tabs.automate">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" data-ribbon-command="script" type="button" :disabled="disabled && !props.onRunScript" @click="onRunScript">
              <RibbonIcon name="script" /><span>{{ tr.script }}</span>
            </button>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="recordActions" type="button" :disabled="disabled" @click="onRecordActions">
              <RibbonIcon name="script" /><span>{{ tr.recordActions }}</span>
            </button>
            <button class="demo__rb demo__rb--wide" data-ribbon-command="allScripts" type="button" @click="onAllScripts">
              <RibbonIcon name="script" /><span>{{ tr.allScripts }}</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ strings.ribbon.tabs.automate }}</div>
        </section>
      </template>

      <template v-else-if="props.activeTab === 'acrobat'">
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.addIn">
          <div class="demo__ribbon-tools">
            <div class="demo__rb-menu" data-ribbon-command="addIn" :class="{ 'demo__rb-menu--open': openDropdown === 'addIn' }" data-dropdown-name="addIn">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled && !props.onAddIn" :title="tr.addIn" :aria-label="tr.addIn" aria-haspopup="menu" :aria-expanded="openDropdown === 'addIn'" @click="toggleDropdown('addIn')">
                <RibbonIcon name="addIn" /><span>{{ tr.addIn }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'addIn'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.addIn">
                <button class="demo__merge-menu__item" data-cell-action="get" type="button" role="menuitem" @click="onAddInAction('get'); closeDropdown()"><RibbonIcon name="addIn" /><span>{{ cellText.addInGet }}</span></button>
                <button class="demo__merge-menu__item" data-cell-action="my" type="button" role="menuitem" @click="onAddInAction('my'); closeDropdown()"><RibbonIcon name="addIn" /><span>{{ cellText.addInMy }}</span></button>
                <div class="demo__cf-menu__sep" role="presentation" />
                <button class="demo__merge-menu__item" data-cell-action="manage" type="button" role="menuitem" @click="onAddInAction('manage'); closeDropdown()"><RibbonIcon name="addIn" /><span>{{ cellText.addInManage }}</span></button>
              </div>
            </div>
          </div>
          <div class="demo__ribbon-label">{{ tr.addIn }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.pdf">
          <div class="demo__ribbon-tools">
            <div class="demo__rb-menu" data-ribbon-command="pdf" :class="{ 'demo__rb-menu--open': openDropdown === 'pdf' }" data-dropdown-name="pdf">
              <button class="demo__rb demo__rb-menu__btn demo__rb--wide" type="button" :disabled="disabled" :title="tr.pdf" :aria-label="tr.pdf" aria-haspopup="menu" :aria-expanded="openDropdown === 'pdf'" @click="toggleDropdown('pdf')">
                <RibbonIcon name="pdf" /><span>{{ tr.pdf }}</span><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
              </button>
              <div v-if="openDropdown === 'pdf'" class="demo__merge-menu demo__cell-menu" role="menu" :aria-label="tr.pdf">
                <button class="demo__merge-menu__item" data-cell-action="create" type="button" role="menuitem" @click="onPdfAction('create'); closeDropdown()"><RibbonIcon name="pdf" /><span>{{ cellText.pdfCreate }}</span></button>
                <button class="demo__merge-menu__item" data-cell-action="share" type="button" role="menuitem" @click="onPdfAction('share'); closeDropdown()"><RibbonIcon name="pdf" /><span>{{ cellText.pdfShare }}</span></button>
                <div class="demo__cf-menu__sep" role="presentation" />
                <button class="demo__merge-menu__item" data-cell-action="preferences" type="button" role="menuitem" @click="onPdfAction('preferences'); closeDropdown()"><RibbonIcon name="pdf" /><span>{{ cellText.pdfPreferences }}</span></button>
              </div>
            </div>
          </div>
          <div class="demo__ribbon-label">{{ tr.pdf }}</div>
        </section>
      </template>

      <template v-else></template>
      <input ref="sheetBackgroundInput" type="file" accept="image/*" hidden data-ribbon-file-input="sheetBackground" @change="onSheetBackgroundFileChange" />
    </div>
    <div v-if="ribbonReportDialog" class="demo__modal" role="presentation">
      <div class="demo__modal-panel demo__modal-panel--narrow" role="dialog" aria-modal="true" :aria-label="ribbonReportDialog.title">
        <header class="demo__modal-header">
          <h2>{{ ribbonReportDialog.title }}</h2>
          <button type="button" class="demo__modal-x" :aria-label="cellText.sortDialogCancel" @click="ribbonReportDialog = null">×</button>
        </header>
        <div class="demo__modal-body demo__sort-dialog">
          <p v-if="ribbonReportDialog.items.length === 0" class="demo__modal-note">{{ strings.reviewReports.noIssues }}</p>
          <div v-else class="demo__report-list">
            <div v-for="(item, index) in ribbonReportDialog.items" :key="`${item.label}-${index}`" class="demo__report-item">
              <strong>{{ item.severity === 'warning' ? strings.reviewReports.warning : strings.reviewReports.info }} - {{ item.label }}</strong>
              <span>{{ item.detail }}</span>
            </div>
          </div>
        </div>
        <footer class="demo__modal-footer">
          <button type="button" class="demo__btn demo__btn--primary" @click="ribbonReportDialog = null">{{ cellText.sortDialogApply }}</button>
        </footer>
      </div>
    </div>
    <div v-if="zoomDialog != null" class="demo__modal" role="presentation">
      <div class="demo__modal-panel demo__modal-panel--narrow" role="dialog" aria-modal="true" :aria-label="tr.zoomDialogTitle">
        <header class="demo__modal-header">
          <h2>{{ tr.zoomDialogTitle }}</h2>
          <button type="button" class="demo__modal-x" :aria-label="cellText.sortDialogCancel" @click="zoomDialog = null">×</button>
        </header>
        <div class="demo__modal-body demo__sort-dialog">
          <label class="demo__modal-field">
            <span>{{ tr.zoomDialogPercent }}</span>
            <input v-model="zoomDialog" type="number" min="10" max="400" />
          </label>
        </div>
        <footer class="demo__modal-footer">
          <button type="button" class="demo__btn" @click="zoomDialog = null">{{ cellText.sortDialogCancel }}</button>
          <button type="button" class="demo__btn demo__btn--primary" @click="applyZoomDialog">{{ cellText.sortDialogApply }}</button>
        </footer>
      </div>
    </div>
    <div v-if="sortDialog" class="demo__modal" role="presentation">
      <div class="demo__modal-panel demo__modal-panel--narrow" role="dialog" aria-modal="true" :aria-label="cellText.sortDialogTitle">
        <header class="demo__modal-header">
          <h2>{{ cellText.sortDialogTitle }}</h2>
          <button type="button" class="demo__modal-x" :aria-label="cellText.sortDialogCancel" @click="sortDialog = null">×</button>
        </header>
        <div class="demo__modal-body demo__sort-dialog">
          <label class="demo__modal-field">
            <span>{{ cellText.sortDialogColumn }}</span>
            <select v-model.number="sortDialog.byCol">
              <option v-for="option in sortColumnOptions" :key="option.value" :value="option.value">{{ option.label }}</option>
            </select>
          </label>
          <label class="demo__modal-field">
            <span>{{ cellText.sortDialogOrder }}</span>
            <select v-model="sortDialog.direction">
              <option value="asc">{{ cellText.sortDialogAscending }}</option>
              <option value="desc">{{ cellText.sortDialogDescending }}</option>
            </select>
          </label>
          <label class="demo__sort-dialog__check">
            <input v-model="sortDialog.hasHeader" type="checkbox" />
            <span>{{ cellText.sortDialogHeader }}</span>
          </label>
        </div>
        <footer class="demo__modal-footer">
          <button type="button" class="demo__btn" @click="sortDialog = null">{{ cellText.sortDialogCancel }}</button>
          <button type="button" class="demo__btn demo__btn--primary" @click="applyCustomSort">{{ cellText.sortDialogApply }}</button>
        </footer>
      </div>
    </div>
    <div v-if="removeDuplicatesDialog" class="demo__modal" role="presentation">
      <div class="demo__modal-panel demo__modal-panel--narrow" role="dialog" aria-modal="true" :aria-label="cellText.removeDuplicatesDialogTitle">
        <header class="demo__modal-header">
          <h2>{{ cellText.removeDuplicatesDialogTitle }}</h2>
          <button type="button" class="demo__modal-x" :aria-label="cellText.sortDialogCancel" @click="removeDuplicatesDialog = null">×</button>
        </header>
        <div class="demo__modal-body demo__sort-dialog">
          <label class="demo__sort-dialog__check">
            <input v-model="removeDuplicatesDialog.hasHeader" type="checkbox" />
            <span>{{ cellText.sortDialogHeader }}</span>
          </label>
          <fieldset class="demo__modal-field">
            <legend>{{ cellText.removeDuplicatesColumns }}</legend>
            <div class="demo__modal-actions">
              <button type="button" class="demo__btn" @click="setRemoveDuplicatesColumns(removeDuplicateColumnOptions.map((option) => option.value))">{{ cellText.removeDuplicatesSelectAll }}</button>
              <button type="button" class="demo__btn" @click="setRemoveDuplicatesColumns([])">{{ cellText.removeDuplicatesUnselectAll }}</button>
            </div>
            <label v-for="option in removeDuplicateColumnOptions" :key="option.value" class="demo__sort-dialog__check">
              <input
                type="checkbox"
                :checked="removeDuplicatesDialog.columns.includes(option.value)"
                @change="toggleRemoveDuplicatesColumn(option.value, ($event.target as HTMLInputElement).checked)"
              />
              <span>{{ option.label }}</span>
            </label>
          </fieldset>
        </div>
        <footer class="demo__modal-footer">
          <button type="button" class="demo__btn" @click="removeDuplicatesDialog = null">{{ cellText.sortDialogCancel }}</button>
          <button type="button" class="demo__btn demo__btn--primary" @click="applyRemoveDuplicatesDialog">{{ cellText.sortDialogApply }}</button>
        </footer>
      </div>
    </div>
    <div v-if="advancedFilterDialog" class="demo__modal" role="presentation">
      <div class="demo__modal-panel demo__modal-panel--narrow" role="dialog" aria-modal="true" :aria-label="cellText.advancedFilterDialogTitle">
        <header class="demo__modal-header">
          <h2>{{ cellText.advancedFilterDialogTitle }}</h2>
          <button type="button" class="demo__modal-x" :aria-label="cellText.sortDialogCancel" @click="advancedFilterDialog = null">×</button>
        </header>
        <div class="demo__modal-body demo__sort-dialog">
          <label class="demo__modal-field">
            <span>{{ cellText.advancedFilterListRange }}</span>
            <input v-model="advancedFilterDialog.listRange" />
          </label>
          <label class="demo__modal-field">
            <span>{{ cellText.advancedFilterCriteriaRange }}</span>
            <input v-model="advancedFilterDialog.criteriaRange" />
          </label>
          <label class="demo__modal-field">
            <span>{{ cellText.advancedFilterCopyTo }}</span>
            <input v-model="advancedFilterDialog.copyTo" />
          </label>
          <label class="demo__sort-dialog__check">
            <input v-model="advancedFilterDialog.uniqueOnly" type="checkbox" />
            <span>{{ cellText.advancedFilterUniqueOnly }}</span>
          </label>
        </div>
        <footer class="demo__modal-footer">
          <button type="button" class="demo__btn" @click="advancedFilterDialog = null">{{ cellText.sortDialogCancel }}</button>
          <button type="button" class="demo__btn demo__btn--primary" @click="applyAdvancedFilterDialog">{{ cellText.sortDialogApply }}</button>
        </footer>
      </div>
    </div>
    <div v-if="dimensionDialog" class="demo__modal" role="presentation">
      <div class="demo__modal-panel demo__modal-panel--narrow" role="dialog" aria-modal="true" :aria-label="dimensionDialog.kind === 'rowHeight' ? cellText.rowHeight : cellText.colWidth">
        <header class="demo__modal-header">
          <h2>{{ dimensionDialog.kind === 'rowHeight' ? cellText.rowHeight : cellText.colWidth }}</h2>
          <button type="button" class="demo__modal-x" :aria-label="cellText.sortDialogCancel" @click="dimensionDialog = null">×</button>
        </header>
        <div class="demo__modal-body demo__sort-dialog">
          <label class="demo__modal-field">
            <span>{{ dimensionDialog.kind === 'rowHeight' ? cellText.rowHeightPrompt : cellText.colWidthPrompt }}</span>
            <input v-model="dimensionDialog.value" type="number" min="1" step="1" />
          </label>
        </div>
        <footer class="demo__modal-footer">
          <button type="button" class="demo__btn" @click="dimensionDialog = null">{{ cellText.sortDialogCancel }}</button>
          <button type="button" class="demo__btn demo__btn--primary" @click="applyDimensionDialog">{{ cellText.sortDialogApply }}</button>
        </footer>
      </div>
    </div>
    <div v-if="sheetRenameDialog" class="demo__modal" role="presentation">
      <div class="demo__modal-panel demo__modal-panel--narrow" role="dialog" aria-modal="true" :aria-label="sheetTabsText.rename">
        <header class="demo__modal-header">
          <h2>{{ sheetTabsText.rename }}</h2>
          <button type="button" class="demo__modal-x" :aria-label="cellText.sortDialogCancel" @click="sheetRenameDialog = null">×</button>
        </header>
        <div class="demo__modal-body demo__sort-dialog">
          <label class="demo__modal-field">
            <span>{{ sheetTabsText.renameSheet.replace('{name}', sheetRenameDialog.value) }}</span>
            <input v-model="sheetRenameDialog.value" />
          </label>
        </div>
        <footer class="demo__modal-footer">
          <button type="button" class="demo__btn" @click="sheetRenameDialog = null">{{ cellText.sortDialogCancel }}</button>
          <button type="button" class="demo__btn demo__btn--primary" @click="applySheetRenameDialog">{{ cellText.sortDialogApply }}</button>
        </footer>
      </div>
    </div>
    <div v-if="scriptDialog" class="demo__modal" role="presentation">
      <div class="demo__modal-panel demo__modal-panel--narrow" role="dialog" aria-modal="true" :aria-label="cellText.scriptDialogTitle">
        <header class="demo__modal-header">
          <h2>{{ cellText.scriptDialogTitle }}</h2>
          <button type="button" class="demo__modal-x" :aria-label="cellText.sortDialogCancel" @click="scriptDialog = null">×</button>
        </header>
        <div class="demo__modal-body demo__sort-dialog">
          <label class="demo__modal-field">
            <span>{{ cellText.scriptDialogCommand }}</span>
            <select v-model="scriptDialog.command">
              <option v-for="option in scriptOptions" :key="option.value" :value="option.value">{{ option.label }}</option>
            </select>
          </label>
        </div>
        <footer class="demo__modal-footer">
          <button type="button" class="demo__btn" @click="scriptDialog = null">{{ cellText.sortDialogCancel }}</button>
          <button type="button" class="demo__btn demo__btn--primary" @click="applyScriptDialog">{{ cellText.sortDialogApply }}</button>
        </footer>
      </div>
    </div>
    <div v-if="textToColumnsDialog" class="demo__modal" role="presentation">
      <div class="demo__modal-panel demo__modal-panel--narrow" role="dialog" aria-modal="true" :aria-label="cellText.textToColumnsDialogTitle">
        <header class="demo__modal-header">
          <h2>{{ cellText.textToColumnsDialogTitle }}</h2>
          <button type="button" class="demo__modal-x" :aria-label="cellText.sortDialogCancel" @click="textToColumnsDialog = null">×</button>
        </header>
        <div class="demo__modal-body demo__sort-dialog">
          <fieldset class="demo__modal-field">
            <legend>{{ cellText.textToColumnsDialogDelimiters }}</legend>
            <label class="demo__sort-dialog__check">
              <input v-model="textToColumnsDialog.comma" type="checkbox" />
              <span>{{ cellText.textToColumnsComma }}</span>
            </label>
            <label class="demo__sort-dialog__check">
              <input v-model="textToColumnsDialog.tab" type="checkbox" />
              <span>{{ cellText.textToColumnsTab }}</span>
            </label>
            <label class="demo__sort-dialog__check">
              <input v-model="textToColumnsDialog.semicolon" type="checkbox" />
              <span>{{ cellText.textToColumnsSemicolon }}</span>
            </label>
            <label class="demo__sort-dialog__check">
              <input v-model="textToColumnsDialog.space" type="checkbox" />
              <span>{{ cellText.textToColumnsSpace }}</span>
            </label>
          </fieldset>
          <label class="demo__sort-dialog__check">
            <input v-model="textToColumnsDialog.collapseConsecutive" type="checkbox" />
            <span>{{ cellText.textToColumnsTreatConsecutive }}</span>
          </label>
        </div>
        <footer class="demo__modal-footer">
          <button type="button" class="demo__btn" @click="textToColumnsDialog = null">{{ cellText.sortDialogCancel }}</button>
          <button type="button" class="demo__btn demo__btn--primary" @click="applyTextToColumnsDialog">{{ cellText.sortDialogApply }}</button>
        </footer>
      </div>
    </div>
  </div>
</template>
