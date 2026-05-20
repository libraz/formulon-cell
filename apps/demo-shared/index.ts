/**
 * Shared configuration for the React and Vue demo apps.
 *
 * The two demos historically maintained their own byte-identical copies of
 * `THEMES`, `LOCALES`, `PRESETS`, `FEATURE_GROUPS`, `DEMO_FUNCTIONS`,
 * `FORMATTERS`, the `formatLoadError` helper, and the `UI` strings table —
 * the only divergence being the framework label ("React" vs "Vue") used in
 * the workbook title and backstage subtitle. That layout invited drift bugs
 * (a translation added to one demo would silently miss the other) and
 * doubled review surface for every UI tweak. This module is the single
 * source of truth so the two demos can stay aligned.
 *
 * Shared helpers keep the React and Vue demos on the same integration line.
 */

import type {
  CellChangeEvent,
  CellRenderInput,
  CellValue,
  FeatureFlags,
  FeatureId,
  PrinterProfile,
  ReviewCell,
  RibbonSearchItem,
  RibbonSearchUsagePrior,
  RibbonTab,
  SpreadsheetInstance,
  SpreadsheetUiOptions,
  ThemeName,
  ToolbarInstance,
  WorkbookHandle,
} from '@libraz/formulon-cell';
import {
  buildPrintDocument,
  buildRibbonSearchIndex,
  EXCEL365_STANDARD_RIBBON_TABS,
  getPageSetup,
  queryRibbonSearchIndex,
  resolvePrinterProfileBounds,
  resolveSpreadsheetUiOptions,
} from '@libraz/formulon-cell';

export type DemoFramework = 'React' | 'Vue';

export const THEMES: { value: ThemeName; label: string }[] = [
  { value: 'paper', label: 'Light' },
  { value: 'ink', label: 'Dark' },
  { value: 'contrast', label: 'Contrast' },
];

export const LOCALES = [
  { value: 'en', label: 'EN' },
  { value: 'ja', label: 'JA' },
] as const;

export const DEMO_PRINTER_PROFILES: readonly PrinterProfile[] = [
  {
    id: 'demo-office-a4',
    name: 'Demo Office Printer - A4',
    paperSize: 'A4',
    orientation: 'portrait',
    printableBounds: { top: 0.16, right: 0.14, bottom: 0.18, left: 0.14 },
  },
  {
    id: 'demo-office-a4-landscape',
    name: 'Demo Office Printer - A4 Landscape',
    paperSize: 'A4',
    orientation: 'landscape',
    printableBounds: { top: 0.14, right: 0.16, bottom: 0.14, left: 0.16 },
  },
  {
    id: 'demo-label-letter',
    name: 'Demo Label Printer - Letter',
    paperSize: 'letter',
    orientation: 'portrait',
    printableBounds: { top: 0.32, right: 0.28, bottom: 0.34, left: 0.28 },
  },
];

export const DEMO_PRINTER_PROFILE_ID = 'demo-office-a4';

export const refreshDemoPrinterProfiles = async (): Promise<readonly PrinterProfile[]> =>
  DEMO_PRINTER_PROFILES.map((profile) => ({
    ...profile,
    printableBounds: { ...profile.printableBounds },
  }));

export type DemoUploadStatus = 'saved' | 'saving' | 'error' | null;

export interface SaveDemoWorkbookOptions {
  instance: Pick<SpreadsheetInstance, 'workbook'> | null;
  bookName: string;
  setUploadStatus: (status: DemoUploadStatus) => void;
  documentRef?: Pick<Document, 'body' | 'createElement'>;
  urlApi?: Pick<typeof URL, 'createObjectURL' | 'revokeObjectURL'>;
  setTimeoutFn?: (handler: () => void, timeout: number) => unknown;
}

export const saveDemoWorkbookToDownload = ({
  instance,
  bookName,
  setUploadStatus,
  documentRef = globalThis.document,
  urlApi = globalThis.URL,
  setTimeoutFn = globalThis.setTimeout.bind(globalThis),
}: SaveDemoWorkbookOptions): void => {
  if (!instance) return;
  setUploadStatus('saving');
  try {
    const bytes = instance.workbook.save();
    const blob = new Blob([bytes as BlobPart], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = urlApi.createObjectURL(blob);
    const a = documentRef.createElement('a');
    a.href = url;
    a.download = `${bookName}.xlsx`;
    documentRef.body.appendChild(a);
    a.click();
    documentRef.body.removeChild(a);
    setTimeoutFn(() => urlApi.revokeObjectURL(url), 1_000);
    setUploadStatus('saved');
  } catch {
    setUploadStatus('error');
  }
};

export const DEMO_ICONS = {
  app: ['M4 5.2 10 3.4l6 1.8v9.6l-6 1.8-6-1.8z', 'M10 3.4v13.2', 'M4 8.6h12', 'M4 11.4h12'],
  save: ['M4 4h10l2 2v10H4z', 'M7 4v5h6V4', 'M7 13h6'],
  undo: [
    'M7.2 5.2H3.8v-3.4',
    'M4 5.2c2.2-2.1 5.7-2.3 8.1-.5 2.7 2.1 3 6.1.7 8.6-1.8 1.9-4.8 2.4-7.1 1.2',
  ],
  redo: [
    'M12.8 5.2h3.4v-3.4',
    'M16 5.2c-2.2-2.1-5.7-2.3-8.1-.5-2.7 2.1-3 6.1-.7 8.6 1.8 1.9 4.8 2.4 7.1 1.2',
  ],
  search: ['M8.5 14a5.5 5.5 0 1 1 0-11 5.5 5.5 0 0 1 0 11z', 'M12.5 12.5L17 17'],
} as const;

export type DemoIconName = keyof typeof DEMO_ICONS;

export type DemoLocale = (typeof LOCALES)[number]['value'];

const isDemoLocale = (value: string | null | undefined): value is DemoLocale =>
  value === 'en' || value === 'ja';

export const resolveInitialLocale = (
  search = globalThis.location?.search ?? '',
  languages = globalThis.navigator?.languages ?? [globalThis.navigator?.language ?? ''],
): DemoLocale => {
  const param = new URLSearchParams(search).get('locale');
  if (isDemoLocale(param)) return param;
  const first = languages.find((lang): lang is string => typeof lang === 'string' && lang !== '');
  if (first?.toLowerCase().startsWith('en')) return 'en';
  return 'ja';
};

export const installDemoSearchShortcut = (
  getInput: () => HTMLInputElement | null | undefined,
): (() => void) => {
  const onKeydown = (event: KeyboardEvent): void => {
    if (!event.altKey || event.ctrlKey || event.metaKey || event.shiftKey) return;
    if (event.key.toLowerCase() !== 'q') return;
    const input = getInput();
    if (!input || input.disabled || input.hidden) return;
    event.preventDefault();
    input.focus();
    input.select();
  };
  document.addEventListener('keydown', onKeydown);
  return () => document.removeEventListener('keydown', onKeydown);
};

export interface DemoF6NavigationTargets {
  getQuickAccess: () => HTMLElement | null | undefined;
  getToolbar: () => ToolbarInstance | null | undefined;
  getInstance: () => SpreadsheetInstance | null | undefined;
}

const isVisibleDemoLandmark = (el: HTMLElement | null | undefined): el is HTMLElement =>
  !!el && el.isConnected && !el.hidden && el.getAttribute('aria-hidden') !== 'true';

const focusFirstFocusable = (root: HTMLElement): boolean => {
  const target = root.matches(
    'button:not(:disabled), input:not(:disabled), textarea:not(:disabled), [tabindex]',
  )
    ? root
    : root.querySelector<HTMLElement>(
        'button:not(:disabled), input:not(:disabled), textarea:not(:disabled), [tabindex]',
      );
  if (!target) return false;
  target.focus({ preventScroll: true });
  return document.activeElement === target;
};

export const installDemoF6Navigation = ({
  getQuickAccess,
  getToolbar,
  getInstance,
}: DemoF6NavigationTargets): (() => void) => {
  const getNameBox = (): HTMLInputElement | null =>
    getInstance()?.host.querySelector<HTMLInputElement>('.fc-host__formulabar-tag') ?? null;
  const getStatusBar = (): HTMLElement | null =>
    getInstance()?.host.querySelector<HTMLElement>('.fc-host__statusbar') ?? null;
  const containsActive = (el: HTMLElement | null | undefined): boolean =>
    !!el && document.activeElement instanceof Node && el.contains(document.activeElement);
  const focusers = [
    () => {
      const quick = getQuickAccess();
      return isVisibleDemoLandmark(quick) && focusFirstFocusable(quick);
    },
    () => getToolbar()?.focusActiveTab() ?? false,
    () => {
      const nameBox = getNameBox();
      if (!isVisibleDemoLandmark(nameBox)) return false;
      nameBox.focus({ preventScroll: true });
      nameBox.select();
      return document.activeElement === nameBox;
    },
    () => {
      const host = getInstance()?.host;
      if (!isVisibleDemoLandmark(host)) return false;
      host.focus({ preventScroll: true });
      return document.activeElement === host;
    },
    () => {
      const status = getStatusBar();
      if (!isVisibleDemoLandmark(status)) return false;
      status.focus({ preventScroll: true });
      return document.activeElement === status;
    },
  ] as const;
  const currentIndex = (): number => {
    if (containsActive(getQuickAccess())) return 0;
    if (containsActive(getToolbar()?.host)) return 1;
    const active = document.activeElement;
    const nameBox = getNameBox();
    if (active === nameBox || containsActive(nameBox?.closest<HTMLElement>('.fc-host__formulabar')))
      return 2;
    const status = getStatusBar();
    if (active === status || containsActive(status)) return 4;
    if (containsActive(getInstance()?.host)) return 3;
    return -1;
  };
  const onKeydown = (event: KeyboardEvent): void => {
    if (event.key !== 'F6' || event.ctrlKey || event.metaKey || event.altKey) return;
    const active = document.activeElement;
    if (
      active instanceof Element &&
      active.closest('.app__dlg, .fc-fmtdlg, .app__menu, .fc-statusbar__chooser')
    )
      return;
    event.preventDefault();
    const start = currentIndex();
    const direction = event.shiftKey ? -1 : 1;
    for (let step = 1; step <= focusers.length; step += 1) {
      const next = (start + direction * step + focusers.length) % focusers.length;
      if (focusers[next]?.()) return;
    }
  };
  document.addEventListener('keydown', onKeydown);
  return () => document.removeEventListener('keydown', onKeydown);
};

export type PresetKey = 'minimal' | 'standard' | 'full';
export const PRESETS: { value: PresetKey; label: string; hint: string }[] = [
  { value: 'minimal', label: 'Minimal', hint: 'bare spreadsheet chrome' },
  { value: 'standard', label: 'Standard', hint: 'lightweight editing chrome' },
  { value: 'full', label: 'Full', hint: 'complete spreadsheet chrome' },
];

export const composeDemoUiOptions = (input: {
  preset: PresetKey;
  overrides: FeatureFlags;
  showRibbon: boolean;
  theme: ThemeName;
}): ReturnType<typeof resolveSpreadsheetUiOptions> => {
  const profile: SpreadsheetUiOptions['profile'] =
    input.preset === 'full' ? 'excel365' : input.preset;
  return resolveSpreadsheetUiOptions({
    profile,
    theme: input.theme,
    features: { ribbon: input.showRibbon },
    advancedFeatures: input.overrides,
  });
};

export const FEATURE_GROUPS: {
  title: string;
  features: { id: FeatureId; label: string }[];
}[] = [
  {
    title: 'Chrome',
    features: [
      { id: 'formulaBar', label: 'Formula bar' },
      { id: 'viewToolbar', label: 'View toolbar' },
      { id: 'sheetTabs', label: 'Sheet tabs' },
      { id: 'statusBar', label: 'Status bar' },
      { id: 'workbookObjects', label: 'Workbook objects' },
      { id: 'contextMenu', label: 'Context menu' },
      { id: 'charts', label: 'Charts' },
      { id: 'watchWindow', label: 'Watch window' },
      { id: 'slicer', label: 'Slicer' },
    ],
  },
  {
    title: 'Editing',
    features: [
      { id: 'clipboard', label: 'Clipboard' },
      { id: 'pasteSpecial', label: 'Paste special' },
      { id: 'quickAnalysis', label: 'Quick Analysis' },
      { id: 'formatPainter', label: 'Format painter' },
      { id: 'autocomplete', label: 'Autocomplete' },
      { id: 'shortcuts', label: 'Shortcuts' },
      { id: 'wheel', label: 'Wheel scroll' },
    ],
  },
  {
    title: 'Dialogs & overlays',
    features: [
      { id: 'findReplace', label: 'Find & replace' },
      { id: 'gotoSpecial', label: 'Go To Special' },
      { id: 'formatDialog', label: 'Format dialog' },
      { id: 'fxDialog', label: 'Function dialog' },
      { id: 'pageSetup', label: 'Page setup' },
      { id: 'iterative', label: 'Iterative calc' },
      { id: 'conditional', label: 'Conditional formatting' },
      { id: 'namedRanges', label: 'Named ranges' },
      { id: 'hyperlink', label: 'Hyperlink' },
      { id: 'commentDialog', label: 'Comment popover' },
      { id: 'pivotTableDialog', label: 'PivotTable dialog' },
      { id: 'validation', label: 'Data validation' },
      { id: 'hoverComment', label: 'Hover comment' },
      { id: 'errorIndicators', label: 'Error indicators' },
    ],
  },
];

export const formatLoadError = (err: unknown): string =>
  err instanceof Error ? err.message : String(err);

export interface DemoUiStrings {
  saved: string;
  search: string;
  searchCommands: string;
  share: string;
  workbook: string;
  demoPane: string;
  open: string;
  save: string;
  file: string;
  info: string;
  print: string;
  pageSetup: string;
  theme: string;
  locale: string;
  signedInUser: string;
  close: string;
  backstageSub: string;
  newWorkbook: string;
  newWorkbookDesc: string;
  openTitle: string;
  openDesc: string;
  saveCopy: string;
  saveDesc: string;
  saveAsDesc: string;
  printDesc: string;
  printPreviewTitle: string;
  printNow: string;
  printToPdf: string;
  printSettings: string;
  printPreviewSheet: string;
  printPreviewOrientation: string;
  printPreviewPaper: string;
  printPreviewPrinter: string;
  printPreviewPrinterMargins: string;
  printPreviewMargins: string;
  printPreviewScale: string;
  printPreviewArea: string;
  printPreviewNoArea: string;
  printPreviewPage: string;
  printPreviewHint: string;
  printPreviewUnavailable: string;
  pageSetupDesc: string;
  editLinks: string;
  linksDesc: string;
  export: string;
  exportDesc: string;
  shareDesc: string;
  options: string;
  optionsDesc: string;
  noCommands: string;
  engineUnavailable: string;
  engineSetup: string;
}

export interface DemoCommandStrings {
  ribbonCommand: string;
  selection: string;
  workbook: string;
  cellsUpdated: string;
  draw: string;
  translate: string;
  addIns: string;
  openFailed: string;
  script: string;
  scriptCommandError: string;
  accessibilityCheck: string;
  inkNotPersisted: string;
  selectInkFirst: string;
  translationUnavailable: string;
  addInsHostCallbacks: string;
  commands: Record<
    | 'open'
    | 'save'
    | 'pageSetup'
    | 'print'
    | 'formatCells'
    | 'conditionalFormatting'
    | 'cellStyles'
    | 'nameManager'
    | 'insertFunction'
    | 'tracePrecedents'
    | 'watchWindow'
    | 'filter'
    | 'sort'
    | 'freezePanes'
    | 'protectSheet'
    | 'options'
    | 'lightTheme'
    | 'darkTheme'
    | 'japaneseLocale'
    | 'englishLocale',
    { label: string; hint: string }
  >;
}

export interface DemoCommandItem {
  readonly id: string;
  readonly label: string;
  readonly hint: string;
  readonly keywords?: string;
  readonly tab?: RibbonTab;
  readonly run: () => void;
}

export interface DemoSearchItem {
  readonly id: string;
  readonly label: string;
  readonly hint: string;
  readonly commandId?: string;
  readonly disabled?: boolean;
  readonly disabledReason?: string;
  readonly keywords?: string;
  readonly tab?: RibbonTab;
  readonly run: () => void;
}

export type DemoSearchUsagePrior = RibbonSearchUsagePrior;

const DEMO_SEARCH_USAGE_KEY = 'formulon-cell.demo.searchUsagePrior';

const isSearchUsagePrior = (value: unknown): value is DemoSearchUsagePrior => {
  if (!value || typeof value !== 'object') return false;
  const boosts = (value as { commandBoosts?: unknown }).commandBoosts;
  return (
    boosts === undefined ||
    (!!boosts &&
      typeof boosts === 'object' &&
      Object.values(boosts).every((entry) => typeof entry === 'number' && Number.isFinite(entry)))
  );
};

export const loadDemoSearchUsagePrior = (): DemoSearchUsagePrior => {
  try {
    const raw = globalThis.localStorage?.getItem(DEMO_SEARCH_USAGE_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw) as unknown;
    return isSearchUsagePrior(parsed) ? parsed : {};
  } catch {
    return {};
  }
};

export const saveDemoSearchUsagePrior = (prior: DemoSearchUsagePrior): void => {
  try {
    globalThis.localStorage?.setItem(DEMO_SEARCH_USAGE_KEY, JSON.stringify(prior));
  } catch {
    // Storage is optional in embedded demos.
  }
};

export const recordDemoSearchUsage = (
  prior: DemoSearchUsagePrior,
  item: DemoSearchItem,
): DemoSearchUsagePrior => {
  const commandId = item.commandId;
  if (!commandId) return prior;
  return {
    commandBoosts: {
      ...(prior.commandBoosts ?? {}),
      [commandId]: Math.min(100, (prior.commandBoosts?.[commandId] ?? 0) + 12),
    },
  };
};

export interface BuildDemoCommandsOptions {
  commandText: DemoCommandStrings;
  instance: SpreadsheetInstance | null;
  openWorkbook: () => void;
  saveWorkbook: () => void;
  setRibbonTab: (tab: RibbonTab) => void;
  togglePanel: () => void;
  setTheme: (theme: ThemeName) => void;
  setLocale: (locale: string) => void;
}

export const buildDemoCommands = ({
  commandText,
  instance,
  openWorkbook,
  saveWorkbook,
  setRibbonTab,
  togglePanel,
  setTheme,
  setLocale,
}: BuildDemoCommandsOptions): readonly DemoCommandItem[] => [
  {
    id: 'open',
    label: commandText.commands.open.label,
    hint: commandText.commands.open.hint,
    tab: 'file',
    run: openWorkbook,
  },
  {
    id: 'save',
    label: commandText.commands.save.label,
    hint: commandText.commands.save.hint,
    tab: 'file',
    run: saveWorkbook,
  },
  {
    id: 'page-setup',
    label: commandText.commands.pageSetup.label,
    hint: commandText.commands.pageSetup.hint,
    tab: 'file',
    run: () => instance?.openPageSetup(),
  },
  {
    id: 'print',
    label: commandText.commands.print.label,
    hint: commandText.commands.print.hint,
    tab: 'file',
    run: () => instance?.print('print'),
  },
  {
    id: 'format-cells',
    label: commandText.commands.formatCells.label,
    hint: commandText.commands.formatCells.hint,
    tab: 'home',
    run: () => instance?.openFormatDialog(),
  },
  {
    id: 'conditional',
    label: commandText.commands.conditionalFormatting.label,
    hint: commandText.commands.conditionalFormatting.hint,
    tab: 'insert',
    run: () => instance?.openConditionalDialog(),
  },
  {
    id: 'cell-styles',
    label: commandText.commands.cellStyles.label,
    hint: commandText.commands.cellStyles.hint,
    tab: 'insert',
    run: () => instance?.openCellStylesGallery(),
  },
  {
    id: 'name-manager',
    label: commandText.commands.nameManager.label,
    hint: commandText.commands.nameManager.hint,
    tab: 'insert',
    run: () => instance?.openNamedRangeDialog(),
  },
  {
    id: 'insert-function',
    label: commandText.commands.insertFunction.label,
    hint: commandText.commands.insertFunction.hint,
    tab: 'formulas',
    run: () => instance?.openFunctionArguments(),
  },
  {
    id: 'trace-precedents',
    label: commandText.commands.tracePrecedents.label,
    hint: commandText.commands.tracePrecedents.hint,
    tab: 'formulas',
    run: () => instance?.tracePrecedents(),
  },
  {
    id: 'watch-window',
    label: commandText.commands.watchWindow.label,
    hint: commandText.commands.watchWindow.hint,
    tab: 'formulas',
    run: () => instance?.toggleWatchWindow(),
  },
  {
    id: 'filter',
    label: commandText.commands.filter.label,
    hint: commandText.commands.filter.hint,
    tab: 'data',
    run: () => setRibbonTab('data'),
  },
  {
    id: 'sort',
    label: commandText.commands.sort.label,
    hint: commandText.commands.sort.hint,
    tab: 'data',
    run: () => setRibbonTab('data'),
  },
  {
    id: 'freeze-panes',
    label: commandText.commands.freezePanes.label,
    hint: commandText.commands.freezePanes.hint,
    tab: 'view',
    run: () => setRibbonTab('view'),
  },
  {
    id: 'protect-sheet',
    label: commandText.commands.protectSheet.label,
    hint: commandText.commands.protectSheet.hint,
    tab: 'view',
    run: () => instance?.toggleSheetProtection(),
  },
  {
    id: 'options-pane',
    label: commandText.commands.options.label,
    hint: commandText.commands.options.hint,
    run: togglePanel,
  },
  {
    id: 'theme-light',
    label: commandText.commands.lightTheme.label,
    hint: commandText.commands.lightTheme.hint,
    run: () => setTheme('paper'),
  },
  {
    id: 'theme-dark',
    label: commandText.commands.darkTheme.label,
    hint: commandText.commands.darkTheme.hint,
    run: () => setTheme('ink'),
  },
  {
    id: 'locale-ja',
    label: commandText.commands.japaneseLocale.label,
    hint: commandText.commands.japaneseLocale.hint,
    run: () => setLocale('ja'),
  },
  {
    id: 'locale-en',
    label: commandText.commands.englishLocale.label,
    hint: commandText.commands.englishLocale.hint,
    run: () => setLocale('en'),
  },
];

const ribbonSearchHint = (item: RibbonSearchItem): string => {
  const hint =
    item.kind === 'tab' || item.kind === 'help' ? item.hint : `${item.hint} · ${item.tab}`;
  return item.disabledReason ? `${hint} · ${item.disabledReason}` : hint;
};

export const buildDemoSearchItems = (
  commandItems: readonly DemoCommandItem[],
  locale: string,
  setRibbonTab: (tab: RibbonTab) => void,
  applyRibbonCommand?: (commandId: string) => boolean,
  ribbonTabs: readonly RibbonTab[] = EXCEL365_STANDARD_RIBBON_TABS,
): readonly DemoSearchItem[] => {
  const commandKeys = new Set(commandItems.flatMap((item) => [item.id, item.label]));
  const ribbonItems = buildRibbonSearchIndex(locale === 'en' ? 'en' : 'ja', {
    includeDisabled: true,
    tabs: ribbonTabs,
  })
    .filter((item) => !commandKeys.has(item.commandId ?? '') && !commandKeys.has(item.label))
    .map(
      (item): DemoSearchItem => ({
        id: `ribbon:${item.id}`,
        label: item.label,
        hint: ribbonSearchHint(item),
        commandId: item.commandId,
        disabled: item.disabled,
        disabledReason: item.disabledReason,
        keywords: item.keywords,
        tab: item.tab,
        run: () => {
          if (item.commandId && applyRibbonCommand?.(item.commandId)) return;
          setRibbonTab(item.tab);
        },
      }),
    );
  return [...commandItems, ...ribbonItems];
};

export const queryDemoSearchItems = (
  items: readonly DemoSearchItem[],
  query: string,
  limit = 8,
  usagePrior: DemoSearchUsagePrior = {},
): readonly DemoSearchItem[] => {
  const q = query.trim().toLowerCase();
  if (!q) return items.slice(0, limit);
  const ribbonBacked = items.filter((item) => item.id.startsWith('ribbon:'));
  const demoBacked = items.filter((item) => !item.id.startsWith('ribbon:'));
  const demoMatches = demoBacked.filter((item) =>
    `${item.label} ${item.hint} ${item.keywords ?? ''}`.toLowerCase().includes(q),
  );
  const ribbonMatches = queryRibbonSearchIndex(
    ribbonBacked.map(
      (item): RibbonSearchItem => ({
        id: item.id,
        kind: 'command',
        label: item.label,
        hint: item.hint,
        tab: item.tab ?? 'home',
        commandId: item.commandId,
        disabled: item.disabled,
        disabledReason: item.disabledReason,
        keywords: `${item.label} ${item.hint} ${item.keywords ?? ''}`.toLowerCase(),
      }),
    ),
    query,
    limit,
    { usagePrior },
  );
  const byRibbonId = new Map(ribbonBacked.map((item) => [item.id, item]));
  const orderedRibbonMatches = ribbonMatches.flatMap((item) => {
    const match = byRibbonId.get(item.id);
    return match ? [match] : [];
  });
  return [...demoMatches, ...orderedRibbonMatches].slice(0, limit);
};

export const demoSearchOptionId = (index: number): string => `demo-search-option-${index}`;

export const nextDemoSearchIndex = (
  current: number,
  count: number,
  direction: 'first' | 'next' | 'previous',
): number => {
  if (count <= 0) return -1;
  if (direction === 'first') return current >= 0 ? current : 0;
  if (direction === 'next') return (current + 1 + count) % count;
  return (current < 0 ? count - 1 : current - 1 + count) % count;
};

export interface DemoStrings {
  en: DemoUiStrings;
  ja: DemoUiStrings;
}

export type DemoBackstageAction =
  | 'info'
  | 'new'
  | 'open'
  | 'save'
  | 'save-as'
  | 'print'
  | 'page-setup'
  | 'edit-links'
  | 'share'
  | 'export'
  | 'options'
  | 'close';

export interface DemoBackstageItem {
  action: DemoBackstageAction;
  label: string;
  desc?: string;
  active?: boolean;
}

export interface DemoPrintPreviewModel {
  title: string;
  subtitle: string;
  printLabel: string;
  pdfLabel: string;
  pageSetupLabel: string;
  previewTitle: string;
  previewHint: string;
  previewHtml: string;
  settings: readonly { label: string; value: string }[];
}

export const DEMO_PRINT_PREVIEW_LINES = [
  'row-1',
  'row-2',
  'row-3',
  'row-4',
  'row-5',
  'row-6',
  'row-7',
  'row-8',
  'row-9',
  'row-10',
  'row-11',
  'row-12',
] as const;

export const demoBackstageRequiresWorkbook = (action: DemoBackstageAction): boolean =>
  action === 'save' ||
  action === 'save-as' ||
  action === 'print' ||
  action === 'page-setup' ||
  action === 'edit-links' ||
  action === 'export';

export const isDemoBackstageActionDisabled = (
  action: DemoBackstageAction,
  instance: SpreadsheetInstance | null | undefined,
): boolean => !instance && demoBackstageRequiresWorkbook(action);

export interface RunDemoBackstageActionOptions {
  action: DemoBackstageAction;
  instance: SpreadsheetInstance | null | undefined;
  ui: Pick<DemoUiStrings, 'share' | 'shareDesc'>;
  newWorkbook: () => void | Promise<void>;
  openWorkbook: () => void;
  saveWorkbook: () => void;
  showNotice: (title: string, detail: string) => void;
  toggleOptions: () => void;
  closeBackstage: () => void;
}

export const runDemoBackstageAction = (opts: RunDemoBackstageActionOptions): void => {
  const { action, instance } = opts;
  if (action === 'info') return;
  if (action === 'new') {
    void opts.newWorkbook();
  } else if (action === 'open') opts.openWorkbook();
  else if (action === 'save' || action === 'save-as') opts.saveWorkbook();
  else if (action === 'print') instance?.print('print');
  else if (action === 'export') instance?.print('pdf');
  else if (action === 'page-setup') instance?.openPageSetup();
  else if (action === 'edit-links') instance?.openExternalLinksDialog();
  else if (action === 'share') opts.showNotice(opts.ui.share, opts.ui.shareDesc);
  else if (action === 'options') opts.toggleOptions();
  else if (action === 'close') opts.closeBackstage();
};

const formatMarginSummary = (margins: {
  top: number;
  right: number;
  bottom: number;
  left: number;
}): string => `${margins.top}" / ${margins.right}" / ${margins.bottom}" / ${margins.left}"`;

const selectedDemoPrinterProfile = (): PrinterProfile | undefined =>
  DEMO_PRINTER_PROFILES.find((profile) => profile.id === DEMO_PRINTER_PROFILE_ID) ??
  DEMO_PRINTER_PROFILES[0];

export const buildDemoPrintPreviewModel = (
  ui: DemoUiStrings,
  instance: SpreadsheetInstance | null | undefined,
  bookName: string,
): DemoPrintPreviewModel => {
  if (!instance) {
    return {
      title: ui.print,
      subtitle: ui.printPreviewUnavailable,
      printLabel: ui.printNow,
      pdfLabel: ui.printToPdf,
      pageSetupLabel: ui.pageSetup,
      previewTitle: ui.printPreviewPage,
      previewHint: ui.printPreviewUnavailable,
      previewHtml: '',
      settings: [],
    };
  }
  const state = instance.store.getState();
  const sheet = state.data.sheetIndex;
  const setup = getPageSetup(state, sheet);
  const scale =
    setup.fitWidth || setup.fitHeight
      ? `${setup.fitWidth || 1} x ${setup.fitHeight || 1}`
      : `${Math.round((setup.scale ?? 1) * 100)}%`;
  const printerProfile = selectedDemoPrinterProfile();
  const printerBounds = resolvePrinterProfileBounds(
    setup,
    DEMO_PRINTER_PROFILES,
    DEMO_PRINTER_PROFILE_ID,
  );
  const printerSettings = [
    printerProfile?.name ? { label: ui.printPreviewPrinter, value: printerProfile.name } : null,
    printerBounds
      ? { label: ui.printPreviewPrinterMargins, value: formatMarginSummary(printerBounds) }
      : null,
  ].filter((item): item is { label: string; value: string } => item !== null);
  const printDocument = buildPrintDocument(
    instance.workbook,
    instance.store,
    sheet,
    ui.printPreviewTitle,
    {
      printableBounds: printerBounds ?? null,
    },
  );
  return {
    title: ui.printPreviewTitle,
    subtitle: bookName,
    printLabel: ui.printNow,
    pdfLabel: ui.printToPdf,
    pageSetupLabel: ui.pageSetup,
    previewTitle: `${ui.printPreviewPage} 1`,
    previewHint: ui.printPreviewHint,
    previewHtml: printDocument.html,
    settings: [
      { label: ui.printPreviewSheet, value: String(sheet + 1) },
      { label: ui.printPreviewOrientation, value: setup.orientation },
      { label: ui.printPreviewPaper, value: setup.paperSize },
      ...printerSettings,
      { label: ui.printPreviewMargins, value: formatMarginSummary(setup.margins) },
      { label: ui.printPreviewScale, value: scale },
      { label: ui.printPreviewArea, value: setup.printArea?.trim() || ui.printPreviewNoArea },
    ],
  };
};

export const buildDemoBackstageNav = (
  ui: DemoUiStrings,
  active: DemoBackstageAction = 'info',
): readonly DemoBackstageItem[] => [
  { action: 'info', label: ui.info, active: active === 'info' },
  { action: 'new', label: ui.newWorkbook },
  { action: 'open', label: ui.openTitle },
  { action: 'save', label: ui.save },
  { action: 'save-as', label: ui.saveCopy },
  { action: 'print', label: ui.print, active: active === 'print' },
  { action: 'share', label: ui.share },
  { action: 'export', label: ui.export },
  { action: 'options', label: ui.options },
  { action: 'close', label: ui.close },
];

export const buildDemoBackstageCards = (ui: DemoUiStrings): readonly DemoBackstageItem[] => [
  { action: 'new', label: ui.newWorkbook, desc: ui.newWorkbookDesc },
  { action: 'open', label: ui.openTitle, desc: ui.openDesc },
  { action: 'save', label: ui.save, desc: ui.saveDesc },
  { action: 'save-as', label: ui.saveCopy, desc: ui.saveAsDesc },
  { action: 'print', label: ui.print, desc: ui.printDesc },
  { action: 'page-setup', label: ui.pageSetup, desc: ui.pageSetupDesc },
  { action: 'edit-links', label: ui.editLinks, desc: ui.linksDesc },
  { action: 'share', label: ui.share, desc: ui.shareDesc },
  { action: 'export', label: ui.export, desc: ui.exportDesc },
  { action: 'options', label: ui.options, desc: ui.optionsDesc },
];

const JA_FRAMEWORK: Record<DemoFramework, string> = {
  React: 'React',
  Vue: 'Vue',
};

/** Build the UI string table for a demo. The English/Japanese tables are
 *  identical between React and Vue except for the framework label embedded
 *  in `workbook` and `backstageSub`. */
export function createDemoStrings(framework: DemoFramework): DemoStrings {
  const ja = JA_FRAMEWORK[framework];
  return {
    en: {
      saved: 'Saved to this device',
      search: 'Search',
      searchCommands: 'Search commands',
      share: 'Share',
      workbook: `${framework} workbook`,
      demoPane: 'Options',
      open: 'Open xlsx…',
      save: 'Save',
      file: 'File',
      info: 'Info',
      print: 'Print',
      pageSetup: 'Page Setup',
      theme: 'Theme',
      locale: 'Locale',
      signedInUser: 'Signed in user',
      close: 'Close',
      backstageSub: `${framework} workbook · full spreadsheet layout`,
      newWorkbook: 'New',
      newWorkbookDesc: 'Start from a blank workbook in this demo session.',
      openTitle: 'Open',
      openDesc: 'Load an .xlsx or .xlsm workbook from this device.',
      saveCopy: 'Save As',
      saveDesc: 'Download the current workbook as an .xlsx file.',
      saveAsDesc: 'Download a separate copy with the current workbook name.',
      printDesc: 'Use the browser print dialog or save as PDF.',
      printPreviewTitle: 'Print',
      printNow: 'Print',
      printToPdf: 'Export to PDF',
      printSettings: 'Settings',
      printPreviewSheet: 'Active sheet',
      printPreviewOrientation: 'Orientation',
      printPreviewPaper: 'Paper size',
      printPreviewPrinter: 'Printer',
      printPreviewPrinterMargins: 'Minimum margins',
      printPreviewMargins: 'Margins',
      printPreviewScale: 'Scaling',
      printPreviewArea: 'Print area',
      printPreviewNoArea: 'No print area set',
      printPreviewPage: 'Page',
      printPreviewHint: 'Preview reflects the active sheet page setup.',
      printPreviewUnavailable: 'Open a workbook to preview print settings.',
      pageSetupDesc: 'Set orientation, margins, paper size, headers, and print titles.',
      editLinks: 'Edit Links',
      linksDesc: 'Inspect external workbook references carried by the file.',
      export: 'Export',
      exportDesc: 'Use the browser print flow to export as PDF.',
      shareDesc: 'Show the sharing status for this host-driven workbook.',
      options: 'Options',
      optionsDesc: 'Show the integration panel and feature toggles.',
      noCommands: 'No commands found',
      engineUnavailable: 'Spreadsheet engine unavailable',
      engineSetup:
        'Serve this demo with COOP: same-origin and COEP: require-corp so SharedArrayBuffer is available.',
    },
    ja: {
      saved: 'このデバイスに保存済み',
      search: '検索',
      searchCommands: 'コマンドの検索',
      share: '共有',
      workbook: `${ja} ブック`,
      demoPane: 'オプション',
      open: 'xlsx を開く…',
      save: '保存',
      file: 'ファイル',
      info: '情報',
      print: '印刷',
      pageSetup: 'ページ設定',
      theme: 'テーマ',
      locale: '表示言語',
      signedInUser: 'サインイン中のユーザー',
      close: '閉じる',
      backstageSub: `${ja} ブック · スプレッドシート レイアウト`,
      newWorkbook: '新規',
      newWorkbookDesc: 'このデモセッションで空のブックを開始します。',
      openTitle: '開く',
      openDesc: '.xlsx または .xlsm ブックをこのデバイスから読み込みます。',
      saveCopy: '名前を付けて保存',
      saveDesc: '現在のブックを .xlsx ファイルとしてダウンロードします。',
      saveAsDesc: '現在のブック名で別コピーをダウンロードします。',
      printDesc: 'ブラウザーの印刷ダイアログ、または PDF 保存を使用します。',
      printPreviewTitle: '印刷',
      printNow: '印刷',
      printToPdf: 'PDF にエクスポート',
      printSettings: '設定',
      printPreviewSheet: 'アクティブ シート',
      printPreviewOrientation: '印刷の向き',
      printPreviewPaper: '用紙サイズ',
      printPreviewPrinter: 'プリンター',
      printPreviewPrinterMargins: '最小余白',
      printPreviewMargins: '余白',
      printPreviewScale: '拡大縮小',
      printPreviewArea: '印刷範囲',
      printPreviewNoArea: '印刷範囲なし',
      printPreviewPage: 'ページ',
      printPreviewHint: 'プレビューはアクティブ シートのページ設定を反映します。',
      printPreviewUnavailable: 'ブックを開くと印刷設定をプレビューできます。',
      pageSetupDesc: '用紙方向、余白、用紙サイズ、ヘッダー、印刷タイトルを設定します。',
      editLinks: 'リンクの編集',
      linksDesc: 'ファイルに含まれる外部ブック参照を確認します。',
      export: 'エクスポート',
      exportDesc: 'ブラウザーの印刷フローを使って PDF として出力します。',
      shareDesc: 'このホスト管理ブックの共有状態を表示します。',
      options: 'オプション',
      optionsDesc: '統合パネルと機能トグルを表示します。',
      noCommands: 'コマンドが見つかりません',
      engineUnavailable: 'スプレッドシートエンジンを起動できません',
      engineSetup:
        'SharedArrayBuffer を有効にするため、COOP: same-origin と COEP: require-corp 付きで配信してください。',
    },
  };
}

export const demoCommandText = (locale: string): DemoCommandStrings =>
  locale === 'ja'
    ? {
        ribbonCommand: 'リボン コマンド',
        selection: '選択範囲',
        workbook: 'ブック',
        cellsUpdated: '{count} 個のセルを更新しました。',
        draw: '描画',
        translate: '翻訳',
        addIns: 'アドイン',
        openFailed: 'ファイルを開けませんでした',
        script: 'スクリプト',
        scriptCommandError: '次のいずれかを使用してください: uppercase, lowercase, trim, clear.',
        accessibilityCheck: 'アクセシビリティ チェック',
        inkNotPersisted: 'このデモ ブックではインク ストロークは保存されません。',
        selectInkFirst: '消しゴムを使うには、先にインク ストロークを選択してください。',
        translationUnavailable: 'このデモには翻訳サービスが接続されていません。',
        addInsHostCallbacks: 'ここでは Office アドインをホスト コールバックで表しています。',
        commands: {
          open: { label: '開く', hint: 'xlsx または xlsm ブックを開きます' },
          save: { label: '保存', hint: 'ブックを xlsx としてダウンロードします' },
          pageSetup: { label: 'ページ設定', hint: 'ページ設定を開きます' },
          print: { label: '印刷', hint: 'ブラウザーの印刷ダイアログを開きます' },
          formatCells: { label: 'セルの書式設定', hint: '書式設定ダイアログを開きます' },
          conditionalFormatting: {
            label: '条件付き書式',
            hint: '条件付き書式を作成または編集します',
          },
          cellStyles: { label: 'セルのスタイル', hint: 'スタイル ギャラリーを開きます' },
          nameManager: { label: '名前の管理', hint: '名前付き範囲を確認します' },
          insertFunction: { label: '関数の挿入', hint: '関数の引数を開きます' },
          tracePrecedents: { label: '参照元のトレース', hint: '参照元矢印を表示します' },
          watchWindow: { label: 'ウォッチ ウィンドウ', hint: 'ウォッチ ウィンドウを切り替えます' },
          filter: { label: 'フィルター', hint: 'データ タブのフィルター ツールを表示します' },
          sort: { label: '並べ替え', hint: '並べ替えボタンを表示します' },
          freezePanes: { label: 'ウィンドウ枠の固定', hint: 'ウィンドウ枠の固定を表示します' },
          protectSheet: { label: 'シートの保護', hint: '表示タブからシート保護を切り替えます' },
          options: { label: 'オプション', hint: '統合パネルの表示を切り替えます' },
          lightTheme: { label: 'ライト テーマ', hint: 'ブックをライト テーマに切り替えます' },
          darkTheme: { label: 'ダーク テーマ', hint: 'ブックをダーク テーマに切り替えます' },
          japaneseLocale: { label: '日本語表示', hint: 'ラベルを日本語に切り替えます' },
          englishLocale: { label: '英語表示', hint: 'ラベルを英語に切り替えます' },
        },
      }
    : {
        ribbonCommand: 'Ribbon command',
        selection: 'Selection',
        workbook: 'Workbook',
        cellsUpdated: '{count} cells updated.',
        draw: 'Draw',
        translate: 'Translate',
        addIns: 'Add-ins',
        openFailed: 'Open failed',
        script: 'Script',
        scriptCommandError: 'Use one of: uppercase, lowercase, trim, clear.',
        accessibilityCheck: 'Accessibility Check',
        inkNotPersisted: 'Ink strokes are not persisted in this demo workbook.',
        selectInkFirst: 'Select an ink stroke first to use the eraser.',
        translationUnavailable: 'No translation service is connected in this demo.',
        addInsHostCallbacks: 'Office add-ins are represented by host callbacks here.',
        commands: {
          open: { label: 'Open', hint: 'Open an xlsx or xlsm workbook' },
          save: { label: 'Save', hint: 'Download the workbook as xlsx' },
          pageSetup: { label: 'Page Setup', hint: 'Open page setup' },
          print: { label: 'Print', hint: 'Open browser print dialog' },
          formatCells: { label: 'Format Cells', hint: 'Open the format dialog' },
          conditionalFormatting: {
            label: 'Conditional Formatting',
            hint: 'Create or edit conditional formatting',
          },
          cellStyles: { label: 'Cell Styles', hint: 'Open the style gallery' },
          nameManager: { label: 'Name Manager', hint: 'Inspect named ranges' },
          insertFunction: { label: 'Insert Function', hint: 'Open function arguments' },
          tracePrecedents: { label: 'Trace Precedents', hint: 'Show precedent arrows' },
          watchWindow: { label: 'Watch Window', hint: 'Toggle Watch Window' },
          filter: { label: 'Filter', hint: 'Show the Data tab filter tools' },
          sort: { label: 'Sort', hint: 'Show sort buttons' },
          freezePanes: { label: 'Freeze Panes', hint: 'Show Freeze Panes' },
          protectSheet: { label: 'Protect Sheet', hint: 'Toggle sheet protection from View' },
          options: { label: 'Options', hint: 'Show or hide the integration panel' },
          lightTheme: { label: 'Light Theme', hint: 'Switch to light workbook theme' },
          darkTheme: { label: 'Dark Theme', hint: 'Switch to dark workbook theme' },
          japaneseLocale: { label: 'Japanese Locale', hint: 'Switch labels to JA' },
          englishLocale: { label: 'English Locale', hint: 'Switch labels to EN' },
        },
      };

/** Sample user-defined functions registered into both demos so testers can
 *  verify host function injection without re-deriving the boilerplate. */
export const DEMO_FUNCTIONS = [
  {
    name: 'GREET',
    impl: (...args: CellValue[]) => {
      const v = args[0];
      const who = v?.kind === 'text' ? v.value : 'World';
      return `Hello, ${who}!`;
    },
    meta: { description: 'Friendly greeting', args: ['name'], returnType: 'text' as const },
  },
  {
    name: 'FAHRENHEIT',
    impl: (...args: CellValue[]) => {
      const v = args[0];
      const c = v?.kind === 'number' ? v.value : Number.NaN;
      return Number.isFinite(c) ? c * 1.8 + 32 : null;
    },
    meta: {
      description: 'Celsius to Fahrenheit',
      args: ['celsius'],
      returnType: 'number' as const,
    },
  },
];

/** Custom cell formatters demonstrating the formatter registry. */
export const FORMATTERS = {
  uppercaseA: {
    id: 'demo:uppercaseA',
    match: (i: CellRenderInput) => i.addr.col === 0 && i.value.kind === 'text',
    format: (i: CellRenderInput) => (i.value.kind === 'text' ? i.value.value.toUpperCase() : null),
  },
  arrowNegatives: {
    id: 'demo:arrowNegatives',
    match: (i: CellRenderInput) => i.value.kind === 'number' && i.value.value < 0,
    format: (i: CellRenderInput) =>
      i.value.kind === 'number' ? `↓ ${Math.abs(i.value.value).toFixed(2)}` : null,
  },
};

// ─── Demo runtime helpers ──────────────────────────────────────────────
// These were previously duplicated byte-for-byte between react-demo and
// vue-demo. Each one is framework-agnostic — the React app wraps the
// modal focus helper with `useEffect`, Vue calls it directly.

export const demoColLabel = (n: number): string => {
  let out = '';
  let v = n;
  do {
    out = String.fromCharCode(65 + (v % 26)) + out;
    v = Math.floor(v / 26) - 1;
  } while (v >= 0);
  return out;
};

const FOCUSABLE_DEMO_MODAL_SELECTOR = [
  'button',
  'input',
  'select',
  'textarea',
  'a[href]',
  '[tabindex]:not([tabindex="-1"])',
].join(',');

const focusableDemoModalItems = (root: HTMLElement): HTMLElement[] =>
  Array.from(root.querySelectorAll<HTMLElement>(FOCUSABLE_DEMO_MODAL_SELECTOR)).filter((el) => {
    if (el.closest('[hidden],[aria-hidden="true"]')) return false;
    if ('disabled' in el && (el as HTMLButtonElement | HTMLInputElement).disabled) return false;
    return el.tabIndex >= 0;
  });

/** Wires Tab/Shift+Tab focus trap + Escape close for a demo modal and
 *  returns a teardown callback. Used directly by Vue; React wraps this
 *  in a `useEffect` to attach/detach with the modal's open state. */
export const activateDemoModal = (root: HTMLElement, onClose: () => void): (() => void) => {
  const restoreFocusEl =
    document.activeElement instanceof HTMLElement ? document.activeElement : null;
  const focusFirst = window.requestAnimationFrame(() => {
    (focusableDemoModalItems(root)[0] ?? root).focus({ preventScroll: true });
  });
  const onKeyDown = (event: KeyboardEvent): void => {
    if (event.key === 'Escape') {
      event.preventDefault();
      onClose();
      return;
    }
    if (event.key !== 'Tab') return;
    const items = focusableDemoModalItems(root);
    if (items.length === 0) {
      event.preventDefault();
      root.focus({ preventScroll: true });
      return;
    }
    const first = items[0];
    const last = items[items.length - 1];
    if (event.shiftKey && document.activeElement === first) {
      event.preventDefault();
      last?.focus({ preventScroll: true });
    } else if (!event.shiftKey && document.activeElement === last) {
      event.preventDefault();
      first?.focus({ preventScroll: true });
    }
  };
  root.addEventListener('keydown', onKeyDown);
  return () => {
    window.cancelAnimationFrame(focusFirst);
    root.removeEventListener('keydown', onKeyDown);
    if (
      restoreFocusEl &&
      (root.contains(document.activeElement) || document.activeElement === document.body)
    ) {
      restoreFocusEl.focus({ preventScroll: true });
    }
  };
};

/** Formats a cell change event for the demo's change log (favours the
 *  raw formula text when present so users see what they typed). */
export const previewCellChange = (e: CellChangeEvent): string => {
  if (e.formula) return e.formula;
  switch (e.value.kind) {
    case 'number':
      return String(e.value.value);
    case 'text':
      return JSON.stringify(e.value.value);
    case 'bool':
      return String(e.value.value);
    case 'error':
      return `#${e.value.code}`;
    case 'blank':
      return '∅';
    default:
      return '?';
  }
};

/** Demo seed — only runs once on the initial blank workbook. Core gates
 *  `seed` on `ownsWb`, so re-mounts and Open xlsx don't re-trigger it. */
export const seedDemoWorkbook = (wb: WorkbookHandle): void => {
  wb.setText({ sheet: 0, row: 0, col: 0 }, 'item');
  wb.setText({ sheet: 0, row: 0, col: 1 }, 'celsius');
  wb.setText({ sheet: 0, row: 0, col: 2 }, 'fahrenheit');
  wb.setText({ sheet: 0, row: 0, col: 3 }, 'greeting');
  const rows: [string, number][] = [
    ['London', 8],
    ['Tokyo', 22],
    ['Reykjavík', -3],
    ['Cairo', 31],
  ];
  rows.forEach(([city, c], i) => {
    const r = i + 1;
    wb.setText({ sheet: 0, row: r, col: 0 }, city);
    wb.setNumber({ sheet: 0, row: r, col: 1 }, c);
    wb.setFormula({ sheet: 0, row: r, col: 2 }, `=B${r + 1}*1.8+32`);
    wb.setFormula({ sheet: 0, row: r, col: 3 }, `=A${r + 1}&" ☼"`);
  });
  wb.recalc();
};

/** Projects a workbook's cells for the demo review dialog. */
export const reviewCellsForInstance = (inst: SpreadsheetInstance): ReviewCell[] => {
  const sheet = inst.store.getState().data.sheetIndex;
  return Array.from(inst.workbook.cells(sheet), (entry) => ({
    label: `${demoColLabel(entry.addr.col)}${entry.addr.row + 1}`,
    value:
      entry.value.kind === 'text'
        ? { kind: 'text' as const, value: entry.value.value }
        : entry.value.kind === 'error'
          ? { kind: 'error' as const, text: entry.value.text }
          : entry.value.kind === 'number'
            ? { kind: 'number' as const }
            : entry.value.kind === 'bool'
              ? { kind: 'bool' as const }
              : { kind: 'blank' as const },
    formula: entry.formula,
  }));
};
