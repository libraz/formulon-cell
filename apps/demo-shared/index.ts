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
  ReviewCell,
  RibbonTab,
  SpreadsheetInstance,
  SpreadsheetUiOptions,
  ThemeName,
  ToolbarInstance,
  WorkbookHandle,
} from '@libraz/formulon-cell';
import {
  EXCEL365_MAC_RIBBON_TABS,
  EXCEL365_STANDARD_RIBBON_TABS,
  resolveSpreadsheetPlatform,
  resolveSpreadsheetUiOptions,
} from '@libraz/formulon-cell';

export * from './demo-backstage.js';
export * from './demo-search.js';
export * from './demo-strings.js';

export type DemoFramework = 'React' | 'Vue';

export type DemoPlatform = 'mac' | 'default';
export type DemoPlatformInput = 'auto' | DemoPlatform;

export const THEMES: { value: ThemeName; label: string }[] = [
  { value: 'paper', label: 'Light' },
  { value: 'ink', label: 'Dark' },
  { value: 'contrast', label: 'Contrast' },
];

export const LOCALES = [
  { value: 'en', label: 'EN' },
  { value: 'ja', label: 'JA' },
] as const;

/** The Microsoft 365 baseline plus the optional tabs this demo actually
 *  wires: the Automate tab is backed by the demo's Script menu and dialog. */
export const DEMO_RIBBON_TABS: readonly RibbonTab[] = [
  ...EXCEL365_STANDARD_RIBBON_TABS,
  'automate',
];

/** The Mac tab strip follows the desktop Microsoft 365 surface.  The title
 * bar owns the backstage entry, so the ribbon deliberately has no File tab. */
export const DEMO_MAC_RIBBON_TABS: readonly RibbonTab[] = EXCEL365_MAC_RIBBON_TABS;

const isDemoPlatformInput = (value: string | null | undefined): value is DemoPlatformInput =>
  value === 'auto' || value === 'mac' || value === 'default';

export const resolveInitialPlatform = (
  search = globalThis.location?.search ?? '',
): DemoPlatform => {
  const requested = new URLSearchParams(search).get('platform');
  return resolveSpreadsheetPlatform(isDemoPlatformInput(requested) ? requested : 'auto');
};

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

export type DemoIconName = 'app' | 'save' | 'undo' | 'redo' | 'search';

export type DemoIconSegment = {
  d: string;
  fill?: string;
  stroke?: string;
};

export const DEMO_ICONS: Record<DemoIconName, readonly DemoIconSegment[]> = {
  app: [
    { d: 'M4.5 4.5h11v11h-11z', stroke: '#ffffff' },
    { d: 'M4.5 10h11M10 4.5v11', stroke: '#ffffff' },
    { d: 'M5.3 5.3h4.2v4.2h-4.2z', fill: '#107c41' },
  ],
  save: [
    { d: 'M4 4h10l2 2v10H4z', fill: '#ffffff', stroke: '#1f1f1f' },
    { d: 'M7 4v5h6V4', fill: '#2f75b5', stroke: '#1f4e79' },
    { d: 'M7 12.6h6v3H7z', fill: '#eef6ee', stroke: '#8a8f98' },
    { d: 'M8.5 14h3', stroke: '#8a8f98' },
  ],
  undo: [
    { d: 'M7.2 5.2H3.8v-3.4', stroke: '#107c41' },
    {
      d: 'M4 5.2c2.2-2.1 5.7-2.3 8.1-.5 2.7 2.1 3 6.1.7 8.6-1.8 1.9-4.8 2.4-7.1 1.2',
      stroke: '#1f1f1f',
    },
  ],
  redo: [
    { d: 'M12.8 5.2h3.4v-3.4', stroke: '#107c41' },
    {
      d: 'M16 5.2c-2.2-2.1-5.7-2.3-8.1-.5-2.7 2.1-3 6.1-.7 8.6 1.8 1.9 4.8 2.4 7.1 1.2',
      stroke: '#1f1f1f',
    },
  ],
  search: [
    { d: 'M8.5 14a5.5 5.5 0 1 1 0-11 5.5 5.5 0 0 1 0 11z', fill: '#ffffff', stroke: '#1f1f1f' },
    { d: 'M12.5 12.5L17 17', stroke: '#107c41' },
    { d: 'M6.2 8.5h4.6', stroke: '#2f75b5' },
  ],
};

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
  platform: DemoPlatform = 'default',
): (() => void) => {
  const onKeydown = (event: KeyboardEvent): void => {
    const isMacSearchShortcut =
      platform === 'mac' &&
      event.metaKey &&
      event.ctrlKey &&
      !event.altKey &&
      !event.shiftKey &&
      event.key.toLowerCase() === 'u';
    const isGenericSearchShortcut =
      event.altKey &&
      !event.ctrlKey &&
      !event.metaKey &&
      !event.shiftKey &&
      event.key.toLowerCase() === 'q';
    if (!isMacSearchShortcut && !isGenericSearchShortcut) return;
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
      active.closest('.fc-tb__dlg, .fc-fmtdlg, .fc-tb__menu, .fc-statusbar__chooser')
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
  platform?: DemoPlatformInput;
}): ReturnType<typeof resolveSpreadsheetUiOptions> => {
  const profile: SpreadsheetUiOptions['profile'] =
    input.preset === 'full' ? 'excel365' : input.preset;
  return resolveSpreadsheetUiOptions({
    profile,
    theme: input.theme,
    platform: input.platform,
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
  // Focus the modal now rather than on the next frame: until focus is inside
  // it, Escape still goes to the launcher and the modal ignores it. Callers
  // already run this after the modal is in the DOM.
  (focusableDemoModalItems(root)[0] ?? root).focus({ preventScroll: true });
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
  // One recalc for the whole seed instead of one per written cell.
  wb.withBatchedRecalc(() => {
    seedRows(wb);
  });
};

const seedRows = (wb: WorkbookHandle): void => {
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
