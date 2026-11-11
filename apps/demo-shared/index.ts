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
  CellChange,
  CellEvent,
  CellFormatter,
  FeatureFlags,
  FeatureId,
  FormulaHint,
  ReviewCell,
  RibbonTab,
  ScriptCommand,
  Sheet,
  SpreadsheetInstance,
  ThemeName,
  WorkbookSaveFormat,
} from '@libraz/formulon-cell';
import {
  analyzeAccessibilityCells,
  analyzeSpellingCells,
  EXCEL365_MAC_RIBBON_TABS,
  EXCEL365_STANDARD_RIBBON_TABS,
  isFeatureDefaultOff,
  parseScriptCommand,
  presets,
  resolveSpreadsheetPlatform,
  resolveSpreadsheetUiOptions,
  WORKBOOK_SAVE_FORMATS,
  Workbook,
} from '@libraz/formulon-cell';
import { saveWithPrompts } from '@libraz/formulon-cell/toolbar';
import { nextDemoSearchIndex } from './demo-search.js';
import type { DemoCommandStrings, DemoStrings, DemoUiStrings } from './demo-strings.js';

export * from './demo-actions.js';
export * from './demo-backstage.js';
export * from './demo-document-menu.js';
export * from './demo-function-help.js';
export * from './demo-glyphs.js';
export * from './demo-quick-access.js';
export * from './demo-search.js';
export * from './demo-select.js';
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
  { value: 'en', label: 'English' },
  { value: 'ja', label: '日本語' },
  { value: 'de', label: 'Deutsch' },
  { value: 'fr', label: 'Français' },
  { value: 'zh', label: '中文' },
  { value: 'ko', label: '한국어' },
  { value: 'th', label: 'ไทย' },
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
  instance: SpreadsheetInstance | null;
  bookName: string;
  setUploadStatus: (status: DemoUploadStatus) => void;
  format?: WorkbookSaveFormat;
  documentRef?: Pick<Document, 'body' | 'createElement'>;
  urlApi?: Pick<typeof URL, 'createObjectURL' | 'revokeObjectURL'>;
  setTimeoutFn?: (handler: () => void, timeout: number) => unknown;
}

/** Writes the workbook (asking for the file type first when `format` is omitted), warns before a
 *  file that leaves content out, and downloads it with the extension of the chosen format. */
const downloadDemoWorkbook = async ({
  instance,
  bookName,
  setUploadStatus,
  format,
  documentRef = globalThis.document,
  urlApi = globalThis.URL,
  setTimeoutFn = globalThis.setTimeout.bind(globalThis),
}: SaveDemoWorkbookOptions): Promise<void> => {
  if (!instance) return;
  setUploadStatus('saving');
  try {
    const outcome = await saveWithPrompts(instance, { format });
    if (!outcome) {
      setUploadStatus(null);
      return;
    }
    const { extension, mimeType } = WORKBOOK_SAVE_FORMATS[outcome.format];
    const blob = new Blob([outcome.bytes as BlobPart], { type: mimeType });
    const url = urlApi.createObjectURL(blob);
    const a = documentRef.createElement('a');
    a.href = url;
    a.download = `${bookName}${extension}`;
    documentRef.body.appendChild(a);
    a.click();
    documentRef.body.removeChild(a);
    setTimeoutFn(() => urlApi.revokeObjectURL(url), 1_000);
    setUploadStatus('saved');
  } catch {
    setUploadStatus('error');
  }
};

/** Save: writes the file type `format` (xlsx by default). */
export const saveDemoWorkbookToDownload = (opts: SaveDemoWorkbookOptions): Promise<void> =>
  downloadDemoWorkbook({ ...opts, format: opts.format ?? 'xlsx' });

/** Save As: ask for the file type first, then save like Save. */
export const saveDemoWorkbookAsToDownload = (
  opts: Omit<SaveDemoWorkbookOptions, 'format'>,
): Promise<void> => downloadDemoWorkbook(opts);

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
    { d: 'M4 4h10l2 2v10H4z', fill: '#ffffff' },
    { d: 'M7 4v5h6V4', fill: '#2f75b5', stroke: '#1f4e79' },
    { d: 'M7 12.6h6v3H7z', fill: '#eef6ee', stroke: '#8a8f98' },
    { d: 'M8.5 14h3', stroke: '#8a8f98' },
  ],
  undo: [
    { d: 'M7.2 5.2H3.8v-3.4', stroke: '#107c41' },
    {
      d: 'M4 5.2c2.2-2.1 5.7-2.3 8.1-.5 2.7 2.1 3 6.1.7 8.6-1.8 1.9-4.8 2.4-7.1 1.2',
    },
  ],
  redo: [
    { d: 'M12.8 5.2h3.4v-3.4', stroke: '#107c41' },
    {
      d: 'M16 5.2c-2.2-2.1-5.7-2.3-8.1-.5-2.7 2.1-3 6.1-.7 8.6 1.8 1.9 4.8 2.4 7.1 1.2',
    },
  ],
  search: [
    { d: 'M8.5 14a5.5 5.5 0 1 1 0-11 5.5 5.5 0 0 1 0 11z' },
    { d: 'M12.5 12.5L17 17', stroke: '#107c41' },
    { d: 'M6.2 8.5h4.6', stroke: '#2f75b5' },
  ],
};

export type DemoLocale = (typeof LOCALES)[number]['value'];

const isDemoLocale = (value: string | null | undefined): value is DemoLocale =>
  LOCALES.some((locale) => locale.value === value);

export const normalizeDemoLocale = (value: string | null | undefined): DemoLocale | undefined => {
  if (!value) return undefined;
  const language = value.trim().toLowerCase().replaceAll('_', '-').split('-')[0];
  return isDemoLocale(language) ? language : undefined;
};

/** Resolve the demo chrome dictionary for a locale already accepted by the
 *  selector, while keeping invalid external values deterministic. */
export const demoUiStringsForLocale = (strings: DemoStrings, locale: string): DemoUiStrings =>
  strings[normalizeDemoLocale(locale) ?? 'ja'];

export const resolveInitialLocale = (
  search = globalThis.location?.search ?? '',
  languages = globalThis.navigator?.languages ?? [globalThis.navigator?.language ?? ''],
): DemoLocale => {
  const param = normalizeDemoLocale(new URLSearchParams(search).get('locale'));
  if (param) return param;
  for (const lang of languages) {
    const resolved = normalizeDemoLocale(lang);
    if (resolved) return resolved;
  }
  return 'ja';
};

export const installDemoSearchShortcut = (
  getInput: () => HTMLInputElement | null | undefined,
  platform: DemoPlatform = 'default',
  /** Shows a search field that is not rendered yet, then focuses it itself. */
  openSearch?: () => void,
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
    if (!input && openSearch) {
      event.preventDefault();
      openSearch();
      return;
    }
    if (!input || input.disabled || input.hidden) return;
    event.preventDefault();
    input.focus();
    input.select();
  };
  document.addEventListener('keydown', onKeydown);
  return () => document.removeEventListener('keydown', onKeydown);
};

/** The ribbon as F6 navigation sees it: the element it lives in and its tab focus. */
export interface DemoF6Toolbar {
  readonly host: HTMLElement;
  focusActiveTab(): boolean;
}

export interface DemoF6NavigationTargets {
  getQuickAccess: () => HTMLElement | null | undefined;
  getToolbar: () => DemoF6Toolbar | null;
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
  platform?: DemoPlatformInput;
}): ReturnType<typeof resolveSpreadsheetUiOptions> =>
  resolveSpreadsheetUiOptions({
    profile: input.preset,
    platform: input.platform,
    features: { ...input.overrides, ribbon: input.showRibbon },
  });

export const FEATURE_GROUPS: {
  title: string;
  features: { id: FeatureId; label: string }[];
}[] = [
  {
    title: 'Chrome',
    features: [
      { id: 'localeSwitcher', label: 'Language selector' },
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

/** Sample formula hints for the host-defined function names `GREET` and `FAHRENHEIT`. They feed
 *  autocomplete and the argument helper only; the calculator does not evaluate them. */
export const DEMO_FORMULA_HINTS: readonly FormulaHint[] = [
  { name: 'GREET', args: ['name'], description: 'Friendly greeting' },
  { name: 'FAHRENHEIT', args: ['celsius'], description: 'Celsius to Fahrenheit' },
];

/** Signature text of a hint, e.g. `GREET(name)`. */
export const demoFormulaHintSignature = (hint: FormulaHint): string =>
  `${hint.name}(${(hint.args ?? []).join(', ')})`;

/** `kind` tags of a cell `Value` the demo reads or writes. */
const POINTS_PER_CSS_PX = 0.75;

const VALUE_KIND = { blank: 0, number: 1, text: 3 } as const;

/** Custom cell formatters demonstrating the formatter registry. */
export const FORMATTERS = {
  uppercaseA: {
    id: 'demo:uppercaseA',
    match: (i) => i.at.col === 0 && i.value.kind === VALUE_KIND.text,
    format: (i) => (i.value.kind === VALUE_KIND.text ? i.value.text.toUpperCase() : null),
  },
  arrowNegatives: {
    id: 'demo:arrowNegatives',
    match: (i) => i.value.kind === VALUE_KIND.number && i.value.number < 0,
    format: (i) =>
      i.value.kind === VALUE_KIND.number ? `↓ ${Math.abs(i.value.number).toFixed(2)}` : null,
  },
} satisfies Record<string, CellFormatter>;

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

/** Formats a cell change for the demo's change log (favours the formula text
 *  when present so users see what they typed). */
export const previewCellChange = (workbook: Workbook, e: CellEvent): string => {
  const sheet = workbook.sheet(e.sheet);
  const formula = sheet.getFormula(e.at.row, e.at.col);
  if (formula) return formula;
  const value = sheet.getValue(e.at.row, e.at.col);
  if (value === null) return '∅';
  switch (typeof value) {
    case 'number':
    case 'boolean':
      return String(value);
    case 'string':
      return JSON.stringify(value);
    default:
      return Array.isArray(value) ? '?' : String(value);
  }
};

/** Demo seed — only runs once on the initial blank workbook. Core gates
 *  `seed` on `ownsWb`, so re-mounts and Open xlsx don't re-trigger it. */
export const seedDemoWorkbook = (wb: Workbook): void => {
  // One recalc for the whole seed instead of one per written cell.
  wb.batch(() => {
    seedRows(wb.sheet(0));
  });
};

const seedRows = (sheet: Sheet): void => {
  sheet.setText(0, 0, 'item');
  sheet.setText(0, 1, 'celsius');
  sheet.setText(0, 2, 'fahrenheit');
  sheet.setText(0, 3, 'greeting');
  const rows: [string, number][] = [
    ['London', 8],
    ['Tokyo', 22],
    ['Reykjavík', -3],
    ['Cairo', 31],
  ];
  rows.forEach(([city, c], i) => {
    const r = i + 1;
    sheet.setText(r, 0, city);
    sheet.setNumber(r, 1, c);
    sheet.setFormula(r, 2, `=B${r + 1}*1.8+32`);
    sheet.setFormula(r, 3, `=A${r + 1}&" ☼"`);
  });
};

/** Projects a workbook's cells for the demo review dialog. */
export const reviewCellsForInstance = (inst: SpreadsheetInstance): ReviewCell[] =>
  inst.workbook
    .sheet(inst.selection.get().sheet)
    .listCells()
    .map((cell) => ({
      label: `${demoColLabel(cell.col)}${cell.row + 1}`,
      value:
        cell.value === null
          ? { kind: 'blank' as const }
          : typeof cell.value === 'string'
            ? { kind: 'text' as const, value: cell.value }
            : typeof cell.value === 'number'
              ? { kind: 'number' as const }
              : typeof cell.value === 'boolean'
                ? { kind: 'bool' as const }
                : { kind: 'error' as const, text: String(cell.value) },
      formula: cell.formula,
    }));

/** Creates the demo's initial workbook; the caller owns it. Core only auto-seeds when it owns the
 *  workbook, so the demo seeds by hand; `?fixture=empty` (deterministic blank
 *  workbook for E2E specs) skips the seed. */
export const createInitialDemoWorkbook = async (
  search = globalThis.location?.search ?? '',
): Promise<Workbook> => {
  const wb = await Workbook.createDefault({
    locale: resolveInitialLocale(search),
    platform: resolveInitialPlatform(search),
  });
  if (new URLSearchParams(search).get('fixture') !== 'empty') {
    seedDemoWorkbook(wb);
    // Seeded widths are CSS px, converted to the character units `setColumnWidth` takes.
    const sheet = wb.sheet(0);
    const seededColumnWidthsPx = [112, 80, 104, 128];
    for (const [col, px] of seededColumnWidthsPx.entries()) {
      sheet.setColumnWidth(col, col, sheet.convertColumnPtToChars(px * POINTS_PER_CSS_PX));
    }
  }
  return wb;
};

/** Reads an xlsx/xlsm/xlsb file into a new workbook the caller owns, and returns it with its
 *  display name (the filename without its extension). Rejects on load failure. */
export const openDemoWorkbookFile = async (
  file: File,
): Promise<{ workbook: Workbook; name: string }> => ({
  workbook: await Workbook.load(new Uint8Array(await file.arrayBuffer())),
  name: file.name.replace(/\.(xlsx|xlsm)$/i, ''),
});

/** Shows `next` in `live`. Settles on the instance's `workbook` event for it, or rejects with the
 *  swap's synchronous refusal or its `error` event. */
const showWorkbook = (live: SpreadsheetInstance, next: Workbook): Promise<void> =>
  new Promise<void>((resolve, reject) => {
    if (live.workbook === next) {
      resolve();
      return;
    }
    const offs: (() => void)[] = [];
    const settle = (done: () => void): void => {
      for (const off of offs.splice(0)) off();
      done();
    };
    offs.push(
      live.on('workbook', (event) => {
        if (event.workbook === next) settle(resolve);
      }),
      // `workbook` is the source the instance reports a refused workbook swap under.
      live.on('error', (event) => {
        if (event.source === 'workbook') settle(() => reject(event.error));
      }),
    );
    try {
      live.update({ workbook: next });
    } catch (error) {
      settle(() => reject(error));
    }
  });

/**
 * Hands a workbook the demo owns to the live instance. Resolves `true` once `live` shows it, after
 * which the caller owns it and releases the workbook it replaced. Otherwise `next` is disposed
 * here: resolves `false` when `isCurrent()` turns false before or after the swap (the demo or
 * the instance went away), and rejects with the swap's refusal.
 */
export const adoptDemoWorkbook = async (
  live: SpreadsheetInstance | null,
  next: Workbook,
  isCurrent: () => boolean,
): Promise<boolean> => {
  try {
    if (!isCurrent()) {
      next.dispose();
      return false;
    }
    if (live) await showWorkbook(live, next);
  } catch (error) {
    next.dispose();
    throw error;
  }
  if (!isCurrent()) {
    next.dispose();
    return false;
  }
  return true;
};

/** Routes clicks on the built-in `#menu-script` items to `onCommand`. The
 *  toolbar owns the rest of the click (closing the menu, restoring focus and
 *  routing `custom` to its own dialog), so `custom` is ignored here. */
export const installDemoScriptMenu = (
  onCommand: (command: ScriptCommand) => void,
): (() => void) => {
  const onMenuClick = (e: MouseEvent): void => {
    const target = e.target;
    if (!(target instanceof Element)) return;
    const btn = target.closest<HTMLButtonElement>('[data-script-action]');
    if (!btn) return;
    if (!btn.closest<HTMLDivElement>('#menu-script')) return;
    const action = btn.dataset.scriptAction ?? '';
    if (action === 'custom') return;
    const command = parseScriptCommand(action);
    if (command) onCommand(command);
  };
  document.addEventListener('click', onMenuClick);
  return () => document.removeEventListener('click', onMenuClick);
};

const transformDemoScriptText = (
  value: string,
  command: Exclude<ScriptCommand, 'clear'>,
): string => {
  switch (command) {
    case 'uppercase':
      return value.toUpperCase();
    case 'lowercase':
      return value.toLowerCase();
    case 'trim':
      return value.trim().replace(/\s+/g, ' ');
  }
};

/** Applies a text-script command to the selection as one undo step (protected
 *  cells are skipped) and returns the number of changed cells. */
export const runDemoScript = (inst: SpreadsheetInstance, command: ScriptCommand): number => {
  const { sheet, range } = inst.selection.get();
  const changes = new Map<string, CellChange>();
  for (const cell of inst.workbook.sheet(sheet).listCells(range)) {
    const at = { row: cell.row, col: cell.col };
    const key = `${cell.row}:${cell.col}`;
    if (command === 'clear') {
      changes.set(key, { sheet, at, value: { kind: VALUE_KIND.blank }, formula: null });
    } else if (typeof cell.value === 'string') {
      const next = transformDemoScriptText(cell.value, command);
      if (next !== cell.value) {
        changes.set(key, {
          sheet,
          at,
          value: { kind: VALUE_KIND.text, text: next },
          formula: null,
        });
      }
    }
  }
  // A protected sheet refuses the whole batch and names the locked cell that decided it; drop
  // that cell and write the rest.
  let result = inst.applyChanges([...changes.values()], { history: 'record' });
  while (result.status === 'refused') {
    const at = result.rejections[0]?.at;
    if (!at || !changes.delete(`${at.row}:${at.col}`) || changes.size === 0) return 0;
    result = inst.applyChanges([...changes.values()], { history: 'record' });
  }
  return result.applied.length;
};

/** A few features ship default-off; everything else is opt-out. */
export const isDemoFeatureOn = (features: FeatureFlags, id: FeatureId): boolean =>
  isFeatureDefaultOff(id) ? features[id] === true : features[id] !== false;

/** Next override map after toggling `id`. Toggling back to the preset's
 *  default drops the override so the preset's value wins. */
export const nextDemoFeatureOverrides = (input: {
  preset: PresetKey;
  overrides: FeatureFlags;
  features: FeatureFlags;
  id: FeatureId;
}): FeatureFlags => {
  const { id } = input;
  const presetFlags = presets[input.preset]();
  const presetDefault = isFeatureDefaultOff(id)
    ? presetFlags[id] === true
    : presetFlags[id] !== false;
  const nextVal = !isDemoFeatureOn(input.features, id);
  const next = { ...input.overrides };
  if (nextVal === presetDefault) {
    delete next[id];
  } else {
    next[id] = nextVal;
  }
  return next;
};

/** Status label for the selection: `B3`, or `B3:D5 · 9` for a range. */
export const demoSelectionLabel = (selection: {
  active: { row: number; col: number };
  range: { r0: number; c0: number; r1: number; c1: number };
}): string => {
  const { active, range } = selection;
  if (range.r0 === range.r1 && range.c0 === range.c1) {
    return `${demoColLabel(active.col)}${active.row + 1}`;
  }
  const tl = `${demoColLabel(range.c0)}${range.r0 + 1}`;
  const br = `${demoColLabel(range.c1)}${range.r1 + 1}`;
  const cells = (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);
  return `${tl}:${br} · ${cells}`;
};

export interface DemoReviewDialogState {
  readonly title: string;
  readonly items: readonly { label: string; detail: string }[];
}

export const buildDemoReviewDialog = (
  title: string,
  label: string,
  detail: string,
): DemoReviewDialogState => ({ title, items: [{ label, detail }] });

/** Review dialog state for the ribbon's spelling / accessibility checks. */
export const buildDemoReviewFindings = (
  kind: 'spelling' | 'accessibility',
  inst: SpreadsheetInstance,
  locale: string,
  commandText: DemoCommandStrings,
): DemoReviewDialogState => {
  const cells = reviewCellsForInstance(inst);
  const lang = normalizeDemoLocale(locale) ?? 'ja';
  return kind === 'spelling'
    ? { title: commandText.spellingReview, items: analyzeSpellingCells(cells, lang) }
    : { title: commandText.accessibilityCheck, items: analyzeAccessibilityCells(cells, lang) };
};

/** Runs a script command on the selection and reports the changed-cell count
 *  as a review dialog. */
export const reportDemoScriptRun = (
  inst: SpreadsheetInstance,
  command: ScriptCommand,
  commandText: DemoCommandStrings,
): DemoReviewDialogState =>
  buildDemoReviewDialog(
    commandText.script,
    commandText.selection,
    commandText.cellsUpdated.replace('{count}', String(runDemoScript(inst, command))),
  );

export interface ChangeLogEntry {
  readonly id: number;
  readonly cell: string;
  readonly preview: string;
}

export const DEMO_CHANGE_LOG_LIMIT = 8;

let demoChangeId = 0;

export const demoChangeLogEntry = (workbook: Workbook, e: CellEvent): ChangeLogEntry => ({
  id: ++demoChangeId,
  cell: `${demoColLabel(e.at.col)}${e.at.row + 1}`,
  preview: previewCellChange(workbook, e),
});

/** Prepends the change to the log, keeping the newest `DEMO_CHANGE_LOG_LIMIT`. */
export const pushDemoChangeLog = (
  log: readonly ChangeLogEntry[],
  entry: ChangeLogEntry,
): ChangeLogEntry[] => [entry, ...log].slice(0, DEMO_CHANGE_LOG_LIMIT);

export type DemoSearchKeyAction =
  | { kind: 'close' }
  | { kind: 'move'; index: number }
  | { kind: 'run'; index: number }
  | null;

/** Maps a keydown on the demo search input to an action. `move` and `run`
 *  keys also need `preventDefault()` from the caller. */
export const resolveDemoSearchKey = (
  key: string,
  activeIndex: number,
  count: number,
): DemoSearchKeyAction => {
  if (key === 'Escape') return { kind: 'close' };
  if (key === 'ArrowDown' || key === 'ArrowUp') {
    return {
      kind: 'move',
      index: nextDemoSearchIndex(activeIndex, count, key === 'ArrowDown' ? 'next' : 'previous'),
    };
  }
  if (key === 'Enter' && count > 0) {
    return { kind: 'run', index: nextDemoSearchIndex(activeIndex, count, 'first') };
  }
  return null;
};

export interface DemoSearchState {
  query: string;
  open: boolean;
  activeIndex: number;
}

export type DemoSearchEvent =
  | { type: 'focus' }
  | { type: 'blur' }
  /** `focused` is whether the input still owns focus; Escape clears a search
   *  input natively and that change must not reopen the list. */
  | { type: 'input'; value: string; focused: boolean }
  | { type: 'hover'; index: number }
  | { type: 'key'; action: NonNullable<DemoSearchKeyAction> }
  /** A command ran: clear the query and close the list. */
  | { type: 'reset' };

export const INITIAL_DEMO_SEARCH_STATE: DemoSearchState = {
  query: '',
  open: false,
  activeIndex: -1,
};

/** Pure transition for the demo search box, shared by the React and Vue title bars. */
export const reduceDemoSearch = (
  state: DemoSearchState,
  event: DemoSearchEvent,
): DemoSearchState => {
  switch (event.type) {
    case 'focus':
      return { ...state, open: true, activeIndex: -1 };
    case 'blur':
      return { ...state, open: false };
    case 'input':
      return { query: event.value, open: event.focused, activeIndex: -1 };
    case 'hover':
      return { ...state, activeIndex: event.index };
    case 'key':
      if (event.action.kind === 'move') {
        return { ...state, open: true, activeIndex: event.action.index };
      }
      if (event.action.kind === 'close') return { ...state, open: false, activeIndex: -1 };
      return state;
    case 'reset':
      return INITIAL_DEMO_SEARCH_STATE;
  }
};
