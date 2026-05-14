<script setup lang="ts">
import {
  type CellChangeEvent,
  type CellRenderInput,
  type CellValue,
  type FeatureFlags,
  type FeatureId,
  presets,
  type SpreadsheetInstance,
  type ThemeName,
  WorkbookHandle,
} from '@libraz/formulon-cell';
import { type RibbonTab, Spreadsheet, useSelection } from '@libraz/formulon-cell-vue';
import SpreadsheetToolbar from '@libraz/formulon-cell-vue/toolbar.vue';
import { computed, onUnmounted, ref, shallowRef, watch } from 'vue';

const THEMES: { value: ThemeName; label: string }[] = [
  { value: 'paper', label: 'Light' },
  { value: 'ink', label: 'Dark' },
  { value: 'contrast', label: 'Contrast' },
];
const LOCALES = [
  { value: 'en', label: 'EN' },
  { value: 'ja', label: 'JA' },
];

const formatLoadError = (err: unknown): string => (err instanceof Error ? err.message : String(err));

const UI = {
  en: {
    saved: 'Saved to this device',
    search: 'Search',
    share: 'Share',
    workbook: 'Vue workbook',
    demoPane: 'Options',
    open: 'Open xlsx…',
    save: 'Save',
    file: 'File',
    info: 'Info',
    print: 'Print',
    pageSetup: 'Page Setup',
    close: 'Close',
    backstageSub: 'Vue workbook · full spreadsheet layout',
    openTitle: 'Open',
    openDesc: 'Load an .xlsx or .xlsm workbook from this device.',
    saveCopy: 'Save a Copy',
    saveDesc: 'Download the current workbook as an .xlsx file.',
    printDesc: 'Use the browser print dialog or save as PDF.',
    pageSetupDesc: 'Set orientation, margins, paper size, headers, and print titles.',
    editLinks: 'Edit Links',
    linksDesc: 'Inspect external workbook references carried by the file.',
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
    share: '共有',
    workbook: 'Vue ブック',
    demoPane: 'オプション',
    open: 'xlsx を開く…',
    save: '保存',
    file: 'ファイル',
    info: '情報',
    print: '印刷',
    pageSetup: 'ページ設定',
    close: '閉じる',
    backstageSub: 'Vue ブック · スプレッドシート レイアウト',
    openTitle: '開く',
    openDesc: '.xlsx または .xlsm ブックをこのデバイスから読み込みます。',
    saveCopy: 'コピーを保存',
    saveDesc: '現在のブックを .xlsx ファイルとしてダウンロードします。',
    printDesc: 'ブラウザーの印刷ダイアログ、または PDF 保存を使用します。',
    pageSetupDesc: '用紙方向、余白、用紙サイズ、ヘッダー、印刷タイトルを設定します。',
    editLinks: 'リンクの編集',
    linksDesc: 'ファイルに含まれる外部ブック参照を確認します。',
    options: 'オプション',
    optionsDesc: '統合パネルと機能トグルを表示します。',
    noCommands: 'コマンドが見つかりません',
    engineUnavailable: 'スプレッドシートエンジンを起動できません',
    engineSetup:
      'SharedArrayBuffer を有効にするため、COOP: same-origin と COEP: require-corp 付きで配信してください。',
  },
} as const;

type PresetKey = 'minimal' | 'standard' | 'full';
const PRESETS: { value: PresetKey; label: string; hint: string }[] = [
  { value: 'minimal', label: 'Minimal', hint: 'bare spreadsheet chrome' },
  { value: 'standard', label: 'Standard', hint: 'lightweight editing chrome' },
  { value: 'full', label: 'Full', hint: 'complete spreadsheet chrome' },
];

const FEATURE_GROUPS: { title: string; features: { id: FeatureId; label: string }[] }[] = [
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

const colLabel = (n: number): string => {
  let out = '';
  let v = n;
  do {
    out = String.fromCharCode(65 + (v % 26)) + out;
    v = Math.floor(v / 26) - 1;
  } while (v >= 0);
  return out;
};

const DEMO_FUNCTIONS = [
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

const FORMATTERS = {
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

interface ChangeLogEntry {
  readonly id: number;
  readonly cell: string;
  readonly preview: string;
}

interface CommandItem {
  readonly id: string;
  readonly label: string;
  readonly hint: string;
  readonly tab?: RibbonTab;
  readonly run: () => void;
}

let changeId = 0;

const previewValue = (e: CellChangeEvent): string => {
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

// Demo seed — only runs once on the initial blank workbook (core gates
// `seed` on `ownsWb`).
const seed = (wb: WorkbookHandle): void => {
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

const composeFeatures = (preset: PresetKey, overrides: FeatureFlags): FeatureFlags => ({
  ...presets[preset](),
  ...overrides,
});

const theme = ref<ThemeName>('paper');
const locale = ref<string>('en');
const workbook = shallowRef<WorkbookHandle | null>(null);
// Vue's reactive proxy walks deeply by default; the spreadsheet instance
// holds a canvas + many internal refs that should not be reactivified.
const instance = shallowRef<SpreadsheetInstance | null>(null);
const log = ref<ChangeLogEntry[]>([]);
const formatters = ref({ uppercase: true, arrows: true });
const probe = ref<{ name: string; result: string } | null>(null);
const fileInput = ref<HTMLInputElement | null>(null);
const preset = ref<PresetKey>('full');
const overrides = ref<FeatureFlags>({});
const showRibbon = ref(true);
const showPanel = ref(false);
const ribbonTab = ref<RibbonTab>('home');
const searchQuery = ref('');
const searchOpen = ref(false);
const bookName = ref('Book1');
const loadError = ref<string | null>(null);

const features = computed<FeatureFlags>(() => composeFeatures(preset.value, overrides.value));
const ui = computed(() => UI[locale.value === 'ja' ? 'ja' : 'en']);

void WorkbookHandle.createDefault()
  .then((wb) => {
    // Core only auto-seeds when it owns the workbook (no `workbook` prop).
    // The demo passes a pre-built handle, so seed by hand here. `?fixture=empty`
    // (used by E2E specs that need a deterministic blank workbook) skips this.
    const fx = new URLSearchParams(window.location.search).get('fixture');
    if (fx !== 'empty') seed(wb);
    loadError.value = null;
    workbook.value = wb;
  })
  .catch((err: unknown) => {
    loadError.value = formatLoadError(err);
  });

watch(
  [instance, () => formatters.value.uppercase, () => formatters.value.arrows],
  (_n, _o, onCleanup) => {
    const inst = instance.value;
    if (!inst) return;
    const disposers: (() => void)[] = [];
    if (formatters.value.uppercase) {
      disposers.push(inst.cells.registerFormatter(FORMATTERS.uppercaseA));
    }
    if (formatters.value.arrows) {
      disposers.push(inst.cells.registerFormatter(FORMATTERS.arrowNegatives));
    }
    onCleanup(() => {
      for (const d of disposers) d();
    });
  },
  { immediate: true },
);

const onCellChange = (e: CellChangeEvent): void => {
  const cell = `${colLabel(e.addr.col)}${e.addr.row + 1}`;
  log.value = [{ id: ++changeId, cell, preview: previewValue(e) }, ...log.value].slice(0, 8);
};

const onReady = (inst: SpreadsheetInstance): void => {
  instance.value = inst;
  // Expose the live instance on `window.__fcInst` so cross-demo E2E scenarios
  // can drive imperative paths without depending on demo-specific UI.
  (window as unknown as { __fcInst?: SpreadsheetInstance | null }).__fcInst = inst;
};

const selection = useSelection(instance);
const selectionLabel = computed(() => {
  const { active, range } = selection.value;
  if (range.r0 === range.r1 && range.c0 === range.c1) {
    return `${colLabel(active.col)}${active.row + 1}`;
  }
  const tl = `${colLabel(range.c0)}${range.r0 + 1}`;
  const br = `${colLabel(range.c1)}${range.r1 + 1}`;
  const cells = (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);
  return `${tl}:${br} · ${cells}`;
});

const runProbe = (name: string, args: CellValue[]): void => {
  const inst = instance.value;
  if (!inst) return;
  try {
    const out = inst.formula.evaluate(name, args);
    const display =
      out.kind === 'number'
        ? out.value.toString()
        : out.kind === 'text'
          ? out.value
          : JSON.stringify(out);
    probe.value = { name, result: display };
  } catch (err) {
    probe.value = { name, result: err instanceof Error ? err.message : String(err) };
  }
};

const onSave = (): void => {
  const inst = instance.value;
  if (!inst) return;
  const bytes = inst.workbook.save();
  const blob = new Blob([bytes as BlobPart], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${bookName.value}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1_000);
};

const onOpenFiles = async (ev: Event): Promise<void> => {
  const target = ev.target as HTMLInputElement;
  const file = target.files?.[0];
  if (!file) return;
  target.value = '';
  const inst = instance.value;
  if (!inst) return;
  try {
    const buf = await file.arrayBuffer();
    const next = await WorkbookHandle.loadBytes(new Uint8Array(buf));
    await inst.setWorkbook(next);
    loadError.value = null;
    bookName.value = file.name.replace(/\.(xlsx|xlsm)$/i, '');
  } catch (err) {
    window.alert(formatLoadError(err));
  }
};

const onPresetChange = (next: PresetKey): void => {
  if (next === preset.value) return;
  preset.value = next;
  overrides.value = {};
};

const onFeatureToggle = (id: FeatureId): void => {
  const presetFlags = presets[preset.value]();
  const defaultOff = id === 'watchWindow' || id === 'slicer';
  const presetDefault = defaultOff ? presetFlags[id] === true : presetFlags[id] !== false;
  const currentVal = isFeatureOn(id);
  const nextVal = !currentVal;
  const nextOverrides: FeatureFlags = { ...overrides.value };
  if (nextVal === presetDefault) {
    delete nextOverrides[id];
  } else {
    nextOverrides[id] = nextVal;
  }
  overrides.value = nextOverrides;
};

// `watchWindow` and `slicer` ship default-off; everything else is opt-out.
const isFeatureOn = (id: FeatureId): boolean =>
  id === 'watchWindow' || id === 'slicer'
    ? features.value[id] === true
    : features.value[id] !== false;

const commands = computed<CommandItem[]>(() => [
  {
    id: 'open',
    label: 'Open',
    hint: 'Open an xlsx or xlsm workbook',
    tab: 'file',
    run: () => fileInput.value?.click(),
  },
  {
    id: 'save',
    label: 'Save',
    hint: 'Download the workbook as xlsx',
    tab: 'file',
    run: onSave,
  },
  {
    id: 'page-setup',
    label: 'Page Setup',
    hint: 'Open page setup',
    tab: 'file',
    run: () => instance.value?.openPageSetup(),
  },
  {
    id: 'print',
    label: 'Print',
    hint: 'Open browser print dialog',
    tab: 'file',
    run: () => instance.value?.print(),
  },
  {
    id: 'format-cells',
    label: 'Format Cells',
    hint: 'Open the format dialog',
    tab: 'home',
    run: () => instance.value?.openFormatDialog(),
  },
  {
    id: 'conditional',
    label: 'Conditional Formatting',
    hint: 'Create or edit conditional formatting',
    tab: 'insert',
    run: () => instance.value?.openConditionalDialog(),
  },
  {
    id: 'cell-styles',
    label: 'Cell Styles',
    hint: 'Open the style gallery',
    tab: 'insert',
    run: () => instance.value?.openCellStylesGallery(),
  },
  {
    id: 'name-manager',
    label: 'Name Manager',
    hint: 'Inspect named ranges',
    tab: 'insert',
    run: () => instance.value?.openNamedRangeDialog(),
  },
  {
    id: 'insert-function',
    label: 'Insert Function',
    hint: 'Open function arguments',
    tab: 'formulas',
    run: () => instance.value?.openFunctionArguments(),
  },
  {
    id: 'trace-precedents',
    label: 'Trace Precedents',
    hint: 'Show precedent arrows',
    tab: 'formulas',
    run: () => instance.value?.tracePrecedents(),
  },
  {
    id: 'watch-window',
    label: 'Watch Window',
    hint: 'Toggle Watch Window',
    tab: 'formulas',
    run: () => instance.value?.toggleWatchWindow(),
  },
  {
    id: 'filter',
    label: 'Filter',
    hint: 'Show the Data tab filter tools',
    tab: 'data',
    run: () => {
      ribbonTab.value = 'data';
    },
  },
  {
    id: 'sort',
    label: 'Sort',
    hint: 'Show sort buttons',
    tab: 'data',
    run: () => {
      ribbonTab.value = 'data';
    },
  },
  {
    id: 'freeze-panes',
    label: 'Freeze Panes',
    hint: 'Show Freeze Panes',
    tab: 'view',
    run: () => {
      ribbonTab.value = 'view';
    },
  },
  {
    id: 'protect-sheet',
    label: 'Protect Sheet',
    hint: 'Toggle sheet protection from View',
    tab: 'view',
    run: () => instance.value?.toggleSheetProtection(),
  },
  {
    id: 'options-pane',
    label: 'Options',
    hint: 'Show or hide the integration panel',
    run: () => {
      showPanel.value = !showPanel.value;
    },
  },
  {
    id: 'theme-light',
    label: 'Light Theme',
    hint: 'Switch to light workbook theme',
    run: () => {
      theme.value = 'paper';
    },
  },
  {
    id: 'theme-dark',
    label: 'Dark Theme',
    hint: 'Switch to dark workbook theme',
    run: () => {
      theme.value = 'ink';
    },
  },
  {
    id: 'locale-ja',
    label: 'Japanese Locale',
    hint: 'Switch labels to JA',
    run: () => {
      locale.value = 'ja';
    },
  },
  {
    id: 'locale-en',
    label: 'English Locale',
    hint: 'Switch labels to EN',
    run: () => {
      locale.value = 'en';
    },
  },
]);

const filteredCommands = computed(() => {
  const q = searchQuery.value.trim().toLowerCase();
  if (!q) return commands.value.slice(0, 8);
  return commands.value
    .filter((cmd) => `${cmd.label} ${cmd.hint}`.toLowerCase().includes(q))
    .slice(0, 8);
});

const runCommand = (cmd: CommandItem): void => {
  if (cmd.tab) ribbonTab.value = cmd.tab;
  cmd.run();
  searchQuery.value = '';
  searchOpen.value = false;
};

const onSearchKeydown = (ev: KeyboardEvent): void => {
  if (ev.key === 'Escape') {
    searchOpen.value = false;
    (ev.currentTarget as HTMLInputElement).blur();
  }
  if (ev.key === 'Enter' && filteredCommands.value[0]) {
    ev.preventDefault();
    runCommand(filteredCommands.value[0]);
  }
};

onUnmounted(() => {
  // The Spreadsheet component disposes itself; nothing extra to clean up.
});
</script>

<template>
  <div v-if="!workbook" class="demo demo--loading">
    <div v-if="loadError" class="demo__load-error" role="alert">
      <strong>{{ ui.engineUnavailable }}</strong>
      <span>{{ ui.engineSetup }}</span>
      <code>{{ loadError }}</code>
    </div>
    <template v-else>Loading engine…</template>
  </div>
  <div v-else class="demo" :data-theme="theme">
    <header class="demo__head">
      <div class="demo__titlebar">
        <div class="demo__quick" role="toolbar" aria-label="Quick access toolbar">
          <span class="demo__brand-mark">⊞</span>
          <button type="button" class="demo__title-icon" aria-label="Save" @click="onSave">
            <svg class="demo__rb-icon" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.45" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
              <path d="M4 4h10l2 2v10H4z" />
              <path d="M7 4v5h6V4" />
              <path d="M7 13h6" />
            </svg>
          </button>
          <button type="button" class="demo__title-icon" aria-label="Undo" @click="instance?.undo()">
            <svg class="demo__rb-icon" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.45" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
              <path d="M7.2 5.2H3.8v-3.4" />
              <path d="M4 5.2c2.2-2.1 5.7-2.3 8.1-.5 2.7 2.1 3 6.1.7 8.6-1.8 1.9-4.8 2.4-7.1 1.2" />
            </svg>
          </button>
          <button type="button" class="demo__title-icon" aria-label="Redo" @click="instance?.redo()">
            <svg class="demo__rb-icon" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.45" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
              <path d="M12.8 5.2h3.4v-3.4" />
              <path d="M16 5.2c-2.2-2.1-5.7-2.3-8.1-.5-2.7 2.1-3 6.1-.7 8.6 1.8 1.9 4.8 2.4 7.1 1.2" />
            </svg>
          </button>
        </div>
        <div class="demo__title">
          <strong>{{ bookName }}</strong>
          <span>{{ ui.saved }}</span>
        </div>
        <div class="demo__search">
          <svg class="demo__rb-icon" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.45" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <path d="M8.5 14a5.5 5.5 0 1 1 0-11 5.5 5.5 0 0 1 0 11z" />
            <path d="M12.5 12.5L17 17" />
          </svg>
          <input
            v-model="searchQuery"
            type="search"
            :placeholder="ui.search"
            aria-label="Search commands"
            @focus="searchOpen = true"
            @input="searchOpen = true"
            @keydown="onSearchKeydown"
            @blur="searchOpen = false"
          />
          <div v-if="searchOpen" class="demo__command-menu">
            <div v-if="filteredCommands.length === 0" class="demo__command-empty">
              {{ ui.noCommands }}
            </div>
            <button
              v-for="cmd in filteredCommands"
              v-else
              :key="cmd.id"
              type="button"
              class="demo__command-item"
              @mousedown.prevent
              @click="runCommand(cmd)"
            >
              <strong>{{ cmd.label }}</strong>
              <span>{{ cmd.hint }}</span>
            </button>
          </div>
        </div>
        <div class="demo__account">
          <button type="button" class="demo__share">
            {{ ui.share }}
          </button>
          <span class="demo__avatar" role="img" aria-label="Signed in user">FC</span>
        </div>
      </div>
      <div class="demo__commandbar">
        <div class="demo__brand">
          <strong>formulon-cell</strong>
          <span class="demo__brand-sep">·</span>
          <span class="demo__brand-tag">{{ ui.workbook }}</span>
        </div>
        <div class="demo__controls">
          <div class="demo__seg" role="group" aria-label="Theme">
            <button
              v-for="t in THEMES"
              :key="t.value"
              type="button"
              :class="['demo__seg-btn', { 'demo__seg-btn--active': theme === t.value }]"
              :aria-pressed="theme === t.value"
              @click="theme = t.value"
            >
              {{ t.label }}
            </button>
          </div>
          <div class="demo__seg" role="group" aria-label="Locale">
            <button
              v-for="l in LOCALES"
              :key="l.value"
              type="button"
              :class="['demo__seg-btn', { 'demo__seg-btn--active': locale === l.value }]"
              :aria-pressed="locale === l.value"
              @click="locale = l.value"
            >
              {{ l.label }}
            </button>
          </div>
          <button
            type="button"
            :class="['demo__btn', { 'demo__btn--active': showPanel }]"
            :aria-pressed="showPanel"
            @click="showPanel = !showPanel"
          >
            {{ ui.demoPane }}
          </button>
          <button type="button" class="demo__btn" @click="fileInput?.click()">
            {{ ui.open }}
          </button>
          <button type="button" class="demo__btn" :disabled="!instance" @click="onSave">
            {{ ui.save }}
          </button>
          <input ref="fileInput" type="file" accept=".xlsx,.xlsm" hidden @change="onOpenFiles" />
        </div>
      </div>
    </header>

    <main :class="['demo__body', { 'demo__body--panel': showPanel }]">
      <div class="demo__sheet-col">
        <SpreadsheetToolbar
          v-if="showRibbon"
          :instance="instance"
          :active-tab="ribbonTab"
          :locale="locale"
          @tab-change="ribbonTab = $event"
        />
        <Spreadsheet
          class="demo__sheet"
          :workbook="workbook"
          :theme="theme"
          :locale="locale"
          :features="features"
          :functions="DEMO_FUNCTIONS"
          @ready="onReady"
          @cell-change="onCellChange"
        />
        <div v-if="ribbonTab === 'file'" class="demo__backstage" role="dialog" :aria-label="ui.file">
          <nav class="demo__backstage-nav" :aria-label="ui.file">
            <strong>{{ ui.file }}</strong>
            <button type="button" class="demo__backstage-navitem demo__backstage-navitem--active">
              {{ ui.info }}
            </button>
            <button type="button" class="demo__backstage-navitem" @click="fileInput?.click()">
              {{ ui.openTitle }}
            </button>
            <button type="button" class="demo__backstage-navitem" @click="onSave">
              {{ ui.save }}
            </button>
            <button
              type="button"
              class="demo__backstage-navitem"
              :disabled="!instance"
              @click="instance?.print()"
            >
              {{ ui.print }}
            </button>
            <button
              type="button"
              class="demo__backstage-navitem"
              :disabled="!instance"
              @click="instance?.openPageSetup()"
            >
              {{ ui.pageSetup }}
            </button>
            <button type="button" class="demo__backstage-navitem" @click="ribbonTab = 'home'">
              {{ ui.close }}
            </button>
          </nav>
          <div class="demo__backstage-main">
            <div class="demo__backstage-title">
              <span class="demo__backstage-xl">⊞</span>
              <div>
                <h1>{{ bookName }}</h1>
                <p>{{ ui.backstageSub }}</p>
              </div>
            </div>
            <div class="demo__backstage-grid">
              <button type="button" class="demo__backstage-card" @click="fileInput?.click()">
                <strong>{{ ui.openTitle }}</strong>
                <span>{{ ui.openDesc }}</span>
              </button>
              <button type="button" class="demo__backstage-card" @click="onSave">
                <strong>{{ ui.saveCopy }}</strong>
                <span>{{ ui.saveDesc }}</span>
              </button>
              <button
                type="button"
                class="demo__backstage-card"
                :disabled="!instance"
                @click="instance?.print()"
              >
                <strong>{{ ui.print }}</strong>
                <span>{{ ui.printDesc }}</span>
              </button>
              <button
                type="button"
                class="demo__backstage-card"
                :disabled="!instance"
                @click="instance?.openPageSetup()"
              >
                <strong>{{ ui.pageSetup }}</strong>
                <span>{{ ui.pageSetupDesc }}</span>
              </button>
              <button
                type="button"
                class="demo__backstage-card"
                :disabled="!instance"
                @click="instance?.openExternalLinksDialog()"
              >
                <strong>{{ ui.editLinks }}</strong>
                <span>{{ ui.linksDesc }}</span>
              </button>
              <button type="button" class="demo__backstage-card" @click="showPanel = !showPanel">
                <strong>{{ ui.options }}</strong>
                <span>{{ ui.optionsDesc }}</span>
              </button>
            </div>
          </div>
        </div>
      </div>
      <aside class="demo__panel" aria-label="Options panel" :hidden="!showPanel">
        <section class="demo__card">
          <h2>Preset</h2>
          <p class="demo__hint">
            Toggle entire feature bundles, or override individual flags below. Changes
            flow through <code>inst.setFeatures()</code> live — edits survive.
          </p>
          <div class="demo__preset">
            <button
              v-for="p in PRESETS"
              :key="p.value"
              type="button"
              :class="['demo__preset-btn', { 'demo__preset-btn--active': preset === p.value }]"
              :aria-pressed="preset === p.value"
              @click="onPresetChange(p.value)"
            >
              <span class="demo__preset-name">{{ p.label }}</span>
              <span class="demo__preset-hint">{{ p.hint }}</span>
            </button>
          </div>
        </section>

        <section class="demo__card">
          <h2>Features</h2>
          <p class="demo__hint">
            Live-toggle individual <code>FeatureFlags</code>. Disabled flags skip their
            <code>attach*</code> in <code>mount.ts</code>.
          </p>
          <div v-for="group in FEATURE_GROUPS" :key="group.title" class="demo__feat-group">
            <h3 class="demo__feat-title">{{ group.title }}</h3>
            <div class="demo__feat-grid">
              <label
                v-for="f in group.features"
                :key="f.id"
                :class="['demo__feat', { 'demo__feat--on': isFeatureOn(f.id) }]"
              >
                <input
                  type="checkbox"
                  :checked="isFeatureOn(f.id)"
                  @change="onFeatureToggle(f.id)"
                />
                <span>{{ f.label }}</span>
              </label>
              <label
                v-if="group.title === 'Chrome'"
                :class="['demo__feat', { 'demo__feat--on': showRibbon }]"
              >
                <input type="checkbox" v-model="showRibbon" />
                <span>Spreadsheet ribbon</span>
              </label>
            </div>
          </div>
        </section>

        <section class="demo__card">
          <h2>Selection</h2>
          <p class="demo__mono">{{ selectionLabel }}</p>
        </section>

        <section class="demo__card">
          <h2>Cell renderers</h2>
          <p class="demo__hint">
            Wired via <code>inst.cells.registerFormatter</code>.
          </p>
          <label class="demo__check">
            <input type="checkbox" v-model="formatters.uppercase" />
            Uppercase column A
          </label>
          <label class="demo__check">
            <input type="checkbox" v-model="formatters.arrows" />
            Arrow-prefix negatives
          </label>
        </section>

        <section class="demo__card">
          <h2>Custom functions</h2>
          <p class="demo__hint">
            Registered via the <code>functions</code> prop. Probe the host-side
            registry directly:
          </p>
          <div class="demo__probe">
            <button
              type="button"
              class="demo__btn demo__btn--ghost"
              :disabled="!instance"
              @click="runProbe('GREET', [{ kind: 'text', value: 'Vue' }])"
            >
              GREET("Vue")
            </button>
            <button
              type="button"
              class="demo__btn demo__btn--ghost"
              :disabled="!instance"
              @click="runProbe('FAHRENHEIT', [{ kind: 'number', value: 100 }])"
            >
              FAHRENHEIT(100)
            </button>
            <p v-if="probe" class="demo__probe-out">
              → <code>{{ probe.result }}</code>
            </p>
          </div>
        </section>

        <section class="demo__card demo__card--log">
          <h2>Cell change log</h2>
          <p class="demo__hint">
            Mirrors the <code>cell-change</code> emit into Vue refs.
          </p>
          <p v-if="log.length === 0" class="demo__empty">
            Edit a cell to see events stream in.
          </p>
          <ul v-else class="demo__log">
            <li v-for="entry in log" :key="entry.id">
              <span class="demo__log-cell">{{ entry.cell }}</span>
              <span class="demo__log-arrow">→</span>
              <span class="demo__mono">{{ entry.preview }}</span>
            </li>
          </ul>
        </section>
      </aside>
    </main>
  </div>
</template>
