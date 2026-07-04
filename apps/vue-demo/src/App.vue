<script setup lang="ts">
import {
  type CellChangeEvent,
  type CellValue,
  type FeatureFlags,
  type FeatureId,
  EXCEL365_STANDARD_RIBBON_TABS,
  analyzeAccessibilityCells,
  analyzeSpellingCells,
  applyTextScript,
  mutators,
  parseScriptCommand,
  presets,
  type SpreadsheetInstance,
  type ThemeName,
  type ToolbarInstance,
  WorkbookHandle,
} from '@libraz/formulon-cell';
import { type RibbonTab, Spreadsheet, useSelection } from '@libraz/formulon-cell-vue';
import SpreadsheetToolbar from '@libraz/formulon-cell-vue/toolbar.vue';
import {
  computed,
  nextTick,
  onBeforeUnmount,
  onMounted,
  onUnmounted,
  ref,
  shallowRef,
  watch,
} from 'vue';
import {
  activateDemoModal,
  buildDemoBackstageCards,
  buildDemoBackstageNav,
  buildDemoCommands,
  buildDemoPrintPreviewModel,
  buildDemoSearchItems,
  composeDemoUiOptions,
  createDemoStrings,
  demoSearchOptionId,
  DEMO_ICONS,
  DEMO_FUNCTIONS,
  DEMO_PRINT_PREVIEW_LINES,
  DEMO_PRINTER_PROFILE_ID,
  DEMO_PRINTER_PROFILES,
  demoColLabel,
  demoCommandText,
  type DemoBackstageAction,
  type DemoSearchItem,
  type DemoSearchUsagePrior,
  FEATURE_GROUPS,
  formatLoadError,
  FORMATTERS,
  installDemoF6Navigation,
  isDemoBackstageActionDisabled,
  loadDemoSearchUsagePrior,
  LOCALES,
  nextDemoSearchIndex,
  type PresetKey,
  PRESETS,
  previewCellChange,
  installDemoSearchShortcut,
  queryDemoSearchItems,
  recordDemoSearchUsage,
  refreshDemoPrinterProfiles,
  resolveInitialLocale,
  reviewCellsForInstance,
  runDemoBackstageAction,
  saveDemoSearchUsagePrior,
  saveDemoWorkbookToDownload,
  seedDemoWorkbook,
  THEMES,
} from '../../demo-shared/index.js';

const UI = createDemoStrings('Vue');



const colLabel = demoColLabel;



interface ChangeLogEntry {
  readonly id: number;
  readonly cell: string;
  readonly preview: string;
}

interface ReviewDialogState {
  readonly title: string;
  readonly items: readonly { label: string; detail: string }[];
}

let changeId = 0;
let disposeSearchShortcut: (() => void) | undefined;
let disposeF6Navigation: (() => void) | undefined;

// Modal focus trap + Esc-to-close + change-log preview + seed +
// review-cell projection all live in demo-shared/index.ts so the React
// and Vue demos stay aligned.
const previewValue = previewCellChange;
const seed = seedDemoWorkbook;

const theme = ref<ThemeName>('paper');
const locale = ref<string>(resolveInitialLocale());
const workbook = shallowRef<WorkbookHandle | null>(null);
// Vue's reactive proxy walks deeply by default; the spreadsheet instance
// holds a canvas + many internal refs that should not be reactivified.
const instance = shallowRef<SpreadsheetInstance | null>(null);
const toolbar = shallowRef<ToolbarInstance | null>(null);
const log = ref<ChangeLogEntry[]>([]);
const formatters = ref({ uppercase: true, arrows: true });
const probe = ref<{ name: string; result: string } | null>(null);
const fileInput = ref<HTMLInputElement | null>(null);
const searchInput = ref<HTMLInputElement | null>(null);
const quickAccess = ref<HTMLElement | null>(null);
const preset = ref<PresetKey>('full');
const overrides = ref<FeatureFlags>({});
const showRibbon = ref(true);
const showPanel = ref(false);
const ribbonTab = ref<RibbonTab>('home');
const backstageAction = ref<DemoBackstageAction>('info');
const searchQuery = ref('');
const searchOpen = ref(false);
const searchActiveIndex = ref(-1);
const searchUsagePrior = ref<DemoSearchUsagePrior>(loadDemoSearchUsagePrior());
const bookName = ref('Book1');
const loadError = ref<string | null>(null);
const reviewDialog = ref<ReviewDialogState | null>(null);
const scriptOpen = ref(false);
const scriptCommand = ref('uppercase');
const scriptError = ref<string | null>(null);
const uploadStatus = ref<'saved' | 'saving' | 'error' | null>(null);
const reviewModalEl = ref<HTMLElement | null>(null);
const scriptModalEl = ref<HTMLElement | null>(null);

const resolvedUi = computed(() =>
  composeDemoUiOptions({
    preset: preset.value,
    overrides: overrides.value,
    showRibbon: showRibbon.value,
    theme: theme.value,
  }),
);
const features = computed<FeatureFlags>(() => resolvedUi.value.features);
const ui = computed(() => UI[locale.value === 'ja' ? 'ja' : 'en']);
const commandText = computed(() => demoCommandText(locale.value));
const backstageNav = computed(() => buildDemoBackstageNav(ui.value, backstageAction.value));
const backstageCards = computed(() => buildDemoBackstageCards(ui.value));
const printPreview = computed(() => {
  void backstageAction.value;
  return buildDemoPrintPreviewModel(ui.value, instance.value, bookName.value);
});

watch(
  locale,
  (next) => {
    document.documentElement.lang = next === 'ja' ? 'ja' : 'en';
  },
  { immediate: true },
);

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

const onSpellingReview = (): void => {
  const inst = instance.value;
  if (!inst) return;
  reviewDialog.value = {
    title: commandText.value.spellingReview,
    items: analyzeSpellingCells(reviewCellsForInstance(inst), locale.value === 'ja' ? 'ja' : 'en'),
  };
};

const onAccessibilityCheck = (): void => {
  const inst = instance.value;
  if (!inst) return;
  reviewDialog.value = {
    title: commandText.value.accessibilityCheck,
    items: analyzeAccessibilityCells(
      reviewCellsForInstance(inst),
      locale.value === 'ja' ? 'ja' : 'en',
    ),
  };
};

const onRunScript = (): void => {
  const inst = instance.value;
  if (!inst) return;
  scriptCommand.value = 'uppercase';
  scriptError.value = null;
  scriptOpen.value = true;
};

const closeReviewDialog = (): void => {
  reviewDialog.value = null;
};

const closeScriptDialog = (): void => {
  scriptOpen.value = false;
};

let demoModalCleanup: (() => void) | null = null;

watch(
  () => (reviewDialog.value ? 'review' : scriptOpen.value ? 'script' : null),
  async (openModal) => {
    demoModalCleanup?.();
    demoModalCleanup = null;
    if (!openModal) return;
    await nextTick();
    const root = openModal === 'review' ? reviewModalEl.value : scriptModalEl.value;
    if (!root) return;
    demoModalCleanup = activateDemoModal(
      root,
      openModal === 'review' ? closeReviewDialog : closeScriptDialog,
    );
  },
);

const showRibbonNotice = (title: string, detail: string): void => {
  reviewDialog.value = { title, items: [{ label: commandText.value.ribbonCommand, detail }] };
};

const applyParsedScript = (command: ReturnType<typeof parseScriptCommand>): void => {
  const inst = instance.value;
  if (!inst || !command) return;
  const range = inst.store.getState().selection.range;
  let changed = 0;
  inst.history.begin();
  try {
    for (let row = range.r0; row <= range.r1; row += 1) {
      for (let col = range.c0; col <= range.c1; col += 1) {
        const addr = { sheet: range.sheet, row, col };
        const value = inst.workbook.getValue(addr);
        if (command === 'clear') {
          if (value.kind !== 'blank' || inst.workbook.cellFormula(addr)) {
            inst.workbook.setBlank(addr);
            changed += 1;
          }
          continue;
        }
        if (value.kind === 'text') {
          const next = applyTextScript(value.value, command);
          if (next !== value.value) {
            inst.workbook.setText(addr, next);
            changed += 1;
          }
        }
      }
    }
  } finally {
    inst.history.end();
  }
  mutators.replaceCells(inst.store, inst.workbook.cells(range.sheet));
  reviewDialog.value = {
    title: commandText.value.script,
    items: [
      {
        label: commandText.value.selection,
        detail: commandText.value.cellsUpdated.replace('{count}', String(changed)),
      },
    ],
  };
};

const applyScriptCommand = (): void => {
  const command = parseScriptCommand(scriptCommand.value);
  if (!command) {
    scriptError.value = commandText.value.scriptCommandError;
    return;
  }
  scriptOpen.value = false;
  applyParsedScript(command);
};

const onScriptMenuClick = (e: MouseEvent): void => {
  const target = e.target;
  if (!(target instanceof Element)) return;
  const btn = target.closest<HTMLButtonElement>('[data-script-action]');
  if (!btn) return;
  const menu = btn.closest<HTMLDivElement>('#menu-script');
  if (!menu) return;
  const action = btn.dataset.scriptAction ?? '';
  menu.hidden = true;
  const opener = menu.previousElementSibling;
  if (opener instanceof HTMLElement) {
    opener.setAttribute('aria-expanded', 'false');
    opener.focus({ preventScroll: true });
  }
  if (action === 'custom') {
    if (!instance.value) return;
    scriptCommand.value = 'uppercase';
    scriptError.value = null;
    scriptOpen.value = true;
    return;
  }
  const command = parseScriptCommand(action);
  if (command) applyParsedScript(command);
};

onMounted(() => {
  document.addEventListener('click', onScriptMenuClick);
});
onBeforeUnmount(() => {
  document.removeEventListener('click', onScriptMenuClick);
});

const onSave = (): void => {
  saveDemoWorkbookToDownload({
    instance: instance.value,
    bookName: bookName.value,
    setUploadStatus: (next) => {
      uploadStatus.value = next;
    },
  });
};

const onNewWorkbook = async (): Promise<void> => {
  const wb = await WorkbookHandle.createDefault();
  workbook.value = wb;
  await instance.value?.setWorkbook(wb);
  bookName.value = 'Book1';
  log.value = [];
  ribbonTab.value = 'home';
};

const backstageActionDisabled = (action: DemoBackstageAction): boolean =>
  isDemoBackstageActionDisabled(action, instance.value);

const runBackstageAction = (action: DemoBackstageAction): void => {
  if (action === 'info' || action === 'print') {
    backstageAction.value = action;
    return;
  }
  runDemoBackstageAction({
    action,
    instance: instance.value,
    ui: ui.value,
    newWorkbook: onNewWorkbook,
    openWorkbook: () => fileInput.value?.click(),
    saveWorkbook: onSave,
    showNotice: showRibbonNotice,
    toggleOptions: () => {
      showPanel.value = !showPanel.value;
    },
    closeBackstage: () => {
      backstageAction.value = 'info';
      ribbonTab.value = 'home';
    },
  });
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
    reviewDialog.value = {
      title: commandText.value.openFailed,
      items: [{ label: commandText.value.workbook, detail: formatLoadError(err) }],
    };
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

const commands = computed(() =>
  buildDemoCommands({
    commandText: commandText.value,
    instance: instance.value,
    openWorkbook: () => fileInput.value?.click(),
    saveWorkbook: onSave,
    setRibbonTab: (tab) => {
      ribbonTab.value = tab;
    },
    togglePanel: () => {
      showPanel.value = !showPanel.value;
    },
    setTheme: (next) => {
      theme.value = next;
    },
    setLocale: (next) => {
      locale.value = next;
    },
  }),
);
const searchItems = computed(() =>
  buildDemoSearchItems(
    commands.value,
    locale.value,
    (tab) => {
      ribbonTab.value = tab;
    },
    (commandId) => toolbar.value?.applyCommand(commandId) ?? false,
    EXCEL365_STANDARD_RIBBON_TABS,
  ),
);

const filteredCommands = computed(() => {
  return queryDemoSearchItems(searchItems.value, searchQuery.value, 8, searchUsagePrior.value);
});

const runCommand = (cmd: DemoSearchItem): void => {
  searchUsagePrior.value = recordDemoSearchUsage(searchUsagePrior.value, cmd);
  if (cmd.tab) ribbonTab.value = cmd.tab;
  cmd.run();
  searchQuery.value = '';
  searchOpen.value = false;
  searchActiveIndex.value = -1;
};

const onToolbarReady = (next: ToolbarInstance | null): void => {
  toolbar.value = next;
  (window as unknown as { __fcToolbar?: ToolbarInstance | null }).__fcToolbar = next;
};

const onSearchKeydown = (ev: KeyboardEvent): void => {
  if (ev.key === 'Escape') {
    searchOpen.value = false;
    searchActiveIndex.value = -1;
    (ev.currentTarget as HTMLInputElement).blur();
  }
  if (ev.key === 'ArrowDown' || ev.key === 'ArrowUp') {
    ev.preventDefault();
    searchOpen.value = true;
    searchActiveIndex.value = nextDemoSearchIndex(
      searchActiveIndex.value,
      filteredCommands.value.length,
      ev.key === 'ArrowDown' ? 'next' : 'previous',
    );
  }
  if (ev.key === 'Enter' && filteredCommands.value.length > 0) {
    ev.preventDefault();
    const index = nextDemoSearchIndex(
      searchActiveIndex.value,
      filteredCommands.value.length,
      'first',
    );
    const command = filteredCommands.value[index];
    if (command) runCommand(command);
  }
};

watch([searchQuery, searchOpen], () => {
  searchActiveIndex.value = -1;
});

watch(searchUsagePrior, (prior) => saveDemoSearchUsagePrior(prior));

onUnmounted(() => {
  demoModalCleanup?.();
  // The Spreadsheet component disposes itself; nothing extra to clean up.
});

onMounted(() => {
  disposeSearchShortcut = installDemoSearchShortcut(() => searchInput.value);
  disposeF6Navigation = installDemoF6Navigation({
    getQuickAccess: () => quickAccess.value,
    getToolbar: () => toolbar.value,
    getInstance: () => instance.value,
  });
});

onBeforeUnmount(() => {
  disposeSearchShortcut?.();
  disposeSearchShortcut = undefined;
  disposeF6Navigation?.();
  disposeF6Navigation = undefined;
});
</script>

<template>
  <div v-if="!workbook" class="demo demo--loading">
    <div v-if="loadError" class="demo__load-error" role="alert">
      <strong>{{ ui.engineUnavailable }}</strong>
      <span>{{ ui.engineSetup }}</span>
      <code>{{ loadError }}</code>
    </div>
    <template v-else>{{ ui.loadingEngine }}</template>
  </div>
  <div v-else class="demo" :data-fc-theme="theme">
    <header class="demo__head">
      <div class="demo__titlebar">
        <div
          ref="quickAccess"
          class="demo__quick"
          role="toolbar"
          :aria-label="ui.quickAccessToolbar"
        >
          <span class="demo__brand-mark" aria-hidden="true">
            <svg class="demo__rb-icon" viewBox="0 0 20 20" stroke-width="1.45" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
              <path v-for="segment in DEMO_ICONS.app" :key="segment.d" :d="segment.d" :fill="segment.fill ?? 'none'" :stroke="segment.stroke ?? 'currentColor'" />
            </svg>
          </span>
          <button type="button" class="demo__title-icon" :aria-label="ui.save" @click="onSave">
            <svg class="demo__rb-icon" viewBox="0 0 20 20" stroke-width="1.45" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
              <path v-for="segment in DEMO_ICONS.save" :key="segment.d" :d="segment.d" :fill="segment.fill ?? 'none'" :stroke="segment.stroke ?? 'currentColor'" />
            </svg>
          </button>
          <button type="button" class="demo__title-icon" :aria-label="ui.undo" @click="instance?.undo()">
            <svg class="demo__rb-icon" viewBox="0 0 20 20" stroke-width="1.45" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
              <path v-for="segment in DEMO_ICONS.undo" :key="segment.d" :d="segment.d" :fill="segment.fill ?? 'none'" :stroke="segment.stroke ?? 'currentColor'" />
            </svg>
          </button>
          <button type="button" class="demo__title-icon" :aria-label="ui.redo" @click="instance?.redo()">
            <svg class="demo__rb-icon" viewBox="0 0 20 20" stroke-width="1.45" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
              <path v-for="segment in DEMO_ICONS.redo" :key="segment.d" :d="segment.d" :fill="segment.fill ?? 'none'" :stroke="segment.stroke ?? 'currentColor'" />
            </svg>
          </button>
        </div>
        <div class="demo__title">
          <strong>{{ bookName }}</strong>
          <span>{{ ui.saved }}</span>
        </div>
        <div class="demo__search">
          <svg class="demo__rb-icon" viewBox="0 0 20 20" stroke-width="1.45" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <path v-for="segment in DEMO_ICONS.search" :key="segment.d" :d="segment.d" :fill="segment.fill ?? 'none'" :stroke="segment.stroke ?? 'currentColor'" />
          </svg>
          <input
            ref="searchInput"
            v-model="searchQuery"
            type="search"
            role="combobox"
            :placeholder="ui.search"
            :aria-label="ui.searchCommands"
            aria-controls="demo-search-results"
            :aria-expanded="searchOpen"
            :aria-activedescendant="searchOpen && searchActiveIndex >= 0 ? demoSearchOptionId(searchActiveIndex) : undefined"
            @focus="searchOpen = true; searchActiveIndex = -1"
            @input="searchOpen = true; searchActiveIndex = -1"
            @keydown="onSearchKeydown"
            @blur="searchOpen = false"
          />
          <div v-if="searchOpen" id="demo-search-results" class="demo__command-menu" role="listbox">
            <div v-if="filteredCommands.length === 0" class="demo__command-empty">
              {{ ui.noCommands }}
            </div>
            <button
              v-for="(cmd, index) in filteredCommands"
              v-else
              :key="cmd.id"
              :id="demoSearchOptionId(index)"
              type="button"
              role="option"
              :aria-selected="index === searchActiveIndex"
              :aria-disabled="cmd.disabled ? 'true' : undefined"
              :data-disabled-reason="cmd.disabledReason"
              :class="[
                'demo__command-item',
                {
                  'demo__command-item--active': index === searchActiveIndex,
                  'demo__command-item--disabled': cmd.disabled,
                },
              ]"
              @mousedown.prevent
              @mouseenter="searchActiveIndex = index"
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
          <button
            type="button"
            :class="['demo__share', { 'demo__share--active': showPanel }]"
            :aria-pressed="showPanel"
            @click="showPanel = !showPanel"
          >
            {{ ui.demoPane }}
          </button>
          <span class="demo__avatar" role="img" :aria-label="ui.signedInUser">FC</span>
        </div>
      </div>
    </header>
    <input ref="fileInput" type="file" accept=".xlsx,.xlsm" hidden @change="onOpenFiles" />

    <main :class="['demo__body', { 'demo__body--panel': showPanel }]">
      <div class="demo__sheet-col">
        <SpreadsheetToolbar
          v-if="resolvedUi.ribbon"
          :instance="instance"
          :active-tab="ribbonTab"
          :locale="locale"
          :ribbon-tabs="EXCEL365_STANDARD_RIBBON_TABS"
          :on-spelling-review="onSpellingReview"
          :on-accessibility-check="onAccessibilityCheck"
          :on-run-script="onRunScript"
          :on-draw-pen="
            () => showRibbonNotice(commandText.draw, commandText.inkNotPersisted)
          "
          :on-draw-eraser="
            () => showRibbonNotice(commandText.draw, commandText.selectInkFirst)
          "
          :on-translate="
            () => showRibbonNotice(commandText.translate, commandText.translationUnavailable)
          "
          :on-add-in="
            () => showRibbonNotice(commandText.addIns, commandText.addInsHostCallbacks)
          "
          :on-toolbar-ready="onToolbarReady"
          @tab-change="ribbonTab = $event"
        />
        <Spreadsheet
          class="demo__sheet"
          :workbook="workbook"
          :theme="theme"
          :locale="locale"
          :features="features"
          :functions="DEMO_FUNCTIONS"
          :printer-profiles="DEMO_PRINTER_PROFILES"
          :printer-profile-id="DEMO_PRINTER_PROFILE_ID"
          :refresh-printer-profiles="refreshDemoPrinterProfiles"
          :upload-status="uploadStatus"
          :macro-recording="scriptOpen"
          @ready="onReady"
          @cell-change="onCellChange"
        />
        <div v-if="ribbonTab === 'file'" class="demo__backstage" role="dialog" :aria-label="ui.file">
          <nav class="demo__backstage-nav" :aria-label="ui.file">
            <strong>{{ ui.file }}</strong>
            <button
              v-for="item in backstageNav"
              :key="item.action"
              type="button"
              :class="[
                'demo__backstage-navitem',
                item.active ? 'demo__backstage-navitem--active' : '',
              ]"
              :disabled="backstageActionDisabled(item.action)"
              @click="runBackstageAction(item.action)"
            >
              {{ item.label }}
            </button>
          </nav>
          <div class="demo__backstage-main">
            <div class="demo__backstage-title">
              <span class="demo__backstage-xl" aria-hidden="true">
                <svg class="demo__rb-icon" viewBox="0 0 20 20" stroke-width="1.45" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                  <path v-for="segment in DEMO_ICONS.app" :key="segment.d" :d="segment.d" :fill="segment.fill ?? 'none'" :stroke="segment.stroke ?? 'currentColor'" />
                </svg>
              </span>
              <div>
                <h1>{{ bookName }}</h1>
                <p>{{ ui.backstageSub }}</p>
              </div>
            </div>
            <div v-if="backstageAction === 'print'" class="demo__print-preview" data-demo-print-preview>
              <section class="demo__print-settings" :aria-label="ui.printSettings">
                <h2>{{ printPreview.title }}</h2>
                <p>{{ printPreview.subtitle }}</p>
                <button
                  type="button"
                  class="demo__print-action demo__print-action--primary"
                  :disabled="!instance"
                  @click="instance?.print('print')"
                >
                  {{ printPreview.printLabel }}
                </button>
                <button
                  type="button"
                  class="demo__print-action"
                  :disabled="!instance"
                  @click="instance?.print('pdf')"
                >
                  {{ printPreview.pdfLabel }}
                </button>
                <button
                  type="button"
                  class="demo__print-action"
                  :disabled="!instance"
                  @click="instance?.openPageSetup()"
                >
                  {{ printPreview.pageSetupLabel }}
                </button>
                <dl class="demo__print-meta">
                  <div v-for="row in printPreview.settings" :key="row.label">
                    <dt>{{ row.label }}</dt>
                    <dd>{{ row.value }}</dd>
                  </div>
                </dl>
              </section>
              <section class="demo__print-paper" :aria-label="printPreview.previewTitle">
                <iframe
                  v-if="printPreview.previewHtml"
                  class="demo__print-frame"
                  :title="printPreview.previewTitle"
                  sandbox=""
                  :srcdoc="printPreview.previewHtml"
                />
                <div v-else class="demo__print-page">
                  <strong>{{ printPreview.previewTitle }}</strong>
                  <div aria-hidden="true" class="demo__print-sheet-lines">
                    <span v-for="line in DEMO_PRINT_PREVIEW_LINES" :key="line" />
                  </div>
                </div>
                <p>{{ printPreview.previewHint }}</p>
              </section>
            </div>
            <div v-else class="demo__backstage-grid">
              <button
                v-for="item in backstageCards"
                :key="item.action"
                type="button"
                class="demo__backstage-card"
                :disabled="backstageActionDisabled(item.action)"
                @click="runBackstageAction(item.action)"
              >
                <strong>{{ item.label }}</strong>
                <span>{{ item.desc }}</span>
              </button>
            </div>
          </div>
        </div>
      </div>
      <aside class="demo__panel" :aria-label="ui.optionsPanel" :hidden="!showPanel">
        <section class="demo__card">
          <h2>{{ ui.demoChrome }}</h2>
          <div class="demo__controls demo__controls--panel">
            <div class="demo__seg" role="group" :aria-label="ui.theme">
              <button
                v-for="t in THEMES"
                :key="t.value"
                type="button"
                :class="['demo__seg-btn', { 'demo__seg-btn--active': theme === t.value }]"
                :aria-pressed="theme === t.value"
                @click="theme = t.value"
              >
                {{ ui.themeLabels[t.value] ?? t.label }}
              </button>
            </div>
            <div class="demo__seg" role="group" :aria-label="ui.locale">
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
          </div>
        </section>

        <section class="demo__card">
          <h2>{{ ui.preset }}</h2>
          <p class="demo__hint">{{ ui.presetHint }}</p>
          <div class="demo__preset">
            <button
              v-for="p in PRESETS"
              :key="p.value"
              type="button"
              :class="['demo__preset-btn', { 'demo__preset-btn--active': preset === p.value }]"
              :aria-pressed="preset === p.value"
              @click="onPresetChange(p.value)"
            >
              <span class="demo__preset-name">{{ ui.presets[p.value]?.label ?? p.label }}</span>
              <span class="demo__preset-hint">{{ ui.presets[p.value]?.hint ?? p.hint }}</span>
            </button>
          </div>
        </section>

        <section class="demo__card">
          <h2>{{ ui.features }}</h2>
          <p class="demo__hint">{{ ui.featuresHint }}</p>
          <div v-for="group in FEATURE_GROUPS" :key="group.title" class="demo__feat-group">
            <h3 class="demo__feat-title">{{ ui.featureGroupLabels[group.title] ?? group.title }}</h3>
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
                <span>{{ ui.featureLabels[f.id] ?? f.label }}</span>
              </label>
              <label
                v-if="group.title === 'Chrome'"
                :class="['demo__feat', { 'demo__feat--on': resolvedUi.ribbon }]"
              >
                <input type="checkbox" v-model="showRibbon" />
                <span>{{ ui.spreadsheetRibbon }}</span>
              </label>
            </div>
          </div>
        </section>

        <section class="demo__card">
          <h2>{{ commandText.selection }}</h2>
          <p class="demo__mono">{{ selectionLabel }}</p>
        </section>

        <section class="demo__card">
          <h2>{{ ui.cellRenderers }}</h2>
          <p class="demo__hint">{{ ui.cellRenderersHint }}</p>
          <label class="demo__check">
            <input type="checkbox" v-model="formatters.uppercase" />
            {{ ui.uppercaseColumnA }}
          </label>
          <label class="demo__check">
            <input type="checkbox" v-model="formatters.arrows" />
            {{ ui.arrowPrefixNegatives }}
          </label>
        </section>

        <section class="demo__card">
          <h2>{{ ui.customFunctions }}</h2>
          <p class="demo__hint">{{ ui.customFunctionsHint }}</p>
          <div class="demo__probe">
            <button
              type="button"
              class="demo__btn demo__btn--ghost"
              :disabled="!instance"
              @click="runProbe('GREET', [{ kind: 'text', value: 'Workbook' }])"
            >
              GREET("Workbook")
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
          <h2>{{ ui.cellChangeLog }}</h2>
          <p class="demo__hint">{{ ui.cellChangeLogHint }}</p>
          <p v-if="log.length === 0" class="demo__empty">
            {{ ui.editCellToSeeEvents }}
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
    <div
      v-if="reviewDialog"
      ref="reviewModalEl"
      class="demo__modal"
      role="dialog"
      aria-modal="true"
      :aria-label="reviewDialog.title"
    >
      <section class="demo__modal-panel">
        <header class="demo__modal-header">
          <h2>{{ reviewDialog.title }}</h2>
          <button
            type="button"
            class="demo__modal-x"
            :aria-label="ui.close"
            @click="closeReviewDialog"
          >
            ×
          </button>
        </header>
        <div class="demo__modal-body">
          <p v-if="reviewDialog.items.length === 0" class="demo__modal-empty">
            {{ ui.noIssuesFound }}
          </p>
          <ul v-else class="demo__modal-list">
            <li v-for="(item, index) in reviewDialog.items" :key="`${item.label}-${index}`">
              <strong>{{ item.label }}</strong>
              <span>{{ item.detail }}</span>
            </li>
          </ul>
        </div>
        <footer class="demo__modal-footer">
          <button type="button" class="demo__btn" @click="closeReviewDialog">{{ ui.ok }}</button>
        </footer>
      </section>
    </div>
    <div
      v-if="scriptOpen"
      ref="scriptModalEl"
      class="demo__modal"
      role="dialog"
      aria-modal="true"
      :aria-label="commandText.script"
    >
      <form class="demo__modal-panel demo__modal-panel--narrow" @submit.prevent="applyScriptCommand">
        <header class="demo__modal-header">
          <h2>{{ commandText.script }}</h2>
          <button
            type="button"
            class="demo__modal-x"
            :aria-label="ui.close"
            @click="closeScriptDialog"
          >
            ×
          </button>
        </header>
        <div class="demo__modal-body">
          <label class="demo__modal-field">
            <span>{{ ui.command }}</span>
            <input v-model="scriptCommand" autofocus @input="scriptError = null" />
          </label>
          <p v-if="scriptError" class="demo__modal-error">{{ scriptError }}</p>
        </div>
        <footer class="demo__modal-footer">
          <button type="button" class="demo__btn" @click="closeScriptDialog">{{ ui.cancel }}</button>
          <button type="submit" class="demo__btn demo__btn--active">{{ ui.run }}</button>
        </footer>
      </form>
    </div>
  </div>
</template>
