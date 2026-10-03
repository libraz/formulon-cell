<script setup lang="ts">
import {
  type CellChangeEvent,
  type FeatureFlags,
  type FeatureId,
  parseScriptCommand,
  type ScriptCommand,
  type SpreadsheetInstance,
  type ThemeName,
  type ToolbarInstance,
  WorkbookHandle,
} from '@libraz/formulon-cell';
import { type RibbonTab, Spreadsheet } from '@libraz/formulon-cell-vue';
import SpreadsheetToolbar from '@libraz/formulon-cell-vue/toolbar.vue';
import {
  computed,
  onBeforeUnmount,
  onMounted,
  ref,
  shallowRef,
  watch,
} from 'vue';
import {
  buildDemoCommands,
  buildDemoReviewDialog,
  buildDemoReviewFindings,
  buildDemoSearchItems,
  composeDemoUiOptions,
  createDemoStrings,
  createInitialDemoWorkbook,
  demoSearchOptionId,
  DEMO_FUNCTIONS,
  demoFunctionArgumentHelp,
  DEMO_MAC_RIBBON_TABS,
  DEMO_PRINTER_PROFILE_ID,
  DEMO_PRINTER_PROFILES,
  DEMO_RIBBON_TABS,
  pushDemoChangeLog,
  resolveDemoSearchKey,
  nextDemoFeatureOverrides,
  openDemoWorkbookFile,
  installDemoScriptMenu,
  type ChangeLogEntry,
  type DemoReviewDialogState,
  demoCommandText,
  type DemoBackstageAction,
  type DemoSearchItem,
  type DemoSearchUsagePrior,
  formatLoadError,
  FORMATTERS,
  installDemoF6Navigation,
  loadDemoSearchUsagePrior,
  type PresetKey,
  installDemoSearchShortcut,
  queryDemoSearchItems,
  recordDemoSearchUsage,
  refreshDemoPrinterProfiles,
  resolveInitialLocale,
  resolveInitialPlatform,
  type DemoPlatform,
  reportDemoScriptRun,
  runDemoBackstageAction,
  saveDemoSearchUsagePrior,
  saveDemoWorkbookToDownload,
} from '../../demo-shared/index.js';
import DemoBackstage from './DemoBackstage.vue';
import DemoIcon from './DemoIcon.vue';
import DemoOptionsPanel from './DemoOptionsPanel.vue';
import DemoReviewDialog from './DemoReviewDialog.vue';
import DemoScriptDialog from './DemoScriptDialog.vue';

const UI = createDemoStrings('Vue');



let disposeSearchShortcut: (() => void) | undefined;
let disposeF6Navigation: (() => void) | undefined;

const theme = ref<ThemeName>('paper');
const locale = ref<string>(resolveInitialLocale());
const platform = ref<DemoPlatform>(resolveInitialPlatform());
const workbook = shallowRef<WorkbookHandle | null>(null);
// Vue's reactive proxy walks deeply by default; the spreadsheet instance
// holds a canvas + many internal refs that should not be reactivified.
const instance = shallowRef<SpreadsheetInstance | null>(null);
const toolbar = shallowRef<ToolbarInstance | null>(null);
const log = ref<ChangeLogEntry[]>([]);
const formatters = ref({ uppercase: true, arrows: true });
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
const reviewDialog = ref<DemoReviewDialogState | null>(null);
const scriptOpen = ref(false);
const uploadStatus = ref<'saved' | 'saving' | 'error' | null>(null);

const resolvedUi = computed(() =>
  composeDemoUiOptions({
    preset: preset.value,
    overrides: overrides.value,
    showRibbon: showRibbon.value,
    theme: theme.value,
    platform: platform.value,
  }),
);
const ribbonTabs = computed(() =>
  platform.value === 'mac' ? DEMO_MAC_RIBBON_TABS : DEMO_RIBBON_TABS,
);
const features = computed<FeatureFlags>(() => resolvedUi.value.features);
const ui = computed(() => UI[locale.value === 'ja' ? 'ja' : 'en']);
const commandText = computed(() => demoCommandText(locale.value));

watch(
  locale,
  (next) => {
    document.documentElement.lang = next === 'ja' ? 'ja' : 'en';
  },
  { immediate: true },
);

void createInitialDemoWorkbook()
  .then((wb) => {
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
  log.value = pushDemoChangeLog(log.value, e);
};

const onReady = (inst: SpreadsheetInstance): void => {
  instance.value = inst;
  // Expose the live instance on `window.__fcInst` so cross-demo E2E scenarios
  // can drive imperative paths without depending on demo-specific UI.
  (window as unknown as { __fcInst?: SpreadsheetInstance | null }).__fcInst = inst;
};

const onSpellingReview = (): void => {
  const inst = instance.value;
  if (!inst) return;
  reviewDialog.value = buildDemoReviewFindings('spelling', inst, locale.value, commandText.value);
};

const onAccessibilityCheck = (): void => {
  const inst = instance.value;
  if (!inst) return;
  reviewDialog.value = buildDemoReviewFindings(
    'accessibility',
    inst,
    locale.value,
    commandText.value,
  );
};

const onRunScript = (): void => {
  const inst = instance.value;
  if (!inst) return;
  scriptOpen.value = true;
};

const closeReviewDialog = (): void => {
  reviewDialog.value = null;
};

const closeScriptDialog = (): void => {
  scriptOpen.value = false;
};

const showRibbonNotice = (title: string, detail: string): void => {
  reviewDialog.value = buildDemoReviewDialog(title, commandText.value.ribbonCommand, detail);
};

const applyParsedScript = (command: ReturnType<typeof parseScriptCommand>): void => {
  const inst = instance.value;
  if (!inst || !command) return;
  reviewDialog.value = reportDemoScriptRun(inst, command, commandText.value);
};

const onScriptSubmit = (command: ScriptCommand): void => {
  scriptOpen.value = false;
  applyParsedScript(command);
};

let disposeScriptMenu: (() => void) | undefined;

onMounted(() => {
  disposeScriptMenu = installDemoScriptMenu(applyParsedScript);
});
onBeforeUnmount(() => {
  disposeScriptMenu?.();
  disposeScriptMenu = undefined;
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
      toolbar.value?.setBackstageOpen(false);
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
    bookName.value = await openDemoWorkbookFile(inst, file);
    loadError.value = null;
  } catch (err) {
    reviewDialog.value = buildDemoReviewDialog(
      commandText.value.openFailed,
      commandText.value.workbook,
      formatLoadError(err),
    );
  }
};

const onFormatterChange = (key: 'uppercase' | 'arrows', checked: boolean): void => {
  formatters.value[key] = checked;
};

const onPresetChange = (next: PresetKey): void => {
  if (next === preset.value) return;
  preset.value = next;
  overrides.value = {};
};

const onFeatureToggle = (id: FeatureId): void => {
  overrides.value = nextDemoFeatureOverrides({
    preset: preset.value,
    overrides: overrides.value,
    features: features.value,
    id,
  });
};

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
    ribbonTabs.value,
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

const openBackstage = (): void => {
  ribbonTab.value = 'file';
  toolbar.value?.setBackstageOpen(true);
};

const onToolbarReady = (next: ToolbarInstance | null): void => {
  if (next) next.host.dataset.fcPlatform = platform.value;
  toolbar.value = next;
  (window as unknown as { __fcToolbar?: ToolbarInstance | null }).__fcToolbar = next;
};

const onSearchKeydown = (ev: KeyboardEvent): void => {
  const action = resolveDemoSearchKey(ev.key, searchActiveIndex.value, filteredCommands.value.length);
  if (action?.kind === 'close') {
    searchOpen.value = false;
    searchActiveIndex.value = -1;
    (ev.currentTarget as HTMLInputElement).blur();
  } else if (action?.kind === 'move') {
    ev.preventDefault();
    searchOpen.value = true;
    searchActiveIndex.value = action.index;
  } else if (action?.kind === 'run') {
    ev.preventDefault();
    const command = filteredCommands.value[action.index];
    if (command) runCommand(command);
  }
};

watch([searchQuery, searchOpen], () => {
  searchActiveIndex.value = -1;
});

watch(searchUsagePrior, (prior) => saveDemoSearchUsagePrior(prior));

onMounted(() => {
  disposeSearchShortcut = installDemoSearchShortcut(() => searchInput.value, platform.value);
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
  <div v-else class="demo" :data-fc-theme="theme" :data-fc-platform="platform">
    <header class="demo__head">
      <div class="fc-tb__titlebar">
        <div
          ref="quickAccess"
          class="demo__quick"
          role="toolbar"
          :aria-label="ui.quickAccessToolbar"
        >
          <button
            v-if="platform === 'mac'"
            type="button"
            class="demo__brand-mark"
            :aria-label="ui.file"
            :title="ui.file"
            @click="openBackstage"
          >
            <DemoIcon name="app" />
          </button>
          <span v-else class="demo__brand-mark" aria-hidden="true">
            <DemoIcon name="app" />
          </span>
          <button type="button" class="demo__title-icon" :aria-label="ui.save" @click="onSave">
            <DemoIcon name="save" />
          </button>
          <button type="button" class="demo__title-icon" :aria-label="ui.undo" @click="instance?.undo()">
            <DemoIcon name="undo" />
          </button>
          <button type="button" class="demo__title-icon" :aria-label="ui.redo" @click="instance?.redo()">
            <DemoIcon name="redo" />
          </button>
        </div>
        <div class="fc-tb__title">
          <strong>{{ bookName }}</strong>
          <span>{{ ui.saved }}</span>
        </div>
        <div class="fc-tb__search">
          <DemoIcon name="search" />
          <input
            ref="searchInput"
            v-model="searchQuery"
            type="search"
            role="combobox"
            :placeholder="ui.search"
            :aria-label="platform === 'mac' ? (locale === 'ja' ? '検索' : 'Search') : ui.searchCommands"
            aria-controls="demo-search-results"
            :aria-expanded="searchOpen"
            :aria-activedescendant="searchOpen && searchActiveIndex >= 0 ? demoSearchOptionId(searchActiveIndex) : undefined"
            @focus="searchOpen = true; searchActiveIndex = -1"
            @input="searchOpen = true; searchActiveIndex = -1"
            @keydown="onSearchKeydown"
            @blur="searchOpen = false"
          />
          <div v-if="searchOpen" id="demo-search-results" class="fc-tb__command-menu" role="listbox">
            <div v-if="filteredCommands.length === 0" class="fc-tb__command-empty">
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
                'fc-tb__command-item',
                {
                  'fc-tb__command-item--active': index === searchActiveIndex,
                  'fc-tb__command-item--disabled': cmd.disabled,
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
      <div class="fc-tb__sheet-col">
        <SpreadsheetToolbar
          v-if="resolvedUi.ribbon"
          :instance="instance"
          :active-tab="ribbonTab"
          :locale="locale"
          :ribbon-tabs="ribbonTabs"
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
          :ui="resolvedUi"
          :toolbar="false"
          :workbook="workbook"
          :theme="theme"
          :locale="locale"
          :features="features"
          :functions="DEMO_FUNCTIONS"
          :printer-profiles="DEMO_PRINTER_PROFILES"
          :printer-profile-id="DEMO_PRINTER_PROFILE_ID"
          :refresh-printer-profiles="refreshDemoPrinterProfiles"
          :get-function-argument-help="demoFunctionArgumentHelp"
          :upload-status="uploadStatus"
          :macro-recording="scriptOpen"
          @ready="onReady"
          @cell-change="onCellChange"
        />
        <DemoBackstage
          v-if="ribbonTab === 'file'"
          :ui="ui"
          :instance="instance"
          :book-name="bookName"
          :action="backstageAction"
          @action="runBackstageAction"
        />
      </div>
      <DemoOptionsPanel
        :ui="ui"
        :command-text="commandText"
        :instance="instance"
        :hidden="!showPanel"
        :theme="theme"
        :locale="locale"
        :preset="preset"
        :features="features"
        :ribbon="resolvedUi.ribbon"
        :show-ribbon="showRibbon"
        :formatters="formatters"
        :log="log"
        @theme-change="theme = $event"
        @locale-change="locale = $event"
        @preset-change="onPresetChange"
        @feature-toggle="onFeatureToggle"
        @ribbon-change="showRibbon = $event"
        @formatter-change="onFormatterChange"
      />
    </main>
    <DemoReviewDialog
      v-if="reviewDialog"
      :dialog="reviewDialog"
      :ui="ui"
      @close="closeReviewDialog"
    />
    <DemoScriptDialog
      v-if="scriptOpen"
      :ui="ui"
      :command-text="commandText"
      @submit="onScriptSubmit"
      @close="closeScriptDialog"
    />
  </div>
</template>
