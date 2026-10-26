<script setup lang="ts">
import {
  type CellChangeEvent,
  type FeatureFlags,
  type FeatureId,
  type LocaleChangeEvent,
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
  DEMO_FUNCTIONS,
  demoFunctionArgumentHelp,
  DEMO_MAC_RIBBON_TABS,
  DEMO_PRINTER_PROFILE_ID,
  DEMO_PRINTER_PROFILES,
  DEMO_RIBBON_TABS,
  pushDemoChangeLog,
  nextDemoFeatureOverrides,
  openDemoWorkbookFile,
  installDemoScriptMenu,
  type ChangeLogEntry,
  type DemoReviewDialogState,
  demoCommandText,
  type DemoBackstageAction,
  type DemoLocale,
  demoUiStringsForLocale,
  formatLoadError,
  FORMATTERS,
  installDemoF6Navigation,
  normalizeDemoLocale,
  type PresetKey,
  refreshDemoPrinterProfiles,
  resolveInitialLocale,
  resolveInitialPlatform,
  type DemoPlatform,
  type DemoUploadStatus,
  reportDemoScriptRun,
  runDemoBackstageAction,
  saveDemoWorkbookAsToDownload,
  saveDemoWorkbookToDownload,
} from '../../demo-shared/index.js';
import DemoBackstage from './DemoBackstage.vue';
import DemoOptionsPanel from './DemoOptionsPanel.vue';
import DemoReviewDialog from './DemoReviewDialog.vue';
import DemoScriptDialog from './DemoScriptDialog.vue';
import DemoTitleBar from './DemoTitleBar.vue';

const UI = createDemoStrings('Vue');



let disposeF6Navigation: (() => void) | undefined;

const theme = ref<ThemeName>('paper');
const locale = ref<DemoLocale>(resolveInitialLocale());
const platform = ref<DemoPlatform>(resolveInitialPlatform());
const workbook = shallowRef<WorkbookHandle | null>(null);
// Vue's reactive proxy walks deeply by default; the spreadsheet instance
// holds a canvas + many internal refs that should not be reactivified.
const instance = shallowRef<SpreadsheetInstance | null>(null);
const toolbar = shallowRef<ToolbarInstance | null>(null);
const log = ref<ChangeLogEntry[]>([]);
const formatters = ref({ uppercase: false, arrows: false });
const fileInput = ref<HTMLInputElement | null>(null);
const titleBar = ref<InstanceType<typeof DemoTitleBar> | null>(null);
const preset = ref<PresetKey>('full');
const overrides = ref<FeatureFlags>({});
const showRibbon = ref(true);
const showPanel = ref(false);
const ribbonTab = ref<RibbonTab>('home');
const backstageAction = ref<DemoBackstageAction>('info');
const bookName = ref('Book1');
const loadError = ref<string | null>(null);
const reviewDialog = ref<DemoReviewDialogState | null>(null);
const scriptOpen = ref(false);
const uploadStatus = ref<'saved' | 'saving' | 'error' | null>(null);
let demoAlive = true;
let workbookOwner: WorkbookHandle | null = null;

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
const ui = computed(() => demoUiStringsForLocale(UI, locale.value));
const commandText = computed(() => demoCommandText(locale.value));

watch(
  locale,
  (next) => {
    document.documentElement.lang = next;
  },
  { immediate: true },
);

void createInitialDemoWorkbook()
  .then((wb) => {
    if (!demoAlive) {
      // A late-created demo workbook is still owned by this app and must be
      // released when the component has already unmounted.
      wb.dispose();
      return;
    }
    workbookOwner = wb;
    loadError.value = null;
    workbook.value = wb;
  })
  .catch((err: unknown) => {
    if (!demoAlive) return;
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

const demoSaveOptions = () => ({
  instance: instance.value,
  bookName: bookName.value,
  setUploadStatus: (next: DemoUploadStatus) => {
    uploadStatus.value = next;
  },
});

const onSave = (): void => {
  void saveDemoWorkbookToDownload(demoSaveOptions());
};

const onSaveAs = (): void => {
  void saveDemoWorkbookAsToDownload(demoSaveOptions());
};

const onNewWorkbook = async (): Promise<void> => {
  const wb = await WorkbookHandle.createDefault({ locale: locale.value, platform: platform.value });
  if (!demoAlive) {
    wb.dispose();
    return;
  }
  const previous = workbookOwner;
  try {
    await instance.value?.setWorkbook(wb);
  } catch (error) {
    if (!demoAlive) {
      wb.dispose();
      return;
    }
    wb.dispose();
    reviewDialog.value = buildDemoReviewDialog(
      commandText.value.openFailed,
      commandText.value.workbook,
      formatLoadError(error),
    );
    return;
  }
  if (!demoAlive) {
    // The host may have unmounted while setWorkbook was pending. The core
    // instance cleanup already ran; do not revive Vue state from this result.
    wb.dispose();
    return;
  }
  workbookOwner = wb;
  workbook.value = wb;
  previous?.dispose();
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
    saveWorkbookAs: onSaveAs,
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
  if (!inst || !demoAlive) return;
  const previous = workbookOwner;
  try {
    const name = await openDemoWorkbookFile(inst, file);
    if (!demoAlive || instance.value !== inst) return;
    const next = inst.workbook;
    workbookOwner = next;
    workbook.value = next;
    // The helper has detached the previous workbook before resolving. Release
    // the app-owned borrowed handle only after that successful handoff.
    previous?.dispose();
    bookName.value = name;
    loadError.value = null;
  } catch (err) {
    if (!demoAlive) return;
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

const setLocale = (next: string): void => {
  const resolved = normalizeDemoLocale(next);
  if (resolved) locale.value = resolved;
};

const onCoreLocaleChange = (event: LocaleChangeEvent): void => {
  setLocale(event.locale);
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
    setLocale,
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

const openBackstage = (): void => {
  ribbonTab.value = 'file';
  toolbar.value?.setBackstageOpen(true);
};

const onToolbarReady = (next: ToolbarInstance | null): void => {
  if (next) next.host.dataset.fcPlatform = platform.value;
  toolbar.value = next;
  (window as unknown as { __fcToolbar?: ToolbarInstance | null }).__fcToolbar = next;
};

onMounted(() => {
  disposeF6Navigation = installDemoF6Navigation({
    getQuickAccess: () => titleBar.value?.quickAccess ?? null,
    getToolbar: () => toolbar.value,
    getInstance: () => instance.value,
  });
});

onBeforeUnmount(() => {
  demoAlive = false;
  // The binding borrows the initial workbook. Dispose the core instance first
  // so a workbook adopted during a New action is released by core, then close
  // the app-owned borrowed handle as well.
  try {
    instance.value?.dispose();
  } catch (error) {
    console.error('formulon demo: spreadsheet cleanup failed', error);
  } finally {
    instance.value = null;
    workbookOwner?.dispose();
    workbookOwner = null;
  }
  disposeF6Navigation?.();
  disposeF6Navigation = undefined;
  delete (window as unknown as { __fcInst?: SpreadsheetInstance | null }).__fcInst;
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
    <DemoTitleBar
      ref="titleBar"
      :ui="ui"
      :platform="platform"
      :book-name="bookName"
      :instance="instance"
      :show-panel="showPanel"
      :search-items="searchItems"
      @save="onSave"
      @open-backstage="openBackstage"
      @toggle-panel="showPanel = !showPanel"
      @tab-change="ribbonTab = $event"
    />
    <input ref="fileInput" type="file" accept=".xlsx,.xlsm,.xlsb" hidden @change="onOpenFiles" />

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
          :get-workbook-name="() => bookName"
          :upload-status="uploadStatus"
          :macro-recording="scriptOpen"
          @ready="onReady"
          @cell-change="onCellChange"
          @locale-change="onCoreLocaleChange"
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
        @locale-change="setLocale"
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
