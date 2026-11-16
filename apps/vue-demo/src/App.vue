<script setup lang="ts">
import {
  type CellEvent,
  type HostProviders,
  type SpreadsheetInstance,
  type ThemeName,
  Workbook,
} from '@libraz/formulon-cell';
import type { FeatureFlags, FeatureId } from '@libraz/formulon-cell/extensions';
import type { RibbonHooks, RibbonTab, Toolbar } from '@libraz/formulon-cell/toolbar';
import { Spreadsheet, SpreadsheetToolbar } from '@libraz/formulon-cell-vue';
import {
  computed,
  onBeforeUnmount,
  onMounted,
  ref,
  shallowRef,
  watch,
} from 'vue';
import {
  adoptDemoWorkbook,
  buildDemoCommands,
  buildDemoReviewDialog,
  buildDemoSearchItems,
  composeDemoUiOptions,
  createDemoStrings,
  createInitialDemoWorkbook,
  DEMO_FORMULA_HINTS,
  demoFunctionArgumentHelp,
  DEMO_MAC_RIBBON_TABS,
  DEMO_PRINTER_PROFILE_ID,
  DEMO_PRINTER_PROFILES,
  DEMO_RIBBON_TABS,
  demoChangeLogEntry,
  pushDemoChangeLog,
  nextDemoFeatureOverrides,
  openDemoWorkbookFile,
  type ChangeLogEntry,
  type DemoReviewDialogState,
  demoCommandText,
  type DemoBackstageAction,
  type DemoLocale,
  demoUiStringsForLocale,
  formatLoadError,
  FORMATTERS,
  normalizeDemoLocale,
  type PresetKey,
  refreshDemoPrinterProfiles,
  resolveInitialLocale,
  resolveInitialPlatform,
  type DemoPlatform,
  type DemoUploadStatus,
  runDemoBackstageAction,
  saveDemoWorkbookAsToDownload,
  saveDemoWorkbookToDownload,
} from '../../demo-shared/index.js';
import DemoBackstage from './DemoBackstage.vue';
import DemoOptionsPanel from './DemoOptionsPanel.vue';
import DemoReviewDialog from './DemoReviewDialog.vue';
import DemoTitleBar from './DemoTitleBar.vue';

const UI = createDemoStrings('Vue');

const theme = ref<ThemeName>('paper');
const locale = ref<DemoLocale>(resolveInitialLocale());
const platform = ref<DemoPlatform>(resolveInitialPlatform());
const workbook = shallowRef<Workbook | null>(null);
// Vue's reactive proxy walks deeply by default; the spreadsheet instance
// holds a canvas + many internal refs that should not be reactivified.
const instance = shallowRef<SpreadsheetInstance | null>(null);
const toolbar = shallowRef<Toolbar | null>(null);
const log = ref<ChangeLogEntry[]>([]);
const formatters = ref({ uppercase: false, arrows: false });
const fileInput = ref<HTMLInputElement | null>(null);
const preset = ref<PresetKey>('full');
const overrides = ref<FeatureFlags>({});
const showRibbon = ref(true);
const showPanel = ref(false);
const ribbonTab = ref<RibbonTab>('home');
const backstageAction = ref<DemoBackstageAction>('info');
const bookName = ref('Book1');
const loadError = ref<string | null>(null);
const reviewDialog = ref<DemoReviewDialogState | null>(null);
const uploadStatus = ref<'saved' | 'saving' | 'error' | null>(null);
let demoAlive = true;
let workbookOwner: Workbook | null = null;

const resolvedUi = computed(() =>
  composeDemoUiOptions({
    preset: preset.value,
    overrides: overrides.value,
    showRibbon: showRibbon.value,
    platform: platform.value,
  }),
);
const ribbonTabs = computed(() =>
  platform.value === 'mac' ? DEMO_MAC_RIBBON_TABS : DEMO_RIBBON_TABS,
);
const features = computed<FeatureFlags>(() => resolvedUi.value.features);
const ui = computed(() => demoUiStringsForLocale(UI, locale.value));
const commandText = computed(() => demoCommandText(locale.value));
const host = computed<HostProviders>(() => {
  const name = bookName.value;
  return {
    workbookName: () => name,
    printerProfiles: DEMO_PRINTER_PROFILES,
    printerProfileId: DEMO_PRINTER_PROFILE_ID,
    refreshPrinterProfiles: refreshDemoPrinterProfiles,
    functionArgumentHelp: demoFunctionArgumentHelp,
  };
});

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

const onCellChange = (live: SpreadsheetInstance, e: CellEvent): void => {
  log.value = pushDemoChangeLog(log.value, demoChangeLogEntry(live.workbook, e));
};

watch(
  instance,
  (target, _o, onCleanup) => {
    if (!target) return;
    onCleanup(target.on('cell', (e) => onCellChange(target, e)));
  },
  { immediate: true },
);

watch(
  [instance, () => formatters.value.uppercase, () => formatters.value.arrows],
  (_n, _o, onCleanup) => {
    const target = instance.value;
    if (!target) return;
    const disposers: (() => void)[] = [];
    if (formatters.value.uppercase) disposers.push(target.cells.add(FORMATTERS.uppercaseA));
    if (formatters.value.arrows) disposers.push(target.cells.add(FORMATTERS.arrowNegatives));
    onCleanup(() => {
      for (const d of disposers) d();
    });
  },
  { immediate: true },
);

// Expose the live instance on `window.__fcInst` so cross-demo E2E scenarios
// can drive imperative paths without depending on demo-specific UI.
watch(
  instance,
  (target) => {
    (window as unknown as { __fcInst?: SpreadsheetInstance | null }).__fcInst = target;
  },
  { immediate: true },
);

const onReady = (inst: SpreadsheetInstance): void => {
  instance.value = inst;
};

const closeReviewDialog = (): void => {
  reviewDialog.value = null;
};

const showRibbonNotice = (title: string, detail: string): void => {
  reviewDialog.value = buildDemoReviewDialog(title, commandText.value.ribbonCommand, detail);
};

const ribbonHooks = computed<RibbonHooks>(() => {
  const text = commandText.value;
  return {
    review: {
      translate: () => showRibbonNotice(text.translate, text.translationUnavailable),
    },
    drawing: {
      setInkMode: (mode) =>
        showRibbonNotice(text.draw, mode === 'pen' ? text.inkNotPersisted : text.selectInkFirst),
    },
    automation: {
      addInManager: () => showRibbonNotice(text.addIns, text.addInsHostCallbacks),
    },
  };
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

const onSpreadsheetError = (error: unknown): void => {
  console.error('formulon demo: spreadsheet error', error);
};

const onNewWorkbook = async (): Promise<void> => {
  const wb = await Workbook.createDefault({ locale: locale.value, platform: platform.value });
  const previous = workbookOwner;
  const live = instance.value;
  try {
    // Swapping here rather than through the prop reports a refused swap to this flow; the prop
    // then already matches the instance.
    if (!(await adoptDemoWorkbook(live, wb, () => demoAlive))) return;
  } catch (error) {
    if (!demoAlive) return;
    reviewDialog.value = buildDemoReviewDialog(
      commandText.value.openFailed,
      commandText.value.workbook,
      formatLoadError(error),
    );
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
      toolbar.value?.update({ backstageOpen: false });
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
  const live = instance.value;
  if (!live || !demoAlive) return;
  const previous = workbookOwner;
  try {
    const { workbook: next, name } = await openDemoWorkbookFile(file, {
      locale: locale.value,
      platform: platform.value,
    });
    const current = () => demoAlive && instance.value === live;
    if (!(await adoptDemoWorkbook(live, next, current))) return;
    workbookOwner = next;
    workbook.value = next;
    // The instance no longer shows the previous workbook; release the one the demo owned.
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

const onCoreLocaleChange = (event: { readonly locale: string }): void => {
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
const applyToolbarCommand = (commandId: string): boolean => {
  const current = toolbar.value;
  if (!current) return false;
  try {
    current.execute(commandId);
    return true;
  } catch {
    // A refused command did not run; the caller falls back to showing its tab.
    return false;
  }
};
const searchItems = computed(() =>
  buildDemoSearchItems(
    commands.value,
    locale.value,
    (tab) => {
      ribbonTab.value = tab;
    },
    applyToolbarCommand,
    ribbonTabs.value,
  ),
);

const openBackstage = (): void => {
  ribbonTab.value = 'file';
  toolbar.value?.update({ backstageOpen: true });
};

const onToolbarReady = (next: Toolbar | null): void => {
  toolbar.value = next;
  (window as unknown as { __fcToolbar?: Toolbar | null }).__fcToolbar = next;
};

onBeforeUnmount(() => {
  demoAlive = false;
  // Dispose the instance first, then the workbook the demo owns.
  try {
    instance.value?.dispose();
  } catch (error) {
    console.error('formulon demo: spreadsheet cleanup failed', error);
  } finally {
    instance.value = null;
    workbookOwner?.dispose();
    workbookOwner = null;
  }
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
      :ui="ui"
      :platform="platform"
      :book-name="bookName"
      :instance="instance"
      :show-panel="showPanel"
      :search-items="searchItems"
      :apply-command="applyToolbarCommand"
      @save="onSave"
      @save-as="onSaveAs"
      @rename="bookName = $event"
      @new-workbook="void onNewWorkbook()"
      @open-backstage="openBackstage"
      @toggle-panel="showPanel = !showPanel"
      @tab-change="ribbonTab = $event"
    />
    <input ref="fileInput" type="file" accept=".xlsx,.xlsm,.xlsb" hidden @change="onOpenFiles" />

    <main :class="['demo__body', { 'demo__body--panel': showPanel }]">
      <div class="fc-tb__sheet-col">
        <SpreadsheetToolbar
          v-if="resolvedUi.ribbon"
          v-model:activeTab="ribbonTab"
          :spreadsheet="instance"
          :ribbon-tabs="ribbonTabs"
          :hooks="ribbonHooks"
          @ready="onToolbarReady"
        />
        <Spreadsheet
          class="demo__sheet"
          :ui="resolvedUi"
          :workbook="workbook ?? undefined"
          :formula-hints="DEMO_FORMULA_HINTS"
          :theme="theme"
          :locale="locale"
          :host="host"
          :upload-status="uploadStatus"
          @ready="onReady"
          @error="onSpreadsheetError"
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
        :spreadsheet="instance"
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
  </div>
</template>
