<script setup lang="ts">
import {
  type CellChangeEvent,
  type FeatureFlags,
  type FeatureId,
  type HostProviders,
  type LegacySpreadsheetInstance,
  type LocaleChangeEvent,
  parseScriptCommand,
  type ScriptCommand,
  type SpreadsheetInstance,
  type ThemeName,
  WorkbookHandle,
} from '@libraz/formulon-cell';
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
import { legacyInstanceOf } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { workbookOf } from '../../../packages/formulon-cell/src/api/workbook.js';
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
  demoChangeLogEntry,
  pushDemoChangeLog,
  nextDemoFeatureOverrides,
  openDemoWorkbookFile,
  installDemoScriptMenu,
  type ChangeLogEntry,
  type DemoReviewDialogState,
  demoCommandText,
  type DemoBackstageAction,
  type DemoF6Toolbar,
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

/** The ribbon as F6 navigation reads it: the wrapper element around it and its tab focus. */
const f6ToolbarOf = (toolbar: Toolbar | null, host: HTMLElement | null): DemoF6Toolbar | null =>
  host ? { host, focusActiveTab: () => toolbar?.focusActiveTab() ?? false } : null;

/** Shows `handle` in `live`. Settles on the instance's `workbook` event for it, or rejects with the
 *  swap's synchronous refusal or its `error` event. */
const swapWorkbook = (live: SpreadsheetInstance, handle: WorkbookHandle): Promise<void> =>
  new Promise<void>((resolve, reject) => {
    const facade = workbookOf(handle);
    if (live.workbook === facade) {
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
        if (event.workbook === facade) settle(resolve);
      }),
      // `workbook` is the source the instance reports a refused workbook swap under.
      live.on('error', (event) => {
        if (event.source === 'workbook') settle(() => reject(event.error));
      }),
    );
    try {
      live.update({ workbook: facade });
    } catch (error) {
      settle(() => reject(error));
    }
  });

let disposeF6Navigation: (() => void) | undefined;

const theme = ref<ThemeName>('paper');
const locale = ref<DemoLocale>(resolveInitialLocale());
const platform = ref<DemoPlatform>(resolveInitialPlatform());
const workbook = shallowRef<WorkbookHandle | null>(null);
// Vue's reactive proxy walks deeply by default; the spreadsheet instance
// holds a canvas + many internal refs that should not be reactivified.
const instance = shallowRef<SpreadsheetInstance | null>(null);
const toolbar = shallowRef<Toolbar | null>(null);
const ribbonHost = ref<HTMLDivElement | null>(null);
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
    platform: platform.value,
  }),
);
const ribbonTabs = computed(() =>
  platform.value === 'mac' ? DEMO_MAC_RIBBON_TABS : DEMO_RIBBON_TABS,
);
const features = computed<FeatureFlags>(() => resolvedUi.value.features);
const ui = computed(() => demoUiStringsForLocale(UI, locale.value));
const commandText = computed(() => demoCommandText(locale.value));
// The demo-shared helpers and the E2E hook still drive the mount behind the public instance.
const legacy = computed<LegacySpreadsheetInstance | null>(() =>
  instance.value ? legacyInstanceOf(instance.value) : null,
);
// The demo keeps owning the engine handle; the facade leaves its disposal to the demo.
const workbookFacade = computed(() => (workbook.value ? workbookOf(workbook.value) : null));
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

const onCellChange = (e: CellChangeEvent): void => {
  log.value = pushDemoChangeLog(log.value, demoChangeLogEntry(e));
};

watch(
  legacy,
  (target, _o, onCleanup) => {
    if (!target) return;
    const disposers = DEMO_FUNCTIONS.map((f) => target.formula.register(f.name, f.impl, f.meta));
    disposers.push(target.on('cellChange', onCellChange));
    onCleanup(() => {
      for (const d of disposers) d();
    });
  },
  { immediate: true },
);

watch(
  [legacy, () => formatters.value.uppercase, () => formatters.value.arrows],
  (_n, _o, onCleanup) => {
    const target = legacy.value;
    if (!target) return;
    const disposers: (() => void)[] = [];
    if (formatters.value.uppercase) {
      disposers.push(target.cells.registerFormatter(FORMATTERS.uppercaseA));
    }
    if (formatters.value.arrows) {
      disposers.push(target.cells.registerFormatter(FORMATTERS.arrowNegatives));
    }
    onCleanup(() => {
      for (const d of disposers) d();
    });
  },
  { immediate: true },
);

// Expose the live instance on `window.__fcInst` so cross-demo E2E scenarios
// can drive imperative paths without depending on demo-specific UI.
watch(
  legacy,
  (target) => {
    (window as unknown as { __fcInst?: LegacySpreadsheetInstance | null }).__fcInst = target;
  },
  { immediate: true },
);

const onReady = (inst: SpreadsheetInstance): void => {
  instance.value = inst;
};

const onSpellingReview = (): void => {
  const inst = legacy.value;
  if (!inst) return;
  reviewDialog.value = buildDemoReviewFindings('spelling', inst, locale.value, commandText.value);
};

const onAccessibilityCheck = (): void => {
  const inst = legacy.value;
  if (!inst) return;
  // The Mac host docks its own checker; other platforms keep the review list.
  if (inst.openAccessibilityChecker()) return;
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

const ribbonHooks = computed<RibbonHooks>(() => {
  const text = commandText.value;
  return {
    review: {
      spelling: onSpellingReview,
      accessibility: onAccessibilityCheck,
      translate: () => showRibbonNotice(text.translate, text.translationUnavailable),
    },
    drawing: {
      setInkMode: (mode) =>
        showRibbonNotice(text.draw, mode === 'pen' ? text.inkNotPersisted : text.selectInkFirst),
    },
    automation: {
      runScript: onRunScript,
      addInManager: () => showRibbonNotice(text.addIns, text.addInsHostCallbacks),
    },
  };
});

const applyParsedScript = (command: ReturnType<typeof parseScriptCommand>): void => {
  const inst = legacy.value;
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
  instance: legacy.value,
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
  const wb = await WorkbookHandle.createDefault({ locale: locale.value, platform: platform.value });
  if (!demoAlive) {
    wb.dispose();
    return;
  }
  const previous = workbookOwner;
  const live = instance.value;
  try {
    // Swapping here rather than through the prop reports a refused swap to this flow; the prop
    // then already matches the instance.
    if (live) await swapWorkbook(live, wb);
  } catch (error) {
    wb.dispose();
    if (!demoAlive) return;
    reviewDialog.value = buildDemoReviewDialog(
      commandText.value.openFailed,
      commandText.value.workbook,
      formatLoadError(error),
    );
    return;
  }
  if (!demoAlive) {
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
    instance: legacy.value,
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
  const mounted = legacy.value;
  if (!live || !mounted || !demoAlive) return;
  const previous = workbookOwner;
  try {
    const name = await openDemoWorkbookFile(mounted, file);
    if (!demoAlive || instance.value !== live) return;
    // The helper swapped the workbook on the mount; bring the public instance to the same one.
    const next = mounted.workbook;
    await swapWorkbook(live, next);
    if (!demoAlive) return;
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
    instance: legacy.value,
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

onMounted(() => {
  disposeF6Navigation = installDemoF6Navigation({
    getQuickAccess: () => titleBar.value?.quickAccess ?? null,
    getToolbar: () => f6ToolbarOf(toolbar.value, ribbonHost.value),
    getInstance: () => legacy.value,
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
  delete (window as unknown as { __fcInst?: LegacySpreadsheetInstance | null }).__fcInst;
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
      :instance="legacy"
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
        <div v-if="resolvedUi.ribbon" ref="ribbonHost" style="display: contents">
          <SpreadsheetToolbar
            v-model:activeTab="ribbonTab"
            :spreadsheet="instance"
            :ribbon-tabs="ribbonTabs"
            :hooks="ribbonHooks"
            @ready="onToolbarReady"
          />
        </div>
        <Spreadsheet
          class="demo__sheet"
          :ui="resolvedUi"
          :workbook="workbookFacade ?? undefined"
          :theme="theme"
          :locale="locale"
          :host="host"
          :upload-status="uploadStatus"
          :macro-recording="scriptOpen"
          @ready="onReady"
          @error="onSpreadsheetError"
          @locale-change="onCoreLocaleChange"
        />
        <DemoBackstage
          v-if="ribbonTab === 'file'"
          :ui="ui"
          :instance="legacy"
          :book-name="bookName"
          :action="backstageAction"
          @action="runBackstageAction"
        />
      </div>
      <DemoOptionsPanel
        :ui="ui"
        :command-text="commandText"
        :spreadsheet="instance"
        :legacy="legacy"
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
