<script setup lang="ts">
import {
  type CellChangeEvent,
  type CellValue,
  type FeatureFlags,
  type FeatureId,
  analyzeAccessibilityCells,
  analyzeSpellingCells,
  applyTextScript,
  mutators,
  parseScriptCommand,
  presets,
  type SpreadsheetInstance,
  type ThemeName,
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
  createDemoStrings,
  DEMO_FUNCTIONS,
  demoColLabel,
  demoCommandText,
  FEATURE_GROUPS,
  formatLoadError,
  FORMATTERS,
  LOCALES,
  type PresetKey,
  PRESETS,
  previewCellChange,
  resolveInitialLocale,
  reviewCellsForInstance,
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

interface CommandItem {
  readonly id: string;
  readonly label: string;
  readonly hint: string;
  readonly tab?: RibbonTab;
  readonly run: () => void;
}

interface ReviewDialogState {
  readonly title: string;
  readonly items: readonly { label: string; detail: string }[];
}

let changeId = 0;

// Modal focus trap + Esc-to-close + change-log preview + seed +
// review-cell projection all live in demo-shared/index.ts so the React
// and Vue demos stay aligned.
const previewValue = previewCellChange;
const seed = seedDemoWorkbook;

const composeFeatures = (preset: PresetKey, overrides: FeatureFlags): FeatureFlags => ({
  ...presets[preset](),
  ...overrides,
});

const theme = ref<ThemeName>('paper');
const locale = ref<string>(resolveInitialLocale());
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
const reviewDialog = ref<ReviewDialogState | null>(null);
const scriptOpen = ref(false);
const scriptCommand = ref('uppercase');
const scriptError = ref<string | null>(null);
const reviewModalEl = ref<HTMLElement | null>(null);
const scriptModalEl = ref<HTMLElement | null>(null);

const features = computed<FeatureFlags>(() => composeFeatures(preset.value, overrides.value));
const ui = computed(() => UI[locale.value === 'ja' ? 'ja' : 'en']);
const commandText = computed(() => demoCommandText(locale.value));

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
    title: 'Spelling Review',
    items: analyzeSpellingCells(reviewCellsForInstance(inst)),
  };
};

const onAccessibilityCheck = (): void => {
  const inst = instance.value;
  if (!inst) return;
  reviewDialog.value = {
    title: commandText.value.accessibilityCheck,
    items: analyzeAccessibilityCells(reviewCellsForInstance(inst)),
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

const commands = computed<CommandItem[]>(() => [
  {
    id: 'open',
    label: commandText.value.commands.open.label,
    hint: commandText.value.commands.open.hint,
    tab: 'file',
    run: () => fileInput.value?.click(),
  },
  {
    id: 'save',
    label: commandText.value.commands.save.label,
    hint: commandText.value.commands.save.hint,
    tab: 'file',
    run: onSave,
  },
  {
    id: 'page-setup',
    label: commandText.value.commands.pageSetup.label,
    hint: commandText.value.commands.pageSetup.hint,
    tab: 'file',
    run: () => instance.value?.openPageSetup(),
  },
  {
    id: 'print',
    label: commandText.value.commands.print.label,
    hint: commandText.value.commands.print.hint,
    tab: 'file',
    run: () => instance.value?.print('print'),
  },
  {
    id: 'format-cells',
    label: commandText.value.commands.formatCells.label,
    hint: commandText.value.commands.formatCells.hint,
    tab: 'home',
    run: () => instance.value?.openFormatDialog(),
  },
  {
    id: 'conditional',
    label: commandText.value.commands.conditionalFormatting.label,
    hint: commandText.value.commands.conditionalFormatting.hint,
    tab: 'insert',
    run: () => instance.value?.openConditionalDialog(),
  },
  {
    id: 'cell-styles',
    label: commandText.value.commands.cellStyles.label,
    hint: commandText.value.commands.cellStyles.hint,
    tab: 'insert',
    run: () => instance.value?.openCellStylesGallery(),
  },
  {
    id: 'name-manager',
    label: commandText.value.commands.nameManager.label,
    hint: commandText.value.commands.nameManager.hint,
    tab: 'insert',
    run: () => instance.value?.openNamedRangeDialog(),
  },
  {
    id: 'insert-function',
    label: commandText.value.commands.insertFunction.label,
    hint: commandText.value.commands.insertFunction.hint,
    tab: 'formulas',
    run: () => instance.value?.openFunctionArguments(),
  },
  {
    id: 'trace-precedents',
    label: commandText.value.commands.tracePrecedents.label,
    hint: commandText.value.commands.tracePrecedents.hint,
    tab: 'formulas',
    run: () => instance.value?.tracePrecedents(),
  },
  {
    id: 'watch-window',
    label: commandText.value.commands.watchWindow.label,
    hint: commandText.value.commands.watchWindow.hint,
    tab: 'formulas',
    run: () => instance.value?.toggleWatchWindow(),
  },
  {
    id: 'filter',
    label: commandText.value.commands.filter.label,
    hint: commandText.value.commands.filter.hint,
    tab: 'data',
    run: () => {
      ribbonTab.value = 'data';
    },
  },
  {
    id: 'sort',
    label: commandText.value.commands.sort.label,
    hint: commandText.value.commands.sort.hint,
    tab: 'data',
    run: () => {
      ribbonTab.value = 'data';
    },
  },
  {
    id: 'freeze-panes',
    label: commandText.value.commands.freezePanes.label,
    hint: commandText.value.commands.freezePanes.hint,
    tab: 'view',
    run: () => {
      ribbonTab.value = 'view';
    },
  },
  {
    id: 'protect-sheet',
    label: commandText.value.commands.protectSheet.label,
    hint: commandText.value.commands.protectSheet.hint,
    tab: 'view',
    run: () => instance.value?.toggleSheetProtection(),
  },
  {
    id: 'options-pane',
    label: commandText.value.commands.options.label,
    hint: commandText.value.commands.options.hint,
    run: () => {
      showPanel.value = !showPanel.value;
    },
  },
  {
    id: 'theme-light',
    label: commandText.value.commands.lightTheme.label,
    hint: commandText.value.commands.lightTheme.hint,
    run: () => {
      theme.value = 'paper';
    },
  },
  {
    id: 'theme-dark',
    label: commandText.value.commands.darkTheme.label,
    hint: commandText.value.commands.darkTheme.hint,
    run: () => {
      theme.value = 'ink';
    },
  },
  {
    id: 'locale-ja',
    label: commandText.value.commands.japaneseLocale.label,
    hint: commandText.value.commands.japaneseLocale.hint,
    run: () => {
      locale.value = 'ja';
    },
  },
  {
    id: 'locale-en',
    label: commandText.value.commands.englishLocale.label,
    hint: commandText.value.commands.englishLocale.hint,
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
  demoModalCleanup?.();
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
            :aria-label="ui.searchCommands"
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
          <span class="demo__avatar" role="img" :aria-label="ui.signedInUser">FC</span>
        </div>
      </div>
      <div class="demo__commandbar">
        <div class="demo__brand">
          <strong>formulon-cell</strong>
          <span class="demo__brand-sep">·</span>
          <span class="demo__brand-tag">{{ ui.workbook }}</span>
        </div>
        <div class="demo__controls">
          <div class="demo__seg" role="group" :aria-label="ui.theme">
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
              @click="instance?.print('print')"
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
                @click="instance?.print('print')"
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
          <h2>{{ commandText.selection }}</h2>
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
            aria-label="Close"
            @click="closeReviewDialog"
          >
            ×
          </button>
        </header>
        <div class="demo__modal-body">
          <p v-if="reviewDialog.items.length === 0" class="demo__modal-empty">
            No issues found.
          </p>
          <ul v-else class="demo__modal-list">
            <li v-for="(item, index) in reviewDialog.items" :key="`${item.label}-${index}`">
              <strong>{{ item.label }}</strong>
              <span>{{ item.detail }}</span>
            </li>
          </ul>
        </div>
        <footer class="demo__modal-footer">
          <button type="button" class="demo__btn" @click="closeReviewDialog">OK</button>
        </footer>
      </section>
    </div>
    <div
      v-if="scriptOpen"
      ref="scriptModalEl"
      class="demo__modal"
      role="dialog"
      aria-modal="true"
      aria-label="Script"
    >
      <form class="demo__modal-panel demo__modal-panel--narrow" @submit.prevent="applyScriptCommand">
        <header class="demo__modal-header">
          <h2>Script</h2>
          <button
            type="button"
            class="demo__modal-x"
            aria-label="Close"
            @click="closeScriptDialog"
          >
            ×
          </button>
        </header>
        <div class="demo__modal-body">
          <label class="demo__modal-field">
            <span>Command</span>
            <input v-model="scriptCommand" autofocus @input="scriptError = null" />
          </label>
          <p v-if="scriptError" class="demo__modal-error">{{ scriptError }}</p>
        </div>
        <footer class="demo__modal-footer">
          <button type="button" class="demo__btn" @click="closeScriptDialog">Cancel</button>
          <button type="submit" class="demo__btn demo__btn--active">Run</button>
        </footer>
      </form>
    </div>
  </div>
</template>
