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
import { Spreadsheet, useSelection } from '@libraz/formulon-cell-vue';
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

type PresetKey = 'minimal' | 'standard' | 'excel';
const PRESETS: { value: PresetKey; label: string; hint: string }[] = [
  { value: 'minimal', label: 'Minimal', hint: 'grid + formula bar only' },
  { value: 'standard', label: 'Standard', hint: 'menus, find/replace, painter' },
  { value: 'excel', label: 'Excel', hint: 'full Excel 365 chrome' },
];

const FEATURE_GROUPS: { title: string; features: { id: FeatureId; label: string }[] }[] = [
  {
    title: 'Chrome',
    features: [
      { id: 'formulaBar', label: 'Formula bar' },
      { id: 'statusBar', label: 'Status bar' },
      { id: 'contextMenu', label: 'Context menu' },
      { id: 'watchWindow', label: 'Watch window' },
    ],
  },
  {
    title: 'Editing',
    features: [
      { id: 'clipboard', label: 'Clipboard' },
      { id: 'pasteSpecial', label: 'Paste special' },
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
      { id: 'formatDialog', label: 'Format dialog' },
      { id: 'fxDialog', label: 'Function dialog' },
      { id: 'conditional', label: 'Conditional formatting' },
      { id: 'namedRanges', label: 'Named ranges' },
      { id: 'hyperlink', label: 'Hyperlink' },
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
const preset = ref<PresetKey>('excel');
const overrides = ref<FeatureFlags>({});

const features = computed<FeatureFlags>(() => composeFeatures(preset.value, overrides.value));

void WorkbookHandle.createDefault().then((wb) => {
  // Core only auto-seeds when it owns the workbook (no `workbook` prop).
  // The demo passes a pre-built handle, so seed by hand here.
  seed(wb);
  workbook.value = wb;
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
  a.download = 'vue-demo.xlsx';
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
  const buf = await file.arrayBuffer();
  const next = await WorkbookHandle.loadBytes(new Uint8Array(buf));
  await inst.setWorkbook(next);
};

const onPresetChange = (next: PresetKey): void => {
  if (next === preset.value) return;
  preset.value = next;
  overrides.value = {};
};

const onFeatureToggle = (id: FeatureId): void => {
  const presetFlags = presets[preset.value]();
  const presetDefault =
    id === 'watchWindow' ? presetFlags[id] === true : presetFlags[id] !== false;
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

// `watchWindow` ships default-off; everything else is opt-out.
const isFeatureOn = (id: FeatureId): boolean =>
  id === 'watchWindow' ? features.value[id] === true : features.value[id] !== false;

onUnmounted(() => {
  // The Spreadsheet component disposes itself; nothing extra to clean up.
});
</script>

<template>
  <div v-if="!workbook" class="demo demo--loading">Loading engine…</div>
  <div v-else class="demo" :data-theme="theme">
    <header class="demo__head">
      <div class="demo__brand">
        <span class="demo__brand-mark">⊞</span>
        <strong>formulon-cell</strong>
        <span class="demo__brand-sep">·</span>
        <span class="demo__brand-tag">vue demo</span>
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
        <button type="button" class="demo__btn" @click="fileInput?.click()">
          Open xlsx…
        </button>
        <button type="button" class="demo__btn" :disabled="!instance" @click="onSave">
          Save
        </button>
        <input ref="fileInput" type="file" accept=".xlsx,.xlsm" hidden @change="onOpenFiles" />
      </div>
    </header>

    <main class="demo__body">
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
      <aside class="demo__panel" aria-label="Demo panel">
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
