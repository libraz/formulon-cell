<script setup lang="ts">
import type {
  CellValue,
  FeatureFlags,
  FeatureId,
  SpreadsheetInstance,
  ThemeName,
} from '@libraz/formulon-cell';
import { useSelection } from '@libraz/formulon-cell-vue';
import { computed, ref, toRef } from 'vue';
import {
  type ChangeLogEntry,
  type DemoCommandStrings,
  type DemoUiStrings,
  demoSelectionLabel,
  evaluateDemoProbe,
  FEATURE_GROUPS,
  isDemoFeatureOn,
  LOCALES,
  PRESETS,
  type PresetKey,
  THEMES,
} from '../../demo-shared/index.js';

const props = defineProps<{
  ui: DemoUiStrings;
  commandText: DemoCommandStrings;
  instance: SpreadsheetInstance | null;
  hidden: boolean;
  theme: ThemeName;
  locale: string;
  preset: PresetKey;
  features: FeatureFlags;
  /** Resolved ribbon visibility (drives the toggle's highlight). */
  ribbon: boolean;
  /** Requested ribbon visibility (drives the checkbox state). */
  showRibbon: boolean;
  formatters: { uppercase: boolean; arrows: boolean };
  log: ChangeLogEntry[];
}>();
const emit = defineEmits<{
  themeChange: [theme: ThemeName];
  localeChange: [locale: string];
  presetChange: [preset: PresetKey];
  featureToggle: [id: FeatureId];
  ribbonChange: [checked: boolean];
  formatterChange: [key: 'uppercase' | 'arrows', checked: boolean];
}>();

const probe = ref<{ name: string; result: string } | null>(null);
const selection = useSelection(toRef(props, 'instance'));
const selectionLabel = computed(() => demoSelectionLabel(selection.value));

const runProbe = (name: string, args: CellValue[]): void => {
  if (!props.instance) return;
  probe.value = evaluateDemoProbe(props.instance, name, args);
};
</script>

<template>
  <aside class="demo__panel" :aria-label="ui.optionsPanel" :hidden="hidden">
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
            @click="emit('themeChange', t.value)"
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
            @click="emit('localeChange', l.value)"
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
          @click="emit('presetChange', p.value)"
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
            :class="['demo__feat', { 'demo__feat--on': isDemoFeatureOn(features, f.id) }]"
          >
            <input
              type="checkbox"
              :checked="isDemoFeatureOn(features, f.id)"
              @change="emit('featureToggle', f.id)"
            />
            <span>{{ ui.featureLabels[f.id] ?? f.label }}</span>
          </label>
          <label
            v-if="group.title === 'Chrome'"
            :class="['demo__feat', { 'demo__feat--on': ribbon }]"
          >
            <input
              type="checkbox"
              :checked="showRibbon"
              @change="emit('ribbonChange', ($event.target as HTMLInputElement).checked)"
            />
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
      <label class="fc-tb__check">
        <input
            type="checkbox"
            :checked="formatters.uppercase"
            @change="emit('formatterChange', 'uppercase', ($event.target as HTMLInputElement).checked)"
          />
        {{ ui.uppercaseColumnA }}
      </label>
      <label class="fc-tb__check">
        <input
            type="checkbox"
            :checked="formatters.arrows"
            @change="emit('formatterChange', 'arrows', ($event.target as HTMLInputElement).checked)"
          />
        {{ ui.arrowPrefixNegatives }}
      </label>
    </section>

    <section class="demo__card">
      <h2>{{ ui.customFunctions }}</h2>
      <p class="demo__hint">{{ ui.customFunctionsHint }}</p>
      <div class="demo__probe">
        <button
          type="button"
          class="fc-tb__btn fc-tb__btn--ghost"
          :disabled="!instance"
          @click="runProbe('GREET', [{ kind: 'text', value: 'Workbook' }])"
        >
          GREET("Workbook")
        </button>
        <button
          type="button"
          class="fc-tb__btn fc-tb__btn--ghost"
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
      <p v-if="log.length === 0" class="fc-tb__empty">
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
</template>
