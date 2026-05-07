<script setup lang="ts">
import {
  applyMerge,
  applyUnmerge,
  autoSum,
  bumpDecimals,
  cycleBorders,
  cycleCurrency,
  cyclePercent,
  mutators,
  recordFormatChange,
  setAlign,
  setFillColor,
  setFreezePanes,
  setFont,
  setFontColor,
  type SpreadsheetInstance,
  toggleBold,
  toggleItalic,
  toggleStrike,
  toggleUnderline,
  toggleWrap,
} from '@libraz/formulon-cell';
import { computed, onUnmounted, ref, watch } from 'vue';

interface Props {
  instance: SpreadsheetInstance | null;
}

const props = defineProps<Props>();

interface ActiveState {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  alignLeft: boolean;
  alignCenter: boolean;
  alignRight: boolean;
  currency: boolean;
  percent: boolean;
  frozen: boolean;
  fontFamily: string;
  fontSize: number;
  fontColor: string;
  fillColor: string;
}

const EMPTY: ActiveState = {
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  alignLeft: false,
  alignCenter: false,
  alignRight: false,
  currency: false,
  percent: false,
  frozen: false,
  fontFamily: 'Aptos',
  fontSize: 11,
  fontColor: '#201f1e',
  fillColor: '#ffffff',
};

const FONT_FAMILIES = ['Aptos', 'Calibri', 'Arial', 'Segoe UI', 'Times New Roman', 'Consolas'];
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36];

const project = (inst: SpreadsheetInstance): ActiveState => {
  const s = inst.store.getState();
  const a = s.selection.active;
  const f = s.format.formats.get(`${a.sheet}:${a.row}:${a.col}`);
  return {
    bold: !!f?.bold,
    italic: !!f?.italic,
    underline: !!f?.underline,
    strike: !!f?.strike,
    alignLeft: f?.align === 'left',
    alignCenter: f?.align === 'center',
    alignRight: f?.align === 'right',
    currency: f?.numFmt?.kind === 'currency',
    percent: f?.numFmt?.kind === 'percent',
    frozen: s.layout.freezeRows > 0 || s.layout.freezeCols > 0,
    fontFamily: f?.fontFamily ?? 'Aptos',
    fontSize: f?.fontSize ?? 11,
    fontColor: f?.color ?? '#201f1e',
    fillColor: f?.fill ?? '#ffffff',
  };
};

const active = ref<ActiveState>(EMPTY);
let unsub: (() => void) | null = null;

watch(
  () => props.instance,
  (inst) => {
    unsub?.();
    unsub = null;
    if (!inst) return;
    active.value = project(inst);
    unsub = inst.store.subscribe(() => {
      active.value = project(inst);
    });
  },
  { immediate: true },
);

onUnmounted(() => {
  unsub?.();
});

const disabled = computed(() => !props.instance);

const wrapFormat = (
  fn: (
    state: ReturnType<SpreadsheetInstance['store']['getState']>,
    store: SpreadsheetInstance['store'],
  ) => void,
): void => {
  const inst = props.instance;
  if (!inst) return;
  recordFormatChange(inst.history, inst.store, () => fn(inst.store.getState(), inst.store));
};

const onUndo = (): void => {
  props.instance?.undo();
};
const onRedo = (): void => {
  props.instance?.redo();
};

const onAutoSum = (): void => {
  const inst = props.instance;
  if (!inst) return;
  const result = autoSum(inst.store.getState(), inst.workbook);
  if (!result) return;
  mutators.replaceCells(inst.store, inst.workbook.cells(result.addr.sheet));
  mutators.setActive(inst.store, result.addr);
};

const onMerge = (): void => {
  const inst = props.instance;
  if (!inst) return;
  const s = inst.store.getState();
  const r = s.selection.range;
  const anchor = s.merges.byAnchor.get(`${r.sheet}:${r.r0}:${r.c0}`);
  const isExact =
    anchor && r.r0 === anchor.r0 && r.c0 === anchor.c0 && r.r1 === anchor.r1 && r.c1 === anchor.c1;
  if (isExact) applyUnmerge(inst.store, inst.workbook, inst.history, r);
  else applyMerge(inst.store, inst.workbook, inst.history, r);
};

const onFreezeToggle = (): void => {
  const inst = props.instance;
  if (!inst) return;
  const s = inst.store.getState();
  if (s.layout.freezeRows > 0 || s.layout.freezeCols > 0) {
    setFreezePanes(inst.store, inst.history, 0, 0, inst.workbook);
  } else {
    const a = s.selection.active;
    const rows = a.row === 0 && a.col === 0 ? 1 : a.row;
    const cols = a.row === 0 && a.col === 0 ? 0 : a.col;
    setFreezePanes(inst.store, inst.history, rows, cols, inst.workbook);
  }
};

const onAlign = (kind: 'left' | 'center' | 'right'): void => {
  wrapFormat((s, st) => setAlign(s, st, kind));
};
const onBumpDecimals = (delta: 1 | -1): void => {
  wrapFormat((s, st) => bumpDecimals(s, st, delta));
};
const onFontFamily = (value: string): void => {
  wrapFormat((s, st) => setFont(s, st, { fontFamily: value }));
};
const onFontSize = (value: string): void => {
  wrapFormat((s, st) => setFont(s, st, { fontSize: Number(value) }));
};
const onFontColor = (value: string): void => {
  wrapFormat((s, st) => setFontColor(s, st, value));
};
const onFillColor = (value: string): void => {
  wrapFormat((s, st) => setFillColor(s, st, value));
};
</script>

<template>
  <div class="demo__ribbon-shell">
    <div class="demo__ribbon-tabs" role="tablist" aria-label="Ribbon tabs">
      <button class="demo__ribbon-tab demo__ribbon-tab--file" type="button" role="tab" aria-selected="false">File</button>
      <button class="demo__ribbon-tab demo__ribbon-tab--active" type="button" role="tab" aria-selected="true">Home</button>
      <button class="demo__ribbon-tab" type="button" role="tab" aria-selected="false" disabled>Insert</button>
      <button class="demo__ribbon-tab" type="button" role="tab" aria-selected="false" disabled>Formulas</button>
      <button class="demo__ribbon-tab" type="button" role="tab" aria-selected="false" disabled>Data</button>
      <button class="demo__ribbon-tab" type="button" role="tab" aria-selected="false" disabled>View</button>
    </div>
    <div class="demo__ribbon" role="toolbar" aria-label="Home ribbon">
      <section class="demo__ribbon-group" aria-label="Clipboard">
        <div class="demo__ribbon-tools">
    <button class="demo__rb" type="button" :disabled="disabled" title="Undo (⌘Z)" aria-label="Undo (⌘Z)" @click="onUndo">
      <svg class="demo__rb-icon" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.45" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
        <path d="M7.2 5.2H3.8v-3.4" />
        <path d="M4 5.2c2.2-2.1 5.7-2.3 8.1-.5 2.7 2.1 3 6.1.7 8.6-1.8 1.9-4.8 2.4-7.1 1.2" />
      </svg>
    </button>
    <button class="demo__rb" type="button" :disabled="disabled" title="Redo (⌘⇧Z)" aria-label="Redo (⌘⇧Z)" @click="onRedo">
      <svg class="demo__rb-icon" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.45" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
        <path d="M12.8 5.2h3.4v-3.4" />
        <path d="M16 5.2c-2.2-2.1-5.7-2.3-8.1-.5-2.7 2.1-3 6.1-.7 8.6 1.8 1.9 4.8 2.4 7.1 1.2" />
      </svg>
    </button>
        </div>
        <div class="demo__ribbon-label">Clipboard</div>
      </section>

      <section class="demo__ribbon-group" aria-label="Number">
        <div class="demo__ribbon-tools">
    <button class="demo__rb demo__rb--mono" :class="{ 'demo__rb--active': active.currency }" type="button" :disabled="disabled" title="Currency" aria-label="Currency" @click="wrapFormat(cycleCurrency)">$</button>
    <button class="demo__rb demo__rb--mono" :class="{ 'demo__rb--active': active.percent }" type="button" :disabled="disabled" title="Percent" aria-label="Percent" @click="wrapFormat(cyclePercent)">%</button>
    <button class="demo__rb" type="button" :disabled="disabled" title="Decrease decimals" aria-label="Decrease decimals" @click="onBumpDecimals(-1)">
      <svg class="demo__rb-icon" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.45" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
        <path d="M3 14.5h5" />
        <path d="M11 5.5h6" />
        <path d="M11 9.5h4" />
        <path d="M11 13.5h2" />
        <path d="M5.5 5.8v6.5" />
        <path d="M3.8 10.5l1.7 1.8 1.7-1.8" />
      </svg>
    </button>
    <button class="demo__rb" type="button" :disabled="disabled" title="Increase decimals" aria-label="Increase decimals" @click="onBumpDecimals(1)">
      <svg class="demo__rb-icon" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.45" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
        <path d="M3 14.5h5" />
        <path d="M11 5.5h2" />
        <path d="M11 9.5h4" />
        <path d="M11 13.5h6" />
        <path d="M5.5 12.2V5.7" />
        <path d="M3.8 7.5l1.7-1.8 1.7 1.8" />
      </svg>
    </button>
    <button class="demo__rb" type="button" :disabled="disabled" title="AutoSum (Σ)" aria-label="AutoSum (Σ)" @click="onAutoSum">
      <svg class="demo__rb-icon" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.45" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
        <path d="M15.5 4.5H5.2l5 5.5-5 5.5h10.3" />
      </svg>
    </button>
        </div>
        <div class="demo__ribbon-label">Number</div>
      </section>

      <section class="demo__ribbon-group" aria-label="Font">
        <div class="demo__ribbon-tools">
    <select class="demo__rb-select demo__rb-select--font" :value="active.fontFamily" :disabled="disabled" title="Font" aria-label="Font" @change="onFontFamily(($event.target as HTMLSelectElement).value)">
      <option v-for="font in FONT_FAMILIES" :key="font" :value="font">{{ font }}</option>
    </select>
    <select class="demo__rb-select" :value="active.fontSize" :disabled="disabled" title="Font size" aria-label="Font size" @change="onFontSize(($event.target as HTMLSelectElement).value)">
      <option v-for="size in FONT_SIZES" :key="size" :value="size">{{ size }}</option>
    </select>
    <button class="demo__rb demo__rb--bold" :class="{ 'demo__rb--active': active.bold }" type="button" :disabled="disabled" title="Bold (⌘B)" aria-label="Bold (⌘B)" @click="wrapFormat(toggleBold)">B</button>
    <button class="demo__rb demo__rb--italic" :class="{ 'demo__rb--active': active.italic }" type="button" :disabled="disabled" title="Italic (⌘I)" aria-label="Italic (⌘I)" @click="wrapFormat(toggleItalic)">I</button>
    <button class="demo__rb demo__rb--underline" :class="{ 'demo__rb--active': active.underline }" type="button" :disabled="disabled" title="Underline (⌘U)" aria-label="Underline (⌘U)" @click="wrapFormat(toggleUnderline)">U</button>
    <button class="demo__rb demo__rb--strike" :class="{ 'demo__rb--active': active.strike }" type="button" :disabled="disabled" title="Strikethrough" aria-label="Strikethrough" @click="wrapFormat(toggleStrike)">S</button>
    <button class="demo__rb" type="button" :disabled="disabled" title="Borders" aria-label="Borders" @click="wrapFormat(cycleBorders)">
      <svg class="demo__rb-icon" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.45" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
        <path d="M4 4h12v12H4z" />
        <path d="M10 4v12" />
        <path d="M4 10h12" />
      </svg>
    </button>
    <label class="demo__rb-color" title="Font color" aria-label="Font color">
      <span>A</span>
      <input type="color" :value="active.fontColor" :disabled="disabled" @change="onFontColor(($event.target as HTMLInputElement).value)" />
    </label>
    <label class="demo__rb-color" title="Fill color" aria-label="Fill color">
      <span>▾</span>
      <input type="color" :value="active.fillColor" :disabled="disabled" @change="onFillColor(($event.target as HTMLInputElement).value)" />
    </label>
        </div>
        <div class="demo__ribbon-label">Font</div>
      </section>

      <section class="demo__ribbon-group" aria-label="Alignment">
        <div class="demo__ribbon-tools">
    <button class="demo__rb" :class="{ 'demo__rb--active': active.alignLeft }" type="button" :disabled="disabled" title="Align left" aria-label="Align left" @click="onAlign('left')">
      <svg class="demo__rb-icon" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.45" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
        <path d="M4 5h12" />
        <path d="M4 8.5h8" />
        <path d="M4 12h12" />
        <path d="M4 15.5h7" />
      </svg>
    </button>
    <button class="demo__rb" :class="{ 'demo__rb--active': active.alignCenter }" type="button" :disabled="disabled" title="Align center" aria-label="Align center" @click="onAlign('center')">
      <svg class="demo__rb-icon" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.45" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
        <path d="M4 5h12" />
        <path d="M6 8.5h8" />
        <path d="M4 12h12" />
        <path d="M6.5 15.5h7" />
      </svg>
    </button>
    <button class="demo__rb" :class="{ 'demo__rb--active': active.alignRight }" type="button" :disabled="disabled" title="Align right" aria-label="Align right" @click="onAlign('right')">
      <svg class="demo__rb-icon" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.45" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
        <path d="M4 5h12" />
        <path d="M8 8.5h8" />
        <path d="M4 12h12" />
        <path d="M9 15.5h7" />
      </svg>
    </button>
    <button class="demo__rb" type="button" :disabled="disabled" title="Merge cells" aria-label="Merge cells" @click="onMerge">
      <svg class="demo__rb-icon" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.45" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
        <path d="M4 5h12v10H4z" />
        <path d="M8 5v10" />
        <path d="M12 5v10" />
        <path d="M7 10h6" />
        <path d="M11.5 8.5L13 10l-1.5 1.5" />
        <path d="M8.5 8.5L7 10l1.5 1.5" />
      </svg>
    </button>
    <button class="demo__rb" type="button" :disabled="disabled" title="Wrap text" aria-label="Wrap text" @click="wrapFormat(toggleWrap)">
      <svg class="demo__rb-icon" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.45" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
        <path d="M4 5h12" />
        <path d="M4 9h9a3 3 0 0 1 0 6H8" />
        <path d="M9.8 12.8L7.6 15l2.2 2.2" />
      </svg>
    </button>
        </div>
        <div class="demo__ribbon-label">Alignment</div>
      </section>

      <section class="demo__ribbon-group" aria-label="View">
        <div class="demo__ribbon-tools">
    <button class="demo__rb" :class="{ 'demo__rb--active': active.frozen }" type="button" :disabled="disabled" title="Freeze panes" aria-label="Freeze panes" @click="onFreezeToggle">
      <svg class="demo__rb-icon" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.45" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
        <path d="M4 4h12v12H4z" />
        <path d="M4 8h12" />
        <path d="M8 4v12" />
        <path d="M8 8h8v8H8z" />
      </svg>
    </button>
        </div>
        <div class="demo__ribbon-label">View</div>
      </section>
    </div>
  </div>
</template>
