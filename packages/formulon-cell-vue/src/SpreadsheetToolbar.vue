<script setup lang="ts">
import {
  applyMerge,
  applyUnmerge,
  autoSum,
  bumpDecimals,
  type CellBorderStyle,
  clearFilter,
  clearFormat,
  cycleBorders,
  cycleCurrency,
  cyclePercent,
  deleteCols,
  deleteRows,
  formatAsTable,
  hiddenInSelection,
  hideCols,
  hideRows,
  insertCols,
  insertRows,
  type MarginPreset,
  mutators,
  type NumFmt,
  type PageOrientation,
  type PaperSize,
  recordFormatChange,
  recordMergesChangeWithEngine,
  recordPageSetupChange,
  removeDuplicates,
  setAlign,
  setAutoFilter,
  setBorderPreset,
  setFillColor,
  setFreezePanes,
  setFont,
  setFontColor,
  setMarginPreset,
  setNumFmt,
  setPageOrientation,
  setPaperSize,
  setSheetZoom,
  showCols,
  showRows,
  sortRange,
  type SpreadsheetInstance,
  setVAlign,
  toggleBold,
  toggleItalic,
  toggleStrike,
  toggleUnderline,
  toggleWrap,
} from '@libraz/formulon-cell';
import { RibbonIcon } from './toolbar/icons.js';
import { computed, nextTick, ref } from 'vue';
import { useToolbarActive } from './toolbar/active.js';
import { useToolbarDropdown } from './toolbar/dropdown.js';
import {
  BORDER_PRESETS,
  BORDER_STYLES,
  type BorderPreset,
  FONT_FAMILIES,
  FONT_SIZES,
  RIBBON_KEYSHORTCUTS,
  RIBBON_TAB_LABELS,
  type RibbonTab,
} from './toolbar/model.js';
import { toolbarTabs } from './toolbar/tabs.js';
import { toolbarText } from './toolbar/translations.js';

interface Props {
  instance: SpreadsheetInstance | null;
  activeTab: RibbonTab;
  locale: string;
  onSpellingReview?: () => void;
  onAccessibilityCheck?: () => void;
  onRunScript?: () => void;
  onDrawPen?: () => void;
  onDrawEraser?: () => void;
  onTranslate?: () => void;
  onAddIn?: () => void;
}

type MergeAction = 'mergeCenter' | 'mergeAcross' | 'mergeCells' | 'unmergeCells';
type NumberFormatAction =
  | 'general'
  | 'fixed'
  | 'currency'
  | 'accounting'
  | 'shortDate'
  | 'longDate'
  | 'time'
  | 'percent'
  | 'fraction'
  | 'scientific'
  | 'text'
  | 'more';

const numberFormatForAction = (action: NumberFormatAction, currentLang: 'ja' | 'en'): NumFmt | null => {
  const symbol = currentLang === 'ja' ? '¥' : '$';
  switch (action) {
    case 'general':
      return { kind: 'general' };
    case 'fixed':
      return { kind: 'fixed', decimals: 0 };
    case 'currency':
      return { kind: 'currency', decimals: 0, symbol };
    case 'accounting':
      return { kind: 'accounting', decimals: 0, symbol };
    case 'shortDate':
      return { kind: 'date', pattern: currentLang === 'ja' ? 'yyyy/m/d' : 'm/d/yyyy' };
    case 'longDate':
      return { kind: 'date', pattern: currentLang === 'ja' ? 'yyyy"年"m"月"d"日' : 'mmmm d, yyyy' };
    case 'time':
      return { kind: 'time', pattern: currentLang === 'ja' ? 'H:MM' : 'h:MM AM/PM' };
    case 'percent':
      return { kind: 'percent', decimals: 0 };
    case 'fraction':
      return { kind: 'custom', pattern: '# ?/?' };
    case 'scientific':
      return { kind: 'scientific', decimals: 2 };
    case 'text':
      return { kind: 'text' };
    case 'more':
      return null;
  }
};

const THEME_COLORS = [
  '#ffffff',
  '#000000',
  '#e7e6e6',
  '#44546a',
  '#5b9bd5',
  '#ed7d31',
  '#70ad47',
  '#4472c4',
  '#a64d79',
  '#70ad47',
  '#f2f2f2',
  '#7f7f7f',
  '#d9e2f3',
  '#d9eaf7',
  '#fce4d6',
  '#e2f0d9',
  '#d9e2f3',
  '#eadcf8',
  '#e2f0d9',
  '#d9d9d9',
  '#595959',
  '#b4c6e7',
  '#bdd7ee',
  '#f8cbad',
  '#c6e0b4',
  '#b4c6e7',
  '#d9bce3',
  '#c6e0b4',
  '#bfbfbf',
  '#404040',
  '#8eaadb',
  '#9dc3e6',
  '#f4b183',
  '#a9d18e',
  '#8eaadb',
  '#c27ba0',
  '#a9d18e',
  '#a6a6a6',
  '#262626',
  '#2f5597',
  '#2e75b6',
  '#c65911',
  '#548235',
  '#2f5597',
  '#741b47',
  '#548235',
] as const;

const STANDARD_COLORS = [
  '#c00000',
  '#ff0000',
  '#ffc000',
  '#ffff00',
  '#92d050',
  '#00b050',
  '#00b0f0',
  '#0070c0',
  '#002060',
  '#7030a0',
] as const;

const props = defineProps<Props>();
const emit = defineEmits<{
  tabChange: [tab: RibbonTab];
}>();

const lang = computed(() => (props.locale === 'ja' ? 'ja' : 'en'));
const tabs = computed(() => toolbarTabs(lang.value));
const tr = computed(() => toolbarText(lang.value));
const tablistRef = ref<HTMLDivElement | null>(null);
const keyShortcuts = (id: string): string | undefined => RIBBON_KEYSHORTCUTS[id];
const borderPresets = computed(() =>
  BORDER_PRESETS.map((preset) => ({
    ...preset,
    label:
      preset.value === 'none'
        ? tr.value.noBorder
        : preset.value === 'outline'
          ? tr.value.outsideBorders
          : preset.value === 'all'
            ? tr.value.allBorders
            : preset.value === 'top'
              ? tr.value.topBorder
              : preset.value === 'bottom'
                ? tr.value.bottomBorder
                : preset.value === 'left'
                  ? tr.value.leftBorder
                  : preset.value === 'right'
                    ? tr.value.rightBorder
                    : tr.value.doubleBottomBorder,
  })),
);
const borderStyles = computed(() =>
  BORDER_STYLES.map((style) => ({
    ...style,
    label:
      style.value === 'thin'
        ? tr.value.thin
        : style.value === 'medium'
          ? tr.value.medium
          : style.value === 'thick'
            ? tr.value.thick
            : style.value === 'dashed'
              ? tr.value.dashed
              : style.value === 'dotted'
                ? tr.value.dotted
                : tr.value.double,
  })),
);

const active = useToolbarActive(() => props.instance);

const disabled = computed(() => !props.instance);
const setActiveTab = (tab: RibbonTab): void => {
  emit('tabChange', tab);
};

const focusRibbonTab = async (tab: RibbonTab): Promise<void> => {
  await nextTick();
  tablistRef.value
    ?.querySelector<HTMLButtonElement>(`[data-ribbon-tab="${tab}"]`)
    ?.focus({ preventScroll: true });
};

const onRibbonTabKeydown = (event: KeyboardEvent): void => {
  const target = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-ribbon-tab]');
  if (!target) return;
  const list = tabs.value;
  const currentId = (target.dataset.ribbonTab as RibbonTab | undefined) ?? props.activeTab;
  const current = Math.max(
    0,
    list.findIndex((tab) => tab.id === currentId),
  );
  let next = current;
  if (event.key === 'ArrowRight') next = (current + 1) % list.length;
  else if (event.key === 'ArrowLeft') next = (current - 1 + list.length) % list.length;
  else if (event.key === 'Home') next = 0;
  else if (event.key === 'End') next = list.length - 1;
  else return;
  event.preventDefault();
  const nextTab = list[next]?.id;
  if (!nextTab) return;
  setActiveTab(nextTab);
  void focusRibbonTab(nextTab);
};

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

const dispatchClipboard = (kind: 'copy' | 'cut' | 'paste'): void => {
  const inst = props.instance;
  if (!inst) return;
  inst.host.focus();
  try {
    document.execCommand(kind);
  } catch {
    // Browser clipboard command support is best-effort from toolbar buttons.
  }
};

const onFormatPainter = (): void => {
  props.instance?.formatPainter?.activate(false);
};

const onAutoSum = (): void => {
  const inst = props.instance;
  if (!inst) return;
  const result = autoSum(inst.store.getState(), inst.workbook);
  if (!result) return;
  mutators.replaceCells(inst.store, inst.workbook.cells(result.addr.sheet));
  mutators.setActive(inst.store, result.addr);
};

const onMergeAction = (action: MergeAction): void => {
  const inst = props.instance;
  if (!inst) return;
  const s = inst.store.getState();
  const r = s.selection.range;
  if (action === 'unmergeCells') {
    applyUnmerge(inst.store, inst.workbook, inst.history, r);
    return;
  }
  if (action === 'mergeAcross') {
    recordMergesChangeWithEngine(inst.history, inst.store, inst.workbook, r.sheet, () => {
      for (let row = r.r0; row <= r.r1; row += 1) {
        if (r.c0 === r.c1) continue;
        mutators.mergeRange(inst.store, { sheet: r.sheet, r0: row, c0: r.c0, r1: row, c1: r.c1 });
      }
    });
    return;
  }
  applyMerge(inst.store, inst.workbook, inst.history, r);
  if (action === 'mergeCenter') {
    recordFormatChange(inst.history, inst.store, () =>
      setAlign(inst.store.getState(), inst.store, 'center'),
    );
  }
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
const onFontSize = (value: string | number): void => {
  wrapFormat((s, st) => setFont(s, st, { fontSize: Number(value) }));
};
const onBorderPreset = (preset: BorderPreset): void => {
  wrapFormat((s, st) => setBorderPreset(s, st, preset, borderStyle.value));
};
const onPageOrientation = (next: PageOrientation): void => {
  const inst = props.instance;
  if (!inst) return;
  const sheet = inst.store.getState().data.sheetIndex;
  recordPageSetupChange(inst.history, inst.store, () => setPageOrientation(inst.store, sheet, next));
};
const onPaperSize = (next: PaperSize): void => {
  const inst = props.instance;
  if (!inst) return;
  const sheet = inst.store.getState().data.sheetIndex;
  recordPageSetupChange(inst.history, inst.store, () => setPaperSize(inst.store, sheet, next));
};
const onMarginPreset = (next: MarginPreset): void => {
  const inst = props.instance;
  if (!inst) return;
  const sheet = inst.store.getState().data.sheetIndex;
  recordPageSetupChange(inst.history, inst.store, () => setMarginPreset(inst.store, sheet, next));
};
const onNumberFormat = (next: string): void => {
  const inst = props.instance;
  if (!inst) return;
  const action = next as NumberFormatAction;
  if (action === 'more') {
    inst.openFormatDialog();
    return;
  }
  const fmt = numberFormatForAction(action, lang.value);
  if (!fmt) return;
  wrapFormat((s, st) => setNumFmt(s, st, fmt));
};

const { borderStyle, closeDropdown, onDropdownKeydown, onDropdownPick, openDropdown, toggleDropdown } =
  useToolbarDropdown({
    onBorderPreset,
    onFontFamily,
    onFontSize,
    onMarginPreset,
    onNumberFormat,
    onOpenPageSetup: () => props.instance?.openPageSetup(),
    onPageOrientation,
    onPaperSize,
  });
const onFontColor = (value: string): void => {
  wrapFormat((s, st) => setFontColor(s, st, value));
};
const onFillColor = (value: string): void => {
  wrapFormat((s, st) => setFillColor(s, st, value));
};
const onPaletteColor = (kind: 'fontColor' | 'fillColor', value: string): void => {
  if (kind === 'fontColor') onFontColor(value);
  else onFillColor(value);
  closeDropdown();
};

const onFormatAsTable = (): void => {
  const inst = props.instance;
  if (!inst) return;
  const r = inst.store.getState().selection.range;
  formatAsTable(inst.store, r);
};

const onInsertRows = (): void => {
  const inst = props.instance;
  if (!inst) return;
  const r = inst.store.getState().selection.range;
  insertRows(inst.store, inst.workbook, inst.history, r.r0, r.r1 - r.r0 + 1);
};

const onDeleteRows = (): void => {
  const inst = props.instance;
  if (!inst) return;
  const r = inst.store.getState().selection.range;
  deleteRows(inst.store, inst.workbook, inst.history, r.r0, r.r1 - r.r0 + 1);
};

const onInsertCols = (): void => {
  const inst = props.instance;
  if (!inst) return;
  const r = inst.store.getState().selection.range;
  insertCols(inst.store, inst.workbook, inst.history, r.c0, r.c1 - r.c0 + 1);
};

const onDeleteCols = (): void => {
  const inst = props.instance;
  if (!inst) return;
  const r = inst.store.getState().selection.range;
  deleteCols(inst.store, inst.workbook, inst.history, r.c0, r.c1 - r.c0 + 1);
};

const onToggleRowsHidden = (): void => {
  const inst = props.instance;
  if (!inst) return;
  const s = inst.store.getState();
  const r = s.selection.range;
  if (hiddenInSelection(s.layout, 'row', r.r0, r.r1).length > 0) {
    showRows(inst.store, inst.history, r.r0, r.r1, inst.workbook);
  } else {
    hideRows(inst.store, inst.history, r.r0, r.r1, inst.workbook);
  }
};

const onToggleColsHidden = (): void => {
  const inst = props.instance;
  if (!inst) return;
  const s = inst.store.getState();
  const r = s.selection.range;
  if (hiddenInSelection(s.layout, 'col', r.c0, r.c1).length > 0) {
    showCols(inst.store, inst.history, r.c0, r.c1, inst.workbook);
  } else {
    hideCols(inst.store, inst.history, r.c0, r.c1, inst.workbook);
  }
};

const onFilterToggle = (): void => {
  const inst = props.instance;
  if (!inst) return;
  const s = inst.store.getState();
  if (s.ui.filterRange) clearFilter(s, inst.store, s.ui.filterRange);
  else setAutoFilter(inst.store, s.selection.range);
};

const onSort = (direction: 'asc' | 'desc'): void => {
  const inst = props.instance;
  if (!inst) return;
  const s = inst.store.getState();
  const ok = sortRange(s, inst.store, inst.workbook, s.selection.range, {
    byCol: s.selection.active.col,
    direction,
    hasHeader: s.selection.range.r0 < s.selection.range.r1,
  });
  if (ok) mutators.replaceCells(inst.store, inst.workbook.cells(s.data.sheetIndex));
};

const onRemoveDuplicates = (): void => {
  const inst = props.instance;
  if (!inst) return;
  const s = inst.store.getState();
  const removed = removeDuplicates(s, inst.store, inst.workbook, s.selection.range);
  if (removed > 0) mutators.replaceCells(inst.store, inst.workbook.cells(s.data.sheetIndex));
};

const onZoom = (zoom: number): void => {
  const inst = props.instance;
  if (!inst) return;
  setSheetZoom(inst.store, zoom, inst.workbook);
};
</script>

<template>
  <div class="demo__ribbon-shell" @keydown="onDropdownKeydown">
    <div
      ref="tablistRef"
      class="demo__ribbon-tabs"
      role="tablist"
      :aria-label="tr.ribbonTabs"
      @keydown="onRibbonTabKeydown"
    >
      <button
        v-for="tab in tabs"
        :key="tab.id"
        :class="[
          'demo__ribbon-tab',
          {
            'demo__ribbon-tab--active': props.activeTab === tab.id,
          },
        ]"
        type="button"
        role="tab"
        :data-ribbon-tab="tab.id"
        :aria-selected="props.activeTab === tab.id"
        :tabindex="props.activeTab === tab.id ? 0 : -1"
        @click="setActiveTab(tab.id)"
      >
        {{ tab.label }}
      </button>
    </div>
    <div class="demo__ribbon" role="toolbar" :aria-label="`${RIBBON_TAB_LABELS[props.activeTab][lang]} ${tr.ribbon}`">
      <template v-if="props.activeTab === 'file'">
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.workbook">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" @click="props.instance?.openPageSetup()">
              <RibbonIcon name="page" /><span>{{ tr.pageSetup }}</span>
            </button>
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" @click="props.instance?.print()">
              <RibbonIcon name="print" /><span>{{ tr.print }}</span>
            </button>
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" @click="props.instance?.openExternalLinksDialog()">
              <RibbonIcon name="link" /><span>{{ tr.links }}</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ tr.workbook }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.inspect">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" :aria-keyshortcuts="keyShortcuts('formatCells')" @click="props.instance?.openFormatDialog()">
              <RibbonIcon name="formatCells" /><span>{{ tr.formatCells }}</span>
            </button>
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" :aria-keyshortcuts="keyShortcuts('gotoSpecial')" @click="props.instance?.openGoToSpecial()">
              <RibbonIcon name="goTo" /><span>{{ tr.goTo }}</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ tr.inspect }}</div>
        </section>
      </template>

      <template v-else-if="props.activeTab === 'home'">
      <section class="demo__ribbon-group demo__ribbon-group--clipboard" :aria-label="tr.clipboard">
        <div class="demo__ribbon-tools">
    <button class="demo__rb demo__rb--large" type="button" :disabled="disabled" :title="tr.paste" :aria-label="tr.paste" :aria-keyshortcuts="keyShortcuts('paste')" @click="dispatchClipboard('paste')">
      <RibbonIcon name="paste" />
      <span>{{ tr.paste }}</span>
    </button>
    <button class="demo__rb" type="button" :disabled="disabled" :title="tr.cut" :aria-label="tr.cut" :aria-keyshortcuts="keyShortcuts('cut')" @click="dispatchClipboard('cut')">
      <RibbonIcon name="cut" />
    </button>
    <button class="demo__rb" type="button" :disabled="disabled" :title="tr.copy" :aria-label="tr.copy" :aria-keyshortcuts="keyShortcuts('copy')" @click="dispatchClipboard('copy')">
      <RibbonIcon name="copy" />
    </button>
    <button class="demo__rb" :class="{ 'demo__rb--active': active.formatPainterArmed }" type="button" :disabled="disabled" :title="tr.formatPainter" :aria-label="tr.formatPainter" @click="onFormatPainter">
      <RibbonIcon name="paint" />
    </button>
    <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" :title="tr.clearFormats" :aria-label="tr.clearFormats" @click="wrapFormat(clearFormat)">
      <RibbonIcon name="clear" />
    </button>
        </div>
        <div class="demo__ribbon-label">{{ tr.clipboard }}</div>
      </section>

      <section class="demo__ribbon-group demo__ribbon-group--font" :aria-label="tr.font">
        <div class="demo__ribbon-tools">
    <div
      class="demo__rb-dd demo__rb-select--font"
      data-dropdown-name="fontFamily"
      :class="{ 'demo__rb-dd--open': openDropdown === 'fontFamily' }"
    >
      <button
        type="button"
        class="demo__rb-dd__btn"
        :disabled="disabled"
        :title="tr.font"
        :aria-label="tr.font"
        aria-haspopup="listbox"
        :aria-expanded="openDropdown === 'fontFamily'"
        @click="toggleDropdown('fontFamily')"
      >
        <span class="demo__rb-dd__value">{{ active.fontFamily }}</span>
        <svg class="demo__rb-dd__chev" viewBox="0 0 12 12" aria-hidden="true">
          <path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" />
        </svg>
      </button>
      <div v-if="openDropdown === 'fontFamily'" class="demo__rb-dd__list" role="listbox" :aria-label="tr.font" tabindex="-1">
        <button
          v-for="font in FONT_FAMILIES"
          :key="font"
          type="button"
          role="option"
          :aria-selected="active.fontFamily === font"
          class="demo__rb-dd__opt"
          :class="{ 'demo__rb-dd__opt--selected': active.fontFamily === font }"
          @click="onDropdownPick('fontFamily', font)"
        >
          <span class="demo__rb-dd__check" aria-hidden="true">
            <svg v-if="active.fontFamily === font" viewBox="0 0 16 16">
              <path d="M3.5 8.5l3 3 6-6.5" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" />
            </svg>
          </span>
          <span class="demo__rb-dd__label">{{ font }}</span>
        </button>
      </div>
    </div>
    <div
      class="demo__rb-dd"
      data-dropdown-name="fontSize"
      :class="{ 'demo__rb-dd--open': openDropdown === 'fontSize' }"
    >
      <button
        type="button"
        class="demo__rb-dd__btn"
        :disabled="disabled"
        :title="tr.fontSize"
        :aria-label="tr.fontSize"
        aria-haspopup="listbox"
        :aria-expanded="openDropdown === 'fontSize'"
        @click="toggleDropdown('fontSize')"
      >
        <span class="demo__rb-dd__value">{{ active.fontSize }}</span>
        <svg class="demo__rb-dd__chev" viewBox="0 0 12 12" aria-hidden="true">
          <path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" />
        </svg>
      </button>
      <div v-if="openDropdown === 'fontSize'" class="demo__rb-dd__list" role="listbox" :aria-label="tr.fontSize" tabindex="-1">
        <button
          v-for="size in FONT_SIZES"
          :key="size"
          type="button"
          role="option"
          :aria-selected="active.fontSize === size"
          class="demo__rb-dd__opt"
          :class="{ 'demo__rb-dd__opt--selected': active.fontSize === size }"
          @click="onDropdownPick('fontSize', size)"
        >
          <span class="demo__rb-dd__check" aria-hidden="true">
            <svg v-if="active.fontSize === size" viewBox="0 0 16 16">
              <path d="M3.5 8.5l3 3 6-6.5" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" />
            </svg>
          </span>
          <span class="demo__rb-dd__label">{{ size }}</span>
        </button>
      </div>
    </div>
    <button class="demo__rb" type="button" :disabled="disabled" :title="tr.increaseFontSize" :aria-label="tr.increaseFontSize" @click="wrapFormat((s, st) => setFont(s, st, { fontSize: active.fontSize + 1 }))">
      <RibbonIcon name="fontGrow" />
    </button>
    <button class="demo__rb" type="button" :disabled="disabled" :title="tr.decreaseFontSize" :aria-label="tr.decreaseFontSize" @click="wrapFormat((s, st) => setFont(s, st, { fontSize: Math.max(1, active.fontSize - 1) }))">
      <RibbonIcon name="fontShrink" />
    </button>
    <span class="demo__rb-break" aria-hidden="true" />
    <button class="demo__rb" :class="{ 'demo__rb--active': active.bold }" type="button" :disabled="disabled" :title="`${tr.bold} (⌘B)`" :aria-label="`${tr.bold} (⌘B)`" @click="wrapFormat(toggleBold)">
      <RibbonIcon name="bold" />
    </button>
    <button class="demo__rb" :class="{ 'demo__rb--active': active.italic }" type="button" :disabled="disabled" :title="`${tr.italic} (⌘I)`" :aria-label="`${tr.italic} (⌘I)`" @click="wrapFormat(toggleItalic)">
      <RibbonIcon name="italic" />
    </button>
    <button class="demo__rb" :class="{ 'demo__rb--active': active.underline }" type="button" :disabled="disabled" :title="`${tr.underline} (⌘U)`" :aria-label="`${tr.underline} (⌘U)`" @click="wrapFormat(toggleUnderline)">
      <RibbonIcon name="underline" />
    </button>
    <button class="demo__rb" :class="{ 'demo__rb--active': active.strike }" type="button" :disabled="disabled" :title="tr.strikethrough" :aria-label="tr.strikethrough" @click="wrapFormat(toggleStrike)">
      <RibbonIcon name="strike" />
    </button>
    <button class="demo__rb" type="button" :disabled="disabled" :title="tr.borders" :aria-label="tr.borders" @click="wrapFormat(cycleBorders)">
      <RibbonIcon name="borders" />
    </button>
    <div
      class="demo__rb-dd demo__rb-select--border"
      data-dropdown-name="borderPreset"
      :class="{ 'demo__rb-dd--open': openDropdown === 'borderPreset' }"
    >
      <button
        type="button"
        class="demo__rb-dd__btn"
        :disabled="disabled"
        :title="tr.borderPattern"
        :aria-label="tr.borderPattern"
        aria-haspopup="listbox"
        :aria-expanded="openDropdown === 'borderPreset'"
        @click="toggleDropdown('borderPreset')"
      >
        <span class="demo__rb-dd__value">{{ tr.outsideBorders }}</span>
        <svg class="demo__rb-dd__chev" viewBox="0 0 12 12" aria-hidden="true">
          <path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" />
        </svg>
      </button>
      <div v-if="openDropdown === 'borderPreset'" class="demo__rb-dd__list" role="listbox" :aria-label="tr.borderPattern" tabindex="-1">
        <button
          v-for="preset in borderPresets"
          :key="preset.value"
          type="button"
          role="option"
          :aria-selected="false"
          class="demo__rb-dd__opt"
          @click="onDropdownPick('borderPreset', preset.value)"
        >
          <span class="demo__rb-dd__check" aria-hidden="true" />
          <span class="demo__rb-dd__label">{{ preset.label }}</span>
        </button>
      </div>
    </div>
    <div
      class="demo__rb-dd demo__rb-select--border-style"
      data-dropdown-name="borderStyle"
      :class="{ 'demo__rb-dd--open': openDropdown === 'borderStyle' }"
    >
      <button
        type="button"
        class="demo__rb-dd__btn"
        :disabled="disabled"
        :title="tr.borderLineStyle"
        :aria-label="tr.borderLineStyle"
        aria-haspopup="listbox"
        :aria-expanded="openDropdown === 'borderStyle'"
        @click="toggleDropdown('borderStyle')"
      >
        <span class="demo__rb-dd__value">{{ borderStyles.find((style) => style.value === borderStyle)?.label ?? borderStyle }}</span>
        <svg class="demo__rb-dd__chev" viewBox="0 0 12 12" aria-hidden="true">
          <path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" />
        </svg>
      </button>
      <div v-if="openDropdown === 'borderStyle'" class="demo__rb-dd__list" role="listbox" :aria-label="tr.borderLineStyle" tabindex="-1">
        <button
          v-for="style in borderStyles"
          :key="style.value"
          type="button"
          role="option"
          :aria-selected="borderStyle === style.value"
          :class="['demo__rb-dd__opt', { 'demo__rb-dd__opt--selected': borderStyle === style.value }]"
          @click="onDropdownPick('borderStyle', style.value)"
        >
          <span class="demo__rb-dd__check" aria-hidden="true" />
          <span class="demo__rb-dd__label">{{ style.label }}</span>
        </button>
      </div>
    </div>
    <div class="demo__rb-color" :class="{ 'demo__rb-color--open': openDropdown === 'fontColor' }" data-dropdown-name="fontColor" :title="tr.fontColor" :aria-label="tr.fontColor">
      <button type="button" class="demo__rb-color__btn" :disabled="disabled" :aria-label="tr.fontColor" aria-haspopup="menu" :aria-expanded="openDropdown === 'fontColor'" @click="toggleDropdown('fontColor')">
        <span class="demo__rb-color__icon"><RibbonIcon name="fontColor" /></span><span class="demo__rb-color__swatch" :style="{ backgroundColor: active.fontColor }" /><svg class="demo__rb-color__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
      </button>
      <div v-if="openDropdown === 'fontColor'" class="demo__color-menu" role="menu" :aria-label="tr.fontColor">
        <label class="demo__color-menu__check"><input type="checkbox" disabled /><span>{{ tr.highContrastOnly }}</span></label>
        <button class="demo__color-menu__auto" type="button" role="menuitem" @click="onPaletteColor('fontColor', '#000000')">{{ tr.automatic }}</button>
        <div class="demo__color-menu__section">{{ tr.themeColors }}</div>
        <div class="demo__color-menu__grid demo__color-menu__grid--theme"><button v-for="(color, index) in THEME_COLORS" :key="`${color}-${index}`" type="button" class="demo__color-menu__chip" :style="{ backgroundColor: color }" :aria-label="color" @click="onPaletteColor('fontColor', color)" /></div>
        <div class="demo__color-menu__section">{{ tr.standardColors }}</div>
        <div class="demo__color-menu__grid demo__color-menu__grid--standard"><button v-for="color in STANDARD_COLORS" :key="color" type="button" class="demo__color-menu__chip" :style="{ backgroundColor: color }" :aria-label="color" @click="onPaletteColor('fontColor', color)" /></div>
        <label class="demo__color-menu__more"><span class="demo__color-menu__wheel" aria-hidden="true" />{{ tr.moreColors }}<input class="demo__color-menu__native" type="color" :value="active.fontColor" @change="onPaletteColor('fontColor', ($event.target as HTMLInputElement).value)" /></label>
      </div>
    </div>
    <div class="demo__rb-color" :class="{ 'demo__rb-color--open': openDropdown === 'fillColor' }" data-dropdown-name="fillColor" :title="tr.fillColor" :aria-label="tr.fillColor">
      <button type="button" class="demo__rb-color__btn" :disabled="disabled" :aria-label="tr.fillColor" aria-haspopup="menu" :aria-expanded="openDropdown === 'fillColor'" @click="toggleDropdown('fillColor')">
        <span class="demo__rb-color__icon"><RibbonIcon name="fillColor" /></span><span class="demo__rb-color__swatch" :style="{ backgroundColor: active.fillColor }" /><svg class="demo__rb-color__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
      </button>
      <div v-if="openDropdown === 'fillColor'" class="demo__color-menu" role="menu" :aria-label="tr.fillColor">
        <label class="demo__color-menu__check"><input type="checkbox" disabled /><span>{{ tr.highContrastOnly }}</span></label>
        <button class="demo__color-menu__auto" type="button" role="menuitem" @click="onPaletteColor('fillColor', '#ffffff')">{{ tr.automatic }}</button>
        <div class="demo__color-menu__section">{{ tr.themeColors }}</div>
        <div class="demo__color-menu__grid demo__color-menu__grid--theme"><button v-for="(color, index) in THEME_COLORS" :key="`${color}-${index}`" type="button" class="demo__color-menu__chip" :style="{ backgroundColor: color }" :aria-label="color" @click="onPaletteColor('fillColor', color)" /></div>
        <div class="demo__color-menu__section">{{ tr.standardColors }}</div>
        <div class="demo__color-menu__grid demo__color-menu__grid--standard"><button v-for="color in STANDARD_COLORS" :key="color" type="button" class="demo__color-menu__chip" :style="{ backgroundColor: color }" :aria-label="color" @click="onPaletteColor('fillColor', color)" /></div>
        <label class="demo__color-menu__more"><span class="demo__color-menu__wheel" aria-hidden="true" />{{ tr.moreColors }}<input class="demo__color-menu__native" type="color" :value="active.fillColor" @change="onPaletteColor('fillColor', ($event.target as HTMLInputElement).value)" /></label>
      </div>
    </div>
        </div>
        <div class="demo__ribbon-label">{{ tr.font }}</div>
      </section>

      <section class="demo__ribbon-group demo__ribbon-group--alignment" :aria-label="tr.alignment">
        <div class="demo__ribbon-tools">
    <button class="demo__rb" type="button" :disabled="disabled" :title="tr.topAlign" :aria-label="tr.topAlign" @click="wrapFormat((s, st) => setVAlign(s, st, 'top'))">
      <RibbonIcon name="top" />
    </button>
    <button class="demo__rb" type="button" :disabled="disabled" :title="tr.middleAlign" :aria-label="tr.middleAlign" @click="wrapFormat((s, st) => setVAlign(s, st, 'middle'))">
      <RibbonIcon name="middle" />
    </button>
    <span class="demo__rb-break" aria-hidden="true" />
    <button class="demo__rb" :class="{ 'demo__rb--active': active.alignLeft }" type="button" :disabled="disabled" :title="tr.alignLeft" :aria-label="tr.alignLeft" @click="onAlign('left')">
      <RibbonIcon name="alignLeft" />
    </button>
    <button class="demo__rb" :class="{ 'demo__rb--active': active.alignCenter }" type="button" :disabled="disabled" :title="tr.alignCenter" :aria-label="tr.alignCenter" @click="onAlign('center')">
      <RibbonIcon name="alignCenter" />
    </button>
    <button class="demo__rb" :class="{ 'demo__rb--active': active.alignRight }" type="button" :disabled="disabled" :title="tr.alignRight" :aria-label="tr.alignRight" @click="onAlign('right')">
      <RibbonIcon name="alignRight" />
    </button>
    <div class="demo__rb-menu" :class="{ 'demo__rb-menu--open': openDropdown === 'merge' }" data-dropdown-name="merge">
      <button class="demo__rb demo__rb-menu__btn" type="button" :disabled="disabled" :title="tr.mergeCells" :aria-label="tr.mergeCells" aria-haspopup="menu" :aria-expanded="openDropdown === 'merge'" @click="toggleDropdown('merge')">
        <RibbonIcon name="merge" /><svg class="demo__rb-menu__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
      </button>
      <div v-if="openDropdown === 'merge'" class="demo__merge-menu" role="menu" :aria-label="tr.mergeCells">
        <button class="demo__merge-menu__item" type="button" role="menuitem" @click="onMergeAction('mergeCenter'); closeDropdown()"><RibbonIcon name="merge" /><span>{{ tr.mergeAndCenter }}</span></button>
        <button class="demo__merge-menu__item" type="button" role="menuitem" @click="onMergeAction('mergeAcross'); closeDropdown()"><RibbonIcon name="merge" /><span>{{ tr.mergeAcross }}</span></button>
        <button class="demo__merge-menu__item" type="button" role="menuitem" @click="onMergeAction('mergeCells'); closeDropdown()"><RibbonIcon name="merge" /><span>{{ tr.mergeCells }}</span></button>
        <button class="demo__merge-menu__item" type="button" role="menuitem" @click="onMergeAction('unmergeCells'); closeDropdown()"><RibbonIcon name="merge" /><span>{{ tr.unmergeCells }}</span></button>
      </div>
    </div>
    <button class="demo__rb" type="button" :disabled="disabled" :title="tr.wrapText" :aria-label="tr.wrapText" @click="wrapFormat(toggleWrap)">
      <RibbonIcon name="wrap" />
    </button>
        </div>
        <div class="demo__ribbon-label">{{ tr.alignment }}</div>
      </section>

      <section class="demo__ribbon-group demo__ribbon-group--number" :aria-label="tr.number">
        <div class="demo__ribbon-tools">
    <div class="demo__rb-dd demo__rb-select--number-format" data-dropdown-name="numberFormat" :class="{ 'demo__rb-dd--open': openDropdown === 'numberFormat' }">
      <button type="button" class="demo__rb-dd__btn" :disabled="disabled" title="Number format" aria-label="Number format" aria-haspopup="listbox" :aria-expanded="openDropdown === 'numberFormat'" @click="toggleDropdown('numberFormat')">
        <span class="demo__rb-dd__value">{{ [
          { value: 'general', label: tr.general },
          { value: 'fixed', label: tr.fixedNumber },
          { value: 'currency', label: tr.currency },
          { value: 'accounting', label: tr.accounting },
          { value: 'shortDate', label: tr.shortDate },
          { value: 'longDate', label: tr.longDate },
          { value: 'time', label: tr.timeFormat },
          { value: 'percent', label: tr.percent },
          { value: 'fraction', label: tr.fraction },
          { value: 'scientific', label: tr.scientific },
          { value: 'text', label: tr.textFormat },
        ].find((option) => option.value === active.numberFormat)?.label ?? tr.general }}</span>
        <svg class="demo__rb-dd__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
      </button>
      <div v-if="openDropdown === 'numberFormat'" class="demo__rb-dd__list" role="listbox" aria-label="Number format" tabindex="-1">
        <button v-for="option in [
          { value: 'general', label: tr.general },
          { value: 'fixed', label: tr.fixedNumber },
          { value: 'currency', label: tr.currency },
          { value: 'accounting', label: tr.accounting },
          { value: 'shortDate', label: tr.shortDate },
          { value: 'longDate', label: tr.longDate },
          { value: 'time', label: tr.timeFormat },
          { value: 'percent', label: tr.percent },
          { value: 'fraction', label: tr.fraction },
          { value: 'scientific', label: tr.scientific },
          { value: 'text', label: tr.textFormat },
          { value: 'more', label: tr.moreNumberFormats },
        ]" :key="option.value" class="demo__rb-dd__opt" :data-fc-value="option.value" type="button" role="option" :aria-selected="active.numberFormat === option.value" @click="onDropdownPick('numberFormat', option.value)">
          <span class="demo__rb-dd__check" aria-hidden="true" />
          <span class="demo__rb-dd__label">{{ option.label }}</span>
        </button>
      </div>
    </div>
    <span class="demo__rb-break" aria-hidden="true" />
    <button class="demo__rb" :class="{ 'demo__rb--active': active.currency }" type="button" :disabled="disabled" :title="tr.currency" :aria-label="tr.currency" @click="wrapFormat(cycleCurrency)">
      <RibbonIcon name="currency" />
    </button>
    <button class="demo__rb" :class="{ 'demo__rb--active': active.percent }" type="button" :disabled="disabled" :title="tr.percent" :aria-label="tr.percent" @click="wrapFormat(cyclePercent)">
      <RibbonIcon name="percent" />
    </button>
    <button class="demo__rb" type="button" :disabled="disabled" :title="tr.commaStyle" :aria-label="tr.commaStyle" @click="wrapFormat((s, st) => setNumFmt(s, st, { kind: 'fixed', decimals: 2 }))">
      <RibbonIcon name="comma" />
    </button>
    <button class="demo__rb" type="button" :disabled="disabled" :title="tr.decreaseDecimals" :aria-label="tr.decreaseDecimals" @click="onBumpDecimals(-1)">
      <RibbonIcon name="decDown" />
    </button>
    <button class="demo__rb" type="button" :disabled="disabled" :title="tr.increaseDecimals" :aria-label="tr.increaseDecimals" @click="onBumpDecimals(1)">
      <RibbonIcon name="decUp" />
    </button>
        </div>
        <div class="demo__ribbon-label">{{ tr.number }}</div>
      </section>

      <section class="demo__ribbon-group demo__ribbon-group--styles" :aria-label="tr.styles">
        <div class="demo__ribbon-tools">
          <button
            class="demo__rb demo__rb--wide"
            type="button"
            :disabled="disabled"
            :title="tr.conditionalFormatting"
            :aria-label="tr.conditionalFormatting"
            @click="props.instance?.openConditionalDialog()"
          >
            <RibbonIcon name="conditional" /><span>{{ tr.conditional }}</span>
          </button>
          <button
            class="demo__rb demo__rb--wide"
            type="button"
            :disabled="disabled"
            :title="tr.cellStyles"
            :aria-label="tr.cellStyles"
            @click="props.instance?.openCellStylesGallery()"
          >
            <RibbonIcon name="tableStyle" /><span>{{ tr.cellStyles }}</span>
          </button>
          <button
            class="demo__rb demo__rb--wide"
            type="button"
            :disabled="disabled"
            :title="tr.manageRules"
            :aria-label="tr.manageRules"
            @click="props.instance?.openCfRulesDialog()"
          >
            <RibbonIcon name="options" /><span>{{ tr.rules }}</span>
          </button>
        </div>
        <div class="demo__ribbon-label">{{ tr.styles }}</div>
      </section>

      <section class="demo__ribbon-group demo__ribbon-group--cells" :aria-label="tr.cells">
        <div class="demo__ribbon-tools">
    <button class="demo__rb" type="button" :disabled="disabled" :title="tr.insertRows" :aria-label="tr.insertRows" @click="onInsertRows">
      <RibbonIcon name="insertRows" />
    </button>
    <button class="demo__rb" type="button" :disabled="disabled" :title="tr.deleteRows" :aria-label="tr.deleteRows" @click="onDeleteRows">
      <RibbonIcon name="deleteRows" />
    </button>
    <button class="demo__rb" type="button" :disabled="disabled" :title="tr.insertCols" :aria-label="tr.insertCols" @click="onInsertCols">
      <RibbonIcon name="insertCols" />
    </button>
    <button class="demo__rb" type="button" :disabled="disabled" :title="tr.deleteCols" :aria-label="tr.deleteCols" @click="onDeleteCols">
      <RibbonIcon name="deleteCols" />
    </button>
    <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" :title="tr.formatCells" :aria-label="tr.formatCells" :aria-keyshortcuts="keyShortcuts('formatCellsHome')" @click="props.instance?.openFormatDialog()">
      <RibbonIcon name="formatCells" /><span>{{ tr.formatCells }}</span>
    </button>
        </div>
        <div class="demo__ribbon-label">{{ tr.cells }}</div>
      </section>

      <section class="demo__ribbon-group demo__ribbon-group--editing" :aria-label="tr.editing">
        <div class="demo__ribbon-tools">
    <button class="demo__rb" type="button" :disabled="disabled" :title="`${tr.autoSum} (Σ)`" :aria-label="`${tr.autoSum} (Σ)`" @click="onAutoSum">
      <RibbonIcon name="autosum" />
    </button>
    <button class="demo__rb" type="button" :disabled="disabled" :title="`${tr.undo} (⌘Z)`" :aria-label="`${tr.undo} (⌘Z)`" :aria-keyshortcuts="keyShortcuts('undoHome')" @click="onUndo">
      <RibbonIcon name="undo" />
    </button>
    <button class="demo__rb" type="button" :disabled="disabled" :title="`${tr.redo} (⌘⇧Z)`" :aria-label="`${tr.redo} (⌘⇧Z)`" :aria-keyshortcuts="keyShortcuts('redoHome')" @click="onRedo">
      <RibbonIcon name="redo" />
    </button>
    <button class="demo__rb" type="button" :disabled="disabled" :title="tr.sortAscending" :aria-label="tr.sortAscending" @click="onSort('asc')">
      <RibbonIcon name="sortAsc" />
    </button>
    <button class="demo__rb" :class="{ 'demo__rb--active': active.filterOn }" type="button" :disabled="disabled" :title="tr.filter" :aria-label="tr.filter" @click="onFilterToggle">
      <RibbonIcon name="filter" />
    </button>
    <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" :title="`${tr.find} (⌘F)`" :aria-label="`${tr.find} (⌘F)`" :aria-keyshortcuts="keyShortcuts('findHome')" @click="props.instance?.openFindReplace()">
      <RibbonIcon name="find" /><span>{{ tr.find }}</span>
    </button>
    <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" :title="tr.gotoSpecial" :aria-label="tr.gotoSpecial" :aria-keyshortcuts="keyShortcuts('gotoSpecialHome')" @click="props.instance?.openGoToSpecial()">
      <RibbonIcon name="goTo" /><span>{{ tr.gotoSpecial }}</span>
    </button>
        </div>
        <div class="demo__ribbon-label">{{ tr.editing }}</div>
      </section>
      </template>

      <template v-else-if="props.activeTab === 'insert'">
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.tables">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" @click="props.instance?.openPivotTableDialog()">
              <RibbonIcon name="table" /><span>{{ tr.pivotTable }}</span>
            </button>
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" @click="onFormatAsTable">
              <RibbonIcon name="tableStyle" /><span>{{ tr.formatTable }}</span>
            </button>
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" :aria-keyshortcuts="keyShortcuts('namedRangesInsert')" @click="props.instance?.openNamedRangeDialog()">
              <RibbonIcon name="names" /><span>{{ tr.names }}</span>
            </button>
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" @click="onRemoveDuplicates">
              <RibbonIcon name="removeDuplicates" /><span>{{ tr.removeDuplicates }}</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ tr.tables }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.charts">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" @click="props.instance?.openQuickAnalysis()">
              <RibbonIcon name="chart" /><span>{{ tr.chart }}</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ tr.charts }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.links">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" :aria-keyshortcuts="keyShortcuts('hyperlinkInsert')" @click="props.instance?.openHyperlinkDialog()">
              <RibbonIcon name="link" /><span>{{ tr.hyperlink }}</span>
            </button>
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" @click="props.instance?.openExternalLinksDialog()">
              <RibbonIcon name="link" /><span>{{ tr.links }}</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ tr.links }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.comments">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" :class="{ 'demo__rb--active': active.hasComment }" type="button" :disabled="disabled" @click="props.instance?.openCommentDialog()">
              <RibbonIcon name="comment" /><span>{{ active.hasComment ? tr.editComment : tr.newComment }}</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ tr.comments }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.symbols">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" :aria-keyshortcuts="keyShortcuts('fxInsert')" @click="props.instance?.openFunctionArguments()">
              <RibbonIcon name="function" /><span>fx</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ tr.symbols }}</div>
        </section>
      </template>

      <template v-else-if="props.activeTab === 'draw'">
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="RIBBON_TAB_LABELS.draw[lang]">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" type="button" :disabled="!props.onDrawPen" @click="props.onDrawPen?.()">
              <RibbonIcon name="pen" /><span>{{ tr.pen }}</span>
            </button>
            <button class="demo__rb demo__rb--wide" type="button" :disabled="!props.onDrawEraser" @click="props.onDrawEraser?.()">
              <RibbonIcon name="eraser" /><span>{{ tr.eraser }}</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ RIBBON_TAB_LABELS.draw[lang] }}</div>
        </section>
      </template>

      <template v-else-if="props.activeTab === 'pageLayout'">
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.pageSetup">
          <div class="demo__ribbon-tools">
            <div class="demo__rb-dd demo__rb-select--border" data-dropdown-name="margins" :class="{ 'demo__rb-dd--open': openDropdown === 'margins' }">
              <button type="button" class="demo__rb-dd__btn" :disabled="disabled" :title="tr.margins" :aria-label="tr.margins" aria-haspopup="listbox" :aria-expanded="openDropdown === 'margins'" @click="toggleDropdown('margins')"><span class="demo__rb-dd__value">{{ active.marginPreset === 'wide' ? tr.marginsWide : active.marginPreset === 'narrow' ? tr.marginsNarrow : active.marginPreset === 'normal' ? tr.marginsNormal : tr.marginsCustom }}</span><svg class="demo__rb-dd__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg></button>
              <div v-if="openDropdown === 'margins'" class="demo__rb-dd__list" role="listbox" :aria-label="tr.margins" tabindex="-1">
                <button class="demo__rb-dd__opt" type="button" role="option" :aria-selected="active.marginPreset === 'normal'" @click="onDropdownPick('margins', 'normal')"><span class="demo__rb-dd__check" aria-hidden="true" /><span class="demo__rb-dd__label">{{ tr.marginsNormal }}</span></button>
                <button class="demo__rb-dd__opt" type="button" role="option" :aria-selected="active.marginPreset === 'wide'" @click="onDropdownPick('margins', 'wide')"><span class="demo__rb-dd__check" aria-hidden="true" /><span class="demo__rb-dd__label">{{ tr.marginsWide }}</span></button>
                <button class="demo__rb-dd__opt" type="button" role="option" :aria-selected="active.marginPreset === 'narrow'" @click="onDropdownPick('margins', 'narrow')"><span class="demo__rb-dd__check" aria-hidden="true" /><span class="demo__rb-dd__label">{{ tr.marginsNarrow }}</span></button>
                <button class="demo__rb-dd__opt" type="button" role="option" :aria-selected="false" @click="onDropdownPick('margins', 'custom')"><span class="demo__rb-dd__check" aria-hidden="true" /><span class="demo__rb-dd__label">{{ tr.marginsCustom }}</span></button>
              </div>
            </div>
            <div class="demo__rb-dd demo__rb-select--border" data-dropdown-name="orientation" :class="{ 'demo__rb-dd--open': openDropdown === 'orientation' }">
              <button type="button" class="demo__rb-dd__btn" :disabled="disabled" :title="tr.orientation" :aria-label="tr.orientation" aria-haspopup="listbox" :aria-expanded="openDropdown === 'orientation'" @click="toggleDropdown('orientation')"><span class="demo__rb-dd__value">{{ active.pageOrientation === 'landscape' ? tr.landscape : tr.portrait }}</span><svg class="demo__rb-dd__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg></button>
              <div v-if="openDropdown === 'orientation'" class="demo__rb-dd__list" role="listbox" :aria-label="tr.orientation" tabindex="-1">
                <button class="demo__rb-dd__opt" type="button" role="option" :aria-selected="active.pageOrientation === 'portrait'" @click="onDropdownPick('orientation', 'portrait')"><span class="demo__rb-dd__check" aria-hidden="true" /><span class="demo__rb-dd__label">{{ tr.portrait }}</span></button>
                <button class="demo__rb-dd__opt" type="button" role="option" :aria-selected="active.pageOrientation === 'landscape'" @click="onDropdownPick('orientation', 'landscape')"><span class="demo__rb-dd__check" aria-hidden="true" /><span class="demo__rb-dd__label">{{ tr.landscape }}</span></button>
              </div>
            </div>
            <div class="demo__rb-dd demo__rb-select--border" data-dropdown-name="paperSize" :class="{ 'demo__rb-dd--open': openDropdown === 'paperSize' }">
              <button type="button" class="demo__rb-dd__btn" :disabled="disabled" :title="tr.paperSize" :aria-label="tr.paperSize" aria-haspopup="listbox" :aria-expanded="openDropdown === 'paperSize'" @click="toggleDropdown('paperSize')"><span class="demo__rb-dd__value">{{ active.paperSize }}</span><svg class="demo__rb-dd__chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5l3.5 3.5 3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg></button>
              <div v-if="openDropdown === 'paperSize'" class="demo__rb-dd__list" role="listbox" :aria-label="tr.paperSize" tabindex="-1">
                <button v-for="paper in ['A4', 'A3', 'A5', 'letter', 'legal', 'tabloid']" :key="paper" class="demo__rb-dd__opt" type="button" role="option" :aria-selected="active.paperSize === paper" @click="onDropdownPick('paperSize', paper)"><span class="demo__rb-dd__check" aria-hidden="true" /><span class="demo__rb-dd__label">{{ paper === 'letter' ? tr.paperLetter : paper === 'legal' ? tr.paperLegal : paper === 'tabloid' ? tr.paperTabloid : paper }}</span></button>
              </div>
            </div>
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" @click="props.instance?.openPageSetup()"><RibbonIcon name="options" /><span>{{ tr.pageSetup }}</span></button>
          </div>
          <div class="demo__ribbon-label">{{ tr.pageSetup }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.print">
          <div class="demo__ribbon-tools"><button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" @click="props.instance?.print()"><RibbonIcon name="print" /><span>{{ tr.print }}</span></button></div>
          <div class="demo__ribbon-label">{{ tr.print }}</div>
        </section>
      </template>
      <template v-else-if="props.activeTab === 'formulas'">
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.functionLibrary">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" :aria-keyshortcuts="keyShortcuts('fx')" @click="props.instance?.openFunctionArguments()">
              <RibbonIcon name="function" />
            </button>
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" @click="onAutoSum">
              <RibbonIcon name="autosum" /><span>{{ tr.autoSum }}</span>
            </button>
            <button class="demo__rb demo__rb--mono" type="button" :disabled="disabled" @click="props.instance?.openFunctionArguments('SUM')">
              <RibbonIcon name="function" /><span>SUM</span>
            </button>
            <button class="demo__rb demo__rb--mono" type="button" :disabled="disabled" @click="props.instance?.openFunctionArguments('AVERAGE')">
              <RibbonIcon name="function" /><span>AVG</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ tr.functionLibrary }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.definedNames">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" :aria-keyshortcuts="keyShortcuts('namedRanges')" @click="props.instance?.openNamedRangeDialog()">
              <RibbonIcon name="names" /><span>{{ tr.names }}</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ tr.definedNames }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.formulaAuditing">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" @click="props.instance?.tracePrecedents()">
              <RibbonIcon name="trace" /><span>{{ tr.tracePrecedents }}</span>
            </button>
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" @click="props.instance?.traceDependents()">
              <RibbonIcon name="dependents" /><span>{{ tr.traceDependents }}</span>
            </button>
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" @click="props.instance?.clearTraces()">
              <RibbonIcon name="clearArrows" /><span>{{ tr.removeArrows }}</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ tr.formulaAuditing }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.calculation">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" :aria-keyshortcuts="keyShortcuts('recalcNow')" @click="props.instance?.recalc()">
              <RibbonIcon name="autosum" /><span>{{ tr.recalc }}</span>
            </button>
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" @click="props.instance?.openIterativeDialog()">
              <RibbonIcon name="options" /><span>{{ tr.options }}</span>
            </button>
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" @click="props.instance?.toggleWatchWindow()">
              <RibbonIcon name="watch" /><span>{{ tr.watch }}</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ tr.calculation }}</div>
        </section>
      </template>

      <template v-else-if="props.activeTab === 'data'">
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.sortFilter">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" :class="{ 'demo__rb--active': active.filterOn }" type="button" :disabled="disabled" @click="onFilterToggle">
              <RibbonIcon name="filter" /><span>{{ tr.filter }}</span>
            </button>
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" @click="onSort('asc')">
              <RibbonIcon name="sortAsc" /><span>A-Z</span>
            </button>
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" @click="onSort('desc')">
              <RibbonIcon name="sortDesc" /><span>Z-A</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ tr.sortFilter }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.dataTools">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" @click="onRemoveDuplicates">
              <RibbonIcon name="removeDuplicates" /><span>{{ tr.removeDuplicates }}</span>
            </button>
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" @click="props.instance?.openExternalLinksDialog()">
              <RibbonIcon name="link" /><span>{{ tr.links }}</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ tr.dataTools }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.outline">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" :class="{ 'demo__rb--active': active.rowsHidden }" type="button" :disabled="disabled" @click="onToggleRowsHidden">
              <RibbonIcon name="table" /><span>{{ active.rowsHidden ? tr.showRows : tr.hideRows }}</span>
            </button>
            <button class="demo__rb demo__rb--wide" :class="{ 'demo__rb--active': active.colsHidden }" type="button" :disabled="disabled" @click="onToggleColsHidden">
              <RibbonIcon name="table" /><span>{{ active.colsHidden ? tr.showCols : tr.hideCols }}</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ tr.outline }}</div>
        </section>
      </template>

      <template v-else-if="props.activeTab === 'review'">
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.proofing">
          <div class="demo__ribbon-tools"><button class="demo__rb demo__rb--wide" type="button" :disabled="!props.onSpellingReview" @click="props.onSpellingReview?.()"><RibbonIcon name="spelling" /><span>{{ tr.spelling }}</span></button></div>
          <div class="demo__ribbon-label">{{ tr.proofing }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.language">
          <div class="demo__ribbon-tools"><button class="demo__rb demo__rb--wide" type="button" :disabled="!props.onTranslate" @click="props.onTranslate?.()"><RibbonIcon name="translate" /><span>{{ tr.translate }}</span></button></div>
          <div class="demo__ribbon-label">{{ tr.language }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.comments">
          <div class="demo__ribbon-tools"><button class="demo__rb demo__rb--wide" :class="{ 'demo__rb--active': active.hasComment }" type="button" :disabled="disabled" @click="props.instance?.openCommentDialog()"><RibbonIcon :name="active.hasComment ? 'commentMultiple' : 'commentAdd'" /><span>{{ active.hasComment ? tr.editComment : tr.newComment }}</span></button></div>
          <div class="demo__ribbon-label">{{ tr.comments }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.find">
          <div class="demo__ribbon-tools"><button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" :aria-keyshortcuts="keyShortcuts('findReview')" @click="props.instance?.openFindReplace()"><RibbonIcon name="find" /><span>{{ tr.find }}</span></button></div>
          <div class="demo__ribbon-label">{{ tr.find }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.protection">
          <div class="demo__ribbon-tools"><button class="demo__rb demo__rb--wide" :class="{ 'demo__rb--active': active.protected }" type="button" :disabled="disabled" @click="props.instance?.toggleSheetProtection()"><RibbonIcon name="protect" /><span>{{ active.protected ? tr.unprotect : tr.protect }}</span></button></div>
          <div class="demo__ribbon-label">{{ tr.protection }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.accessibility">
          <div class="demo__ribbon-tools"><button class="demo__rb demo__rb--wide" type="button" :disabled="!props.onAccessibilityCheck" @click="props.onAccessibilityCheck?.()"><RibbonIcon name="accessibility" /><span>{{ tr.accessibility }}</span></button></div>
          <div class="demo__ribbon-label">{{ tr.accessibility }}</div>
        </section>
      </template>
      <template v-else-if="props.activeTab === 'view'">
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.workbookViews">
          <div class="demo__ribbon-tools"><button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" @click="props.instance?.toggleWatchWindow()"><RibbonIcon name="watch" /><span>{{ tr.watch }}</span></button></div>
          <div class="demo__ribbon-label">{{ tr.workbookViews }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.window">
          <div class="demo__ribbon-tools"><button class="demo__rb" :class="{ 'demo__rb--active': active.frozen }" type="button" :disabled="disabled" @click="onFreezeToggle"><RibbonIcon name="freeze" /><span>{{ tr.freeze }}</span></button></div>
          <div class="demo__ribbon-label">{{ tr.window }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.zoom">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--mono" :class="{ 'demo__rb--active': active.zoom === 0.75 }" type="button" :disabled="disabled" @click="onZoom(0.75)"><RibbonIcon name="zoom" /><span>75%</span></button>
            <button class="demo__rb demo__rb--mono" :class="{ 'demo__rb--active': active.zoom === 1 }" type="button" :disabled="disabled" @click="onZoom(1)"><RibbonIcon name="zoom" /><span>100%</span></button>
            <button class="demo__rb demo__rb--mono" :class="{ 'demo__rb--active': active.zoom === 1.25 }" type="button" :disabled="disabled" @click="onZoom(1.25)"><RibbonIcon name="zoom" /><span>125%</span></button>
          </div>
          <div class="demo__ribbon-label">{{ tr.zoom }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.protection">
          <div class="demo__ribbon-tools"><button class="demo__rb demo__rb--wide" :class="{ 'demo__rb--active': active.protected }" type="button" :disabled="disabled" @click="props.instance?.toggleSheetProtection()"><RibbonIcon name="protect" /><span>{{ active.protected ? tr.unprotect : tr.protect }}</span></button></div>
          <div class="demo__ribbon-label">{{ tr.protection }}</div>
        </section>
      </template>

      <template v-else-if="props.activeTab === 'automate'">
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="RIBBON_TAB_LABELS.automate[lang]">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" type="button" :disabled="!props.onRunScript" @click="props.onRunScript?.()">
              <RibbonIcon name="script" /><span>{{ tr.script }}</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ RIBBON_TAB_LABELS.automate[lang] }}</div>
        </section>
      </template>

      <template v-else-if="props.activeTab === 'acrobat'">
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.addIn">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" type="button" :disabled="!props.onAddIn" @click="props.onAddIn?.()">
              <RibbonIcon name="addIn" /><span>{{ tr.addIn }}</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ tr.addIn }}</div>
        </section>
        <section class="demo__ribbon-group demo__ribbon-group--tiles" :aria-label="tr.pdf">
          <div class="demo__ribbon-tools">
            <button class="demo__rb demo__rb--wide" type="button" :disabled="disabled" @click="props.instance?.print()">
              <RibbonIcon name="pdf" /><span>{{ tr.pdf }}</span>
            </button>
          </div>
          <div class="demo__ribbon-label">{{ tr.pdf }}</div>
        </section>
      </template>

      <template v-else></template>
    </div>
  </div>
</template>
