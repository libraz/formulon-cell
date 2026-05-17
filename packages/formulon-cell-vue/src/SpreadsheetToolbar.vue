<script setup lang="ts">
// Vue adapter on top of `Spreadsheet.mountToolbar`. Mirrors the React adapter
// in `formulon-cell-react` — same prop shape, same delegation strategy. The
// component just owns a host div and mounts / disposes the core toolbar; all
// DOM, helpers, menus, and hook defaults come from core.
import {
  type FeatureFlags,
  Spreadsheet,
  type SpreadsheetInstance,
  type ToolbarInstance,
} from '@libraz/formulon-cell';
import { onBeforeUnmount, onMounted, ref, watch } from 'vue';
import type { RibbonTab } from './toolbar/model.js';

interface Props {
  instance: SpreadsheetInstance | null;
  features?: FeatureFlags;
  activeTab: RibbonTab;
  locale: string;
  onSpellingReview?: () => void;
  onAccessibilityCheck?: () => void;
  onRunScript?: () => void;
  onDrawPen?: () => void;
  onDrawEraser?: () => void;
  onTranslate?: () => void;
  onAddIn?: () => void;
  onNewWorkbook?: () => void;
  onOpenWorkbook?: () => void;
  onSaveWorkbook?: () => void;
  onSaveWorkbookAs?: () => void;
}

const props = defineProps<Props>();
const emit = defineEmits<{
  tabChange: [tab: RibbonTab];
}>();

const hostEl = ref<HTMLDivElement | null>(null);
let toolbar: ToolbarInstance | null = null;

const mountToolbarFor = (instance: SpreadsheetInstance): void => {
  const host = hostEl.value;
  if (!host) return;
  toolbar?.dispose();
  toolbar = Spreadsheet.mountToolbar(host, instance, {
    lang: props.locale === 'en' ? 'en' : 'ja',
    activeTab: props.activeTab,
    onTabChange: (tab) => emit('tabChange', tab),
    hooks: {
      review: {
        spelling: () => props.onSpellingReview?.(),
        accessibility: () => props.onAccessibilityCheck?.(),
        translate: () => props.onTranslate?.(),
      },
      automation: {
        runScript: () => props.onRunScript?.(),
        addInManager: () => props.onAddIn?.(),
      },
      drawing: {
        setInkMode: (mode) => {
          if (mode === 'pen') props.onDrawPen?.();
          else props.onDrawEraser?.();
        },
      },
    },
  });
};

// Initial mount has to wait until `hostEl` is bound, so we use `onMounted`.
// Subsequent prop changes are picked up by the watcher below.
onMounted(() => {
  if (props.instance) mountToolbarFor(props.instance);
});

watch(
  () => [props.instance, props.locale] as const,
  ([instance]) => {
    if (!hostEl.value) return;
    if (instance) mountToolbarFor(instance);
    else {
      toolbar?.dispose();
      toolbar = null;
    }
  },
  { flush: 'post' },
);

// Forward external tab changes without re-mounting.
watch(
  () => props.activeTab,
  (next) => {
    if (toolbar && toolbar.getActiveTab() !== next) toolbar.setActiveTab(next);
  },
);

onBeforeUnmount(() => {
  toolbar?.dispose();
  toolbar = null;
});
</script>

<template>
  <!-- `display: contents` keeps the wrapper out of the layout tree so the
       core ribbon-shell (`flex: 0 0 auto`) sees the parent flex column
       directly. Without this the extra div breaks the flex chain and the
       sibling sheet element collapses to zero height. -->
  <div ref="hostEl" style="display: contents"></div>
</template>
