<script setup lang="ts">
// Vue adapter on top of `Spreadsheet.mountToolbar`. Mirrors the React adapter
// in `formulon-cell-react` — same prop shape, same delegation strategy. The
// component just owns a host div and mounts / disposes the core toolbar; all
// DOM, helpers, menus, and hook defaults come from core.
import {
  type DynamicDropdownsCtx,
  type FeatureFlags,
  type RibbonTab,
  Spreadsheet,
  type SpreadsheetInstance,
  type ToolbarInstance,
} from '@libraz/formulon-cell';
import { onBeforeUnmount, onMounted, ref, watch } from 'vue';

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
  /** Receives the mounted core toolbar instance so hosts can dispatch shared
   *  commands from titlebar search / Tell me without querying DOM buttons. */
  onToolbarReady?: (toolbar: ToolbarInstance | null) => void;
  /** Override one or more entries in the core's default dynamic-dropdowns
   *  context. Use for dialog-opening handlers (sort, protect, file picker,
   *  etc.) that the wrapper can't represent as a named prop. Handlers
   *  supplied here win over the wrapper's built-in script/addIn wiring. */
  dropdownActions?: Partial<DynamicDropdownsCtx>;
  /** Shared tab surface. Use `EXCEL365_STANDARD_RIBBON_TABS` for the
   *  baseline Excel profile; append optional tabs only when those
   *  add-in/automation surfaces are intentionally exposed. */
  ribbonTabs?: readonly RibbonTab[];
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
  if (toolbar) props.onToolbarReady?.(null);
  toolbar?.dispose();
  // Host-supplied `dropdownActions` win over the wrapper's built-in
  // script/addIn wiring — consumers can fully replace those if they want.
  const dropdownOverrides: Partial<DynamicDropdownsCtx> = {
    applyScriptAction: (action) => {
      if (action === 'custom') props.onRunScript?.();
    },
    applyAddInAction: (action) => {
      if (action === 'manage') props.onAddIn?.();
    },
    ...props.dropdownActions,
  };
  toolbar = Spreadsheet.mountToolbar(host, instance, {
    lang: props.locale === 'en' ? 'en' : 'ja',
    activeTab: props.activeTab,
    ribbonTabs: props.ribbonTabs,
    onTabChange: (tab) => emit('tabChange', tab),
    // Opt into core's default dropdown-menu click delegator so Fill / Clear
    // / AutoSum / etc. work without each consumer reimplementing the
    // playground's `createDynamicDropdowns` wiring.
    dynamicDropdowns: dropdownOverrides,
    hooks: {
      review: {
        spelling: () => props.onSpellingReview?.(),
        accessibility: () => props.onAccessibilityCheck?.(),
        translate: () => props.onTranslate?.(),
      },
      drawing: {
        setInkMode: (mode) => {
          if (mode === 'pen') props.onDrawPen?.();
          else props.onDrawEraser?.();
        },
      },
    },
  });
  props.onToolbarReady?.(toolbar);
};

// Initial mount has to wait until `hostEl` is bound, so we use `onMounted`.
// Subsequent prop changes are picked up by the watcher below.
onMounted(() => {
  if (props.instance) mountToolbarFor(props.instance);
});

watch(
  () => [props.instance, props.locale, props.dropdownActions, props.ribbonTabs] as const,
  ([instance]) => {
    if (!hostEl.value) return;
    if (instance) mountToolbarFor(instance);
    else {
      props.onToolbarReady?.(null);
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
  props.onToolbarReady?.(null);
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
