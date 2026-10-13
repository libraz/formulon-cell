<script setup lang="ts">
// Vue adapter on top of `Spreadsheet.mountToolbar`. Mirrors the React adapter
// in `formulon-cell-react` — same prop shape, same delegation strategy. The
// component just owns a host div and mounts / disposes the core toolbar; all
// DOM, helpers, menus, and hook defaults come from core.
import {
  createDefaultDynamicDropdownsCtx,
  dictionaryLocaleFor,
  type DynamicDropdownsCtx,
  openHostMenuFirstDropdown,
  type RibbonHooks,
  type RibbonTab,
  Spreadsheet,
  type SpreadsheetInstance,
  type ToolbarInstance,
} from '@libraz/formulon-cell';
import { onBeforeUnmount, onMounted, ref, watch } from 'vue';
import type { SpreadsheetToolbarProps } from './toolbar.js';

const props = defineProps<SpreadsheetToolbarProps>();
const emit = defineEmits<{
  tabChange: [tab: RibbonTab];
  error: [error: unknown];
}>();

const hostEl = ref<HTMLDivElement | null>(null);
let toolbar: ToolbarInstance | null = null;

const mountToolbarFor = (instance: SpreadsheetInstance): void => {
  const host = hostEl.value;
  if (!host) return;
  if (toolbar) props.onToolbarReady?.(null);
  toolbar?.dispose();
  // A provided prop handles only its own menu item (`custom` for scripts,
  // `manage` for add-ins); every other action, and every action when the
  // prop is omitted, reaches core's default handler. Likewise review hooks
  // are registered only for provided props so core's built-in reports stay.
  // Host-supplied `dropdownActions` win over the wrapper's wiring.
  let coreDefaults: DynamicDropdownsCtx | null = null;
  const getCoreDefaults = (): DynamicDropdownsCtx => {
    coreDefaults ??= createDefaultDynamicDropdownsCtx(instance);
    return coreDefaults;
  };
  const dropdownOverrides: Partial<DynamicDropdownsCtx> = {};
  if (props.onRunScript) {
    dropdownOverrides.applyScriptAction = (action) =>
      action === 'custom' ? props.onRunScript?.() : getCoreDefaults().applyScriptAction(action);
  }
  if (props.onAddIn) {
    dropdownOverrides.applyAddInAction = (action) =>
      action === 'manage' ? props.onAddIn?.() : getCoreDefaults().applyAddInAction(action);
  }
  Object.assign(dropdownOverrides, props.dropdownActions);
  const review: NonNullable<RibbonHooks['review']> = {};
  if (props.onSpellingReview) review.spelling = () => props.onSpellingReview?.();
  if (props.onAccessibilityCheck) review.accessibility = () => props.onAccessibilityCheck?.();
  if (props.onTranslate) review.translate = () => props.onTranslate?.();
  try {
    toolbar = Spreadsheet.mountToolbar(host, instance, {
      lang: dictionaryLocaleFor(props.locale),
      activeTab: props.activeTab,
      ribbonTabs: props.ribbonTabs,
      onTabChange: (tab) => emit('tabChange', tab),
      interceptCommand: (id, button) => openHostMenuFirstDropdown(toolbar, id, button),
      // Opt into core's default dropdown-menu click delegator so Fill / Clear
      // / AutoSum / etc. work without each consumer reimplementing the
      // playground's `createDynamicDropdowns` wiring.
      dynamicDropdowns: dropdownOverrides,
      hooks: {
        review,
        drawing: {
          setInkMode: (mode) => {
            if (mode === 'pen') props.onDrawPen?.();
            else props.onDrawEraser?.();
          },
        },
      },
    });
  } catch (error) {
    props.onToolbarReady?.(null);
    props.onError?.(error);
    emit('error', error);
    return;
  }
  props.onToolbarReady?.(toolbar);
};

// Initial mount has to wait until `hostEl` is bound, so we use `onMounted`.
// Subsequent prop changes are picked up by the watcher below.
onMounted(() => {
  if (props.instance) mountToolbarFor(props.instance);
});

watch(
  // The `Boolean` entries track presence only: an omitted prop falls back to
  // core's default, so toggling presence re-mounts while identity changes don't.
  () =>
    [
      props.instance,
      props.locale,
      props.dropdownActions,
      props.ribbonTabs,
      Boolean(props.onSpellingReview),
      Boolean(props.onAccessibilityCheck),
      Boolean(props.onTranslate),
      Boolean(props.onRunScript),
      Boolean(props.onAddIn),
    ] as const,
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

<script lang="ts">
// `Toolbar` is the React-parity alias for this component. It has to be a
// re-export rather than `const Toolbar = <self import>`: the SFC referencing
// itself is a module cycle, and reading the default binding while the module
// is still evaluating throws before the component is ever created.
export { default as Toolbar } from './SpreadsheetToolbar.vue';
</script>
