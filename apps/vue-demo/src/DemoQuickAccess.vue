<script setup lang="ts">
import type { LegacySpreadsheetInstance as SpreadsheetInstance } from '@libraz/formulon-cell';
import { onBeforeUnmount, ref, watch } from 'vue';
import {
  DEFAULT_DEMO_QUICK_ACCESS,
  DEMO_QUICK_ACCESS_GLYPHS,
  type DemoQuickAccessDeps,
  type DemoQuickAccessId,
  type DemoUiStrings,
  demoQuickAccessEnabled,
  demoQuickAccessMenu,
  runDemoQuickAccess,
  toggleDemoQuickAccess,
} from '../../demo-shared/index.js';
import DemoGlyph from './DemoGlyph.vue';

/** The Mac title bar's Quick Access commands and their customize menu. */
const props = defineProps<{
  ui: DemoUiStrings;
  instance: SpreadsheetInstance | null;
  deps: DemoQuickAccessDeps;
}>();

const shown = ref<readonly DemoQuickAccessId[]>(DEFAULT_DEMO_QUICK_ACCESS);
const autosave = ref(false);
const menuOpen = ref(false);
const menu = ref<HTMLElement | null>(null);
const chevron = ref<HTMLButtonElement | null>(null);
// Bumped on every history change so Undo enables and disables with it.
const historyRevision = ref(0);

let unsubscribeHistory: (() => void) | undefined;
watch(
  () => props.instance,
  (instance) => {
    unsubscribeHistory?.();
    unsubscribeHistory = instance?.history.subscribe(() => {
      historyRevision.value += 1;
    });
  },
  { immediate: true },
);

const enabled = (id: DemoQuickAccessId): boolean => {
  void historyRevision.value;
  return demoQuickAccessEnabled(id, props.instance);
};

const closeMenu = (): void => {
  menuOpen.value = false;
};
const onPointer = (event: PointerEvent): void => {
  const target = event.target as Node;
  if (!menu.value?.contains(target) && !chevron.value?.contains(target)) closeMenu();
};
const onKey = (event: KeyboardEvent): void => {
  if (event.key !== 'Escape') return;
  closeMenu();
  chevron.value?.focus();
};
watch(menuOpen, (open) => {
  if (open) {
    document.addEventListener('pointerdown', onPointer, true);
    document.addEventListener('keydown', onKey, true);
  } else {
    document.removeEventListener('pointerdown', onPointer, true);
    document.removeEventListener('keydown', onKey, true);
  }
});
onBeforeUnmount(() => {
  unsubscribeHistory?.();
  closeMenu();
  document.removeEventListener('pointerdown', onPointer, true);
  document.removeEventListener('keydown', onKey, true);
});

const toggle = (id: DemoQuickAccessId): void => {
  shown.value = toggleDemoQuickAccess(shown.value, id);
  closeMenu();
};
</script>

<template>
  <template v-for="id in shown" :key="id">
    <label v-if="id === 'autosave'" class="demo__autosave">
      <span>{{ ui.quickAccess.autosave }}</span>
      <input v-model="autosave" type="checkbox" role="switch" />
    </label>
    <button
      v-else
      type="button"
      :class="['demo__title-icon', `demo__qat-${id}`]"
      :aria-label="ui.quickAccess.buttons[id]"
      :title="ui.quickAccess.buttons[id]"
      :disabled="!enabled(id)"
      @click="runDemoQuickAccess(id, deps)"
    >
      <DemoGlyph v-if="DEMO_QUICK_ACCESS_GLYPHS[id]" :name="DEMO_QUICK_ACCESS_GLYPHS[id]!" />
      <span v-if="id === 'undo'" class="demo__qat-caret" aria-hidden="true" />
    </button>
  </template>
  <div class="demo__qat-customize">
    <button
      ref="chevron"
      type="button"
      class="demo__title-icon demo__qat-more"
      :aria-label="ui.quickAccess.customize"
      :title="ui.quickAccess.customize"
      aria-haspopup="menu"
      :aria-expanded="menuOpen"
      @click="menuOpen = !menuOpen"
    >
      <DemoGlyph name="moreCommands" />
    </button>
    <div v-if="menuOpen" ref="menu" class="demo__qat-menu" role="menu" :aria-label="ui.quickAccess.customize">
      <div class="demo__qat-menu-heading" role="presentation">{{ ui.quickAccess.customize }}</div>
      <button
        v-for="item in demoQuickAccessMenu(ui, shown)"
        :key="item.id"
        type="button"
        role="menuitemcheckbox"
        :aria-checked="item.checked"
        class="demo__qat-menu-item"
        @click="toggle(item.id)"
      >
        {{ item.label }}
      </button>
      <div class="demo__qat-menu-separator" role="separator" />
      <button type="button" role="menuitem" class="demo__qat-menu-item" disabled>
        {{ ui.quickAccess.moreCommands }}
      </button>
    </div>
  </div>
</template>
