<script setup lang="ts">
import type { RibbonTab, SpreadsheetInstance } from '@libraz/formulon-cell';
import { computed, onBeforeUnmount, onMounted, ref, watch } from 'vue';
import {
  type DemoPlatform,
  type DemoSearchItem,
  type DemoSearchUsagePrior,
  type DemoUiStrings,
  demoSearchOptionId,
  installDemoSearchShortcut,
  loadDemoSearchUsagePrior,
  queryDemoSearchItems,
  recordDemoSearchUsage,
  resolveDemoSearchKey,
  saveDemoSearchUsagePrior,
} from '../../demo-shared/index.js';
import DemoIcon from './DemoIcon.vue';

const props = defineProps<{
  ui: DemoUiStrings;
  platform: DemoPlatform;
  locale: string;
  bookName: string;
  instance: SpreadsheetInstance | null;
  showPanel: boolean;
  searchItems: readonly DemoSearchItem[];
}>();
const emit = defineEmits<{
  save: [];
  openBackstage: [];
  togglePanel: [];
  tabChange: [tab: RibbonTab];
}>();

const searchQuery = ref('');
const searchOpen = ref(false);
const searchActiveIndex = ref(-1);
const searchUsagePrior = ref<DemoSearchUsagePrior>(loadDemoSearchUsagePrior());
const searchInput = ref<HTMLInputElement | null>(null);
/** Quick Access toolbar element, exposed so the host can route F6 into it. */
const quickAccess = ref<HTMLElement | null>(null);
defineExpose({ quickAccess });

let disposeSearchShortcut: (() => void) | undefined;

const filteredCommands = computed(() => {
  return queryDemoSearchItems(props.searchItems, searchQuery.value, 8, searchUsagePrior.value);
});

const runCommand = (cmd: DemoSearchItem): void => {
  searchUsagePrior.value = recordDemoSearchUsage(searchUsagePrior.value, cmd);
  if (cmd.tab) emit('tabChange', cmd.tab);
  cmd.run();
  searchQuery.value = '';
  searchOpen.value = false;
  searchActiveIndex.value = -1;
};

const onSearchKeydown = (ev: KeyboardEvent): void => {
  const action = resolveDemoSearchKey(ev.key, searchActiveIndex.value, filteredCommands.value.length);
  if (action?.kind === 'close') {
    searchOpen.value = false;
    searchActiveIndex.value = -1;
    (ev.currentTarget as HTMLInputElement).blur();
  } else if (action?.kind === 'move') {
    ev.preventDefault();
    searchOpen.value = true;
    searchActiveIndex.value = action.index;
  } else if (action?.kind === 'run') {
    ev.preventDefault();
    const command = filteredCommands.value[action.index];
    if (command) runCommand(command);
  }
};

watch([searchQuery, searchOpen], () => {
  searchActiveIndex.value = -1;
});

watch(searchUsagePrior, (prior) => saveDemoSearchUsagePrior(prior));

onMounted(() => {
  disposeSearchShortcut = installDemoSearchShortcut(() => searchInput.value, props.platform);
});
onBeforeUnmount(() => {
  disposeSearchShortcut?.();
  disposeSearchShortcut = undefined;
});
</script>

<template>
  <header class="demo__head">
    <div class="fc-tb__titlebar">
      <div
        ref="quickAccess"
        class="demo__quick"
        role="toolbar"
        :aria-label="ui.quickAccessToolbar"
      >
        <button
          v-if="platform === 'mac'"
          type="button"
          class="demo__brand-mark"
          :aria-label="ui.file"
          :title="ui.file"
          @click="emit('openBackstage')"
        >
          <DemoIcon name="app" />
        </button>
        <span v-else class="demo__brand-mark" aria-hidden="true">
          <DemoIcon name="app" />
        </span>
        <button type="button" class="demo__title-icon" :aria-label="ui.save" @click="emit('save')">
          <DemoIcon name="save" />
        </button>
        <button type="button" class="demo__title-icon" :aria-label="ui.undo" @click="instance?.undo()">
          <DemoIcon name="undo" />
        </button>
        <button type="button" class="demo__title-icon" :aria-label="ui.redo" @click="instance?.redo()">
          <DemoIcon name="redo" />
        </button>
      </div>
      <div class="fc-tb__title">
        <strong>{{ bookName }}</strong>
        <span>{{ ui.saved }}</span>
      </div>
      <div class="fc-tb__search">
        <DemoIcon name="search" />
        <input
          ref="searchInput"
          v-model="searchQuery"
          type="search"
          role="combobox"
          :placeholder="ui.search"
          :aria-label="platform === 'mac' ? (locale === 'ja' ? '検索' : 'Search') : ui.searchCommands"
          aria-controls="demo-search-results"
          :aria-expanded="searchOpen"
          :aria-activedescendant="searchOpen && searchActiveIndex >= 0 ? demoSearchOptionId(searchActiveIndex) : undefined"
          @focus="searchOpen = true; searchActiveIndex = -1"
          @input="searchOpen = true; searchActiveIndex = -1"
          @keydown="onSearchKeydown"
          @blur="searchOpen = false"
        />
        <div v-if="searchOpen" id="demo-search-results" class="fc-tb__command-menu" role="listbox">
          <div v-if="filteredCommands.length === 0" class="fc-tb__command-empty">
            {{ ui.noCommands }}
          </div>
          <button
            v-for="(cmd, index) in filteredCommands"
            v-else
            :key="cmd.id"
            :id="demoSearchOptionId(index)"
            type="button"
            role="option"
            :aria-selected="index === searchActiveIndex"
            :aria-disabled="cmd.disabled ? 'true' : undefined"
            :data-disabled-reason="cmd.disabledReason"
            :class="[
              'fc-tb__command-item',
              {
                'fc-tb__command-item--active': index === searchActiveIndex,
                'fc-tb__command-item--disabled': cmd.disabled,
              },
            ]"
            @mousedown.prevent
            @mouseenter="searchActiveIndex = index"
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
        <button
          type="button"
          :class="['demo__share', { 'demo__share--active': showPanel }]"
          :aria-pressed="showPanel"
          @click="emit('togglePanel')"
        >
          {{ ui.demoPane }}
        </button>
        <span class="demo__avatar" role="img" :aria-label="ui.signedInUser">FC</span>
      </div>
    </div>
  </header>
</template>
