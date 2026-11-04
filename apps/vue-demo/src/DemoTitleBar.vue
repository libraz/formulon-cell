<script setup lang="ts">
import type { RibbonTab, SpreadsheetInstance } from '@libraz/formulon-cell';
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue';
import {
  buildDemoSearchSuggestions,
  type DemoPlatform,
  type DemoSearchItem,
  type DemoSearchUsagePrior,
  type DemoSearchEvent,
  type DemoSearchState,
  type DemoUiStrings,
  demoSearchOptionId,
  INITIAL_DEMO_SEARCH_STATE,
  installDemoSearchShortcut,
  loadDemoSearchUsagePrior,
  queryDemoSearchItems,
  recordDemoSearchUsage,
  reduceDemoSearch,
  resolveDemoSearchKey,
  saveDemoSearchUsagePrior,
} from '../../demo-shared/index.js';
import DemoDocumentMenu from './DemoDocumentMenu.vue';
import DemoGlyph from './DemoGlyph.vue';
import DemoIcon from './DemoIcon.vue';
import DemoQuickAccess from './DemoQuickAccess.vue';

const props = defineProps<{
  ui: DemoUiStrings;
  platform: DemoPlatform;
  bookName: string;
  instance: SpreadsheetInstance | null;
  showPanel: boolean;
  searchItems: readonly DemoSearchItem[];
  /** Runs a ribbon command id through the toolbar; false when none matched. */
  applyCommand: (commandId: string) => boolean;
}>();
const emit = defineEmits<{
  save: [];
  saveAs: [];
  rename: [name: string];
  newWorkbook: [];
  openBackstage: [];
  togglePanel: [];
  tabChange: [tab: RibbonTab];
}>();

const search = ref<DemoSearchState>(INITIAL_DEMO_SEARCH_STATE);
const dispatchSearch = (event: DemoSearchEvent): void => {
  search.value = reduceDemoSearch(search.value, event);
};
const searchQuery = computed(() => search.value.query);
const searchOpen = computed(() => search.value.open);
const searchActiveIndex = computed(() => search.value.activeIndex);
const searchUsagePrior = ref<DemoSearchUsagePrior>(loadDemoSearchUsagePrior());
const searchInput = ref<HTMLInputElement | null>(null);
/** Quick Access toolbar element, exposed so the host can route F6 into it. */
const quickAccess = ref<HTMLElement | null>(null);
defineExpose({ quickAccess });

let disposeSearchShortcut: (() => void) | undefined;

const filteredCommands = computed(() => {
  return queryDemoSearchItems(props.searchItems, searchQuery.value, 8, searchUsagePrior.value);
});

const mac = computed(() => props.platform === 'mac');
// The Mac title bar shows a magnifier button until the search is opened.
const searchActive = ref(false);
const openSearch = (): void => {
  searchActive.value = true;
  void nextTick(() => searchInput.value?.focus());
};

const quickAccessDeps = computed(() => ({
  instance: props.instance,
  newWorkbook: () => emit('newWorkbook'),
  saveWorkbook: () => emit('save'),
  saveWorkbookAs: () => emit('saveAs'),
  setRibbonTab: (tab: RibbonTab) => emit('tabChange', tab),
  applyCommand: props.applyCommand,
}));

const suggestionRows = computed(() =>
  mac.value
    ? buildDemoSearchSuggestions(props.ui, {
        instance: props.instance,
        setRibbonTab: (tab) => emit('tabChange', tab),
        applyCommand: props.applyCommand,
      })
    : [],
);
const suggestions = computed(() => suggestionRows.value.flatMap((section) => section.items));
// An empty Mac query lists the suggestions; a typed one lists the matches.
const showSuggestions = computed(() => mac.value && searchQuery.value.trim() === '');
const rowCount = computed(() =>
  showSuggestions.value ? suggestions.value.length : filteredCommands.value.length,
);

const closeSearch = (): void => {
  dispatchSearch({ type: 'reset' });
  searchActive.value = false;
};

const runCommand = (cmd: DemoSearchItem): void => {
  searchUsagePrior.value = recordDemoSearchUsage(searchUsagePrior.value, cmd);
  if (cmd.tab) emit('tabChange', cmd.tab);
  cmd.run();
  dispatchSearch({ type: 'reset' });
};

const onSearchInput = (ev: Event): void => {
  const input = ev.currentTarget as HTMLInputElement;
  dispatchSearch({ type: 'input', value: input.value, focused: document.activeElement === input });
};

const runRow = (index: number): void => {
  if (!showSuggestions.value) {
    const command = filteredCommands.value[index];
    if (command) runCommand(command);
    return;
  }
  const suggestion = suggestions.value[index];
  if (!suggestion || suggestion.disabled) return;
  closeSearch();
  suggestion.run();
};

const onSearchBlur = (): void => {
  dispatchSearch({ type: 'blur' });
  if (mac.value && searchQuery.value.trim() === '') searchActive.value = false;
};

const onSearchKeydown = (ev: KeyboardEvent): void => {
  const action = resolveDemoSearchKey(ev.key, searchActiveIndex.value, rowCount.value);
  if (action?.kind === 'close') {
    dispatchSearch({ type: 'key', action });
    (ev.currentTarget as HTMLInputElement).blur();
  } else if (action?.kind === 'move') {
    ev.preventDefault();
    dispatchSearch({ type: 'key', action });
  } else if (action?.kind === 'run') {
    ev.preventDefault();
    runRow(action.index);
  }
};

watch(searchUsagePrior, (prior) => saveDemoSearchUsagePrior(prior));

onMounted(() => {
  disposeSearchShortcut = installDemoSearchShortcut(
    () => searchInput.value,
    props.platform,
    mac.value ? openSearch : undefined,
  );
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
        <DemoQuickAccess v-if="mac" :ui="ui" :instance="instance" :deps="quickAccessDeps" />
        <template v-else>
          <button type="button" class="demo__title-icon" :aria-label="ui.save" :title="ui.save" @click="emit('save')">
            <DemoIcon name="save" />
          </button>
          <button type="button" class="demo__title-icon" :aria-label="ui.undo" :title="ui.undo" @click="instance?.undo()">
            <DemoIcon name="undo" />
          </button>
          <button type="button" class="demo__title-icon" :aria-label="ui.redo" :title="ui.redo" @click="instance?.redo()">
            <DemoIcon name="redo" />
          </button>
        </template>
      </div>
      <DemoDocumentMenu v-if="mac" :ui="ui" :book-name="bookName" @rename="emit('rename', $event)" />
      <div v-else class="fc-tb__title">
        <strong>{{ bookName }}</strong>
        <span>{{ ui.saved }}</span>
      </div>
      <button
        v-if="mac && !searchActive"
        type="button"
        class="demo__search-button"
        :aria-label="ui.searchSuggestions.button"
        :title="ui.searchSuggestions.button"
        @click="openSearch"
      >
        <DemoGlyph name="find" />
      </button>
      <div v-else class="fc-tb__search">
        <DemoGlyph v-if="mac" name="find" />
        <DemoIcon v-else name="search" />
        <input
          ref="searchInput"
          :value="searchQuery"
          type="search"
          role="combobox"
          :placeholder="mac ? ui.searchSuggestions.button : ui.search"
          :aria-label="mac ? ui.search : ui.searchCommands"
          aria-controls="demo-search-results"
          :aria-expanded="searchOpen"
          :aria-activedescendant="searchOpen && searchActiveIndex >= 0 ? demoSearchOptionId(searchActiveIndex) : undefined"
          @focus="dispatchSearch({ type: 'focus' })"
          @input="onSearchInput"
          @keydown="onSearchKeydown"
          @blur="onSearchBlur"
        />
        <div
          v-if="searchOpen && showSuggestions"
          id="demo-search-results"
          class="fc-tb__command-menu demo__search-suggestions"
          role="listbox"
          :aria-label="ui.searchSuggestions.popup"
        >
          <div
            v-for="section in suggestionRows"
            :key="section.id"
            :class="['demo__search-section', `demo__search-section--${section.id}`]"
            role="group"
            :aria-label="section.heading ?? undefined"
          >
            <div v-if="section.heading" class="demo__search-heading" role="presentation">
              {{ section.heading }}
            </div>
            <button
              v-for="item in section.items"
              :key="item.id"
              :id="demoSearchOptionId(suggestions.indexOf(item))"
              type="button"
              role="option"
              :aria-selected="suggestions.indexOf(item) === searchActiveIndex"
              :aria-disabled="item.disabled ? 'true' : undefined"
              :class="[
                'demo__search-suggestion',
                { 'demo__search-suggestion--active': suggestions.indexOf(item) === searchActiveIndex },
              ]"
              @mousedown.prevent
              @mouseenter="dispatchSearch({ type: 'hover', index: suggestions.indexOf(item) })"
              @click="runRow(suggestions.indexOf(item))"
            >
              <DemoGlyph :name="item.glyph" />
              <span class="demo__search-suggestion-label">{{ item.label }}</span>
              <span v-if="item.submenu" class="demo__search-submenu" aria-hidden="true" />
            </button>
          </div>
        </div>
        <div v-else-if="searchOpen" id="demo-search-results" class="fc-tb__command-menu" role="listbox">
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
            @mouseenter="dispatchSearch({ type: 'hover', index })"
            @click="runCommand(cmd)"
          >
            <strong>{{ cmd.label }}</strong>
            <span>{{ cmd.hint }}</span>
          </button>
        </div>
      </div>
      <div class="demo__account">
        <button v-if="!mac" type="button" class="demo__share">
          {{ ui.share }}
        </button>
        <button
          type="button"
          :class="['demo__share', 'demo__pane-toggle', { 'demo__share--active': showPanel }]"
          :aria-pressed="showPanel"
          @click="emit('togglePanel')"
        >
          {{ ui.demoPane }}
        </button>
        <template v-if="mac">
          <button
            type="button"
            class="demo__title-icon demo__comments"
            :aria-label="ui.comments"
            :title="ui.comments"
            @click="applyCommand('mac.review.showComments')"
          >
            <DemoGlyph name="comments" />
          </button>
          <button type="button" class="demo__share-pill">
            <DemoGlyph name="share" />
            <span>{{ ui.share }}</span>
            <span class="demo__share-caret" aria-hidden="true" />
          </button>
        </template>
        <span v-else class="demo__avatar" role="img" :aria-label="ui.signedInUser">FC</span>
      </div>
    </div>
  </header>
</template>
