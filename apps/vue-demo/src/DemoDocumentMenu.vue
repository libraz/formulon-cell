<script setup lang="ts">
import { nextTick, onBeforeUnmount, ref, watch } from 'vue';
import { commitDemoBookName, type DemoUiStrings } from '../../demo-shared/index.js';
import DemoGlyph from './DemoGlyph.vue';

/** The Mac title: book name, the document-actions chevron and its popover. */
const props = defineProps<{ ui: DemoUiStrings; bookName: string }>();
const emit = defineEmits<{ rename: [name: string] }>();

const open = ref(false);
const draft = ref(props.bookName);
const root = ref<HTMLElement | null>(null);
const nameInput = ref<HTMLInputElement | null>(null);

const onPointer = (event: PointerEvent): void => {
  if (!root.value?.contains(event.target as Node)) open.value = false;
};
watch(open, (isOpen) => {
  if (isOpen) {
    draft.value = props.bookName;
    // The desktop selects the name for editing as the popover opens.
    void nextTick(() => nameInput.value?.select());
    document.addEventListener('pointerdown', onPointer, true);
  } else {
    document.removeEventListener('pointerdown', onPointer, true);
  }
});
onBeforeUnmount(() => document.removeEventListener('pointerdown', onPointer, true));

const commit = (): void => {
  const next = commitDemoBookName(draft.value, props.bookName);
  if (next !== props.bookName) emit('rename', next);
  draft.value = next;
};
const onKeydown = (event: KeyboardEvent): void => {
  if (event.key === 'Enter') {
    event.preventDefault();
    commit();
    open.value = false;
  } else if (event.key === 'Escape') {
    event.preventDefault();
    draft.value = props.bookName;
    open.value = false;
  }
};
</script>

<template>
  <div ref="root" class="fc-tb__title demo__document">
    <strong>{{ bookName }}</strong>
    <button
      type="button"
      class="demo__document-chevron"
      :aria-label="ui.documentMenu.button"
      :title="ui.documentMenu.button"
      aria-haspopup="dialog"
      :aria-expanded="open"
      @click="open = !open"
    />
    <div v-if="open" class="demo__document-menu" role="dialog" :aria-label="ui.documentMenu.button">
      <label class="demo__document-row">
        <span>{{ ui.documentMenu.name }}</span>
        <input ref="nameInput" v-model="draft" type="text" @blur="commit" @keydown="onKeydown" />
      </label>
      <div class="demo__document-row">
        <span>{{ ui.documentMenu.location }}</span>
        <button type="button" class="demo__document-location" disabled>
          <DemoGlyph name="folder" />
        </button>
      </div>
      <div class="demo__document-separator" role="separator" />
      <button type="button" class="demo__document-history" disabled>
        <DemoGlyph name="versionHistory" />
        <span>{{ ui.documentMenu.versionHistory }}</span>
      </button>
    </div>
  </div>
</template>
