<script setup lang="ts">
import { onBeforeUnmount, onMounted, ref } from 'vue';
import {
  activateDemoModal,
  type DemoReviewDialogState,
  type DemoUiStrings,
} from '../../demo-shared/index.js';

defineProps<{ dialog: DemoReviewDialogState; ui: DemoUiStrings }>();
const emit = defineEmits<{ close: [] }>();

const modalEl = ref<HTMLElement | null>(null);
let modalCleanup: (() => void) | null = null;

onMounted(() => {
  if (modalEl.value) modalCleanup = activateDemoModal(modalEl.value, () => emit('close'));
});
onBeforeUnmount(() => {
  modalCleanup?.();
  modalCleanup = null;
});
</script>

<template>
  <div
    ref="modalEl"
    class="fc-tb__modal"
    role="dialog"
    aria-modal="true"
    :aria-label="dialog.title"
  >
    <section class="fc-tb__modal-panel">
      <header class="fc-tb__modal-header">
        <h2>{{ dialog.title }}</h2>
        <button
          type="button"
          class="fc-tb__modal-x"
          :aria-label="ui.close"
          @click="emit('close')"
        >
          ×
        </button>
      </header>
      <div class="fc-tb__modal-body">
        <p v-if="dialog.items.length === 0" class="fc-tb__modal-empty">
          {{ ui.noIssuesFound }}
        </p>
        <ul v-else class="fc-tb__modal-list">
          <li v-for="(item, index) in dialog.items" :key="`${item.label}-${index}`">
            <strong>{{ item.label }}</strong>
            <span>{{ item.detail }}</span>
          </li>
        </ul>
      </div>
      <footer class="fc-tb__modal-footer">
        <button type="button" class="fc-tb__btn" @click="emit('close')">{{ ui.ok }}</button>
      </footer>
    </section>
  </div>
</template>
