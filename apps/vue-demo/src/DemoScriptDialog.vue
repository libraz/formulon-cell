<script setup lang="ts">
import { parseScriptCommand, type ScriptCommand } from '@libraz/formulon-cell';
import { onBeforeUnmount, onMounted, ref } from 'vue';
import {
  activateDemoModal,
  type DemoCommandStrings,
  type DemoUiStrings,
} from '../../demo-shared/index.js';

const props = defineProps<{ ui: DemoUiStrings; commandText: DemoCommandStrings }>();
// `submit` carries the parsed command; the host decides when to close.
const emit = defineEmits<{ submit: [command: ScriptCommand]; close: [] }>();

const command = ref('uppercase');
const error = ref<string | null>(null);
const modalEl = ref<HTMLElement | null>(null);
let modalCleanup: (() => void) | null = null;

const submit = (): void => {
  const parsed = parseScriptCommand(command.value);
  if (!parsed) {
    error.value = props.commandText.scriptCommandError;
    return;
  }
  emit('submit', parsed);
};

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
    :aria-label="commandText.script"
  >
    <form class="fc-tb__modal-panel fc-tb__modal-panel--narrow" @submit.prevent="submit">
      <header class="fc-tb__modal-header">
        <h2>{{ commandText.script }}</h2>
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
        <label class="fc-tb__modal-field">
          <span>{{ ui.command }}</span>
          <!-- No `autofocus`: the modal's focus trap owns the initial focus,
               and WebKit applies autofocus after it, which breaks the
               Tab/Shift+Tab wrap contract (and the React demo's parity). -->
          <input v-model="command" @input="error = null" />
        </label>
        <p v-if="error" class="fc-tb__modal-error">{{ error }}</p>
      </div>
      <footer class="fc-tb__modal-footer">
        <button type="button" class="fc-tb__btn" @click="emit('close')">{{ ui.cancel }}</button>
        <button type="submit" class="fc-tb__btn fc-tb__btn--active">{{ ui.run }}</button>
      </footer>
    </form>
  </div>
</template>
