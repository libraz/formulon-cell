import type { SpreadsheetInstance } from '@libraz/formulon-cell';
import { onUnmounted, ref, watch } from 'vue';
import { type ActiveState, EMPTY_ACTIVE_STATE, projectActiveState } from './model.js';

export function useToolbarActive(getInstance: () => SpreadsheetInstance | null) {
  const active = ref<ActiveState>(EMPTY_ACTIVE_STATE);
  let unsub: (() => void) | null = null;

  watch(
    getInstance,
    (inst) => {
      unsub?.();
      unsub = null;
      if (!inst) return;
      active.value = projectActiveState(inst);
      unsub = inst.store.subscribe(() => {
        active.value = projectActiveState(inst);
      });
    },
    { immediate: true },
  );

  onUnmounted(() => {
    unsub?.();
  });

  return active;
}
