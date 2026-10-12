<script setup lang="ts">
import { computed, ref, useId } from 'vue';
import { type DemoSelectOption, resolveDemoSelectKey } from '../../demo-shared/index.js';

const props = defineProps<{
  labelId: string;
  value: string;
  options: readonly DemoSelectOption[];
}>();

const emit = defineEmits<{
  change: [value: string];
}>();

const listId = useId();
const open = ref(false);
const activeIndex = ref(-1);
const selectedIndex = computed(() => props.options.findIndex((o) => o.value === props.value));
const optionId = (index: number): string => `${listId}-opt-${index}`;

const openAt = (index: number): void => {
  activeIndex.value = index;
  open.value = true;
};
const commit = (index: number): void => {
  const option = props.options[index];
  open.value = false;
  if (option && option.value !== props.value) emit('change', option.value);
};
const toggle = (): void => {
  if (open.value) open.value = false;
  else openAt(Math.max(0, selectedIndex.value));
};

const onKeyDown = (event: KeyboardEvent): void => {
  const action = resolveDemoSelectKey(
    event.key,
    open.value,
    activeIndex.value,
    selectedIndex.value,
    props.options.length,
  );
  if (!action) return;
  if (action.kind !== 'commit' || !action.keepDefault) event.preventDefault();
  if (action.kind === 'open') openAt(action.index);
  else if (action.kind === 'move') activeIndex.value = action.index;
  else if (action.kind === 'commit') commit(action.index);
  else open.value = false;
};
</script>

<template>
  <div class="demo__select">
    <button
      type="button"
      role="combobox"
      class="demo__select-trigger"
      :aria-labelledby="labelId"
      aria-haspopup="listbox"
      :aria-expanded="open"
      :aria-controls="listId"
      :aria-activedescendant="open && activeIndex >= 0 ? optionId(activeIndex) : undefined"
      :data-value="value"
      @click="toggle"
      @keydown="onKeyDown"
      @blur="open = false"
    >
      {{ options[selectedIndex]?.label ?? value }}
    </button>
    <div
      :id="listId"
      role="listbox"
      :aria-labelledby="labelId"
      class="demo__select-list"
      :hidden="!open"
    >
      <button
        v-for="(option, index) in options"
        :id="optionId(index)"
        :key="option.value"
        type="button"
        tabindex="-1"
        role="option"
        :aria-selected="option.value === value"
        :class="['demo__select-option', { 'demo__select-option--active': index === activeIndex }]"
        @mousedown.prevent
        @mouseenter="activeIndex = index"
        @click="commit(index)"
      >
        {{ option.label }}
      </button>
    </div>
  </div>
</template>
