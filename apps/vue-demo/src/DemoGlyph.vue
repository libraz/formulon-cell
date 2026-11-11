<script setup lang="ts">
import type { ExcelRibbonIconName } from '@libraz/formulon-cell/toolbar';
import { onMounted, ref, watch } from 'vue';
import { createDemoGlyph } from '../../demo-shared/index.js';

/** Mounts a core icon-kit glyph; the element is built by the core renderer. */
const props = defineProps<{ name: ExcelRibbonIconName }>();
const host = ref<HTMLSpanElement | null>(null);
const mount = (): void => {
  const svg = createDemoGlyph(props.name);
  host.value?.replaceChildren(...(svg ? [svg] : []));
};
onMounted(mount);
watch(() => props.name, mount);
</script>

<template>
  <span ref="host" class="demo__glyph" aria-hidden="true" />
</template>
