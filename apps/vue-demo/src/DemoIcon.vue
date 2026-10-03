<script setup lang="ts">
import { DEMO_ICONS, type DemoIconName } from '../../demo-shared/index.js';

defineProps<{ name: DemoIconName }>();
</script>

<template>
  <svg class="fc-tb__rb-icon" viewBox="0 0 20 20" stroke-width="1.45" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
    <path v-for="segment in DEMO_ICONS[name]" :key="segment.d" :d="segment.d" :fill="segment.fill ?? 'none'" :stroke="segment.stroke ?? 'currentColor'" />
  </svg>
</template>
