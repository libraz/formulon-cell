# @libraz/formulon-cell-vue

[![npm version](https://img.shields.io/npm/v/@libraz/formulon-cell-vue.svg)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![license](https://img.shields.io/npm/l/@libraz/formulon-cell-vue.svg)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![bundle size](https://img.shields.io/bundlephobia/minzip/@libraz/formulon-cell-vue)](https://bundlephobia.com/package/@libraz/formulon-cell-vue)

Vue 3 component + composables for
[`@libraz/formulon-cell`](https://www.npmjs.com/package/@libraz/formulon-cell)
— the spreadsheet UI for the [formulon](https://github.com/libraz/formulon)
WASM calc engine.

## Install

```sh
npm install @libraz/formulon-cell-vue @libraz/formulon-cell vue zustand
```

## Quick Start

```vue
<script setup lang="ts">
import { Spreadsheet, presets, type SpreadsheetInstance } from '@libraz/formulon-cell-vue';
import '@libraz/formulon-cell/styles.css';

const ready = (inst: SpreadsheetInstance) => {
  console.log('mounted', inst.workbook.version);
};
</script>

<template>
  <Spreadsheet
    :features="presets.full()"
    locale="en"
    style="width: 100%; height: 100vh"
    @ready="ready"
  />
</template>
```

## Composables

```ts
import { computed, ref } from 'vue';
import { type SpreadsheetExposed, useI18n, useSelection } from '@libraz/formulon-cell-vue';

const sheetRef = ref<SpreadsheetExposed | null>(null);
const instance = computed(() => sheetRef.value?.instance.value ?? null);
const sel = useSelection(instance);
const { locale, strings } = useI18n(instance);
```

| Composable | Description |
|------------|-------------|
| `useSelection(instance)` | Subscribe to the active selection |
| `useI18n(instance)` | Read current locale + strings, reactive to runtime swaps |

## Toolbar

`SpreadsheetToolbar` is published as an SFC subpath so Vue bundlers can
compile it with the same pipeline as application components.

```vue
<script setup lang="ts">
import { type RibbonTab } from '@libraz/formulon-cell-vue';
import SpreadsheetToolbar from '@libraz/formulon-cell-vue/toolbar.vue';
import '@libraz/formulon-cell-vue/toolbar.css';
</script>
```

## Runtime prop updates

`theme`, `locale`, `strings`, `workbook`, `features`, and `extensions`
update the running spreadsheet through the core imperative API. The
component does **not** re-mount the canvas, so selection, focus, and
host event subscriptions stay intact.

## Core helpers

This package re-exports core command helpers and types — `createSessionChart`,
`saveSheetView`, `activateSheetView`, `listDefinedNames`,
`upsertDefinedName`, etc. — so Vue apps can type host chrome from a single
import.

## Documentation

For the complete API reference and bundler integration notes, see the
[project README](https://github.com/libraz/formulon-cell).

## License

[Apache License 2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
