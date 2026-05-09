# @libraz/formulon-cell-vue

Vue 3 component + composables for [`@libraz/formulon-cell`](../formulon-cell/README.md).

## Install

```sh
npm install @libraz/formulon-cell-vue @libraz/formulon-cell vue zustand
```

## Quick start

```vue
<script setup lang="ts">
import { ref } from 'vue';
import { Spreadsheet, presets, type SpreadsheetInstance } from '@libraz/formulon-cell-vue';
import '@libraz/formulon-cell/styles.css';

const ready = (inst: SpreadsheetInstance) => {
  console.log('mounted', inst.workbook.version);
};
</script>

<template>
  <Spreadsheet
    :features="presets.full()"
    locale="ja"
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

## Core helpers

The package re-exports the core command helpers and types, including
`createSessionChart`, `saveSheetView`, `activateSheetView`,
`listDefinedNames`, and `upsertDefinedName`, so Vue apps can type their host
chrome from one import.

## Runtime prop updates

`theme`, `locale`, `strings`, `workbook`, `features`, and `extensions` update
the running spreadsheet through the core imperative API. The component does
not re-mount the canvas when these props change, so selection, focus, and
host event subscriptions stay intact.

## License

[Apache-2.0](./LICENSE)
