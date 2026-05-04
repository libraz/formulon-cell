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
    :features="presets.excel()"
    locale="ja"
    style="width: 100%; height: 100vh"
    @ready="ready"
  />
</template>
```

## Composables

```ts
import { ref } from 'vue';
import { useI18n, useSelection } from '@libraz/formulon-cell-vue';

const sheetRef = ref<{ instance: any } | null>(null);
const sel = useSelection(computed(() => sheetRef.value?.instance ?? null));
const { locale, strings } = useI18n(computed(() => sheetRef.value?.instance ?? null));
```

## License

[Apache-2.0](./LICENSE)
