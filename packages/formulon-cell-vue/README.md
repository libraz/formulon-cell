# @libraz/formulon-cell-vue

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![npm — core](https://img.shields.io/npm/v/@libraz/formulon-cell?label=core)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — react](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![Vue](https://img.shields.io/badge/Vue-3-blue?logo=vuedotjs)](https://vuejs.org/)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-blue)](https://formulon.libraz.net/cell/)
[![Demo](https://img.shields.io/badge/demo-live-brightgreen)](https://formulon.libraz.net/cell/demo)

**A working spreadsheet in a Vue app.** `<Spreadsheet>` and `SpreadsheetToolbar` wrap [`@libraz/formulon-cell`](https://www.npmjs.com/package/@libraz/formulon-cell), a canvas-rendered grid with desktop-spreadsheet chrome over the [formulon](https://github.com/libraz/formulon) WASM calc engine. Prop changes go through the core imperative API, so the canvas is never re-mounted and selection, focus, and event subscriptions survive.

Try it in the [live demo](https://formulon.libraz.net/cell/demo).

The grid, ribbon, menus, commands, and dialogs all live in core; this package only binds them to Vue.

## Install

```sh
npm install @libraz/formulon-cell-vue @libraz/formulon-cell vue zustand
```

Bundler settings are in the [project README](https://github.com/libraz/formulon-cell#bundler-integration).

## Quick start

```vue
<script setup lang="ts">
import { Spreadsheet, presets, type SpreadsheetInstance } from '@libraz/formulon-cell-vue';
import '@libraz/formulon-cell/styles.css';

const ready = (instance: SpreadsheetInstance) => {
  console.log('mounted', instance.workbook.version);
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
const sheetRef = ref<SpreadsheetExposed | null>(null);
const instance = computed(() => sheetRef.value?.instance.value ?? null);
const selection = useSelection(instance);
const { locale, strings } = useI18n(instance);
```

| composable | description |
|------------|-------------|
| `useSelection(instance)` | the active selection |
| `useSpreadsheet(instance, selector, fallback)` | any slice of the store state |
| `useI18n(instance)` | current locale and strings, updated on runtime swaps |
| `useSpreadsheetEvent(instance, event, handler)` | subscribe to `cellChange`, `selectionChange`, `recalc`, … |

## Ribbon toolbar

`SpreadsheetToolbar` ships as an SFC subpath so it compiles with your own Vue pipeline. It adapts core `Spreadsheet.mountToolbar`; pass `dropdownActions` to replace individual dropdown handlers without forking the ribbon.

```vue
<script setup lang="ts">
import SpreadsheetToolbar from '@libraz/formulon-cell-vue/toolbar.vue';
import '@libraz/formulon-cell-vue/toolbar.css';
</script>

<template>
  <SpreadsheetToolbar :instance="instance" :active-tab="tab" locale="en" @tab-change="tab = $event" />
</template>
```

## Restricted embedding

`ui`, `policy`, `viewport`, `context-menu`, `overlays`, and `toolbar` props apply at runtime without re-mounting. `fixedFormPolicy` and `viewerPolicy` are re-exported, and successful batch updates are emitted as `change-batch`.

```vue
<Spreadsheet
  :ui="{ profile: 'embedded' }"
  :policy="fixedFormPolicy([{ sheet: 0, r0: 1, c0: 1, r1: 9, c1: 2 }])"
  @change-batch="save"
/>
```

See [Embedding and use cases](https://formulon.libraz.net/cell/embedding) for viewport limits, context-menu composition, and host updates.

## Documentation

- [Live demo](https://formulon.libraz.net/cell/demo)
- [React / Vue](https://formulon.libraz.net/cell/frameworks)
- [Options](https://formulon.libraz.net/cell/options)
- [Embedding and use cases](https://formulon.libraz.net/cell/embedding)
- [Modals and fullscreen](https://formulon.libraz.net/cell/modals)
- [Theming](https://formulon.libraz.net/cell/theming)
- [Language and labels](https://formulon.libraz.net/cell/i18n)

## License

[Apache-2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
