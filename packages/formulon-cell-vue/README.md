# @libraz/formulon-cell-vue

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![npm — core](https://img.shields.io/npm/v/@libraz/formulon-cell?label=core)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — react](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![Vue](https://img.shields.io/badge/Vue-3-blue?logo=vuedotjs)](https://vuejs.org/)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-blue)](https://formulon.libraz.net/cell/)
[![Demo](https://img.shields.io/badge/demo-live-brightgreen)](https://formulon.libraz.net/cell/demo)

**A working spreadsheet in a Vue app.** `<Spreadsheet>` and `SpreadsheetToolbar` wrap [`@libraz/formulon-cell`](https://www.npmjs.com/package/@libraz/formulon-cell), a canvas-rendered grid with chrome built to match desktop Excel, over the [formulon](https://github.com/libraz/formulon) WASM calc engine. Prop changes go through the core imperative API, so the canvas is never re-mounted and selection, focus, and event subscriptions survive.

Try it in the [live demo](https://formulon.libraz.net/cell/demo).

The grid, ribbon, menus, commands, and dialogs all live in core; this package only binds them to Vue.

> **Beta, pre-1.0.** The UI is built to match desktop Excel and is compared against it one feature at a time. The public API may still change between minor versions, and behaviour that has not been compared yet may differ; please report a divergence as an [issue](https://github.com/libraz/formulon-cell/issues).

## Install

```sh
npm install @libraz/formulon-cell-vue @libraz/formulon-cell vue
```

Bundler settings are in the [project README](https://github.com/libraz/formulon-cell#bundler-integration).

## Quick start

```vue
<script setup lang="ts">
import { Spreadsheet, type SpreadsheetInstance } from '@libraz/formulon-cell-vue';
import '@libraz/formulon-cell/styles.css';

const ready = (instance: SpreadsheetInstance) => {
  console.log('mounted', instance.workbook.engine);
};
</script>

<template>
  <Spreadsheet
    :ui="{ profile: 'full' }"
    locale="en"
    style="width: 100%; height: 100vh"
    @ready="ready"
  />
</template>
```

## Composables

```ts
const sheetRef = ref<SpreadsheetExposed | null>(null);
const instance = computed(() => sheetRef.value?.spreadsheet ?? null);
const selection = useSelection(instance);
const { locale, strings } = useI18n(instance);
```

| composable | description |
|------------|-------------|
| `useSelection(instance)` | the active selection |
| `useSpreadsheet(instance, selector)` | a value derived from the instance, re-run after every change |
| `useI18n(instance)` | current locale and strings, updated on runtime swaps |
| `useSpreadsheetEvent(instance, event, handler)` | subscribe to `change`, `cell`, `recalc`, … |

## Ribbon toolbar

`SpreadsheetToolbar` adapts core `mountToolbar`. Pass `hooks` to supply host actions for ribbon commands without forking the ribbon.

```vue
<script setup lang="ts">
import { SpreadsheetToolbar } from '@libraz/formulon-cell-vue';
import '@libraz/formulon-cell/styles/toolbar.css';
</script>

<template>
  <SpreadsheetToolbar :spreadsheet="instance" v-model:active-tab="tab" />
</template>
```

## Embedding

`ui`, `policy`, `viewport`, `context-menu`, `overlays`, and `toolbar` props apply at runtime without re-mounting. `fixedFormPolicy` and `viewerPolicy` come from `@libraz/formulon-cell`, a policy applies to every input path, and each batch of changes is emitted as `change` with its `status`.

```vue
<Spreadsheet
  :ui="{ profile: 'embedded' }"
  :policy="fixedFormPolicy(['B2:C10'])"
  @change="save"
/>
```

See [Embedding and use cases](https://formulon.libraz.net/cell/embedding) for viewport limits, context-menu composition, and host updates. Under a policy, ribbon commands that are not classified by operation stay disabled, so a restricted sheet is best shown without the ribbon or with host-built controls.

## Documentation

- [Live demo](https://formulon.libraz.net/cell/demo)
- [React / Vue](https://formulon.libraz.net/cell/frameworks)
- [Options](https://formulon.libraz.net/cell/options)
- [Embedding and use cases](https://formulon.libraz.net/cell/embedding)
- [Modals and fullscreen](https://formulon.libraz.net/cell/modals)
- [Theming](https://formulon.libraz.net/cell/theming)
- [Language and labels](https://formulon.libraz.net/cell/i18n)
- [Changelog](https://github.com/libraz/formulon-cell/blob/main/packages/formulon-cell/CHANGELOG.md) (this package's changes are listed with the core's)

## License

[Apache-2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
