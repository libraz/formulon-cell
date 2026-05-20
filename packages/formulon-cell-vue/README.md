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
compile it with the same pipeline as application components. It is a thin
adapter over core `Spreadsheet.mountToolbar`; the ribbon DOM, menu factories,
activation model, and dynamic dropdown dispatcher live in
`@libraz/formulon-cell`.

For host audits and custom chrome, use the core exports such as
`ribbonActivationEntries`, `ribbonSurfaceCommandIds`,
`DYNAMIC_RIBBON_DROPDOWN_HANDLER_ATTRS`, `attachRangePickerButton`,
`appendConditionalApplyFormatControls`, `conditionalStyleOptions`,
`showReport`, `reportDialogLabels`, `projectDisabledReason`, and `projectDisabledState`. Do not recreate ribbon command sets or
Excel-style dialog/report controls in Vue.

```vue
<script setup lang="ts">
import { type RibbonTab } from '@libraz/formulon-cell-vue';
import SpreadsheetToolbar from '@libraz/formulon-cell-vue/toolbar.vue';
import '@libraz/formulon-cell-vue/toolbar.css';
</script>
```

Use `dropdownActions` to override specific core dropdown handlers without
forking the ribbon.

## Runtime prop updates

`theme`, `locale`, `strings`, `workbook`, `features`, `extensions`,
`printerProfiles`, `printerProfileId`, `uploadStatus`, and `macroRecording`
update the running spreadsheet through the core imperative API. The component
does **not** re-mount the canvas, so selection, focus, and host event
subscriptions stay intact.

Host-only capabilities can be passed as props without reimplementing ribbon
behavior in Vue:

```vue
<template>
  <Spreadsheet
    :capture-screen-clip="captureScreenClip"
    :refresh-printer-profiles="refreshPrinterProfiles"
  />
</template>
```

`captureScreenClip` backs Insert > Screenshot > Screen Clipping. Printer
profile props feed Page Setup / print preview minimum-margin handling.

## Core helpers

This package re-exports core command helpers and types — `createSessionChart`,
`saveSheetView`, `activateSheetView`, `listDefinedNames`, `upsertDefinedName`,
`ribbonActivationEntries`, `attachRangePickerButton`,
`appendConditionalApplyFormatControls`, `conditionalStyleOptions`, `showReport`,
`reportDialogLabels`, `projectDisabledReason`, `projectDisabledState`, `ScreenClipCapture`, `ScreenClipResult`, etc. — so Vue
apps can type host chrome from a single import.

## Documentation

For the complete API reference and bundler integration notes, see the
[project README](https://github.com/libraz/formulon-cell).

## License

[Apache License 2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
