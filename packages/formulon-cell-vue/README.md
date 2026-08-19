# @libraz/formulon-cell-vue

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![npm — core](https://img.shields.io/npm/v/@libraz/formulon-cell?label=core)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — react](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![Vue](https://img.shields.io/badge/Vue-3-blue?logo=vuedotjs)](https://vuejs.org/)

**Drop a working spreadsheet into a Vue app.** `<Spreadsheet>` and
`SpreadsheetToolbar` wrap
[`@libraz/formulon-cell`](https://www.npmjs.com/package/@libraz/formulon-cell) —
a canvas-rendered grid with desktop-spreadsheet chrome over the
[formulon](https://github.com/libraz/formulon) WASM calc engine. Props update
the running instance through the core imperative API instead of re-mounting the
canvas.

This package is a thin adapter: the grid, ribbon, menus, commands, and dialogs
all live in core, so Vue and React hosts cannot drift apart.

## Installation

```sh
npm install @libraz/formulon-cell-vue @libraz/formulon-cell vue zustand
```

## Quick start

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

## Also available

```sh
npm install @libraz/formulon-cell         # vanilla TypeScript / DOM core
npm install @libraz/formulon-cell-react   # React 18+ component + hooks
```

## License

[Apache-2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
