# @libraz/formulon-cell

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — react](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![npm — vue](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-6-blue?logo=typescript)](https://www.typescriptlang.org/)

**Put a working spreadsheet inside a web page, with no framework required.**
Mount it into a DOM element and you get a canvas-rendered grid with
desktop-spreadsheet chrome — formula bar, ribbon, sheet tabs, context menu —
over the [formulon](https://github.com/libraz/formulon) WASM calc engine, which
evaluates formulas off the main thread. Features are composed from presets and
extension factories, and locales swap at runtime without re-mounting.

This is the vanilla TypeScript / DOM core. For framework bindings use
[`@libraz/formulon-cell-react`](https://www.npmjs.com/package/@libraz/formulon-cell-react)
or [`@libraz/formulon-cell-vue`](https://www.npmjs.com/package/@libraz/formulon-cell-vue);
both are thin adapters over the same mount call.

> **Excel compatibility.** `formulon-cell` is being built toward an
> Excel-compatible spreadsheet experience while exercising
> [**formulon**](https://github.com/libraz/formulon) in a real browser. Its
> basic workbook workflows — selection, editing, formula entry,
> recalculation, and file handling — are ready to use. UI coverage is still
> incomplete: detailed control behaviour, dialogs, keyboard interaction,
> accessibility, and other UI/UX do not yet guarantee Excel-equivalent
> behaviour, and bugs may remain. Do not present it as a drop-in Excel
> replacement or a complete end-user spreadsheet product yet.

## Installation

```sh
npm install @libraz/formulon-cell zustand
```

`zustand` is a peer dependency. The WASM engine requires a
[crossOriginIsolated](https://developer.mozilla.org/docs/Web/API/crossOriginIsolated)
context (`COOP: same-origin` + `COEP: require-corp`); without it,
`WorkbookHandle.createDefault()` rejects before mounting. The in-memory stub
engine is opt-in via `preferStub: true` for tests and explicit demos.

See [bundler integration](https://github.com/libraz/formulon-cell#bundler-integration)
for Vite / webpack / esbuild setup notes.

## Quick start

```ts
import { Spreadsheet, WorkbookHandle, presets } from '@libraz/formulon-cell';
import '@libraz/formulon-cell/styles.css';

const host = document.getElementById('sheet')!;
const wb = await WorkbookHandle.createDefault();
const sheet = await Spreadsheet.mount(host, {
  workbook: wb,
  features: presets.full(),
  locale: 'en',
  toolbar: true,                 // ribbon in the same call → sheet.toolbar
});

sheet.i18n.setLocale('ja');     // runtime locale swap
sheet.setTheme('ink');           // dark mode — themes grid and toolbar together
```

## Host integrations

Browser APIs do not expose every desktop Excel integration point. Hosts can
provide those capabilities through `MountOptions` and keep the shared ribbon /
Backstage behavior in core:

```ts
const sheet = await Spreadsheet.mount(host, {
  workbook: wb,
  captureScreenClip: async () => ({
    src: await nativeCaptureRegionAsDataUrl(),
    alt: 'Screen clipping',
  }),
  printerProfiles: [
    {
      id: 'office-printer',
      name: 'Office Printer',
      paperSize: 'A4',
      orientation: 'portrait',
      printableBounds: { top: 0.16, right: 0.16, bottom: 0.16, left: 0.16 },
    },
  ],
  refreshPrinterProfiles: () => nativeListPrinterProfiles(),
  uploadStatus: 'saving',
  macroRecording: false,
});
```

`captureScreenClip` powers Insert > Screenshot > Screen Clipping. When omitted,
that command reports that native screen clipping is host-provided. Printer
profiles supply physical printer minimum margins for Page Setup / print
preview; browser-only hosts can omit them. `uploadStatus` accepts `saved`,
`saving`, `error`, or `null`; `macroRecording` accepts `true`, `false`, or
`null`.

## Presets

| preset | what's in it |
|--------|--------------|
| `presets.minimal()`  | formula bar, status bar, basic keymap |
| `presets.standard()` | + View toolbar, Quick Analysis, context menu, find/replace, clipboard, format painter |
| `presets.full()`     | + format dialog, paste-special, conditional formatting, named ranges, hyperlink dialog, PivotTable creation, validation, autocomplete, hover comments |

## Subpath exports

| Import path | Description |
|---|---|
| `@libraz/formulon-cell` | Core: `Spreadsheet`, `WorkbookHandle`, `presets`, extension factories |
| `@libraz/formulon-cell/extensions` | All extension factories (re-export) |
| `@libraz/formulon-cell/extensions/*` | Individual extensions (`statusBar`, `findReplace`, `contextMenu`, …) |
| `@libraz/formulon-cell/i18n/ja` | Japanese locale dictionary |
| `@libraz/formulon-cell/i18n/en` | English locale dictionary |
| `@libraz/formulon-cell/styles.css` | All-in-one bundle: grid, dialogs, overlays, ribbon toolbar, and all three palettes. This is the only stylesheet most embedders need |
| `@libraz/formulon-cell/styles/toolbar.css` | Ribbon styles only — for granular setups that don't want the all-in-one bundle |
| `@libraz/formulon-cell/styles/paper.css` | Paper (light) theme only |
| `@libraz/formulon-cell/styles/ink.css` | Ink (dark) theme only |
| `@libraz/formulon-cell/styles/contrast.css` | High-contrast theme only |
| `@libraz/formulon-cell/styles/tokens.css` | Theme tokens only |

## Key APIs

| API | Description |
|-----|-------------|
| `Spreadsheet.mount(host, opts)` | Mount the spreadsheet UI into a DOM element |
| `WorkbookHandle.createDefault()` | Create a workbook backed by the WASM engine |
| `isUsingStub()` | Detect whether the explicit stub engine is in use |
| `presets.{minimal,standard,full}()` | Built-in feature presets |
| `instance.i18n.setLocale(loc)` | Swap locale at runtime — no remount |
| `instance.setTheme(theme)` | Swap theme at runtime |
| `createSessionChart(store, range, options)` | Create session column/line chart overlays |
| `saveSheetView` / `activateSheetView` | Manage session Sheet Views |
| `listDefinedNames` / `upsertDefinedName` | Headless Name Manager API |
| `ribbonActivationEntries` / `ribbonSurfaceCommandIds` | Shared ribbon command manifests for host audits and wrapper parity |
| `attachRangePickerButton` | Shared dialog range-picker control used by Excel-style dialogs |
| `appendConditionalApplyFormatControls` / `conditionalStyleOptions` | Shared conditional-format rule UI helpers |
| `showReport` / `reportDialogLabels` | Shared report dialog and localized label mapping for host-dependent compatibility reports |
| `projectDisabledReason` / `projectDisabledState` | Shared disabled/read-only reason projection for aria, title, dataset attributes, and control state |

For the complete API reference, see the [project README](https://github.com/libraz/formulon-cell).

## Also available

| Package | Description |
|---------|-------------|
| [`@libraz/formulon-cell-react`](https://www.npmjs.com/package/@libraz/formulon-cell-react) | `<Spreadsheet>` React component + hooks + `SpreadsheetToolbar` ribbon |
| [`@libraz/formulon-cell-vue`](https://www.npmjs.com/package/@libraz/formulon-cell-vue) | `<Spreadsheet>` Vue component + composables + `SpreadsheetToolbar` ribbon |

React:

```tsx
import { SpreadsheetToolbar, type RibbonTab } from '@libraz/formulon-cell-react';
import '@libraz/formulon-cell-react/toolbar.css';
```

Vue:

```vue
<script setup lang="ts">
import { type RibbonTab } from '@libraz/formulon-cell-vue';
import SpreadsheetToolbar from '@libraz/formulon-cell-vue/toolbar.vue';
import '@libraz/formulon-cell-vue/toolbar.css';
</script>
```

Both wrappers are thin adapters over core `Spreadsheet.mountToolbar`; the
ribbon DOM, menu factories, activation model, and dynamic dropdown dispatcher
are shared by `@libraz/formulon-cell`.

## License

[Apache-2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
