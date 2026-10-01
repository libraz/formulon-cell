# @libraz/formulon-cell

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — react](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![npm — vue](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-6-blue?logo=typescript)](https://www.typescriptlang.org/)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-blue)](https://formulon.libraz.net/cell/)
[![Demo](https://img.shields.io/badge/demo-live-brightgreen)](https://formulon.libraz.net/cell/demo)

**A working spreadsheet for any web page, no framework required.** Mount it into a DOM element and you get a canvas-rendered grid with desktop-spreadsheet chrome (formula bar, ribbon, sheet tabs, context menu) over the [formulon](https://github.com/libraz/formulon) WASM calc engine. Features are composed from presets and extension factories, and locales and themes swap at runtime without re-mounting.

Try it in the [live demo](https://formulon.libraz.net/cell/demo).

This is the vanilla TypeScript / DOM core. [`@libraz/formulon-cell-react`](https://www.npmjs.com/package/@libraz/formulon-cell-react) and [`@libraz/formulon-cell-vue`](https://www.npmjs.com/package/@libraz/formulon-cell-vue) are thin adapters over the same mount call.

> Basic workbook workflows are ready to use. Dialogs, detailed control behaviour, keyboard interaction, and accessibility do not yet match desktop spreadsheets in every case, so do not present this as a finished end-user product yet.

## Install

```sh
npm install @libraz/formulon-cell zustand
```

`zustand` is a peer dependency. The default WASM engine is single-threaded and needs no COOP/COEP headers. `WorkbookHandle.createDefault()` rejects if the engine fails to initialize; the in-memory stub is used only when you pass `preferStub: true`. Bundler settings are in the [project README](https://github.com/libraz/formulon-cell#bundler-integration).

## Quick start

```ts
import { Spreadsheet, WorkbookHandle, presets } from '@libraz/formulon-cell';
import '@libraz/formulon-cell/styles.css';

const wb = await WorkbookHandle.createDefault();
const sheet = await Spreadsheet.mount(document.getElementById('sheet')!, {
  workbook: wb,
  features: presets.full(),
  locale: 'en',
  toolbar: true, // ribbon mounted in the same call, available as sheet.toolbar
});

sheet.i18n.setLocale('ja');
sheet.setTheme('ink'); // themes the grid and the ribbon together
```

For a grid-only form or viewer, mount with `ui: { profile: 'embedded' }` and a policy such as `fixedFormPolicy(ranges)` or `viewerPolicy()`. See [Embedding and use cases](https://formulon.libraz.net/cell/embedding).

## Presets

| preset | what's in it |
|--------|--------------|
| `presets.minimal()`  | formula bar, status bar, basic keymap |
| `presets.standard()` | + Quick Analysis, context menu, find/replace, clipboard, format painter |
| `presets.full()`     | + format dialog, paste special, conditional formatting, named ranges, hyperlink dialog, PivotTable creation, validation, autocomplete, hover comments |

## Subpath exports

| import path | contents |
|---|---|
| `@libraz/formulon-cell` | `Spreadsheet`, `WorkbookHandle`, `presets`, policies, extension factories |
| `@libraz/formulon-cell/extensions` | all extension factories |
| `@libraz/formulon-cell/extensions/*` | one extension (`statusBar`, `findReplace`, `contextMenu`, …) |
| `@libraz/formulon-cell/i18n/ja`, `/i18n/en` | locale dictionaries |
| `@libraz/formulon-cell/styles.css` | grid, dialogs, overlays, ribbon, and all themes in one file |
| `@libraz/formulon-cell/styles/{tokens,paper,ink,contrast,toolbar}.css` | individual stylesheets for granular setups |

## Key APIs

| API | description |
|-----|-------------|
| `Spreadsheet.mount(host, opts)` | mount the spreadsheet into a DOM element |
| `WorkbookHandle.createDefault()` | create a workbook backed by the WASM engine |
| `workbook.withBatchedRecalc(fn)` | run a multi-cell write with a single recalc |
| `instance.i18n.setLocale(loc)` / `instance.setTheme(theme)` | swap locale or theme at runtime |
| `instance.setPolicy` / `setViewportOptions` / `setContextMenu` / `setUi` / `setToolbar` | change embedding settings at runtime |
| `instance.applyChanges(changes)` | trusted host update that bypasses the user policy |
| `instance.on(event, handler)` | `cellChange`, `selectionChange`, `changeBatch`, `recalc`, … |

The full option and API reference is on the [documentation site](https://formulon.libraz.net/cell/).

## Documentation

- [Live demo](https://formulon.libraz.net/cell/demo)
- [Getting started](https://formulon.libraz.net/cell/install)
- [Options](https://formulon.libraz.net/cell/options)
- [Embedding and use cases](https://formulon.libraz.net/cell/embedding)
- [Modals and fullscreen](https://formulon.libraz.net/cell/modals)
- [React / Vue](https://formulon.libraz.net/cell/frameworks)
- [Theming](https://formulon.libraz.net/cell/theming)
- [Language and labels](https://formulon.libraz.net/cell/i18n)

## License

[Apache-2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
