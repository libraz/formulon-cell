# formulon-cell

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell?label=%40libraz%2Fformulon-cell)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — react](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![npm — vue](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![codecov](https://codecov.io/gh/libraz/formulon-cell/branch/main/graph/badge.svg)](https://codecov.io/gh/libraz/formulon-cell)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-6-blue?logo=typescript)](https://www.typescriptlang.org/)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-blue)](https://formulon.libraz.net/cell/)
[![Demo](https://img.shields.io/badge/demo-live-brightgreen)](https://formulon.libraz.net/cell/demo)

**formulon-cell puts a working spreadsheet inside a web page.** Mount it into a DOM node, a React component, or a Vue component and you get a canvas-rendered grid with desktop-spreadsheet chrome (formula bar, ribbon, sheet tabs, context menu) over the [formulon](https://github.com/libraz/formulon) WASM calc engine. The core is framework-free TypeScript; the React and Vue packages are thin adapters over the same mount call.

Try it in the [live demo](https://formulon.libraz.net/cell/demo).

Reach for it when you need to:

- **Embed a spreadsheet, not a data grid.** Selection, editing, formula entry, recalculation, and file handling work from the first mount.
- **Run formulas in a real engine.** Evaluation goes to a C++ engine compiled to WebAssembly, not to a JavaScript parser bolted onto the UI.
- **Ship only the features you use.** Presets and extension factories decide what mounts. You can drop the chrome and keep only the canvas and the store.
- **Lock a sheet down for a form or a viewer.** Interaction policies limit which cells and operations a user may touch, independently of which UI is visible.

> **Compatibility status.** formulon-cell aims at desktop-spreadsheet (Excel-compatible) behaviour. Basic workbook workflows are ready to use, but dialogs, detailed control behaviour, keyboard interaction, and accessibility do not yet guarantee equivalent behaviour, and bugs may remain. Do not present it as a drop-in replacement or a finished end-user product yet.

## Packages

| package | what it is |
|---------|------------|
| [`@libraz/formulon-cell`](./packages/formulon-cell) | Vanilla TS / DOM core: grid, chrome, extensions, i18n, themes |
| [`@libraz/formulon-cell-react`](./packages/formulon-cell-react) | React 18+ component, hooks, and ribbon toolbar |
| [`@libraz/formulon-cell-vue`](./packages/formulon-cell-vue) | Vue 3 component, composables, and ribbon toolbar |

## Install

```sh
npm install @libraz/formulon-cell zustand
# React: npm install @libraz/formulon-cell-react @libraz/formulon-cell react react-dom zustand
# Vue:   npm install @libraz/formulon-cell-vue @libraz/formulon-cell vue zustand
```

`zustand` is a peer dependency so the host can read the same store the chrome subscribes to.

The default WASM engine in formulon 0.18.0 is single-threaded and needs neither COOP/COEP headers nor `SharedArrayBuffer`. `WorkbookHandle.createDefault()` rejects if the engine fails to initialize; there is no silent fallback. Node.js integrations require Node 22 or later. The in-memory stub engine is opt-in through `preferStub: true`, for tests and explicit demos only.

## Quick start

```ts
import { Spreadsheet, WorkbookHandle, presets } from '@libraz/formulon-cell';
import '@libraz/formulon-cell/styles.css';

const wb = await WorkbookHandle.createDefault();
const sheet = await Spreadsheet.mount(document.getElementById('sheet')!, {
  workbook: wb,
  features: { ...presets.full(), localeSwitcher: true },
  locale: 'en',
  toolbar: true,
});

sheet.setLocale('ja'); // swap locale without re-mounting
sheet.setTheme('ink');      // dark theme
```

Pass `getWorkbookName: () => currentFilename` to show the current workbook name in Watch Window. The React and Vue components accept the same callback.

The language selector is hidden by default. Enable it with `features.localeSwitcher: true`, or toggle it with `sheet.setFeatures({ localeSwitcher: false })`. `sheet.setLocale()` updates the dictionary and calculation profile together; it returns `false` when the change is rejected and keeps the current locale. Close open dialogs, Find/Replace, comment edits and filter menus before switching; their drafts are preserved when a switch is rejected.

React and Vue take the same options as component props:

```tsx
import { Spreadsheet, presets } from '@libraz/formulon-cell-react';
import '@libraz/formulon-cell/styles.css';

<Spreadsheet features={presets.full()} locale="en" style={{ height: '100vh' }} />;
```

```vue
<script setup lang="ts">
import { Spreadsheet, presets } from '@libraz/formulon-cell-vue';
import '@libraz/formulon-cell/styles.css';
</script>

<template>
  <Spreadsheet :features="presets.full()" locale="en" style="height: 100vh" />
</template>
```

## Restricted embedding

The `embedded` UI profile mounts only the grid, and a policy decides what users may change. Hiding the ribbon does not make a cell readonly; the policy does.

```ts
import { Spreadsheet, fixedFormPolicy } from '@libraz/formulon-cell';

const sheet = await Spreadsheet.mount(host, {
  ui: { profile: 'embedded', features: { clipboard: true, shortcuts: true } },
  policy: fixedFormPolicy([{ sheet: 0, r0: 1, c0: 1, r1: 9, c1: 2 }]),
  viewport: { range: { sheet: 0, r0: 0, c0: 0, r1: 9, c1: 3 } },
  contextMenu: { mode: 'builtIn', items: ['copy', 'paste', 'clear'] },
});
```

`fixedFormPolicy` allows value edits, clears, pastes, and fills inside the given ranges; `viewerPolicy()` makes the sheet readonly while keeping selection and copy. Every input path (editor, formula bar, clipboard, keyboard, pointer, context menu, ribbon) goes through the same check, and undo/redo re-check the current policy. Runtime setters, host updates, and the list of supported operations are covered in [Embedding and use cases](https://formulon.libraz.net/cell/embedding).

## Bundler integration

The engine ships its WASM through an Emscripten wrapper, so the bundler has to leave it alone. For Vite:

```ts
// vite.config.ts
export default defineConfig({
  build: { target: 'es2022' }, // the engine factory uses top-level await
  optimizeDeps: { exclude: ['@libraz/formulon-cell', '@libraz/formulon'] },
});
```

The `.wasm` file is fetched from next to the engine's JS. To serve it from a CDN or another path, create the workbook yourself and pass a resolver; the engine loads once per page, so the first load decides the URL:

```ts
const workbook = await WorkbookHandle.createDefault({
  locateFile: (path) => `https://cdn.example.com/formulon/${path}`,
});
await Spreadsheet.mount(host, { workbook });
```

Only a direct import of `@libraz/formulon/threads` needs `worker: { format: 'es' }` and cross-origin isolation (`Cross-Origin-Opener-Policy: same-origin`, `Cross-Origin-Embedder-Policy: require-corp`). The default formulon-cell loader starts no workers.

## What's inside

- Desktop-spreadsheet chrome: formula bar, status bar, ribbon, sheet tabs, context menu, and dialogs (format cells, paste special, find/replace, Go To, conditional formatting, named ranges, validation, PivotTable creation, and more).
- Formula editing as on the desktop: arrow keys and F4 pick and cycle references, Ctrl+arrow jumps to the data edge while pointing, Ctrl+Z undoes inside the draft, F9 evaluates the selected part, and the R1C1 reference style works throughout. A typed web address becomes a hyperlink when the cell is committed.
- A canvas-rendered grid themed through CSS tokens. `paper` (light), `ink` (dark), and `contrast` ship in the box.
- Charts (column, bar, line, area, pie, doughnut and scatter) with a settings panel. Charts in loaded workbooks are edited in place and saved back.
- XLSX persistence for threaded comments and notes, images, charts, shapes, sparklines, tables and table styles, PivotTables, slicers, workbook themes, filters, sheet and workbook protection, view settings and calculation options. Loaded pictures can be moved, resized and reordered unless they use an absolute anchor. Save As can also write `.xlsb`, with a warning that lists what the binary format drops.
- Feature flags and extension factories. `presets.minimal()`, `presets.standard()`, and `presets.full()` cover the common sets; any piece can be swapped or left out.
- An opt-in Mac profile. `ui: { platform: 'mac' }` (or `'auto'` to detect it from the browser) switches to a Mac ribbon tab set, Mac shortcuts, and a nonmodal formula argument palette.
- Runtime i18n. `en`, `ja`, `de`, `fr`, `zh`, `ko`, and `th` ship by default, and `sheet.i18n.register()` adds more without re-mounting.
- A headless mode that keeps only the canvas and the store, for hosts that bring their own chrome.

## Documentation

- [Live demo](https://formulon.libraz.net/cell/demo)
- [Getting started](https://formulon.libraz.net/cell/install)
- [Options](https://formulon.libraz.net/cell/options)
- [Embedding and use cases](https://formulon.libraz.net/cell/embedding)
- [Modals and fullscreen](https://formulon.libraz.net/cell/modals)
- [React / Vue](https://formulon.libraz.net/cell/frameworks)
- [Theming](https://formulon.libraz.net/cell/theming)
- [Language and labels](https://formulon.libraz.net/cell/i18n)

## Development

```sh
yarn install
yarn dev:react   # apps/react-demo
yarn dev:vue     # apps/vue-demo
yarn build
yarn test
```

Pull requests are welcome; [`CONTRIBUTING.md`](./CONTRIBUTING.md) lists the checks CI runs. Releases follow the manual tag-based flow in [`docs/releasing.md`](./docs/releasing.md).

## Non-goals

formulon-cell is the UI layer, not the calculation engine and not an application. Formula evaluation, recalculation, and xlsx parsing belong to [formulon](https://github.com/libraz/formulon). There is no server, persistence layer, collaboration transport, or authentication; the host owns those. The React and Vue packages stay thin, with menus, commands, and dialogs living in core so the three hosts cannot drift apart.

## License

[Apache-2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
