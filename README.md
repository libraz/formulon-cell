# formulon-cell

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell?label=%40libraz%2Fformulon-cell)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — react](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![npm — vue](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![codecov](https://codecov.io/gh/libraz/formulon-cell/branch/main/graph/badge.svg)](https://codecov.io/gh/libraz/formulon-cell)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-6-blue?logo=typescript)](https://www.typescriptlang.org/)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-2563eb)](https://formulon.libraz.net/cell/)
[![Demo](https://img.shields.io/badge/demo-live-brightgreen)](https://formulon.libraz.net/cell/demo)

**formulon-cell is a spreadsheet component for web apps, built to match desktop Excel.** It mounts into a DOM node, a React component or a Vue component and renders a canvas grid with the formula bar, ribbon, sheet tabs, context menu and dialogs; formulas run in [formulon](https://github.com/libraz/formulon), a C++ calc engine compiled to WebAssembly. Cell rendering, dialogs, keyboard shortcuts and clipboard behaviour are compared against the Excel desktop app one feature at a time, and corrected where they differ.

🧪 **[Live demo](https://formulon.libraz.net/cell/demo)** &nbsp;·&nbsp; 📖 **[Documentation](https://formulon.libraz.net/cell/)** &nbsp;·&nbsp; **[Getting started](https://formulon.libraz.net/cell/install)**

> **Beta, pre-1.0.** The public API may still change between minor versions. Behaviour that has not been compared yet may differ from Excel; please report a divergence as an [issue](https://github.com/libraz/formulon-cell/issues).

## What's inside

- **Desktop spreadsheet chrome** — formula bar, status bar, ribbon, sheet tabs, context menu and dialogs. [Ribbon toolbar](https://formulon.libraz.net/cell/toolbar) · [Modals and dialogs](https://formulon.libraz.net/cell/modals)
- **Desktop formula editing** — F4 and arrow-key reference picking, Ctrl+arrow while pointing, undo inside the draft, F9 on a selection, R1C1 references.
- **Canvas grid with themes** — styled through CSS tokens; `paper`, `ink` and `contrast` ship in the box. [Theming](https://formulon.libraz.net/cell/theming)
- **Charts** — seven chart types with a settings panel, edited in place and saved back. [Charts](https://formulon.libraz.net/cell/charts)
- **XLSX persistence** — comments, images, charts, tables, PivotTables, slicers, protection and more survive a save; Save As can write `.xlsb`, and offers `.xlsx` instead when the binary writer cannot carry a feature. [Host integration](https://formulon.libraz.net/cell/host-integration)
- **Restricted embedding** — `fixedFormPolicy(ranges)` and `viewerPolicy()` limit what users may change on every input path, independently of which UI is shown. [Embedding and use cases](https://formulon.libraz.net/cell/embedding)
- **UI profiles and extensions** — `ui.profile` (`embedded`, `minimal`, `standard`, `full`), with any piece swappable or left out through extensions. [Extensions](https://formulon.libraz.net/cell/extensions)
- **Mac profile** — `ui: { platform: 'mac' }` switches to the Mac ribbon, shortcuts and argument palette. [Platform and Mac UI](https://formulon.libraz.net/cell/platform)
- **Runtime i18n** — seven languages built in, strings overridable through `instance.i18n.override()`. [Internationalization](https://formulon.libraz.net/cell/i18n)

## Installation

```sh
npm install @libraz/formulon-cell
npm install @libraz/formulon-cell-react @libraz/formulon-cell react react-dom  # React 18+
npm install @libraz/formulon-cell-vue @libraz/formulon-cell vue                # Vue 3
```

The default engine is single-threaded and needs neither COOP/COEP headers nor `SharedArrayBuffer`. `Workbook.createDefault()` rejects if the engine fails to initialize; there is no silent fallback, and the in-memory stub engine runs only when you pass `engine: 'stub'`. Node.js integrations require Node 22 or later.

### Bundler integration

The engine loads its WASM through an Emscripten wrapper, so the bundler has to leave it alone. For Vite:

```ts
// vite.config.ts
export default defineConfig({
  build: { target: 'es2022' }, // the engine factory uses top-level await
  optimizeDeps: { exclude: ['@libraz/formulon-cell', '@libraz/formulon'] },
});
```

Serving the `.wasm` from a CDN (`locateFile`) and the threaded engine build are covered in [Bundler setup](https://formulon.libraz.net/cell/bundler).

## Quick start

### JavaScript / TypeScript

```ts
import { mount, Workbook } from '@libraz/formulon-cell';
import { ribbon } from '@libraz/formulon-cell/toolbar';
import '@libraz/formulon-cell/styles.css';

const wb = await Workbook.createDefault();
const sheet = await mount(document.getElementById('sheet')!, {
  workbook: wb,
  ui: { profile: 'full', features: { localeSwitcher: true } },
  locale: 'en',
  toolbar: ribbon(),
});

sheet.update({ locale: 'ja' }); // swap locale without re-mounting
sheet.update({ theme: 'ink' }); // dark theme
```

→ [Options](https://formulon.libraz.net/cell/options)

### React

```tsx
import { Spreadsheet } from '@libraz/formulon-cell-react';
import '@libraz/formulon-cell/styles.css';

<Spreadsheet ui={{ profile: 'full' }} locale="en" style={{ height: '100vh' }} />;
```

→ [React / Vue](https://formulon.libraz.net/cell/frameworks)

### Vue

```vue
<script setup lang="ts">
import { Spreadsheet } from '@libraz/formulon-cell-vue';
import '@libraz/formulon-cell/styles.css';
</script>

<template>
  <Spreadsheet :ui="{ profile: 'full' }" locale="en" style="height: 100vh" />
</template>
```

→ [React / Vue](https://formulon.libraz.net/cell/frameworks)

## Build from source

```sh
yarn install
yarn dev:react   # React demo (apps/react-demo)
yarn dev:vue     # Vue demo (apps/vue-demo)
yarn build       # libraries
yarn test        # core unit tests
```

The checks CI runs are listed in [`CONTRIBUTING.md`](./CONTRIBUTING.md); releases follow [`docs/releasing.md`](./docs/releasing.md).

## Embedding

Three independent options decide what an embedded sheet shows and what users may do with it, and each one can be changed later without re-mounting (`update({ ui, policy, viewport })`).

```ts
import { mount, Workbook, fixedFormPolicy, viewerPolicy } from '@libraz/formulon-cell';

const sheet = await mount(host, {
  workbook: await Workbook.createDefault(),
  ui: { profile: 'embedded', features: { formulaBar: true } }, // grid plus formula bar
  policy: fixedFormPolicy(['B2:B11']), // only B2:B11 accept input, on every sheet
  viewport: { mode: 'fixed', range: 'A1:D12', tabNavigation: 'editable' },
});

sheet.update({ policy: viewerPolicy() }); // after submit: select and copy only
```

- **`ui`** — what is shown: a profile (`embedded`, `minimal`, `standard`, `full`) plus about thirty switches such as `sheetTabs`, `contextMenu`, `charts` or `clipboard`.
- **`policy`** — what users may change: 27 operations (value edit, paste, format, row insert, sort, ...) allowed or denied one by one, editable cells as A1 ranges (`'B2:B11'` applies on every sheet, `'Input!B2:B11'` on one) or a predicate, and a `restrict` hook that can deny but never grant. It applies to every input path, so hiding the ribbon is not what makes a cell readonly.
- **`viewport`** — which cells are visible and selectable, and whether Tab walks only the editable cells.
- **`contextMenu`** and **`overlays`** — pick or replace the right-click menu, and keep popups inside a host modal or fullscreen element.

Under a policy, ribbon commands that are not classified by operation stay disabled, so a restricted sheet is best shown without the ribbon or with host-built controls. See [Embedding and use cases](https://formulon.libraz.net/cell/embedding).

## Documentation

The documentation lives at [formulon.libraz.net/cell](https://formulon.libraz.net/cell/).

- **Learn** — [Live demo](https://formulon.libraz.net/cell/demo) · [Getting started](https://formulon.libraz.net/cell/install) · [Embedding and use cases](https://formulon.libraz.net/cell/embedding)
- **Guides** — [Ribbon toolbar](https://formulon.libraz.net/cell/toolbar) · [Modals and dialogs](https://formulon.libraz.net/cell/modals) · [Charts](https://formulon.libraz.net/cell/charts) · [Theming](https://formulon.libraz.net/cell/theming) · [Internationalization](https://formulon.libraz.net/cell/i18n) · [Language switching](https://formulon.libraz.net/cell/language-switching) · [Platform and Mac UI](https://formulon.libraz.net/cell/platform) · [React / Vue](https://formulon.libraz.net/cell/frameworks) · [Bundler setup](https://formulon.libraz.net/cell/bundler)
- **Reference** — [Options](https://formulon.libraz.net/cell/options) · [Extensions](https://formulon.libraz.net/cell/extensions) · [Host integration](https://formulon.libraz.net/cell/host-integration)

## Non-goals

formulon-cell is the UI layer, not the calculation engine and not an application. Formula evaluation, recalculation and xlsx parsing belong to [formulon](https://github.com/libraz/formulon). There is no server, persistence layer, collaboration transport or authentication; the host owns those. The React and Vue packages stay thin, with menus, commands and dialogs living in core so the three hosts cannot drift apart.

## License

[Apache-2.0](LICENSE)
