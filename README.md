# formulon-cell

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![codecov](https://codecov.io/gh/libraz/formulon-cell/branch/main/graph/badge.svg)](https://codecov.io/gh/libraz/formulon-cell)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell?label=%40libraz%2Fformulon-cell)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — react](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![npm — vue](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-6-blue?logo=typescript)](https://www.typescriptlang.org/)

Spreadsheet UI library for the [formulon](https://github.com/libraz/formulon)
WASM calc engine. Desktop-spreadsheet-style chrome, canvas-rendered grid,
extension-based feature composition, runtime i18n.

> **β (beta).** `formulon-cell` is built primarily as a demonstration host
> for [**formulon**](https://github.com/libraz/formulon) — a headless,
> Excel-compatible calculation engine in C++17 that ships a single WASM /
> Python / CLI core. Engine docs live at
> [formulon.libraz.net](https://formulon.libraz.net). The UI surface is
> still evolving; pin a version range you can upgrade on purpose.

## Packages

| package | npm | what it is |
|---------|-----|------------|
| [`@libraz/formulon-cell`](./packages/formulon-cell)             | [![npm](https://img.shields.io/npm/v/@libraz/formulon-cell?label=)](https://www.npmjs.com/package/@libraz/formulon-cell)             | Vanilla TS / DOM core |
| [`@libraz/formulon-cell-react`](./packages/formulon-cell-react) | [![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=)](https://www.npmjs.com/package/@libraz/formulon-cell-react) | React 18+ component, hooks, and ribbon toolbar |
| [`@libraz/formulon-cell-vue`](./packages/formulon-cell-vue)     | [![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)     | Vue 3 component, composables, and ribbon toolbar |

## Install

```sh
npm install @libraz/formulon-cell zustand
# or yarn / pnpm
```

`zustand` is a peer dependency — exposed because consumers can read from
the same store the chrome subscribes to.

The WASM engine ships pthread-enabled and requires a
[crossOriginIsolated context](https://developer.mozilla.org/docs/Web/API/crossOriginIsolated)
(`Cross-Origin-Opener-Policy: same-origin` + `Cross-Origin-Embedder-Policy:
require-corp`). Without it, `WorkbookHandle.createDefault()` rejects before
mounting so a host configuration issue cannot masquerade as a working
spreadsheet. The in-memory stub engine is opt-in via `preferStub: true` for
tests and explicit demos.

## Quick Start

```ts
import { Spreadsheet, WorkbookHandle, presets } from '@libraz/formulon-cell';
import '@libraz/formulon-cell/styles.css';

const host = document.getElementById('sheet')!;
const wb = await WorkbookHandle.createDefault();
const sheet = await Spreadsheet.mount(host, {
  workbook: wb,
  features: presets.full(),
  locale: 'en',
});

sheet.i18n.setLocale('ja');     // runtime locale swap
sheet.setTheme('ink');           // dark mode
```

## Bundler integration

formulon-cell re-uses `@libraz/formulon`'s pthread-enabled WASM module, so
the bundler hygiene rules from the engine package apply here too. Four
things matter:

**1. Workers must ship as ES modules.** The recalc scheduler runs on Web
Workers spawned by Emscripten with
`new Worker(new URL(...), { type: 'module' })`. Bundlers default to
classic (IIFE) workers and must be told otherwise:

```ts
// vite.config.ts
export default defineConfig({
  worker: { format: 'es' },
});
```

webpack 5 picks up `{ type: 'module' }` automatically when
`output.module: true`. esbuild needs `--format=esm` for the worker chunk.

**2. Top-level await + dynamic node imports need an es2022 target.** The
engine factory uses TLA and conditional `await import('node:...')`. Lift
both the main and worker target:

```ts
// vite.config.ts
export default defineConfig({
  build: { target: 'es2022' },
});
```

**3. Keep the engine out of dependency pre-bundling.** formulon-cell imports
`@libraz/formulon`, whose Emscripten wrapper owns the worker/WASM asset
resolution. Keep both packages out of dependency pre-bundling so those
assets stay under the app bundler's control:

```ts
// vite.config.ts
export default defineConfig({
  optimizeDeps: { exclude: ['@libraz/formulon-cell', '@libraz/formulon'] },
});
```

**4. SharedArrayBuffer requires cross-origin isolation.** Serve your page
with `Cross-Origin-Opener-Policy: same-origin` + `Cross-Origin-Embedder-Policy:
require-corp`. Without these headers, `SharedArrayBuffer` is undefined and
`WorkbookHandle.createDefault()` rejects instead of silently falling back to
the in-memory **stub engine**. The stub is reserved for tests and explicit
demos via `preferStub: true`, because formula evaluation, recalc, and xlsx
round-trip are intentionally incomplete there.

```ts
import { WorkbookHandle, isUsingStub } from '@libraz/formulon-cell';

const wb = await WorkbookHandle.createDefault();
if (isUsingStub()) {
  console.warn('formulon-cell: explicit stub engine selected');
}
```

## Features

- **Desktop-spreadsheet-style** chrome out of the box (formula bar, status bar,
  context menu, sheet tabs, View toolbar).
- **Canvas-rendered** grid with theme tokens — `paper` (light) and `ink`
  (dark) ship in the box; bring your own with the documented CSS variables.
- **Extension-based** API: built-ins are controlled with feature flags, and
  replaceable pieces (find/replace, format dialog, paste-special, hyperlink
  dialog, hover comments, View toolbar, Quick Analysis, PivotTable creation,
  …) are available as extension factories you can compose into the mount call.
- **Runtime i18n** — swap locales without re-mounting; `ja` and `en` ship
  by default, register more at runtime.
- **Headless option** — keep just the canvas + store and provide your own
  chrome.

### Presets

| preset | what's in it |
|--------|--------------|
| `presets.minimal()`  | formula bar, status bar, basic keymap |
| `presets.standard()` | + View toolbar, Quick Analysis, session chart overlays, workbook object inspector, context menu, find/replace, clipboard, format painter, wheel scroll |
| `presets.full()`     | + format dialog, paste-special, conditional formatting, iterative calculation settings, Go To Special, page setup, named ranges, hyperlink dialog, PivotTable creation, validation, autocomplete, hover comments, spreadsheet keymap |

### i18n

```ts
import { Spreadsheet } from '@libraz/formulon-cell';

const sheet = await Spreadsheet.mount(host, { locale: 'en' });

// Swap locale at runtime — every label updates in place.
sheet.i18n.setLocale('ja');

// Override a few strings without forking the dictionary.
sheet.i18n.extend('ja', { contextMenu: { copy: 'コピーする' } });

// Register a brand new locale.
import fr from './fr.js';
sheet.i18n.register('fr', fr);
sheet.i18n.setLocale('fr');
```

## Demo apps

| app | run | what it shows |
|-----|-----|---------------|
| `apps/playground`  | `yarn dev`        | Vanilla DOM playground (spreadsheet keymap) |
| `apps/react-demo`  | `yarn dev:react`  | Same surface as `<Spreadsheet>` React component |
| `apps/vue-demo`    | `yarn dev:vue`    | Same surface as `<Spreadsheet>` Vue component |

## Framework Ribbon Toolbars

The React and Vue packages publish the demo ribbon as reusable framework
chrome. Both implementations expose the same ribbon tab model and command
surface; import the matching toolbar CSS alongside the component.

```tsx
import { SpreadsheetToolbar, type RibbonTab } from '@libraz/formulon-cell-react';
import '@libraz/formulon-cell-react/toolbar.css';
```

```vue
<script setup lang="ts">
import { type RibbonTab } from '@libraz/formulon-cell-vue';
import SpreadsheetToolbar from '@libraz/formulon-cell-vue/toolbar.vue';
import '@libraz/formulon-cell-vue/toolbar.css';
</script>
```

## Releasing

See [`docs/releasing.md`](./docs/releasing.md) for the manual tag-based
release flow.

## License

[Apache-2.0](LICENSE)
