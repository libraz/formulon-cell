# formulon-cell

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell?label=%40libraz%2Fformulon-cell)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — react](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![npm — vue](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![codecov](https://codecov.io/gh/libraz/formulon-cell/branch/main/graph/badge.svg)](https://codecov.io/gh/libraz/formulon-cell)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-6-blue?logo=typescript)](https://www.typescriptlang.org/)

**formulon-cell puts a working spreadsheet inside a web page.** Mount it into a
DOM node, a React component, or a Vue component and you get a canvas-rendered
grid with desktop-spreadsheet chrome — formula bar, ribbon, sheet tabs, context
menu — over the [formulon](https://github.com/libraz/formulon) WASM calc engine,
which evaluates formulas off the main thread. The core is framework-free
TypeScript; the React and Vue packages are thin adapters over the same mount
call.

**Reach for it when you need to:**

- **Embed a spreadsheet, not a data grid** — selection, editing, formula entry, recalculation, and file handling are there from the first mount.
- **Run the calc in a real engine** — formulas go to the C++ engine compiled to WebAssembly, not to a JavaScript formula parser bolted onto the UI.
- **Ship only the features you use** — presets and extension factories decide what mounts; drop the chrome entirely and keep the canvas plus the store.
- **Switch languages at runtime** — `ja` and `en` ship in the box, and locales swap in place without re-mounting.

> **Excel compatibility.** `formulon-cell` is being built toward an
> Excel-compatible spreadsheet experience while exercising
> [**formulon**](https://github.com/libraz/formulon) in a real browser. Its
> basic workbook workflows — selection, editing, formula entry,
> recalculation, and file handling — are ready to use. UI coverage is still
> incomplete: detailed control behaviour, dialogs, keyboard interaction,
> accessibility, and other UI/UX do not yet guarantee Excel-equivalent
> behaviour, and bugs may remain. Do not present it as a drop-in Excel
> replacement or a complete end-user spreadsheet product yet.

## Packages

| package | npm | what it is |
|---------|-----|------------|
| [`@libraz/formulon-cell`](./packages/formulon-cell)             | [![npm](https://img.shields.io/npm/v/@libraz/formulon-cell?label=)](https://www.npmjs.com/package/@libraz/formulon-cell)             | Vanilla TS / DOM core |
| [`@libraz/formulon-cell-react`](./packages/formulon-cell-react) | [![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=)](https://www.npmjs.com/package/@libraz/formulon-cell-react) | React 18+ component, hooks, and ribbon toolbar |
| [`@libraz/formulon-cell-vue`](./packages/formulon-cell-vue)     | [![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)     | Vue 3 component, composables, and ribbon toolbar |

## Installation

```sh
npm install @libraz/formulon-cell zustand
# or yarn / pnpm
```

`zustand` is a peer dependency — exposed because consumers can read from
the same store the chrome subscribes to.

The default WASM engine in formulon 0.12.0 is single-threaded and loads without COOP/COEP headers or `SharedArrayBuffer`. `WorkbookHandle.createDefault()` rejects if initialization fails. The in-memory stub engine is opt-in via `preferStub: true` for tests and explicit demos.

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
});

sheet.i18n.setLocale('ja');     // runtime locale swap
sheet.setTheme('ink');           // dark mode
```

## Bundler integration

formulon-cell uses the default single-threaded WASM from `@libraz/formulon`. Configure the bundler to resolve its WASM assets.

**1. Worker settings apply to the threaded entry.** If your app imports `@libraz/formulon/threads` directly, set `worker: { format: 'es' }` in Vite. The default formulon-cell loader starts no workers.

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
`@libraz/formulon`, whose Emscripten wrapper owns the WASM asset
resolution. Keep both packages out of dependency pre-bundling so those
assets stay under the app bundler's control:

```ts
// vite.config.ts
export default defineConfig({
  optimizeDeps: { exclude: ['@libraz/formulon-cell', '@libraz/formulon'] },
});
```

**4. The default WASM needs no cross-origin isolation.** Only a direct import of `@libraz/formulon/threads` requires `Cross-Origin-Opener-Policy: same-origin` and `Cross-Origin-Embedder-Policy: require-corp`. The stub is reserved for tests and explicit demos via `preferStub: true`; formula evaluation, recalc, and xlsx round-trip are incomplete there.

```ts
import { WorkbookHandle, isUsingStub } from '@libraz/formulon-cell';

const wb = await WorkbookHandle.createDefault();
if (isUsingStub()) {
  console.warn('formulon-cell: explicit stub engine selected');
}
```

## What's inside

- **Desktop-spreadsheet-style** chrome out of the box (formula bar, status bar,
  context menu, sheet tabs, ribbon).
- **Canvas-rendered** grid with theme tokens — `paper` (light) and `ink`
  (dark) ship in the box; bring your own with the documented CSS variables.
- **Extension-based** API: built-ins are controlled with feature flags, and
  replaceable pieces (find/replace, format dialog, paste-special, hyperlink
  dialog, hover comments, Quick Analysis, PivotTable creation,
  …) are available as extension factories you can compose into the mount call.
- **Runtime i18n** — swap locales without re-mounting; `ja` and `en` ship
  by default, register more at runtime.
- **Headless option** — keep just the canvas + store and provide your own
  chrome.

### Presets

| preset | what's in it |
|--------|--------------|
| `presets.minimal()`  | formula bar, status bar, basic keymap |
| `presets.standard()` | + Quick Analysis, session chart overlays, workbook object inspector, context menu, find/replace, clipboard, format painter, wheel scroll |
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
| `apps/react-demo`  | `yarn dev` / `yarn dev:react` | Same surface as `<Spreadsheet>` React component |
| `apps/vue-demo`    | `yarn dev:vue`    | Same surface as `<Spreadsheet>` Vue component |

## Framework ribbon toolbars

The React and Vue packages publish thin adapters over the core
`Spreadsheet.mountToolbar` ribbon. The ribbon DOM, menu factories, activation
model, and dynamic dropdown dispatcher all live in `@libraz/formulon-cell`, so
framework wrappers do not carry separate menu implementations. Import the
matching toolbar CSS alongside the component.

Hosts that need to audit or extend the ribbon should import the shared
manifests from core (`ribbonActivationEntries`, `ribbonSurfaceCommandIds`,
`DYNAMIC_RIBBON_DROPDOWN_HANDLER_ATTRS`) instead of reconstructing command
sets in React or Vue. Dialog-level helpers such as `attachRangePickerButton`,
`appendConditionalApplyFormatControls`, `conditionalStyleOptions`,
`showReport`, `reportDialogLabels`, `projectDisabledReason`, and `projectDisabledState` are also published from core for host UI
that must stay aligned with the built-in Excel 365-style dialogs and reports.

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

## Non-goals

formulon-cell is the UI layer, not the calculation engine and not an
application. Formula evaluation, recalculation, and xlsx parsing belong to
[formulon](https://github.com/libraz/formulon); this repository does not
reimplement them. It ships no server, no persistence layer, no collaboration
transport, and no authentication — the host owns all of those. It also refuses
to degrade quietly: without cross-origin isolation the mount fails rather than
falling back to the in-memory stub. The React and Vue packages stay thin
adapters, with menus, commands, and dialogs living in core, so the three hosts
cannot drift apart.

## Releasing

See [`docs/releasing.md`](./docs/releasing.md) for the manual tag-based
release flow.

## License

[Apache-2.0](LICENSE)
