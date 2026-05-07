# @libraz/formulon-cell

Spreadsheet UI for the [formulon](https://github.com/libraz/formulon) WASM
calc engine.

- **Excel 365-flavored** chrome out of the box (formula bar, status bar,
  context menu, sheet tabs).
- **Canvas-rendered** grid with theme tokens — `paper` (light) and `ink`
  (dark) ship in the box; bring your own with the documented CSS variables.
- **Extension-based** API: every feature (find/replace, format dialog,
  paste-special, hyperlink dialog, hover comments, …) is an opt-in
  extension you compose into the mount call.
- **Runtime i18n** — swap locales without re-mounting; `ja` and `en` ship
  by default, register more at runtime.
- **Headless option** — keep just the canvas + store and provide your own
  chrome.

> **Status:** v0.1 — public API stabilizing. Until v1.0 minor bumps may
> reshape extension contracts. Pin a version range you can update on
> purpose.

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
require-corp`). Without it, formulon-cell falls back to an in-memory stub
engine — the UI keeps working, formulas degrade gracefully.

## Bundler integration (Vite, webpack, esbuild)

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

**3. Browser builds use the web-safe engine entry.** formulon-cell imports
the browser wrapper for the vendored engine, so Vite/webpack apps should not
see `node:*` externalization warnings from the spreadsheet package. Keep the
package out of dependency pre-bundling so the worker/WASM assets stay under
the app bundler's control:

```ts
// vite.config.ts
export default defineConfig({
  optimizeDeps: { exclude: ['@libraz/formulon-cell', '@libraz/formulon'] },
});
```

**4. SharedArrayBuffer requires cross-origin isolation.** Serve your page
with `Cross-Origin-Opener-Policy: same-origin` + `Cross-Origin-Embedder-Policy:
require-corp`. Without these headers, `SharedArrayBuffer` is undefined and
formulon-cell drops to an in-memory **stub engine**: the canvas, formula
bar, and editing affordances all keep working, but formula evaluation,
recalc, and xlsx round-trip degrade to no-ops. Detect at runtime via
`crossOriginIsolated` or via `isUsingStub()` after `WorkbookHandle.createDefault()`:

```ts
import { WorkbookHandle, isUsingStub } from '@libraz/formulon-cell';

const wb = await WorkbookHandle.createDefault();
if (isUsingStub()) {
  console.warn('formulon-cell: running on stub engine — recalc disabled');
}
```

## Quick start

```ts
import { Spreadsheet, WorkbookHandle, presets } from '@libraz/formulon-cell';
import '@libraz/formulon-cell/styles.css';

const host = document.getElementById('sheet')!;
const wb = await WorkbookHandle.createDefault();
const sheet = await Spreadsheet.mount(host, {
  workbook: wb,
  extensions: presets.excel(),
  locale: 'en',
});

sheet.i18n.setLocale('ja');     // runtime locale swap
sheet.setTheme('ink');           // dark mode
```

## Presets

| preset | what's in it |
|--------|--------------|
| `presets.minimal()`  | formula bar, status bar, basic keymap |
| `presets.standard()` | + context menu, find/replace, clipboard, format painter, wheel scroll |
| `presets.excel()`    | + format dialog, paste-special, conditional formatting, named ranges, hyperlink dialog, validation, autocomplete, hover comments, full Excel keymap |

Compose your own:

```ts
import {
  Spreadsheet,
  formulaBar, statusBar, contextMenu, findReplace, theme, i18n, keymap,
} from '@libraz/formulon-cell';

await Spreadsheet.mount(host, {
  extensions: [
    formulaBar({ expandable: true }),
    statusBar({ aggregations: ['sum', 'avg', 'count'] }),
    contextMenu(),
    findReplace(),
    keymap.basic,
    i18n({ locale: 'ja' }),
    theme('paper'),
  ],
});
```

## i18n

```ts
import { Spreadsheet } from '@libraz/formulon-cell';
import ja from '@libraz/formulon-cell/i18n/ja';
import en from '@libraz/formulon-cell/i18n/en';

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

## React / Vue

```sh
npm install @libraz/formulon-cell-react react react-dom
# or
npm install @libraz/formulon-cell-vue vue
```

See [`@libraz/formulon-cell-react`](../formulon-cell-react/README.md) and
[`@libraz/formulon-cell-vue`](../formulon-cell-vue/README.md).

## License

[Apache-2.0](./LICENSE)
