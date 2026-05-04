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
