# @libraz/formulon-cell-react

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![npm — core](https://img.shields.io/npm/v/@libraz/formulon-cell?label=core)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — vue](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![React](https://img.shields.io/badge/React-18%2B-blue?logo=react)](https://react.dev/)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-blue)](https://formulon.libraz.net/cell/)
[![Demo](https://img.shields.io/badge/demo-live-brightgreen)](https://formulon.libraz.net/cell/demo)

**An Excel-like spreadsheet component for React.** `<Spreadsheet>` and `<SpreadsheetToolbar>` wrap [`@libraz/formulon-cell`](https://www.npmjs.com/package/@libraz/formulon-cell), a canvas-rendered grid with chrome built to match desktop Excel, over the [formulon](https://github.com/libraz/formulon) WASM calc engine. Prop changes go through the core imperative API, so the canvas is never re-mounted and selection, focus, and event subscriptions survive.

Try it in the [live demo](https://formulon.libraz.net/cell/demo).

The grid, ribbon, menus, commands, and dialogs all live in core; this package only binds them to React.

> **Beta, pre-1.0.** The UI is built to match desktop Excel and is compared against it one feature at a time. The public API may still change between minor versions, and behaviour that has not been compared yet may differ; please report a divergence as an [issue](https://github.com/libraz/formulon-cell/issues).

## Install

```sh
npm install @libraz/formulon-cell-react @libraz/formulon-cell react react-dom
```

Bundler settings are in the [project README](https://github.com/libraz/formulon-cell#bundler-integration).

## Quick start

```tsx
import { Spreadsheet } from '@libraz/formulon-cell-react';
import '@libraz/formulon-cell/styles.css';

export function MySheet() {
  return (
    <Spreadsheet
      style={{ width: '100%', height: '100vh' }}
      ui={{ profile: 'full' }}
      locale="en"
      onReady={(instance) => console.log('mounted', instance.workbook.engine)}
    />
  );
}
```

`ref` exposes the running instance:

```tsx
const ref = useRef<SpreadsheetRef>(null);
ref.current?.spreadsheet?.history.undo();
```

## Hooks

| hook | description |
|------|-------------|
| `useSelection(instance)` | the active selection |
| `useSpreadsheet(instance, selector)` | a value derived from the instance, re-run after every change |
| `useI18n(instance)` | current locale and strings, updated on runtime swaps |
| `useSpreadsheetEvent(instance, event, handler)` | subscribe to `change`, `cell`, `recalc`, … |

## Ribbon toolbar

`SpreadsheetToolbar` adapts core `mountToolbar`. Use `hooks` to supply host actions for ribbon commands without forking the ribbon.

```tsx
import { SpreadsheetToolbar } from '@libraz/formulon-cell-react';
import type { RibbonTab } from '@libraz/formulon-cell/toolbar';
import '@libraz/formulon-cell/styles/toolbar.css';

<SpreadsheetToolbar
  spreadsheet={instance}
  activeTab={tab}
  onTabChange={setTab}
  hooks={{ protection: { runSheet: openProtectDialog } }}
/>;
```

## Embedding

`ui`, `policy`, `viewport`, `contextMenu`, `overlays`, and `toolbar` props apply at runtime without re-mounting. `fixedFormPolicy` and `viewerPolicy` come from `@libraz/formulon-cell`, a policy applies to every input path, and each batch of changes arrives through `onChange` with its `status`.

```tsx
<Spreadsheet
  ui={{ profile: 'embedded' }}
  policy={fixedFormPolicy(['B2:C10'])}
  onChange={(change) => save(change)}
/>
```

See [Embedding and use cases](https://formulon.libraz.net/cell/embedding) for viewport limits, context-menu composition, and host updates. Under a policy, ribbon commands that are not classified by operation stay disabled, so a restricted sheet is best shown without the ribbon or with host-built controls.

## Documentation

- [Live demo](https://formulon.libraz.net/cell/demo)
- [React / Vue](https://formulon.libraz.net/cell/frameworks)
- [Options](https://formulon.libraz.net/cell/options)
- [Embedding and use cases](https://formulon.libraz.net/cell/embedding)
- [Modals and fullscreen](https://formulon.libraz.net/cell/modals)
- [Theming](https://formulon.libraz.net/cell/theming)
- [Language and labels](https://formulon.libraz.net/cell/i18n)
- [Changelog](https://github.com/libraz/formulon-cell/blob/main/packages/formulon-cell/CHANGELOG.md) (this package's changes are listed with the core's)
- **Related** — [Formulon](https://formulon.libraz.net/): Excel 365 calculation engine (WASM, Python, native) · [formulon-mcp](https://formulon.libraz.net/mcp/): MCP server that gives AI agents workbook tools

## License

[Apache-2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
