# @libraz/formulon-cell-react

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![npm — core](https://img.shields.io/npm/v/@libraz/formulon-cell?label=core)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — vue](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![React](https://img.shields.io/badge/React-18%2B-blue?logo=react)](https://react.dev/)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-blue)](https://formulon.libraz.net/cell/)
[![Demo](https://img.shields.io/badge/demo-live-brightgreen)](https://formulon.libraz.net/cell/demo)

**A working spreadsheet in a React tree.** `<Spreadsheet>` and `<SpreadsheetToolbar>` wrap [`@libraz/formulon-cell`](https://www.npmjs.com/package/@libraz/formulon-cell), a canvas-rendered grid with desktop-spreadsheet chrome over the [formulon](https://github.com/libraz/formulon) WASM calc engine. Prop changes go through the core imperative API, so the canvas is never re-mounted and selection, focus, and event subscriptions survive.

Try it in the [live demo](https://formulon.libraz.net/cell/demo).

The grid, ribbon, menus, commands, and dialogs all live in core; this package only binds them to React.

## Install

```sh
npm install @libraz/formulon-cell-react @libraz/formulon-cell react react-dom zustand
```

Bundler settings are in the [project README](https://github.com/libraz/formulon-cell#bundler-integration).

## Quick start

```tsx
import { Spreadsheet, presets } from '@libraz/formulon-cell-react';
import '@libraz/formulon-cell/styles.css';

export function MySheet() {
  return (
    <Spreadsheet
      style={{ width: '100%', height: '100vh' }}
      features={presets.full()}
      locale="en"
      onReady={(instance) => console.log('mounted', instance.workbook.version)}
    />
  );
}
```

`ref` exposes the running instance:

```tsx
const ref = useRef<SpreadsheetRef>(null);
ref.current?.instance?.undo();
```

## Hooks

| hook | description |
|------|-------------|
| `useSelection(instance)` | the active selection |
| `useSpreadsheet(instance, selector, fallback)` | any slice of the store state |
| `useI18n(instance)` | current locale and strings, updated on runtime swaps |
| `useSpreadsheetEvent(instance, event, handler)` | subscribe to `cellChange`, `selectionChange`, `recalc`, … |

## Ribbon toolbar

`SpreadsheetToolbar` adapts core `Spreadsheet.mountToolbar`. Use `dropdownActions` to replace individual dropdown handlers without forking the ribbon.

```tsx
import { SpreadsheetToolbar, type RibbonTab } from '@libraz/formulon-cell-react';
import '@libraz/formulon-cell-react/toolbar.css';

<SpreadsheetToolbar
  instance={instance}
  activeTab={tab}
  onTabChange={setTab}
  locale="en"
  dropdownActions={{ applyProtectAction: openProtectDialog }}
/>;
```

## Restricted embedding

`ui`, `policy`, `viewport`, `contextMenu`, `overlays`, and `toolbar` props apply at runtime without re-mounting. `fixedFormPolicy` and `viewerPolicy` are re-exported, and successful batch updates arrive through `onChangeBatch`.

```tsx
<Spreadsheet
  ui={{ profile: 'embedded' }}
  policy={fixedFormPolicy([{ sheet: 0, r0: 1, c0: 1, r1: 9, c1: 2 }])}
  onChangeBatch={(batch) => save(batch)}
/>
```

See [Embedding and use cases](https://formulon.libraz.net/cell/embedding) for viewport limits, context-menu composition, and host updates.

## Documentation

- [Live demo](https://formulon.libraz.net/cell/demo)
- [React / Vue](https://formulon.libraz.net/cell/frameworks)
- [Options](https://formulon.libraz.net/cell/options)
- [Embedding and use cases](https://formulon.libraz.net/cell/embedding)
- [Modals and fullscreen](https://formulon.libraz.net/cell/modals)
- [Theming](https://formulon.libraz.net/cell/theming)
- [Language and labels](https://formulon.libraz.net/cell/i18n)

## License

[Apache-2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
