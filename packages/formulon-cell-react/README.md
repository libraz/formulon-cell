# @libraz/formulon-cell-react

[![npm version](https://img.shields.io/npm/v/@libraz/formulon-cell-react.svg)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![license](https://img.shields.io/npm/l/@libraz/formulon-cell-react.svg)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![bundle size](https://img.shields.io/bundlephobia/minzip/@libraz/formulon-cell-react)](https://bundlephobia.com/package/@libraz/formulon-cell-react)

React 18+ component + hooks for
[`@libraz/formulon-cell`](https://www.npmjs.com/package/@libraz/formulon-cell)
— the spreadsheet UI for the [formulon](https://github.com/libraz/formulon)
WASM calc engine.

## Install

```sh
npm install @libraz/formulon-cell-react @libraz/formulon-cell react react-dom zustand
```

## Quick Start

```tsx
import { Spreadsheet, presets } from '@libraz/formulon-cell-react';
import '@libraz/formulon-cell/styles.css';

export function MySheet() {
  return (
    <Spreadsheet
      style={{ width: '100%', height: '100vh' }}
      features={presets.full()}
      locale="en"
      onReady={(inst) => {
        console.log('mounted', inst.workbook.version);
      }}
    />
  );
}
```

## Imperative ref

```tsx
import { useRef } from 'react';
import { Spreadsheet, type SpreadsheetRef } from '@libraz/formulon-cell-react';

const ref = useRef<SpreadsheetRef>(null);
ref.current?.instance?.undo();
```

## Hooks

| Hook | Description |
|------|-------------|
| `useSelection(instance)` | Subscribe to the active selection |
| `useI18n(instance)` | Read current locale + strings, reactive to runtime swaps |

## Toolbar

`SpreadsheetToolbar` provides the ribbon chrome used by the React demo.

```tsx
import { SpreadsheetToolbar, type RibbonTab } from '@libraz/formulon-cell-react';
import '@libraz/formulon-cell-react/toolbar.css';
```

## Runtime prop updates

`theme`, `locale`, `strings`, `workbook`, `features`, and `extensions`
update the running spreadsheet through the core imperative API. The
component does **not** re-mount the canvas, so selection, focus, and
host event subscriptions stay intact.

## Core helpers

This package re-exports core command helpers and types — `createSessionChart`,
`saveSheetView`, `activateSheetView`, `listDefinedNames`,
`upsertDefinedName`, etc. — so React apps can type host chrome from a single
import.

## Documentation

For the complete API reference and bundler integration notes, see the
[project README](https://github.com/libraz/formulon-cell).

## License

[Apache License 2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
