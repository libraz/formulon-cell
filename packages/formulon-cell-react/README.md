# @libraz/formulon-cell-react

React 18+ component + hooks for [`@libraz/formulon-cell`](../formulon-cell/README.md).

## Install

```sh
npm install @libraz/formulon-cell-react @libraz/formulon-cell react react-dom zustand
```

## Quick start

```tsx
import { Spreadsheet, presets } from '@libraz/formulon-cell-react';
import '@libraz/formulon-cell/styles.css';

export function MySheet() {
  return (
    <Spreadsheet
      style={{ width: '100%', height: '100vh' }}
      features={presets.full()}
      locale="ja"
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

function App() {
  const ref = useRef<SpreadsheetRef>(null);
  return (
    <>
      <button onClick={() => ref.current?.instance?.undo()}>Undo</button>
      <Spreadsheet ref={ref} />
    </>
  );
}
```

## Hooks

```tsx
import { useI18n, useSelection } from '@libraz/formulon-cell-react';

const sel = useSelection(instance);
const { locale, strings } = useI18n(instance);
```

## Core helpers

The package re-exports the core command helpers and types, including
`createSessionChart`, `saveSheetView`, `activateSheetView`,
`listDefinedNames`, and `upsertDefinedName`, so React apps can type their
host chrome from one import.

## Runtime prop updates

`theme`, `locale`, `strings`, `workbook`, `features`, and `extensions` update
the running spreadsheet through the core imperative API. The component does
not re-mount the canvas when these props change, so selection, focus, and
host event subscriptions stay intact.

## License

[Apache-2.0](./LICENSE)
