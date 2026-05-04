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
      features={presets.excel()}
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

## License

[Apache-2.0](./LICENSE)
