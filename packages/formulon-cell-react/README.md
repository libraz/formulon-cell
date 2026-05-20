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

`SpreadsheetToolbar` is a thin adapter over core `Spreadsheet.mountToolbar`.
The ribbon DOM, menu factories, activation model, and dynamic dropdown
dispatcher live in `@libraz/formulon-cell`, so React does not carry a separate
ribbon implementation.

For host audits and custom chrome, use the core exports such as
`ribbonActivationEntries`, `ribbonSurfaceCommandIds`,
`DYNAMIC_RIBBON_DROPDOWN_HANDLER_ATTRS`, `attachRangePickerButton`,
`appendConditionalApplyFormatControls`, `conditionalStyleOptions`,
`showReport`, `reportDialogLabels`, `projectDisabledReason`, and `projectDisabledState`. Do not recreate ribbon command sets or
Excel-style dialog/report controls in React.

```tsx
import { SpreadsheetToolbar, type RibbonTab } from '@libraz/formulon-cell-react';
import '@libraz/formulon-cell-react/toolbar.css';
```

Use `dropdownActions` to override specific core dropdown handlers without
forking the ribbon:

```tsx
<SpreadsheetToolbar
  instance={instance}
  activeTab="home"
  locale="en"
  onTabChange={setActiveTab}
  dropdownActions={{ applyProtectAction: openProtectDialog }}
/>
```

## Runtime prop updates

`theme`, `locale`, `strings`, `workbook`, `features`, `extensions`,
`printerProfiles`, `printerProfileId`, `uploadStatus`, and `macroRecording`
update the running spreadsheet through the core imperative API. The component
does **not** re-mount the canvas, so selection, focus, and host event
subscriptions stay intact.

Host-only capabilities can be passed as props without reimplementing ribbon
behavior in React:

```tsx
<Spreadsheet
  captureScreenClip={async () => ({
    src: await nativeCaptureRegionAsDataUrl(),
    alt: 'Screen clipping',
  })}
  refreshPrinterProfiles={() => nativeListPrinterProfiles()}
/>
```

`captureScreenClip` backs Insert > Screenshot > Screen Clipping. Printer
profile props feed Page Setup / print preview minimum-margin handling.

## Core helpers

This package re-exports core command helpers and types — `createSessionChart`,
`saveSheetView`, `activateSheetView`, `listDefinedNames`, `upsertDefinedName`,
`ribbonActivationEntries`, `attachRangePickerButton`,
`appendConditionalApplyFormatControls`, `conditionalStyleOptions`, `showReport`,
`reportDialogLabels`, `projectDisabledReason`, `projectDisabledState`, `ScreenClipCapture`, `ScreenClipResult`, etc. — so React
apps can type host chrome from a single import.

## Documentation

For the complete API reference and bundler integration notes, see the
[project README](https://github.com/libraz/formulon-cell).

## License

[Apache License 2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
