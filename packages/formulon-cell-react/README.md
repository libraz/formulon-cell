# @libraz/formulon-cell-react

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![npm — core](https://img.shields.io/npm/v/@libraz/formulon-cell?label=core)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — vue](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![React](https://img.shields.io/badge/React-18%2B-blue?logo=react)](https://react.dev/)

**Drop a working spreadsheet into a React tree.** `<Spreadsheet>` and
`<SpreadsheetToolbar>` wrap
[`@libraz/formulon-cell`](https://www.npmjs.com/package/@libraz/formulon-cell) —
a canvas-rendered grid with desktop-spreadsheet chrome over the
[formulon](https://github.com/libraz/formulon) WASM calc engine. Props update
the running instance through the core imperative API instead of re-mounting the
canvas.

This package is a thin adapter: the grid, ribbon, menus, commands, and dialogs
all live in core, so React and Vue hosts cannot drift apart.

## Installation

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

## Also available

```sh
npm install @libraz/formulon-cell       # vanilla TypeScript / DOM core
npm install @libraz/formulon-cell-vue   # Vue 3 component + composables
```

## License

[Apache-2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
