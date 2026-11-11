# @libraz/formulon-cell

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — react](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![npm — vue](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-6-blue?logo=typescript)](https://www.typescriptlang.org/)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-blue)](https://formulon.libraz.net/cell/)
[![Demo](https://img.shields.io/badge/demo-live-brightgreen)](https://formulon.libraz.net/cell/demo)

**A working spreadsheet for any web page, no framework required.** Mount it into a DOM element and you get a canvas-rendered grid with chrome built to match desktop Excel (formula bar, ribbon, sheet tabs, context menu) over the [formulon](https://github.com/libraz/formulon) WASM calc engine. Features are composed from UI profiles and extension factories, and locales and themes swap at runtime without re-mounting.

Try it in the [live demo](https://formulon.libraz.net/cell/demo).

This is the vanilla TypeScript / DOM core. [`@libraz/formulon-cell-react`](https://www.npmjs.com/package/@libraz/formulon-cell-react) and [`@libraz/formulon-cell-vue`](https://www.npmjs.com/package/@libraz/formulon-cell-vue) are thin adapters over the same mount call.

> **Beta, pre-1.0.** The UI is built to match desktop Excel and is compared against it one feature at a time. The public API may still change between minor versions, and behaviour that has not been compared yet may differ; please report a divergence as an [issue](https://github.com/libraz/formulon-cell/issues).

## Install

```sh
npm install @libraz/formulon-cell
```

The default WASM engine is single-threaded and needs no COOP/COEP headers. `Workbook.createDefault()` rejects if the engine fails to initialize; the in-memory stub is used only when you pass `engine: 'stub'`. Bundler settings are in the [project README](https://github.com/libraz/formulon-cell#bundler-integration).

## Quick start

```ts
import { mount, Workbook } from '@libraz/formulon-cell';
import { ribbon } from '@libraz/formulon-cell/toolbar';
import '@libraz/formulon-cell/styles.css';

const wb = await Workbook.createDefault();
const sheet = await mount(document.getElementById('sheet')!, {
  workbook: wb,
  ui: { profile: 'full', features: { localeSwitcher: true } },
  locale: 'en',
  toolbar: ribbon(), // ribbon mounted in the same call, available as sheet.toolbar
});

sheet.update({ locale: 'ja' });
sheet.update({ theme: 'ink' }); // themes the grid and the ribbon together
```

Pass `host: { workbookName: () => currentFilename }` to show the current workbook name in Watch Window. The React and Vue components accept the same option.

The language selector is hidden by default. Enable it with `ui.features.localeSwitcher: true`, or toggle it with `sheet.update({ ui: { features: { localeSwitcher: false } } })`, which replaces `ui` as a whole. `update({ locale })` updates the dictionary and calculation profile together; it throws a `FormulonError` when the change is rejected and keeps the current locale. Close open dialogs, Find/Replace, comment edits and filter menus before switching; their drafts are preserved when a switch is rejected.

## Embedding

`ui` decides what is shown, `policy` what users may change, and `viewport` which cells are visible and selectable; each can be changed at runtime with `update({ ui, policy, viewport })`.

```ts
import { fixedFormPolicy, mount, viewerPolicy } from '@libraz/formulon-cell';

const form = await mount(host, {
  workbook: wb,
  ui: { profile: 'embedded', features: { formulaBar: true } }, // grid plus formula bar
  policy: fixedFormPolicy(['B2:B11']), // only B2:B11 accept input, on every sheet
  viewport: { mode: 'fixed', range: 'A1:D12', tabNavigation: 'editable' },
});

form.update({ policy: viewerPolicy() }); // after submit: select and copy only
```

A policy allows or denies 27 operations one by one, takes editable cells as A1 ranges (`'B2:B11'` applies on every sheet, `'Input!B2:B11'` on one) or a predicate, and applies to every input path: editor, clipboard, fill handle, context menu, ribbon and instance API. Its `restrict` hook can add a denial but never a grant. Under a policy, ribbon commands that are not classified by operation stay disabled, so a restricted sheet is best shown without the ribbon or with host-built controls. See [Embedding and use cases](https://formulon.libraz.net/cell/embedding).

## UI profiles

`ui.profile` picks the starting set of features (default `'full'`); `ui.features` switches individual ones on or off over it.

| profile | what's in it |
|---------|--------------|
| `embedded`  | the grid alone, with every feature, the ribbon and printing turned off |
| `minimal`   | formula bar, status bar, basic keymap |
| `standard`  | + Quick Analysis, context menu, find/replace, clipboard, format painter |
| `full`      | + format dialog, paste special, conditional formatting, named ranges, hyperlink dialog, PivotTable creation, validation, autocomplete, hover comments |

## Subpath exports

| import path | contents |
|---|---|
| `@libraz/formulon-cell` | `mount`, `Workbook`, policies and the public types |
| `@libraz/formulon-cell/toolbar` | `ribbon()` and the ribbon model |
| `@libraz/formulon-cell/charts` | chart layout helpers |
| `@libraz/formulon-cell/extensions` | all extension factories |
| `@libraz/formulon-cell/i18n/{en,ja,de,fr,zh,ko,th}` | seven locale dictionaries |
| `@libraz/formulon-cell/styles.css` | grid, dialogs, overlays, ribbon, and all themes in one file |
| `@libraz/formulon-cell/styles/{tokens,paper,ink,contrast,toolbar}.css` | individual stylesheets for granular setups |

## Key APIs

| API | description |
|-----|-------------|
| `mount(host, options)` | mount the spreadsheet into a DOM element |
| `Workbook.createDefault()` / `Workbook.load(bytes)` | create or open a workbook backed by the WASM engine |
| `workbook.batch(fn)` | run a multi-cell write with a single recalc |
| `instance.update(patch)` | change `locale`, `theme`, `ui`, `policy`, `viewport`, `contextMenu`, `toolbar` and other options at runtime |
| `instance.applyChanges(changes, { history })` | batch cell update, all or nothing, decided by the policy and sheet protection |
| `instance.on(event, handler)` | `change`, `cell`, `workbook`, `recalc`, `error`, `update` |

The full option and API reference is on the [documentation site](https://formulon.libraz.net/cell/).

CSV export accepts an explicit locale and delimiter: `wb.sheet(0).exportCsv({ locale: 'de', delimiter: ';' })`. Pass the same settings to `wb.sheet(0).importCsv(text, { at: 'A1', locale: 'de', delimiter: ';' })` for matching numeric and boolean input. Export defaults to the workbook's locale and comma delimiters; import uses the active workbook profile unless `locale` is supplied.

Formula editors and external clipboard input use the active profile's syntax. For example, German `=SUMME(1,5;2)` is stored as `=SUM(1.5,2)`. Programmatic `Sheet.setFormula()` accepts canonical syntax; use `workbook.canonicalizeFormula()` first when supplying localized text. `workbook.localizeFormula()` converts display text, and `workbook.getLocaleFacts()` returns the active profile's separators and scalar names.

Review > Spelling walks the sheet word by word in the desktop Spelling dialog. The built-in checker knows only a short list of common English typos and passes every word of other languages. To check against a real dictionary, pass `spelling: { check(word, locale), suggest(word, locale), addWord?(word, locale) }`. `addWord` receives the words the user adds, so the host can keep them; the library itself stores nothing beyond the session.

## Documentation

- [Live demo](https://formulon.libraz.net/cell/demo)
- [Getting started](https://formulon.libraz.net/cell/install)
- [Options](https://formulon.libraz.net/cell/options)
- [Embedding and use cases](https://formulon.libraz.net/cell/embedding)
- [Modals and fullscreen](https://formulon.libraz.net/cell/modals)
- [React / Vue](https://formulon.libraz.net/cell/frameworks)
- [Theming](https://formulon.libraz.net/cell/theming)
- [Language and labels](https://formulon.libraz.net/cell/i18n)

## License

[Apache-2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
