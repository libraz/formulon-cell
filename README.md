# formulon-cell

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell?label=%40libraz%2Fformulon-cell)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — react](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![npm — vue](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-6-blue?logo=typescript)](https://www.typescriptlang.org/)

Spreadsheet UI library for the [formulon](https://github.com/libraz/formulon)
WASM calc engine. Desktop-spreadsheet-style chrome, canvas-rendered grid,
extension-based feature composition, runtime i18n.

## Packages

| package | npm | what it is |
|---------|-----|------------|
| [`@libraz/formulon-cell`](./packages/formulon-cell)             | [![npm](https://img.shields.io/npm/v/@libraz/formulon-cell?label=)](https://www.npmjs.com/package/@libraz/formulon-cell)             | Vanilla TS / DOM core |
| [`@libraz/formulon-cell-react`](./packages/formulon-cell-react) | [![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=)](https://www.npmjs.com/package/@libraz/formulon-cell-react) | React 18+ component + hooks |
| [`@libraz/formulon-cell-vue`](./packages/formulon-cell-vue)     | [![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)     | Vue 3 component + composables |

## Install

```sh
npm install @libraz/formulon-cell zustand
# or yarn / pnpm
```

## Demo apps

| app | run | what it shows |
|-----|-----|---------------|
| `apps/playground`  | `yarn dev`        | Vanilla DOM playground (spreadsheet keymap) |
| `apps/react-demo`  | `yarn dev:react`  | Same surface as `<Spreadsheet>` React component |
| `apps/vue-demo`    | `yarn dev:vue`    | Same surface as `<Spreadsheet>` Vue component |

## Status

`v0.1.x` — public API stabilizing. Until v1.0 minor bumps may reshape
extension contracts. Pin a version range you can upgrade on purpose.

## License

[Apache-2.0](LICENSE)
