# formulon-cell

Spreadsheet UI library for the [formulon](https://github.com/libraz/formulon)
WASM calc engine. Desktop-spreadsheet-style chrome, canvas-rendered grid,
extension-based feature composition, runtime i18n.

## Packages

| package | npm | what it is |
|---------|-----|------------|
| [`@libraz/formulon-cell`](./packages/formulon-cell)             | `@libraz/formulon-cell`             | Vanilla TS / DOM core |
| [`@libraz/formulon-cell-react`](./packages/formulon-cell-react) | `@libraz/formulon-cell-react`       | React 18+ component + hooks |
| [`@libraz/formulon-cell-vue`](./packages/formulon-cell-vue)     | `@libraz/formulon-cell-vue`         | Vue 3 component + composables |

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

[Apache-2.0](./LICENSE) © libraz
