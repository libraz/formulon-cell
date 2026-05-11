# formulon-cell

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell?label=%40libraz%2Fformulon-cell)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — react](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![npm — vue](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-6-blue?logo=typescript)](https://www.typescriptlang.org/)

[formulon](https://github.com/libraz/formulon) WASM 計算エンジン向けの
スプレッドシート UI ライブラリ。デスクトップ表計算ソフト風のクロム、
Canvas レンダリングのグリッド、拡張ベースの機能構成、ランタイム i18n を
提供します。

## パッケージ

| パッケージ | npm | 概要 |
|---------|-----|------|
| [`@libraz/formulon-cell`](./packages/formulon-cell)             | [![npm](https://img.shields.io/npm/v/@libraz/formulon-cell?label=)](https://www.npmjs.com/package/@libraz/formulon-cell)             | Vanilla TS / DOM コア |
| [`@libraz/formulon-cell-react`](./packages/formulon-cell-react) | [![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=)](https://www.npmjs.com/package/@libraz/formulon-cell-react) | React 18+ コンポーネント + フック |
| [`@libraz/formulon-cell-vue`](./packages/formulon-cell-vue)     | [![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)     | Vue 3 コンポーネント + コンポーザブル |

## インストール

```sh
npm install @libraz/formulon-cell zustand
# または yarn / pnpm
```

## デモアプリ

| アプリ | 起動 | 内容 |
|-----|-----|------|
| `apps/playground`  | `yarn dev`        | Vanilla DOM プレイグラウンド（表計算キーマップ） |
| `apps/react-demo`  | `yarn dev:react`  | React コンポーネント `<Spreadsheet>` と同じ機能面 |
| `apps/vue-demo`    | `yarn dev:vue`    | Vue コンポーネント `<Spreadsheet>` と同じ機能面 |

## ステータス

`v0.1.x` — 公開 API は安定化途上です。v1.0 までは minor バンプで拡張
コントラクトが変わる可能性があります。意図的にアップグレードできる
バージョンレンジを固定してください。

## ライセンス

[Apache-2.0](LICENSE)
