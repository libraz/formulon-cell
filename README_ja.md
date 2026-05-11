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

> **β（ベータ）。** `formulon-cell` は WASM 計算エンジン
> [**formulon**](https://github.com/libraz/formulon) のデモ用ホストとして
> 開発されたパッケージです。formulon は C++17 製の Excel 互換ヘッドレス
> 計算エンジンで、ブラウザ (WASM) / Python / CLI に同一エンジンを配布
> します。エンジン本体のドキュメントは
> [formulon.libraz.net](https://formulon.libraz.net) を参照してください。
> 本パッケージの UI 表層は安定化途上です。意図的にアップグレードできる
> バージョンレンジで固定してください。

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

## ライセンス

[Apache-2.0](LICENSE)
