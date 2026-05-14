# formulon-cell

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![codecov](https://codecov.io/gh/libraz/formulon-cell/branch/main/graph/badge.svg)](https://codecov.io/gh/libraz/formulon-cell)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell?label=%40libraz%2Fformulon-cell)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — react](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![npm — vue](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-6-blue?logo=typescript)](https://www.typescriptlang.org/)

[formulon](https://github.com/libraz/formulon) WASM 計算エンジン向けの
スプレッドシート UI ライブラリ。デスクトップ表計算ソフト風の UI 表層、
Canvas 描画によるグリッド、拡張ベースの機能構成、実行時ロケール切替を
提供します。

> **β（ベータ）。** `formulon-cell` は WASM 計算エンジン
> [**formulon**](https://github.com/libraz/formulon) のデモ用ホストとして
> 開発されたパッケージです。formulon は C++17 製の Excel 互換ヘッドレス
> 計算エンジンで、ブラウザ (WASM) / Python / CLI に同一エンジンを配布
> します。エンジン本体のドキュメントは
> [formulon.libraz.net](https://formulon.libraz.net) を参照してください。
> 本パッケージの UI 表層はまだ安定化の途上にあるため、意図したタイミング
> でのみ更新できるよう、依存バージョンは範囲指定で固定してください。

## パッケージ

| パッケージ | npm | 概要 |
|---------|-----|------|
| [`@libraz/formulon-cell`](./packages/formulon-cell)             | [![npm](https://img.shields.io/npm/v/@libraz/formulon-cell?label=)](https://www.npmjs.com/package/@libraz/formulon-cell)             | Vanilla TS / DOM コア |
| [`@libraz/formulon-cell-react`](./packages/formulon-cell-react) | [![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=)](https://www.npmjs.com/package/@libraz/formulon-cell-react) | React 18+ コンポーネント・フック・リボンツールバー |
| [`@libraz/formulon-cell-vue`](./packages/formulon-cell-vue)     | [![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)     | Vue 3 コンポーネント・コンポーザブル・リボンツールバー |

## インストール

```sh
npm install @libraz/formulon-cell zustand
# または yarn / pnpm
```

`zustand` はピア依存として公開しています。UI 表層が購読しているストアに、
利用者側からも同じインスタンスでアクセスできるようにするためです。

WASM エンジンは pthread 有効版を同梱しており、
[crossOriginIsolated コンテキスト](https://developer.mozilla.org/docs/Web/API/crossOriginIsolated)
（`Cross-Origin-Opener-Policy: same-origin` + `Cross-Origin-Embedder-Policy:
require-corp`）を必要とします。これが満たされない環境では formulon-cell は
インメモリのスタブエンジンにフォールバックします。UI はそのまま動作し、
数式計算のみが段階的に機能を落とす形になります。

## クイックスタート

```ts
import { Spreadsheet, WorkbookHandle, presets } from '@libraz/formulon-cell';
import '@libraz/formulon-cell/styles.css';

const host = document.getElementById('sheet')!;
const wb = await WorkbookHandle.createDefault();
const sheet = await Spreadsheet.mount(host, {
  workbook: wb,
  features: presets.full(),
  locale: 'ja',
});

sheet.i18n.setLocale('en');     // 実行時にロケールを切り替え
sheet.setTheme('ink');           // ダークテーマへ切り替え
```

## バンドラ統合

formulon-cell は `@libraz/formulon` の pthread 有効 WASM モジュールを
再利用するため、エンジンパッケージと同じバンドラ設定がそのまま必要です。
押さえておくべきポイントは 4 点です。

**1. ワーカーは ES モジュール形式で出力する。** 再計算スケジューラは
Emscripten が生成する Web Worker 上で動作し、
`new Worker(new URL(...), { type: 'module' })` で起動します。多くのバンドラは
既定でクラシック (IIFE) 形式のワーカーを生成するため、明示的に切り替える
必要があります。

```ts
// vite.config.ts
export default defineConfig({
  worker: { format: 'es' },
});
```

webpack 5 は `output.module: true` のとき `{ type: 'module' }` を自動的に
認識します。esbuild ではワーカーチャンクに `--format=esm` を指定してください。

**2. トップレベル await と動的な Node モジュール読み込みには es2022
ターゲットが必要。** エンジンファクトリはトップレベル await と条件付きの
`await import('node:...')` を利用します。メイン側・ワーカー側の双方で
ビルドターゲットを `es2022` 以上に引き上げてください。

```ts
// vite.config.ts
export default defineConfig({
  build: { target: 'es2022' },
});
```

**3. 依存関係の事前バンドル対象からエンジンを除外する。** formulon-cell は
`@libraz/formulon` を経由してロードし、その Emscripten ラッパーが
ワーカーと WASM アセットの解決を担当します。両パッケージを事前バンドルの
対象から外し、アセット解決はアプリ側のバンドラに委ねてください。

```ts
// vite.config.ts
export default defineConfig({
  optimizeDeps: { exclude: ['@libraz/formulon-cell', '@libraz/formulon'] },
});
```

**4. SharedArrayBuffer はクロスオリジン分離を要求する。** ページに
`Cross-Origin-Opener-Policy: same-origin` と
`Cross-Origin-Embedder-Policy: require-corp` ヘッダを付与してください。
これらのヘッダが無い環境では `SharedArrayBuffer` が未定義となり、
formulon-cell はインメモリの **スタブエンジン** にフォールバックします。
キャンバス・数式バー・編集系の操作は引き続き動作しますが、数式評価・
再計算・xlsx の読み書きは無効化され、呼び出しても何も起こりません。
実行時に判定するには `crossOriginIsolated` を参照するか、
`WorkbookHandle.createDefault()` の後に `isUsingStub()` を呼び出してください。

```ts
import { WorkbookHandle, isUsingStub } from '@libraz/formulon-cell';

const wb = await WorkbookHandle.createDefault();
if (isUsingStub()) {
  console.warn('formulon-cell: スタブエンジンで実行中 — 再計算は無効');
}
```

## 機能

- **デスクトップ表計算ソフト風の UI 表層** を標準装備（数式バー、
  ステータスバー、コンテキストメニュー、シートタブ、ビューツールバー）。
- **Canvas 描画によるグリッド** とテーマトークン。`paper`（ライト）と
  `ink`（ダーク）を同梱しており、ドキュメント化された CSS 変数で独自テーマ
  も作成可能。
- **拡張ベース API**: ビルトイン機能は機能フラグで制御。差し替え可能な
  パーツ（検索／置換、書式ダイアログ、形式を選択して貼り付け、
  ハイパーリンクダイアログ、ホバーコメント、ビューツールバー、
  クイック分析、ピボットテーブル作成 など）は拡張ファクトリとして提供。
- **実行時 i18n** — 再マウントせずにロケールを切り替え可能。`ja` と `en`
  を同梱しており、実行時にロケールを追加登録することもできます。
- **ヘッドレスモード** — キャンバスとストアのみを利用し、UI 表層を
  独自に実装することもできます。

### プリセット

| プリセット | 含まれる機能 |
|----------|------------|
| `presets.minimal()`  | 数式バー、ステータスバー、基本キーマップ |
| `presets.standard()` | + ビューツールバー、クイック分析、セッションチャートオーバーレイ、ワークブックオブジェクトインスペクター、コンテキストメニュー、検索／置換、クリップボード、書式コピー、ホイールスクロール |
| `presets.full()`     | + 書式ダイアログ、形式を選択して貼り付け、条件付き書式、反復計算設定、ジャンプ — セル選択、ページ設定、名前付き範囲、ハイパーリンクダイアログ、ピボットテーブル作成、入力規則、オートコンプリート、ホバーコメント、表計算キーマップ |

### i18n

```ts
import { Spreadsheet } from '@libraz/formulon-cell';

const sheet = await Spreadsheet.mount(host, { locale: 'en' });

// 実行時にロケールを切り替え — すべてのラベルがその場で更新される
sheet.i18n.setLocale('ja');

// 辞書をフォークせず、一部のキーだけ上書きする
sheet.i18n.extend('ja', { contextMenu: { copy: 'コピーする' } });

// 新しいロケールを登録する
import fr from './fr.js';
sheet.i18n.register('fr', fr);
sheet.i18n.setLocale('fr');
```

## デモアプリ

| アプリ | 起動 | 内容 |
|-----|-----|------|
| `apps/playground`  | `yarn dev`        | Vanilla DOM プレイグラウンド（表計算キーマップ） |
| `apps/react-demo`  | `yarn dev:react`  | React コンポーネント `<Spreadsheet>` と同じ機能面 |
| `apps/vue-demo`    | `yarn dev:vue`    | Vue コンポーネント `<Spreadsheet>` と同じ機能面 |

## フレームワーク向けリボンツールバー

React 版・Vue 版のパッケージは、デモアプリで使用しているリボンを再利用可能
な UI 表層として公開しています。どちらの実装も同じタブ構成・同じコマンド
ラインナップを提供します。コンポーネントと一緒にツールバー用 CSS も
読み込んでください。

```tsx
import { SpreadsheetToolbar, type RibbonTab } from '@libraz/formulon-cell-react';
import '@libraz/formulon-cell-react/toolbar.css';
```

```vue
<script setup lang="ts">
import { type RibbonTab } from '@libraz/formulon-cell-vue';
import SpreadsheetToolbar from '@libraz/formulon-cell-vue/toolbar.vue';
import '@libraz/formulon-cell-vue/toolbar.css';
</script>
```

## リリース

タグベースのリリースフローは [`docs/releasing.md`](./docs/releasing.md)
を参照してください。

## ライセンス

[Apache-2.0](LICENSE)
