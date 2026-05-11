# @libraz/formulon-cell

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm downloads](https://img.shields.io/npm/dm/@libraz/formulon-cell)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-6-blue?logo=typescript)](https://www.typescriptlang.org/)

[formulon](https://github.com/libraz/formulon) WASM 計算エンジン向けの
スプレッドシート UI。

> **β（ベータ）。** 本パッケージは WASM 計算エンジン
> [**formulon**](https://github.com/libraz/formulon) のデモ用ホストとして
> 開発されました。formulon は C++17 製の Excel 互換ヘッドレス計算エンジン
> で、ブラウザ (WASM) / Python / CLI に同一エンジンを配布します。エンジン
> 本体のドキュメントは [formulon.libraz.net](https://formulon.libraz.net)
> を参照してください。UI 表層は安定化途上です。意図的にアップグレード
> できるバージョンレンジで固定してください。

- **デスクトップ表計算ソフト風のクロム** を標準装備（数式バー、ステータス
  バー、コンテキストメニュー、シートタブ）。
- **Canvas レンダリングのグリッド** とテーマトークン。`paper`（ライト）と
  `ink`（ダーク）が同梱。ドキュメント化された CSS 変数で独自テーマも
  作成可能。
- **拡張ベース API**: ビルトインは feature フラグで制御。差し替え可能な
  パーツ（検索／置換、書式ダイアログ、形式を選択して貼り付け、
  ハイパーリンクダイアログ、ホバーコメント、View ツールバー、
  クイック分析、ピボットテーブル作成 など）は拡張ファクトリとして提供。
- **ランタイム i18n** — 再マウントせずにロケール切替。`ja` と `en` を同梱、
  実行時に追加ロケールを登録可能。
- **ヘッドレスモード** — キャンバス + ストアのみを利用し、独自クロムを
  実装することもできます。

## インストール

```sh
npm install @libraz/formulon-cell zustand
# または yarn / pnpm
```

`zustand` は peer dependency です。クロムが購読しているストアを利用者側からも
読めるよう公開しています。

WASM エンジンは pthread 有効版を同梱しており、
[crossOriginIsolated コンテキスト](https://developer.mozilla.org/docs/Web/API/crossOriginIsolated)
（`Cross-Origin-Opener-Policy: same-origin` + `Cross-Origin-Embedder-Policy:
require-corp`）を必要とします。これが満たされない環境では formulon-cell は
インメモリのスタブエンジンにフォールバックします。UI は動作し続け、
数式計算はグレースフルに縮退します。

## バンドラ統合（Vite / webpack / esbuild）

formulon-cell は `@libraz/formulon` の pthread 有効 WASM モジュールを
再利用するため、エンジンパッケージのバンドラ作法がそのまま適用されます。
重要なポイントは 4 つです。

**1. Worker は ES モジュールとして出力する。** 再計算スケジューラは
Emscripten が生成する Web Worker 上で動作し、
`new Worker(new URL(...), { type: 'module' })` で起動します。バンドラの
デフォルトはクラシック (IIFE) Worker なので明示的に切り替えが必要です。

```ts
// vite.config.ts
export default defineConfig({
  worker: { format: 'es' },
});
```

webpack 5 は `output.module: true` のとき `{ type: 'module' }` を自動的に
拾います。esbuild では Worker チャンクに `--format=esm` を指定します。

**2. Top-level await と動的 node import には es2022 ターゲットが必要。**
エンジンファクトリは TLA と条件付きの `await import('node:...')` を使用します。
メインとワーカーの両方でターゲットを引き上げてください。

```ts
// vite.config.ts
export default defineConfig({
  build: { target: 'es2022' },
});
```

**3. エンジンを依存プリバンドルから除外する。** formulon-cell は
`@libraz/formulon` を読み込み、その Emscripten ラッパーが Worker/WASM
アセットの解決を担います。両パッケージを依存プリバンドルから外し、
アセット解決をアプリのバンドラに任せます。

```ts
// vite.config.ts
export default defineConfig({
  optimizeDeps: { exclude: ['@libraz/formulon-cell', '@libraz/formulon'] },
});
```

**4. SharedArrayBuffer はクロスオリジン分離を要求する。** ページに
`Cross-Origin-Opener-Policy: same-origin` と
`Cross-Origin-Embedder-Policy: require-corp` ヘッダを付与してください。
ヘッダが無い環境では `SharedArrayBuffer` が未定義となり、
formulon-cell はインメモリの **スタブエンジン** にフォールバックします。
キャンバス・数式バー・編集操作は機能し続けますが、数式評価・再計算・
xlsx ラウンドトリップは no-op に縮退します。実行時の判定は
`crossOriginIsolated` あるいは `WorkbookHandle.createDefault()` 後の
`isUsingStub()` で行えます。

```ts
import { WorkbookHandle, isUsingStub } from '@libraz/formulon-cell';

const wb = await WorkbookHandle.createDefault();
if (isUsingStub()) {
  console.warn('formulon-cell: スタブエンジンで実行中 — 再計算は無効');
}
```

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

sheet.i18n.setLocale('en');     // ランタイムロケール切替
sheet.setTheme('ink');           // ダークモード
```

## プリセット

| プリセット | 含まれる機能 |
|----------|------------|
| `presets.minimal()`  | 数式バー、ステータスバー、基本キーマップ |
| `presets.standard()` | + View ツールバー、クイック分析、セッションチャートオーバーレイ、ワークブックオブジェクトインスペクタ、コンテキストメニュー、検索／置換、クリップボード、書式コピー、ホイールスクロール |
| `presets.full()`    | + 書式ダイアログ、形式を選択して貼り付け、条件付き書式、反復計算設定、ジャンプ — セル選択、ページ設定、名前付き範囲、ハイパーリンクダイアログ、ピボットテーブル作成、入力規則、オートコンプリート、ホバーコメント、表計算キーマップ |

独自構成:

```ts
import {
  Spreadsheet,
  contextMenu,
  findReplace,
  pasteSpecial,
  presets,
  quickAnalysis,
  statusBar,
  viewToolbar,
  workbookObjects,
} from '@libraz/formulon-cell';

await Spreadsheet.mount(host, {
  features: {
    ...presets.minimal(),
    statusBar: false,
    viewToolbar: false,
    workbookObjects: false,
    quickAnalysis: false,
    contextMenu: false,
    findReplace: false,
    pasteSpecial: false,
  },
  extensions: [
    statusBar(),
    workbookObjects(),
    viewToolbar(),
    quickAnalysis(),
    contextMenu(),
    findReplace(),
    pasteSpecial(),
  ],
});
```

`allBuiltIns()` はデフォルト有効の差し替え可能ビルトインを拡張
ファクトリとして返します。`watchWindow()` や `slicer()` などデフォルト
無効のパネルは個別に export されており、アプリ側で意図的にオプトイン
できます。

## i18n

```ts
import { Spreadsheet } from '@libraz/formulon-cell';
import ja from '@libraz/formulon-cell/i18n/ja';
import en from '@libraz/formulon-cell/i18n/en';

const sheet = await Spreadsheet.mount(host, { locale: 'en' });

// ランタイムでロケール切替 — すべてのラベルがその場で更新される
sheet.i18n.setLocale('ja');

// 辞書を fork せずに一部だけ上書き
sheet.i18n.extend('ja', { contextMenu: { copy: 'コピーする' } });

// 新しいロケールを登録
import fr from './fr.js';
sheet.i18n.register('fr', fr);
sheet.i18n.setLocale('fr');
```

## React / Vue

```sh
npm install @libraz/formulon-cell-react react react-dom
# または
npm install @libraz/formulon-cell-vue vue
```

詳細は [`@libraz/formulon-cell-react`](../formulon-cell-react/README_ja.md) と
[`@libraz/formulon-cell-vue`](../formulon-cell-vue/README_ja.md) を参照。

## ライセンス

[Apache-2.0](LICENSE)
