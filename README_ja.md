# formulon-cell

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell?label=%40libraz%2Fformulon-cell)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — react](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![npm — vue](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![codecov](https://codecov.io/gh/libraz/formulon-cell/branch/main/graph/badge.svg)](https://codecov.io/gh/libraz/formulon-cell)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-6-blue?logo=typescript)](https://www.typescriptlang.org/)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-blue)](https://formulon.libraz.net/ja/cell/)
[![Demo](https://img.shields.io/badge/demo-live-brightgreen)](https://formulon.libraz.net/ja/cell/demo)

**formulon-cell は、Web ページの中で動く Excel 風の表計算を提供します。** DOM 要素、React コンポーネント、Vue コンポーネントのどれにマウントしても、Canvas 描画のグリッドとデスクトップ表計算ソフト風の UI（数式バー、リボン、シートタブ、コンテキストメニュー）が表示され、計算は [formulon](https://github.com/libraz/formulon) の WASM エンジンが受け持ちます。コアはフレームワーク非依存の TypeScript で、React 版と Vue 版は同じマウント処理を包む薄いアダプタです。

[デモ](https://formulon.libraz.net/ja/cell/demo) でブラウザ上ですぐに試せます。

次のような用途に向いています。

- **データグリッドではなく表計算を組み込みたい。** 選択、編集、数式入力、再計算、ファイルの読み書きが最初のマウントから動きます。
- **数式を本物のエンジンで計算したい。** 評価するのは WebAssembly にコンパイルした C++ エンジンで、UI に後付けした JavaScript の数式パーサーではありません。
- **使う機能だけを載せたい。** プリセットと拡張ファクトリでマウントする機能を選べます。UI を外してキャンバスとストアだけを使うこともできます。
- **入力フォームや閲覧専用ビューとして固定したい。** 操作ポリシーで、ユーザーが触れるセルと操作を制限できます。表示する UI とは独立に設定します。

> **互換性の現状。** formulon-cell はデスクトップ表計算ソフト（Excel 互換）の操作感を目指しています。基本的なブック操作は使えますが、ダイアログ、細かなコントロールの挙動、キーボード操作、アクセシビリティはまだ同等の振る舞いを保証しておらず、不具合が残っている可能性もあります。現時点では、そのまま置き換えられる完成品として案内しないでください。

## パッケージ

| パッケージ | 概要 |
|---------|------|
| [`@libraz/formulon-cell`](./packages/formulon-cell) | Vanilla TS / DOM コア（グリッド、UI、拡張、i18n、テーマ） |
| [`@libraz/formulon-cell-react`](./packages/formulon-cell-react) | React 18+ のコンポーネント、フック、リボンツールバー |
| [`@libraz/formulon-cell-vue`](./packages/formulon-cell-vue) | Vue 3 のコンポーネント、コンポーザブル、リボンツールバー |

## インストール

```sh
npm install @libraz/formulon-cell zustand
# React: npm install @libraz/formulon-cell-react @libraz/formulon-cell react react-dom zustand
# Vue:   npm install @libraz/formulon-cell-vue @libraz/formulon-cell vue zustand
```

`zustand` はピア依存です。UI が購読しているストアを、ホスト側からも同じインスタンスで読めるようにしています。

formulon 0.18.0 の標準 WASM エンジンは単一スレッド版で、COOP/COEP ヘッダも `SharedArrayBuffer` も要りません。エンジンの初期化に失敗すると `WorkbookHandle.createDefault()` は reject し、別のエンジンに黙って切り替えることはありません。Node.js で使う場合は22以降が必要です。インメモリのスタブエンジンは `preferStub: true` を渡したときだけ使われ、テストと明示的なデモ専用です。

## クイックスタート

```ts
import { Spreadsheet, WorkbookHandle, presets } from '@libraz/formulon-cell';
import '@libraz/formulon-cell/styles.css';

const wb = await WorkbookHandle.createDefault();
const sheet = await Spreadsheet.mount(document.getElementById('sheet')!, {
  workbook: wb,
  features: { ...presets.full(), localeSwitcher: true },
  locale: 'ja',
  toolbar: true,
});

sheet.setLocale('en'); // 再マウントせずにロケールを切り替え
sheet.setTheme('ink');      // ダークテーマ
```

`getWorkbookName` や言語選択 UI（`features.localeSwitcher`）などのオプションは、[オプション](https://formulon.libraz.net/ja/cell/options)と[言語の切り替え](https://formulon.libraz.net/ja/cell/language-switching)にまとめています。

React と Vue では、同じオプションをコンポーネントの props で渡します。

```tsx
import { Spreadsheet, presets } from '@libraz/formulon-cell-react';
import '@libraz/formulon-cell/styles.css';

<Spreadsheet features={presets.full()} locale="ja" style={{ height: '100vh' }} />;
```

```vue
<script setup lang="ts">
import { Spreadsheet, presets } from '@libraz/formulon-cell-vue';
import '@libraz/formulon-cell/styles.css';
</script>

<template>
  <Spreadsheet :features="presets.full()" locale="ja" style="height: 100vh" />
</template>
```

## 制限付きの埋め込み

UI プロファイル `embedded` はグリッドだけをマウントし、ユーザーが変更できる範囲はポリシーで決めます。`fixedFormPolicy(ranges)` は指定した範囲での編集を許可し、`viewerPolicy()` は選択とコピーを残してシートを読み取り専用にします。リボンを隠してもセルは読み取り専用になりません。どの入力経路でも、読み取り専用にするのはポリシーです。詳しくは [埋め込みと利用例](https://formulon.libraz.net/ja/cell/embedding) を参照してください。

## バンドラの設定

エンジンは Emscripten のラッパー経由で WASM を読み込むため、バンドラ側で手を加えないようにします。Vite の場合は次のとおりです。

```ts
// vite.config.ts
export default defineConfig({
  build: { target: 'es2022' }, // エンジンのファクトリがトップレベル await を使う
  optimizeDeps: { exclude: ['@libraz/formulon-cell', '@libraz/formulon'] },
});
```

CDN からの `.wasm` の配信（`locateFile`）とスレッド版エンジンについては、[バンドラの設定](https://formulon.libraz.net/ja/cell/bundler)にまとめています。

## 主な機能

- **デスクトップ風の UI** — 数式バー、ステータスバー、リボン、シートタブ、コンテキストメニュー、各種ダイアログ。[リボンツールバー](https://formulon.libraz.net/ja/cell/toolbar) · [モーダルとダイアログ](https://formulon.libraz.net/ja/cell/modals)
- **デスクトップ版と同じ数式編集** — 矢印キーと F4 での参照の選択と切り替え、参照中の Ctrl+矢印、入力途中の取り消し、選択部分の F9 評価、R1C1 参照形式。
- **テーマ付きの Canvas グリッド** — CSS トークンで見た目を変えられ、`paper`、`ink`、`contrast` を同梱しています。[テーマ](https://formulon.libraz.net/ja/cell/theming)
- **グラフ** — 7 種類のグラフを設定パネルで扱い、読み込んだブックのグラフもその場で編集して保存できます。[グラフ](https://formulon.libraz.net/ja/cell/charts)
- **XLSX への保存** — コメント、画像、グラフ、テーブル、ピボットテーブル、スライサー、保護などを保存でき、名前を付けて保存では `.xlsb` も書き出せます。[ホストとの連携](https://formulon.libraz.net/ja/cell/host-integration)
- **プリセットと拡張** — `presets.minimal()`、`standard()`、`full()` から選び、個々の部品は差し替えも省略もできます。[拡張](https://formulon.libraz.net/ja/cell/extensions)
- **Mac プロファイル** — `ui: { platform: 'mac' }` で Mac 向けのリボン、ショートカット、関数引数パレットに切り替わります。[プラットフォームと Mac UI](https://formulon.libraz.net/ja/cell/platform)
- **実行時の言語切り替え** — 7 言語を同梱し、`sheet.i18n.register()` で追加できます。[国際化](https://formulon.libraz.net/ja/cell/i18n)
- **ヘッドレス構成** — UI をホスト側で用意する場合に、キャンバスとストアだけを残せます。

## ドキュメント

- [デモ](https://formulon.libraz.net/ja/cell/demo)
- [インストール](https://formulon.libraz.net/ja/cell/install)
- [設定オプション](https://formulon.libraz.net/ja/cell/options)
- [埋め込みと利用例](https://formulon.libraz.net/ja/cell/embedding)
- [モーダルと全画面表示](https://formulon.libraz.net/ja/cell/modals)
- [React / Vue](https://formulon.libraz.net/ja/cell/frameworks)
- [テーマ](https://formulon.libraz.net/ja/cell/theming)
- [言語と表示文言](https://formulon.libraz.net/ja/cell/i18n)

## 開発

```sh
yarn install
yarn dev:react   # apps/react-demo
yarn dev:vue     # apps/vue-demo
yarn build
yarn test
```

プルリクエストを歓迎します。CI が実行するチェックは [`CONTRIBUTING.md`](./CONTRIBUTING.md) にあります。リリースは [`docs/releasing.md`](./docs/releasing.md) のタグを使った手動フローで行います。

## 対象外

formulon-cell は UI 層です。計算エンジンでもアプリケーションでもありません。数式の評価、再計算、xlsx の解析は [formulon](https://github.com/libraz/formulon) の担当です。サーバー、永続化、共同編集の通信、認証も持たず、これらはホスト側が用意します。React 版と Vue 版は薄いまま保ち、メニュー、コマンド、ダイアログはコアに置いているため、3 つのホストの挙動がずれることはありません。

## ライセンス

[Apache-2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
