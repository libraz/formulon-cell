# formulon-cell

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell?label=%40libraz%2Fformulon-cell)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — react](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![npm — vue](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![codecov](https://codecov.io/gh/libraz/formulon-cell/branch/main/graph/badge.svg)](https://codecov.io/gh/libraz/formulon-cell)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-6-blue?logo=typescript)](https://www.typescriptlang.org/)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-2563eb)](https://formulon.libraz.net/ja/cell/)
[![Demo](https://img.shields.io/badge/demo-live-brightgreen)](https://formulon.libraz.net/ja/cell/demo)

**formulon-cell は、デスクトップ版 Excel に合わせて作っている Web アプリ向けの表計算コンポーネントです。** DOM 要素、React コンポーネント、Vue コンポーネントのどれにでもマウントでき、Canvas 描画のグリッドに数式バー、リボン、シートタブ、コンテキストメニュー、ダイアログが付きます。計算は [formulon](https://github.com/libraz/formulon)（WebAssembly にコンパイルした C++ の計算エンジン）が受け持ちます。セルの描画、ダイアログ、キーボードショートカット、クリップボードの挙動は、機能ごとに Excel のデスクトップ版と突き合わせ、違いがあれば直しています。

🧪 **[デモ](https://formulon.libraz.net/ja/cell/demo)** &nbsp;·&nbsp; 📖 **[ドキュメント](https://formulon.libraz.net/ja/cell/)** &nbsp;·&nbsp; **[インストール](https://formulon.libraz.net/ja/cell/install)**

> **ベータ版（1.0 未満）です。** 公開 API はマイナーバージョンの間でも変わることがあります。まだ突き合わせていない挙動は Excel と異なる場合があります。違いを見つけたら [issue](https://github.com/libraz/formulon-cell/issues) で知らせてください。

## できること

- **デスクトップ版の UI** — 数式バー、ステータスバー、リボン、シートタブ、コンテキストメニュー、各種ダイアログ。[リボンツールバー](https://formulon.libraz.net/ja/cell/toolbar) · [モーダルとダイアログ](https://formulon.libraz.net/ja/cell/modals)
- **デスクトップ版と同じ数式編集** — 矢印キーと F4 での参照の選択と切り替え、参照中の Ctrl+矢印、入力途中の取り消し、選択部分の F9 評価、R1C1 参照形式。
- **テーマ付きの Canvas グリッド** — CSS トークンで見た目を変えられ、`paper`、`ink`、`contrast` を同梱しています。[テーマ](https://formulon.libraz.net/ja/cell/theming)
- **グラフ** — 7 種類のグラフを設定パネルで扱い、その場で編集して保存できます。[グラフ](https://formulon.libraz.net/ja/cell/charts)
- **XLSX への保存** — コメント、画像、グラフ、テーブル、ピボットテーブル、スライサー、保護などを保存でき、名前を付けて保存では `.xlsb` も書き出せます。[ホストとの連携](https://formulon.libraz.net/ja/cell/host-integration)
- **制限付きの埋め込み** — `fixedFormPolicy(ranges)` と `viewerPolicy()` で、ユーザーが変更できる範囲をすべての入力経路で制限します。表示する UI とは独立に設定します。[埋め込みと利用例](https://formulon.libraz.net/ja/cell/embedding)
- **UI プロファイルと拡張** — `ui.profile`（`embedded`、`minimal`、`standard`、`full`）から選び、個々の部品は拡張で差し替えも省略もできます。[拡張](https://formulon.libraz.net/ja/cell/extensions)
- **Mac プロファイル** — `ui: { platform: 'mac' }` で Mac 向けのリボン、ショートカット、関数引数パレットに切り替わります。[プラットフォームと Mac UI](https://formulon.libraz.net/ja/cell/platform)
- **実行時の言語切り替え** — 7 言語を同梱し、`instance.i18n.override()` で文言を上書きできます。[国際化](https://formulon.libraz.net/ja/cell/i18n)

## インストール

```sh
npm install @libraz/formulon-cell zustand
npm install @libraz/formulon-cell-react @libraz/formulon-cell react react-dom zustand  # React 18+
npm install @libraz/formulon-cell-vue @libraz/formulon-cell vue zustand                # Vue 3
```

`zustand` はピア依存です。UI が購読しているストアを、ホスト側からも同じインスタンスで読めるようにしています。標準のエンジンは単一スレッド版で、COOP/COEP ヘッダも `SharedArrayBuffer` も要りません。エンジンの初期化に失敗すると `Workbook.createDefault()` は reject し、別のエンジンに黙って切り替えることはありません。インメモリのスタブエンジンは `engine: 'stub'` を渡したときだけ使われます。Node.js で使う場合は 22 以降が必要です。

### バンドラの設定

エンジンは Emscripten のラッパー経由で WASM を読み込むため、バンドラ側で手を加えないようにします。Vite の場合は次のとおりです。

```ts
// vite.config.ts
export default defineConfig({
  build: { target: 'es2022' }, // エンジンのファクトリがトップレベル await を使う
  optimizeDeps: { exclude: ['@libraz/formulon-cell', '@libraz/formulon'] },
});
```

CDN からの `.wasm` の配信（`locateFile`）とスレッド版エンジンについては、[バンドラの設定](https://formulon.libraz.net/ja/cell/bundler)にまとめています。

## クイックスタート

### JavaScript / TypeScript

```ts
import { mount, Workbook } from '@libraz/formulon-cell';
import { ribbon } from '@libraz/formulon-cell/toolbar';
import '@libraz/formulon-cell/styles.css';

const wb = await Workbook.createDefault();
const sheet = await mount(document.getElementById('sheet')!, {
  workbook: wb,
  ui: { profile: 'full', features: { localeSwitcher: true } },
  locale: 'ja',
  toolbar: ribbon(),
});

sheet.update({ locale: 'en' }); // 再マウントせずにロケールを切り替え
sheet.update({ theme: 'ink' }); // ダークテーマ
```

→ [設定オプション](https://formulon.libraz.net/ja/cell/options)

### React

```tsx
import { Spreadsheet } from '@libraz/formulon-cell-react';
import '@libraz/formulon-cell/styles.css';

<Spreadsheet ui={{ profile: 'full' }} locale="ja" style={{ height: '100vh' }} />;
```

→ [React / Vue](https://formulon.libraz.net/ja/cell/frameworks)

### Vue

```vue
<script setup lang="ts">
import { Spreadsheet } from '@libraz/formulon-cell-vue';
import '@libraz/formulon-cell/styles.css';
</script>

<template>
  <Spreadsheet :ui="{ profile: 'full' }" locale="ja" style="height: 100vh" />
</template>
```

→ [React / Vue](https://formulon.libraz.net/ja/cell/frameworks)

## ソースからビルド

```sh
yarn install
yarn dev:react   # React デモ（apps/react-demo）
yarn dev:vue     # Vue デモ（apps/vue-demo）
yarn build       # ライブラリ
yarn test        # コアの単体テスト
```

CI が実行するチェックは [`CONTRIBUTING.md`](./CONTRIBUTING.md) に、リリース手順は [`docs/releasing.md`](./docs/releasing.md) にあります。

## 埋め込み

埋め込んだシートに何を表示し、ユーザーに何をさせるかは、独立した3つのオプションで決めます。どれもあとから再マウントせずに変更できます（`update({ ui, policy, viewport })`）。

```ts
import { mount, Workbook, fixedFormPolicy, viewerPolicy } from '@libraz/formulon-cell';

const sheet = await mount(host, {
  workbook: await Workbook.createDefault(),
  ui: { profile: 'embedded', features: { formulaBar: true } }, // グリッドと数式バーだけ
  policy: fixedFormPolicy(['B2:B11']), // どのシートでも B2:B11 だけ入力可能
  viewport: { mode: 'fixed', range: 'A1:D12', tabNavigation: 'editable' },
});

sheet.update({ policy: viewerPolicy() }); // 送信後は選択とコピーだけ
```

- **`ui`** — 何を表示するか。プロファイル（`embedded`、`minimal`、`standard`、`full`）と、`sheetTabs`、`contextMenu`、`charts`、`clipboard` など約 30 個のスイッチで指定します。
- **`policy`** — ユーザーが何を変更できるか。値の編集、貼り付け、書式、行の挿入、並べ替えなど 27 種類の操作を個別に許可・拒否し、編集できるセルは A1 形式の範囲（`'B2:B11'` は全シート、`'Input!B2:B11'` はそのシートだけ）か述語で指定します。`restrict` フックは拒否を追加できますが、許可を足すことはできません。どの入力経路にも効くので、リボンを隠すことでセルが読み取り専用になるわけではありません。
- **`viewport`** — 表示・選択できるセルの範囲と、Tab で編集可能なセルだけを巡回するかどうか。
- **`contextMenu`** と **`overlays`** — 右クリックメニューの項目選択や差し替えと、ポップアップをホストのモーダルや全画面要素の中に収める設定。

policy を設定すると、操作の分類がないリボンコマンドは無効のままになります。制限付きのシートは、リボンを出さずに表示するか、ホスト側で用意した操作 UI と組み合わせてください。詳しくは [埋め込みと利用例](https://formulon.libraz.net/ja/cell/embedding) を参照してください。

## ドキュメント

ドキュメントは [formulon.libraz.net/ja/cell](https://formulon.libraz.net/ja/cell/) にあります。

- **学ぶ** — [デモ](https://formulon.libraz.net/ja/cell/demo) · [インストール](https://formulon.libraz.net/ja/cell/install) · [埋め込みと利用例](https://formulon.libraz.net/ja/cell/embedding)
- **ガイド** — [リボンツールバー](https://formulon.libraz.net/ja/cell/toolbar) · [モーダルとダイアログ](https://formulon.libraz.net/ja/cell/modals) · [グラフ](https://formulon.libraz.net/ja/cell/charts) · [テーマ](https://formulon.libraz.net/ja/cell/theming) · [国際化](https://formulon.libraz.net/ja/cell/i18n) · [言語の切り替え](https://formulon.libraz.net/ja/cell/language-switching) · [プラットフォームと Mac UI](https://formulon.libraz.net/ja/cell/platform) · [React / Vue](https://formulon.libraz.net/ja/cell/frameworks) · [バンドラの設定](https://formulon.libraz.net/ja/cell/bundler)
- **リファレンス** — [設定オプション](https://formulon.libraz.net/ja/cell/options) · [拡張](https://formulon.libraz.net/ja/cell/extensions) · [ホストとの連携](https://formulon.libraz.net/ja/cell/host-integration)

## 含まないもの（Non-goals）

formulon-cell は UI 層です。計算エンジンでもアプリケーションでもありません。数式の評価、再計算、xlsx の解析は [formulon](https://github.com/libraz/formulon) の担当です。サーバー、永続化、共同編集の通信、認証も持たず、これらはホスト側が用意します。React 版と Vue 版は薄いまま保ち、メニュー、コマンド、ダイアログはコアに置いているため、3 つのホストの挙動がずれることはありません。

## ライセンス

[Apache-2.0](LICENSE)
