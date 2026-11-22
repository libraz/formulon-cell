# @libraz/formulon-cell-vue

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![npm — core](https://img.shields.io/npm/v/@libraz/formulon-cell?label=core)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — react](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![Vue](https://img.shields.io/badge/Vue-3-blue?logo=vuedotjs)](https://vuejs.org/)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-blue)](https://formulon.libraz.net/ja/cell/)
[![Demo](https://img.shields.io/badge/demo-live-brightgreen)](https://formulon.libraz.net/ja/cell/demo)

**Vue 向けの Excel ライクな表計算コンポーネントです。** `<Spreadsheet>` と `SpreadsheetToolbar` は [`@libraz/formulon-cell`](https://www.npmjs.com/package/@libraz/formulon-cell) を包むコンポーネントで、Canvas 描画のグリッドとデスクトップ版 Excel に合わせた UI を、[formulon](https://github.com/libraz/formulon) の WASM エンジンの上で動かします。props の変更はコアの命令型 API で反映するため、キャンバスは再マウントされず、選択状態、フォーカス、イベント購読はそのまま残ります。

[デモ](https://formulon.libraz.net/ja/cell/demo) でブラウザ上ですぐに試せます。

グリッド、リボン、メニュー、コマンド、ダイアログはすべてコアにあり、このパッケージは Vue との接続だけを受け持ちます。

> **ベータ版（1.0 未満）です。** UI はデスクトップ版 Excel に合わせて作っており、機能ごとに突き合わせています。公開 API はマイナーバージョンの間でも変わることがあり、まだ突き合わせていない挙動は異なる場合があります。違いを見つけたら [issue](https://github.com/libraz/formulon-cell/issues) で知らせてください。

## インストール

```sh
npm install @libraz/formulon-cell-vue @libraz/formulon-cell vue
```

バンドラの設定は [プロジェクトの README](https://github.com/libraz/formulon-cell/blob/main/README_ja.md#バンドラの設定) を参照してください。

## クイックスタート

```vue
<script setup lang="ts">
import { Spreadsheet, type SpreadsheetInstance } from '@libraz/formulon-cell-vue';
import '@libraz/formulon-cell/styles.css';

const ready = (instance: SpreadsheetInstance) => {
  console.log('mounted', instance.workbook.engine);
};
</script>

<template>
  <Spreadsheet
    :ui="{ profile: 'full' }"
    locale="ja"
    style="width: 100%; height: 100vh"
    @ready="ready"
  />
</template>
```

## コンポーザブル

```ts
const sheetRef = ref<SpreadsheetExposed | null>(null);
const instance = computed(() => sheetRef.value?.spreadsheet ?? null);
const selection = useSelection(instance);
const { locale, strings } = useI18n(instance);
```

| コンポーザブル | 説明 |
|------------|------|
| `useSelection(instance)` | 現在の選択範囲 |
| `useSpreadsheet(instance, selector)` | インスタンスから導いた値。変更のたびに再計算される |
| `useI18n(instance)` | 現在のロケールと文言。実行時の切り替えに追従する |
| `useSpreadsheetEvent(instance, event, handler)` | `change`、`cell`、`recalc` などを購読する |

## リボンツールバー

`SpreadsheetToolbar` はコアの `mountToolbar` を Vue から使うためのアダプタです。`hooks` を渡すと、リボンを複製せずにリボンコマンドのホスト側の処理を差し込めます。

```vue
<script setup lang="ts">
import { SpreadsheetToolbar } from '@libraz/formulon-cell-vue';
import '@libraz/formulon-cell/styles/toolbar.css';
</script>

<template>
  <SpreadsheetToolbar :spreadsheet="instance" v-model:active-tab="tab" />
</template>
```

## 埋め込み

`ui`、`policy`、`viewport`、`context-menu`、`overlays`、`toolbar` の各 props は、再マウントせずに実行時に反映されます。`fixedFormPolicy` と `viewerPolicy` は `@libraz/formulon-cell` から import できます。ポリシーはどの入力経路にも効き、変更のバッチは `status` 付きの `change` イベントで受け取れます。

```vue
<Spreadsheet
  :ui="{ profile: 'embedded' }"
  :policy="fixedFormPolicy(['B2:C10'])"
  @change="save"
/>
```

表示範囲の制限、コンテキストメニューの組み立て、ホストからの更新は [埋め込みと利用例](https://formulon.libraz.net/ja/cell/embedding) を参照してください。policy を設定すると、操作の分類がないリボンコマンドは無効のままになります。制限付きのシートは、リボンを出さずに表示するか、ホスト側で用意した操作 UI と組み合わせてください。

## ドキュメント

- [デモ](https://formulon.libraz.net/ja/cell/demo)
- [React / Vue](https://formulon.libraz.net/ja/cell/frameworks)
- [設定オプション](https://formulon.libraz.net/ja/cell/options)
- [埋め込みと利用例](https://formulon.libraz.net/ja/cell/embedding)
- [モーダルと全画面表示](https://formulon.libraz.net/ja/cell/modals)
- [テーマ](https://formulon.libraz.net/ja/cell/theming)
- [言語と表示文言](https://formulon.libraz.net/ja/cell/i18n)
- [変更履歴](https://github.com/libraz/formulon-cell/blob/main/packages/formulon-cell/CHANGELOG.md)（このパッケージの変更もコアと同じファイルに記載しています）
- **関連** — [Formulon](https://formulon.libraz.net/ja/)：Excel 365 の計算エンジン（WASM、Python、ネイティブ） · [formulon-mcp](https://formulon.libraz.net/ja/mcp/)：AI エージェントにワークブック操作ツールを渡す MCP サーバー

## ライセンス

[Apache-2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
