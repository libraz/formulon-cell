# @libraz/formulon-cell-vue

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![npm — core](https://img.shields.io/npm/v/@libraz/formulon-cell?label=core)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — react](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![Vue](https://img.shields.io/badge/Vue-3-blue?logo=vuedotjs)](https://vuejs.org/)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-blue)](https://formulon.libraz.net/ja/cell/)
[![Demo](https://img.shields.io/badge/demo-live-brightgreen)](https://formulon.libraz.net/ja/cell/demo)

**Vue アプリに表計算を組み込めます。** `<Spreadsheet>` と `SpreadsheetToolbar` は [`@libraz/formulon-cell`](https://www.npmjs.com/package/@libraz/formulon-cell) を包むコンポーネントで、Canvas 描画のグリッドとデスクトップ表計算ソフト風の UI を、[formulon](https://github.com/libraz/formulon) の WASM エンジンの上で動かします。props の変更はコアの命令型 API で反映するため、キャンバスは再マウントされず、選択状態、フォーカス、イベント購読はそのまま残ります。

[デモ](https://formulon.libraz.net/ja/cell/demo) でブラウザ上ですぐに試せます。

グリッド、リボン、メニュー、コマンド、ダイアログはすべてコアにあり、このパッケージは Vue との接続だけを受け持ちます。

## インストール

```sh
npm install @libraz/formulon-cell-vue @libraz/formulon-cell vue zustand
```

バンドラの設定は [プロジェクトの README](https://github.com/libraz/formulon-cell/blob/main/README_ja.md#バンドラの設定) を参照してください。

## クイックスタート

```vue
<script setup lang="ts">
import { Spreadsheet, presets, type SpreadsheetInstance } from '@libraz/formulon-cell-vue';
import '@libraz/formulon-cell/styles.css';

const ready = (instance: SpreadsheetInstance) => {
  console.log('mounted', instance.workbook.version);
};
</script>

<template>
  <Spreadsheet
    :features="presets.full()"
    locale="ja"
    style="width: 100%; height: 100vh"
    @ready="ready"
  />
</template>
```

## コンポーザブル

```ts
const sheetRef = ref<SpreadsheetExposed | null>(null);
const instance = computed(() => sheetRef.value?.instance.value ?? null);
const selection = useSelection(instance);
const { locale, strings } = useI18n(instance);
```

| コンポーザブル | 説明 |
|------------|------|
| `useSelection(instance)` | 現在の選択範囲 |
| `useSpreadsheet(instance, selector, fallback)` | ストアの状態から任意の値を取り出す |
| `useI18n(instance)` | 現在のロケールと文言。実行時の切り替えに追従する |
| `useSpreadsheetEvent(instance, event, handler)` | `cellChange`、`selectionChange`、`recalc` などを購読する |

## リボンツールバー

`SpreadsheetToolbar` は SFC のサブパスとして配布しているので、アプリ側の Vue ビルドでそのままコンパイルされます。中身はコアの `Spreadsheet.mountToolbar` のアダプタで、`dropdownActions` を渡すとリボンを複製せずに個別のドロップダウン処理だけを差し替えられます。

```vue
<script setup lang="ts">
import SpreadsheetToolbar from '@libraz/formulon-cell-vue/toolbar.vue';
import '@libraz/formulon-cell-vue/toolbar.css';
</script>

<template>
  <SpreadsheetToolbar :instance="instance" :active-tab="tab" locale="ja" @tab-change="tab = $event" />
</template>
```

## 制限付きの埋め込み

`ui`、`policy`、`viewport`、`context-menu`、`overlays`、`toolbar` の各 props は、再マウントせずに実行時に反映されます。`fixedFormPolicy` と `viewerPolicy` はこのパッケージからも import でき、成功したバッチ更新は `change-batch` イベントで受け取れます。

```vue
<Spreadsheet
  :ui="{ profile: 'embedded' }"
  :policy="fixedFormPolicy([{ sheet: 0, r0: 1, c0: 1, r1: 9, c1: 2 }])"
  @change-batch="save"
/>
```

表示範囲の制限、コンテキストメニューの組み立て、ホストからの更新は [埋め込みと利用例](https://formulon.libraz.net/ja/cell/embedding) を参照してください。

## ドキュメント

- [デモ](https://formulon.libraz.net/ja/cell/demo)
- [React / Vue](https://formulon.libraz.net/ja/cell/frameworks)
- [設定オプション](https://formulon.libraz.net/ja/cell/options)
- [埋め込みと利用例](https://formulon.libraz.net/ja/cell/embedding)
- [モーダルと全画面表示](https://formulon.libraz.net/ja/cell/modals)
- [テーマ](https://formulon.libraz.net/ja/cell/theming)
- [言語と表示文言](https://formulon.libraz.net/ja/cell/i18n)

## ライセンス

[Apache-2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
