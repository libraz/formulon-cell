# @libraz/formulon-cell-vue

[![npm version](https://img.shields.io/npm/v/@libraz/formulon-cell-vue.svg)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![license](https://img.shields.io/npm/l/@libraz/formulon-cell-vue.svg)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![bundle size](https://img.shields.io/bundlephobia/minzip/@libraz/formulon-cell-vue)](https://bundlephobia.com/package/@libraz/formulon-cell-vue)

[`@libraz/formulon-cell`](https://www.npmjs.com/package/@libraz/formulon-cell)
を Vue 3 向けにラップしたコンポーネントとコンポーザブル。
[formulon](https://github.com/libraz/formulon) WASM 計算エンジン向けの
スプレッドシート UI です。

## インストール

```sh
npm install @libraz/formulon-cell-vue @libraz/formulon-cell vue zustand
```

## クイックスタート

```vue
<script setup lang="ts">
import { Spreadsheet, presets, type SpreadsheetInstance } from '@libraz/formulon-cell-vue';
import '@libraz/formulon-cell/styles.css';

const ready = (inst: SpreadsheetInstance) => {
  console.log('mounted', inst.workbook.version);
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
import { computed, ref } from 'vue';
import { type SpreadsheetExposed, useI18n, useSelection } from '@libraz/formulon-cell-vue';

const sheetRef = ref<SpreadsheetExposed | null>(null);
const instance = computed(() => sheetRef.value?.instance.value ?? null);
const sel = useSelection(instance);
const { locale, strings } = useI18n(instance);
```

| コンポーザブル | 説明 |
|--------------|------|
| `useSelection(instance)` | アクティブな選択範囲を購読 |
| `useI18n(instance)` | 現在のロケールと文字列を取得（実行時の切替に追従） |

## ツールバー

`SpreadsheetToolbar` は SFC のサブパスとして公開されており、Vue のバンドラ
がアプリ本体のコンポーネントと同じパイプラインでコンパイルできます。

```vue
<script setup lang="ts">
import { type RibbonTab } from '@libraz/formulon-cell-vue';
import SpreadsheetToolbar from '@libraz/formulon-cell-vue/toolbar.vue';
import '@libraz/formulon-cell-vue/toolbar.css';
</script>
```

## 実行時の props 更新

`theme`・`locale`・`strings`・`workbook`・`features`・`extensions` の各
プロパティは、コアの命令的 API を経由して稼働中のスプレッドシートに
反映されます。コンポーネントは **再マウントを行いません** ので、
選択範囲・フォーカス・ホスト側のイベント購読はそのまま維持されます。

## コアヘルパー

このパッケージは、コア側のコマンドヘルパーと型（`createSessionChart`・
`saveSheetView`・`activateSheetView`・`listDefinedNames`・
`upsertDefinedName` など）を再エクスポートしています。Vue アプリの
ホスト側 UI 表層に必要な型を、単一のインポート元から取り込めます。

## ドキュメント

完全な API リファレンスとバンドラ統合は
[プロジェクト README](https://github.com/libraz/formulon-cell/blob/main/README_ja.md)
を参照してください。

## ライセンス

[Apache License 2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
