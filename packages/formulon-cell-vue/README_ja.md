# @libraz/formulon-cell-vue

[![npm version](https://img.shields.io/npm/v/@libraz/formulon-cell-vue.svg)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![license](https://img.shields.io/npm/l/@libraz/formulon-cell-vue.svg)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![bundle size](https://img.shields.io/bundlephobia/minzip/@libraz/formulon-cell-vue)](https://bundlephobia.com/package/@libraz/formulon-cell-vue)

[`@libraz/formulon-cell`](https://www.npmjs.com/package/@libraz/formulon-cell)
の Vue 3 コンポーネント + コンポーザブル。
[formulon](https://github.com/libraz/formulon) WASM 計算エンジン向け
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
| `useI18n(instance)` | 現在のロケール + 文字列を取得 (ランタイム切替に追従) |

## ランタイム props 更新

`theme`、`locale`、`strings`、`workbook`、`features`、`extensions` は
コアの命令的 API 経由で動作中のスプレッドシートを更新します。コンポーネント
は **再マウントしません** — 選択・フォーカス・ホスト側イベント購読は
維持されます。

## コアヘルパー

このパッケージはコアのコマンドヘルパーと型 (`createSessionChart`、
`saveSheetView`、`activateSheetView`、`listDefinedNames`、
`upsertDefinedName` など) を再 export しています。Vue アプリの
ホストクロムを 1 つの import から型付けできます。

## ドキュメント

完全な API リファレンスとバンドラ統合は
[プロジェクト README](https://github.com/libraz/formulon-cell/blob/main/README_ja.md)
を参照してください。

## ライセンス

[Apache License 2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
