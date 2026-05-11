# @libraz/formulon-cell-vue

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![npm downloads](https://img.shields.io/npm/dm/@libraz/formulon-cell-vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-6-blue?logo=typescript)](https://www.typescriptlang.org/)

[`@libraz/formulon-cell`](../formulon-cell/README_ja.md) の Vue 3
コンポーネント + コンポーザブル。

## インストール

```sh
npm install @libraz/formulon-cell-vue @libraz/formulon-cell vue zustand
```

## クイックスタート

```vue
<script setup lang="ts">
import { ref } from 'vue';
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

## コアヘルパー

このパッケージは `createSessionChart`、`saveSheetView`、`activateSheetView`、
`listDefinedNames`、`upsertDefinedName` などコアのコマンドヘルパーと型を
再 export しているため、Vue アプリのホストクロムを 1 つの import から
型付けできます。

## ランタイム props 更新

`theme`、`locale`、`strings`、`workbook`、`features`、`extensions` は
コアの命令的 API 経由で動作中のスプレッドシートを更新します。これらの
prop が変わってもキャンバスは再マウントされないため、選択・フォーカス・
ホスト側のイベント購読が維持されます。

## ライセンス

[Apache-2.0](LICENSE)
