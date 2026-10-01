# @libraz/formulon-cell-vue

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![npm — core](https://img.shields.io/npm/v/@libraz/formulon-cell?label=core)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — react](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![Vue](https://img.shields.io/badge/Vue-3-blue?logo=vuedotjs)](https://vuejs.org/)

**Vue アプリの中に、動く表計算をそのまま置けます。** `<Spreadsheet>` と
`SpreadsheetToolbar` は
[`@libraz/formulon-cell`](https://www.npmjs.com/package/@libraz/formulon-cell)
をラップしたもので、[formulon](https://github.com/libraz/formulon) の WASM
計算エンジンの上に、Canvas 描画のグリッドとデスクトップ表計算ソフト風の UI 表層を
提供します。props の変更はキャンバスを再マウントせず、コアの命令的 API を通して
稼働中のインスタンスへ反映されます。

本パッケージは薄いアダプタです。グリッド、リボン、メニュー、コマンド、
ダイアログはすべてコアにあるため、Vue 版と React 版の実装がずれることは
ありません。

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

`SpreadsheetToolbar` は SFC のサブパスとして公開しており、Vue のバンドラが
アプリ本体のコンポーネントと同じパイプラインでコンパイルできます。実体は
コアの `Spreadsheet.mountToolbar` に対する薄いアダプタで、リボンの DOM、
メニューファクトリ、アクティベーションモデル、動的ドロップダウンの
ディスパッチャは `@libraz/formulon-cell` に集約されています。

ホスト側の監査や独自の UI 表層では、`ribbonActivationEntries`、
`ribbonSurfaceCommandIds`、`DYNAMIC_RIBBON_DROPDOWN_HANDLER_ATTRS`、
`attachRangePickerButton`、`appendConditionalApplyFormatControls`、
`conditionalStyleOptions`、`showReport`、`reportDialogLabels`、
`projectDisabledReason`、`projectDisabledState` といったコアのエクスポートを
使ってください。リボンのコマンド一覧、Excel 風のダイアログやレポート、
無効／読み取り専用の理由の射影を Vue 側で作り直す必要はありません。

```vue
<script setup lang="ts">
import { type RibbonTab } from '@libraz/formulon-cell-vue';
import SpreadsheetToolbar from '@libraz/formulon-cell-vue/toolbar.vue';
import '@libraz/formulon-cell-vue/toolbar.css';
</script>
```

個別のドロップダウンの動作だけ差し替えたい場合は、リボンをフォークせず
`dropdownActions` を使ってください。

## 実行時の props 更新

`theme`・`locale`・`strings`・`workbook`・`features`・`extensions`・
`printerProfiles`・`printerProfileId`・`uploadStatus`・`macroRecording` の各
プロパティは、コアの命令的 API を経由して稼働中のスプレッドシートに
反映されます。コンポーネントは **再マウントを行いません** ので、
選択範囲・フォーカス・ホスト側のイベント購読はそのまま維持されます。

ホストにしか持てない機能も、リボンの挙動を Vue 側で作り直すことなく
props として渡せます。

```vue
<template>
  <Spreadsheet
    :capture-screen-clip="captureScreenClip"
    :refresh-printer-profiles="refreshPrinterProfiles"
  />
</template>
```

`captureScreenClip` は「挿入 > スクリーンショット > 画面の領域切り取り」を
担い、プリンタープロファイル系の props はページ設定と印刷プレビューの
最小余白の扱いに使われます。

## 制限付きの埋め込み

コンポーネントは `policy`・`viewport`・`contextMenu`・`overlays`・`toolbar` props を受け取り、実行中の変更を再マウントせずに反映します。React では `ui={{ profile: 'embedded' }}`、Vue では `:ui="{ profile: 'embedded' }"` でグリッドだけを表示できます。`toolbar={false}` / `:toolbar="false"` は UI プロファイルより優先します。`fixedFormPolicy` と `viewerPolicy` はこのパッケージからも import できます。

更新成功後の hook は、React の `onChangeBatch`、Vue の `@change-batch` で受け取れます。その他の hook には instance のイベント API を使えます。

編集可能範囲、ホストからの更新、独自メニュー、対応する操作の制限は [core の埋め込み例](https://github.com/libraz/formulon-cell#restricted-embedding)を参照してください。

## コアヘルパー

このパッケージは、コア側のコマンドヘルパーと型（`createSessionChart`・
`saveSheetView`・`activateSheetView`・`listDefinedNames`・
`upsertDefinedName`・`ribbonActivationEntries`・`attachRangePickerButton`・
`appendConditionalApplyFormatControls`・`conditionalStyleOptions`・
`showReport`・`reportDialogLabels`・`projectDisabledReason`・
`projectDisabledState`・`ScreenClipCapture`・`ScreenClipResult` など）を
再エクスポートしています。Vue アプリのホスト側 UI 表層に必要な型を、
単一のインポート元から取り込めます。

## ドキュメント

完全な API リファレンスとバンドラ統合は
[プロジェクト README](https://github.com/libraz/formulon-cell/blob/main/README_ja.md)
を参照してください。

## 併せて使えるパッケージ

```sh
npm install @libraz/formulon-cell         # Vanilla TypeScript / DOM コア
npm install @libraz/formulon-cell-react   # React 18+ コンポーネント + フック
```

## ライセンス

[Apache-2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
