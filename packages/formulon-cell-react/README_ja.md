# @libraz/formulon-cell-react

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![npm — core](https://img.shields.io/npm/v/@libraz/formulon-cell?label=core)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — vue](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![React](https://img.shields.io/badge/React-18%2B-blue?logo=react)](https://react.dev/)

**React のツリーの中に、動く表計算をそのまま置けます。** `<Spreadsheet>` と
`<SpreadsheetToolbar>` は
[`@libraz/formulon-cell`](https://www.npmjs.com/package/@libraz/formulon-cell)
をラップしたもので、[formulon](https://github.com/libraz/formulon) の WASM
計算エンジンの上に、Canvas 描画のグリッドとデスクトップ表計算ソフト風の UI 表層を
提供します。props の変更はキャンバスを再マウントせず、コアの命令的 API を通して
稼働中のインスタンスへ反映されます。

本パッケージは薄いアダプタです。グリッド、リボン、メニュー、コマンド、
ダイアログはすべてコアにあるため、React 版と Vue 版の実装がずれることは
ありません。

## インストール

```sh
npm install @libraz/formulon-cell-react @libraz/formulon-cell react react-dom zustand
```

## クイックスタート

```tsx
import { Spreadsheet, presets } from '@libraz/formulon-cell-react';
import '@libraz/formulon-cell/styles.css';

export function MySheet() {
  return (
    <Spreadsheet
      style={{ width: '100%', height: '100vh' }}
      features={presets.full()}
      locale="ja"
      onReady={(inst) => {
        console.log('mounted', inst.workbook.version);
      }}
    />
  );
}
```

## 命令的 ref

```tsx
import { useRef } from 'react';
import { Spreadsheet, type SpreadsheetRef } from '@libraz/formulon-cell-react';

const ref = useRef<SpreadsheetRef>(null);
ref.current?.instance?.undo();
```

## フック

| フック | 説明 |
|------|------|
| `useSelection(instance)` | アクティブな選択範囲を購読 |
| `useI18n(instance)` | 現在のロケールと文字列を取得（実行時の切替に追従） |

## ツールバー

`SpreadsheetToolbar` は、コアの `Spreadsheet.mountToolbar` に対する薄い
アダプタです。リボンの DOM、メニューファクトリ、アクティベーションモデル、
動的ドロップダウンのディスパッチャは `@libraz/formulon-cell` に集約されており、
React 側が別個のリボン実装を持つことはありません。

ホスト側の監査や独自の UI 表層では、`ribbonActivationEntries`、
`ribbonSurfaceCommandIds`、`DYNAMIC_RIBBON_DROPDOWN_HANDLER_ATTRS`、
`attachRangePickerButton`、`appendConditionalApplyFormatControls`、
`conditionalStyleOptions`、`showReport`、`reportDialogLabels`、
`projectDisabledReason`、`projectDisabledState` といったコアのエクスポートを
使ってください。リボンのコマンド一覧、Excel 風のダイアログやレポート、
無効／読み取り専用の理由の射影を React 側で作り直す必要はありません。

```tsx
import { SpreadsheetToolbar, type RibbonTab } from '@libraz/formulon-cell-react';
import '@libraz/formulon-cell-react/toolbar.css';
```

個別のドロップダウンの動作だけ差し替えたい場合は、リボンをフォークせず
`dropdownActions` を使ってください。

```tsx
<SpreadsheetToolbar
  instance={instance}
  activeTab="home"
  locale="ja"
  onTabChange={setActiveTab}
  dropdownActions={{ applyProtectAction: openProtectDialog }}
/>
```

## 実行時の props 更新

`theme`・`locale`・`strings`・`workbook`・`features`・`extensions`・
`printerProfiles`・`printerProfileId`・`uploadStatus`・`macroRecording` の各
プロパティは、コアの命令的 API を経由して稼働中のスプレッドシートに
反映されます。コンポーネントは **再マウントを行いません** ので、
選択範囲・フォーカス・ホスト側のイベント購読はそのまま維持されます。

ホストにしか持てない機能も、リボンの挙動を React 側で作り直すことなく
props として渡せます。

```tsx
<Spreadsheet
  captureScreenClip={async () => ({
    src: await nativeCaptureRegionAsDataUrl(),
    alt: '画面の領域切り取り',
  })}
  refreshPrinterProfiles={() => nativeListPrinterProfiles()}
/>
```

`captureScreenClip` は「挿入 > スクリーンショット > 画面の領域切り取り」を
担い、プリンタープロファイル系の props はページ設定と印刷プレビューの
最小余白の扱いに使われます。

## コアヘルパー

このパッケージは、コア側のコマンドヘルパーと型（`createSessionChart`・
`saveSheetView`・`activateSheetView`・`listDefinedNames`・
`upsertDefinedName`・`ribbonActivationEntries`・`attachRangePickerButton`・
`appendConditionalApplyFormatControls`・`conditionalStyleOptions`・
`showReport`・`reportDialogLabels`・`projectDisabledReason`・
`projectDisabledState`・`ScreenClipCapture`・`ScreenClipResult` など）を
再エクスポートしています。React アプリのホスト側 UI 表層に必要な型を、
単一のインポート元から取り込めます。

## ドキュメント

完全な API リファレンスとバンドラ統合は
[プロジェクト README](https://github.com/libraz/formulon-cell/blob/main/README_ja.md)
を参照してください。

## 併せて使えるパッケージ

```sh
npm install @libraz/formulon-cell       # Vanilla TypeScript / DOM コア
npm install @libraz/formulon-cell-vue   # Vue 3 コンポーネント + コンポーザブル
```

## ライセンス

[Apache-2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
