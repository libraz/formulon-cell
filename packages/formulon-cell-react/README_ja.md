# @libraz/formulon-cell-react

[![npm version](https://img.shields.io/npm/v/@libraz/formulon-cell-react.svg)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![license](https://img.shields.io/npm/l/@libraz/formulon-cell-react.svg)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![bundle size](https://img.shields.io/bundlephobia/minzip/@libraz/formulon-cell-react)](https://bundlephobia.com/package/@libraz/formulon-cell-react)

[`@libraz/formulon-cell`](https://www.npmjs.com/package/@libraz/formulon-cell)
を React 18+ 向けにラップしたコンポーネントとフック。
[formulon](https://github.com/libraz/formulon) WASM 計算エンジン向けの
スプレッドシート UI です。

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

`SpreadsheetToolbar` は、React デモが採用しているリボン UI 表層を
そのまま再利用できるコンポーネントです。

```tsx
import { SpreadsheetToolbar, type RibbonTab } from '@libraz/formulon-cell-react';
import '@libraz/formulon-cell-react/toolbar.css';
```

## 実行時の props 更新

`theme`・`locale`・`strings`・`workbook`・`features`・`extensions` の各
プロパティは、コアの命令的 API を経由して稼働中のスプレッドシートに
反映されます。コンポーネントは **再マウントを行いません** ので、
選択範囲・フォーカス・ホスト側のイベント購読はそのまま維持されます。

## コアヘルパー

このパッケージは、コア側のコマンドヘルパーと型（`createSessionChart`・
`saveSheetView`・`activateSheetView`・`listDefinedNames`・
`upsertDefinedName` など）を再エクスポートしています。React アプリの
ホスト側 UI 表層に必要な型を、単一のインポート元から取り込めます。

## ドキュメント

完全な API リファレンスとバンドラ統合は
[プロジェクト README](https://github.com/libraz/formulon-cell/blob/main/README_ja.md)
を参照してください。

## ライセンス

[Apache License 2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
