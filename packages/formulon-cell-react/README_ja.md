# @libraz/formulon-cell-react

[![npm version](https://img.shields.io/npm/v/@libraz/formulon-cell-react.svg)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![license](https://img.shields.io/npm/l/@libraz/formulon-cell-react.svg)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![bundle size](https://img.shields.io/bundlephobia/minzip/@libraz/formulon-cell-react)](https://bundlephobia.com/package/@libraz/formulon-cell-react)

[`@libraz/formulon-cell`](https://www.npmjs.com/package/@libraz/formulon-cell)
の React 18+ コンポーネント + フック。[formulon](https://github.com/libraz/formulon)
WASM 計算エンジン向けスプレッドシート UI です。

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
| `useI18n(instance)` | 現在のロケール + 文字列を取得 (ランタイム切替に追従) |

## ランタイム props 更新

`theme`、`locale`、`strings`、`workbook`、`features`、`extensions` は
コアの命令的 API 経由で動作中のスプレッドシートを更新します。コンポーネント
は **再マウントしません** — 選択・フォーカス・ホスト側イベント購読は
維持されます。

## コアヘルパー

このパッケージはコアのコマンドヘルパーと型 (`createSessionChart`、
`saveSheetView`、`activateSheetView`、`listDefinedNames`、
`upsertDefinedName` など) を再 export しています。React アプリの
ホストクロムを 1 つの import から型付けできます。

## ドキュメント

完全な API リファレンスとバンドラ統合は
[プロジェクト README](https://github.com/libraz/formulon-cell/blob/main/README_ja.md)
を参照してください。

## ライセンス

[Apache License 2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
