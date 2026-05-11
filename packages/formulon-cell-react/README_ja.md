# @libraz/formulon-cell-react

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![npm downloads](https://img.shields.io/npm/dm/@libraz/formulon-cell-react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-6-blue?logo=typescript)](https://www.typescriptlang.org/)

[`@libraz/formulon-cell`](../formulon-cell/README_ja.md) の React 18+
コンポーネント + フック。

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

function App() {
  const ref = useRef<SpreadsheetRef>(null);
  return (
    <>
      <button onClick={() => ref.current?.instance?.undo()}>元に戻す</button>
      <Spreadsheet ref={ref} />
    </>
  );
}
```

## フック

```tsx
import { useI18n, useSelection } from '@libraz/formulon-cell-react';

const sel = useSelection(instance);
const { locale, strings } = useI18n(instance);
```

## コアヘルパー

このパッケージは `createSessionChart`、`saveSheetView`、`activateSheetView`、
`listDefinedNames`、`upsertDefinedName` などコアのコマンドヘルパーと型を
再 export しているため、React アプリのホストクロムを 1 つの import から
型付けできます。

## ランタイム props 更新

`theme`、`locale`、`strings`、`workbook`、`features`、`extensions` は
コアの命令的 API 経由で動作中のスプレッドシートを更新します。これらの
prop が変わってもキャンバスは再マウントされないため、選択・フォーカス・
ホスト側のイベント購読が維持されます。

## ライセンス

[Apache-2.0](LICENSE)
