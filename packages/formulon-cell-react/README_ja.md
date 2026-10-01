# @libraz/formulon-cell-react

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![npm — core](https://img.shields.io/npm/v/@libraz/formulon-cell?label=core)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — vue](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![React](https://img.shields.io/badge/React-18%2B-blue?logo=react)](https://react.dev/)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-blue)](https://formulon.libraz.net/ja/cell/)
[![Demo](https://img.shields.io/badge/demo-live-brightgreen)](https://formulon.libraz.net/ja/cell/demo)

**React のツリーに表計算を組み込めます。** `<Spreadsheet>` と `<SpreadsheetToolbar>` は [`@libraz/formulon-cell`](https://www.npmjs.com/package/@libraz/formulon-cell) を包むコンポーネントで、Canvas 描画のグリッドとデスクトップ表計算ソフト風の UI を、[formulon](https://github.com/libraz/formulon) の WASM エンジンの上で動かします。props の変更はコアの命令型 API で反映するため、キャンバスは再マウントされず、選択状態、フォーカス、イベント購読はそのまま残ります。

[デモ](https://formulon.libraz.net/ja/cell/demo) でブラウザ上ですぐに試せます。

グリッド、リボン、メニュー、コマンド、ダイアログはすべてコアにあり、このパッケージは React との接続だけを受け持ちます。

## インストール

```sh
npm install @libraz/formulon-cell-react @libraz/formulon-cell react react-dom zustand
```

バンドラの設定は [プロジェクトの README](https://github.com/libraz/formulon-cell/blob/main/README_ja.md#バンドラの設定) を参照してください。

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
      onReady={(instance) => console.log('mounted', instance.workbook.version)}
    />
  );
}
```

`ref` から動作中のインスタンスを参照できます。

```tsx
const ref = useRef<SpreadsheetRef>(null);
ref.current?.instance?.undo();
```

## フック

| フック | 説明 |
|------|------|
| `useSelection(instance)` | 現在の選択範囲 |
| `useSpreadsheet(instance, selector, fallback)` | ストアの状態から任意の値を取り出す |
| `useI18n(instance)` | 現在のロケールと文言。実行時の切り替えに追従する |
| `useSpreadsheetEvent(instance, event, handler)` | `cellChange`、`selectionChange`、`recalc` などを購読する |

## リボンツールバー

`SpreadsheetToolbar` はコアの `Spreadsheet.mountToolbar` を React から使うためのアダプタです。`dropdownActions` を渡すと、リボンを複製せずに個別のドロップダウン処理だけを差し替えられます。

```tsx
import { SpreadsheetToolbar, type RibbonTab } from '@libraz/formulon-cell-react';
import '@libraz/formulon-cell-react/toolbar.css';

<SpreadsheetToolbar
  instance={instance}
  activeTab={tab}
  onTabChange={setTab}
  locale="ja"
  dropdownActions={{ applyProtectAction: openProtectDialog }}
/>;
```

## 制限付きの埋め込み

`ui`、`policy`、`viewport`、`contextMenu`、`overlays`、`toolbar` の各 props は、再マウントせずに実行時に反映されます。`fixedFormPolicy` と `viewerPolicy` はこのパッケージからも import でき、成功したバッチ更新は `onChangeBatch` で受け取れます。

```tsx
<Spreadsheet
  ui={{ profile: 'embedded' }}
  policy={fixedFormPolicy([{ sheet: 0, r0: 1, c0: 1, r1: 9, c1: 2 }])}
  onChangeBatch={(batch) => save(batch)}
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
