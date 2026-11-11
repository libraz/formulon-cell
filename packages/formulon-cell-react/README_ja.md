# @libraz/formulon-cell-react

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![npm — core](https://img.shields.io/npm/v/@libraz/formulon-cell?label=core)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — vue](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![React](https://img.shields.io/badge/React-18%2B-blue?logo=react)](https://react.dev/)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-blue)](https://formulon.libraz.net/ja/cell/)
[![Demo](https://img.shields.io/badge/demo-live-brightgreen)](https://formulon.libraz.net/ja/cell/demo)

**React のツリーに表計算を組み込めます。** `<Spreadsheet>` と `<SpreadsheetToolbar>` は [`@libraz/formulon-cell`](https://www.npmjs.com/package/@libraz/formulon-cell) を包むコンポーネントで、Canvas 描画のグリッドとデスクトップ版 Excel に合わせた UI を、[formulon](https://github.com/libraz/formulon) の WASM エンジンの上で動かします。props の変更はコアの命令型 API で反映するため、キャンバスは再マウントされず、選択状態、フォーカス、イベント購読はそのまま残ります。

[デモ](https://formulon.libraz.net/ja/cell/demo) でブラウザ上ですぐに試せます。

グリッド、リボン、メニュー、コマンド、ダイアログはすべてコアにあり、このパッケージは React との接続だけを受け持ちます。

> **ベータ版（1.0 未満）です。** UI はデスクトップ版 Excel に合わせて作っており、機能ごとに突き合わせています。公開 API はマイナーバージョンの間でも変わることがあり、まだ突き合わせていない挙動は異なる場合があります。違いを見つけたら [issue](https://github.com/libraz/formulon-cell/issues) で知らせてください。

## インストール

```sh
npm install @libraz/formulon-cell-react @libraz/formulon-cell react react-dom
```

バンドラの設定は [プロジェクトの README](https://github.com/libraz/formulon-cell/blob/main/README_ja.md#バンドラの設定) を参照してください。

## クイックスタート

```tsx
import { Spreadsheet } from '@libraz/formulon-cell-react';
import '@libraz/formulon-cell/styles.css';

export function MySheet() {
  return (
    <Spreadsheet
      style={{ width: '100%', height: '100vh' }}
      ui={{ profile: 'full' }}
      locale="ja"
      onReady={(instance) => console.log('mounted', instance.workbook.engine)}
    />
  );
}
```

`ref` から動作中のインスタンスを参照できます。

```tsx
const ref = useRef<SpreadsheetRef>(null);
ref.current?.spreadsheet?.history.undo();
```

## フック

| フック | 説明 |
|------|------|
| `useSelection(instance)` | 現在の選択範囲 |
| `useSpreadsheet(instance, selector)` | インスタンスから導いた値。変更のたびに再計算される |
| `useI18n(instance)` | 現在のロケールと文言。実行時の切り替えに追従する |
| `useSpreadsheetEvent(instance, event, handler)` | `change`、`cell`、`recalc` などを購読する |

## リボンツールバー

`SpreadsheetToolbar` はコアの `mountToolbar` を React から使うためのアダプタです。`hooks` を渡すと、リボンを複製せずにリボンコマンドのホスト側の処理を差し込めます。

```tsx
import { SpreadsheetToolbar } from '@libraz/formulon-cell-react';
import type { RibbonTab } from '@libraz/formulon-cell/toolbar';
import '@libraz/formulon-cell/styles/toolbar.css';

<SpreadsheetToolbar
  spreadsheet={instance}
  activeTab={tab}
  onTabChange={setTab}
  hooks={{ protection: { runSheet: openProtectDialog } }}
/>;
```

## 埋め込み

`ui`、`policy`、`viewport`、`contextMenu`、`overlays`、`toolbar` の各 props は、再マウントせずに実行時に反映されます。`fixedFormPolicy` と `viewerPolicy` は `@libraz/formulon-cell` から import できます。ポリシーはどの入力経路にも効き、変更のバッチは `status` 付きで `onChange` に届きます。

```tsx
<Spreadsheet
  ui={{ profile: 'embedded' }}
  policy={fixedFormPolicy(['B2:C10'])}
  onChange={(change) => save(change)}
/>
```

表示範囲の制限、コンテキストメニューの組み立て、ホストからの更新は [埋め込みと利用例](https://formulon.libraz.net/ja/cell/embedding) を参照してください。policy を設定すると、操作の分類がないリボンコマンドは無効のままになります。制限付きのシートは、リボンを出さずに表示するか、ホスト側で用意した操作 UI と組み合わせてください。

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
