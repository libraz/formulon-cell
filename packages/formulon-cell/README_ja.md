# @libraz/formulon-cell

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — react](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![npm — vue](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-6-blue?logo=typescript)](https://www.typescriptlang.org/)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-blue)](https://formulon.libraz.net/ja/cell/)
[![Demo](https://img.shields.io/badge/demo-live-brightgreen)](https://formulon.libraz.net/ja/cell/demo)

**フレームワークなしで、Web ページに表計算を組み込めます。** DOM 要素にマウントすると、Canvas 描画のグリッドとデスクトップ表計算ソフト風の UI（数式バー、リボン、シートタブ、コンテキストメニュー）が表示され、計算は [formulon](https://github.com/libraz/formulon) の WASM エンジンが受け持ちます。機能はプリセットと拡張ファクトリで組み合わせ、ロケールとテーマは再マウントせずに切り替えられます。

[デモ](https://formulon.libraz.net/ja/cell/demo) でブラウザ上ですぐに試せます。

このパッケージは Vanilla TypeScript / DOM のコアです。[`@libraz/formulon-cell-react`](https://www.npmjs.com/package/@libraz/formulon-cell-react) と [`@libraz/formulon-cell-vue`](https://www.npmjs.com/package/@libraz/formulon-cell-vue) は、同じマウント処理を包む薄いアダプタです。

> 基本的なブック操作は使えます。ダイアログ、細かなコントロールの挙動、キーボード操作、アクセシビリティは、まだデスクトップ表計算ソフトと同じにならない場合があります。現時点では完成したエンドユーザー向け製品として案内しないでください。

## インストール

```sh
npm install @libraz/formulon-cell zustand
```

`zustand` はピア依存です。標準の WASM エンジンは単一スレッド版で、COOP/COEP ヘッダは要りません。エンジンの初期化に失敗すると `WorkbookHandle.createDefault()` は reject します。インメモリのスタブは `preferStub: true` を渡したときだけ使われます。バンドラの設定は [プロジェクトの README](https://github.com/libraz/formulon-cell/blob/main/README_ja.md#バンドラの設定) を参照してください。

## クイックスタート

```ts
import { Spreadsheet, WorkbookHandle, presets } from '@libraz/formulon-cell';
import '@libraz/formulon-cell/styles.css';

const wb = await WorkbookHandle.createDefault();
const sheet = await Spreadsheet.mount(document.getElementById('sheet')!, {
  workbook: wb,
  features: presets.full(),
  locale: 'ja',
  toolbar: true, // リボンも同時にマウントし、sheet.toolbar で参照できる
});

sheet.i18n.setLocale('en');
sheet.setTheme('ink'); // グリッドとリボンのテーマをまとめて切り替え
```

グリッドだけの入力フォームや閲覧ビューにするときは、`ui: { profile: 'embedded' }` と、`fixedFormPolicy(ranges)` や `viewerPolicy()` などのポリシーを渡してマウントします。詳しくは [埋め込みと利用例](https://formulon.libraz.net/ja/cell/embedding) を参照してください。

## プリセット

| プリセット | 内容 |
|--------|------|
| `presets.minimal()`  | 数式バー、ステータスバー、基本キーマップ |
| `presets.standard()` | + クイック分析、コンテキストメニュー、検索と置換、クリップボード、書式のコピー/貼り付け |
| `presets.full()`     | + セルの書式設定、形式を選択して貼り付け、条件付き書式、名前の定義、ハイパーリンク、ピボットテーブルの作成、データの入力規則、オートコンプリート、コメントのホバー表示 |

## サブパスエクスポート

| import パス | 内容 |
|---|---|
| `@libraz/formulon-cell` | `Spreadsheet`、`WorkbookHandle`、`presets`、ポリシー、拡張ファクトリ |
| `@libraz/formulon-cell/extensions` | 拡張ファクトリ一式 |
| `@libraz/formulon-cell/extensions/*` | 個別の拡張（`statusBar`、`findReplace`、`contextMenu` など） |
| `@libraz/formulon-cell/i18n/ja`、`/i18n/en` | ロケール辞書 |
| `@libraz/formulon-cell/styles.css` | グリッド、ダイアログ、オーバーレイ、リボン、全テーマをまとめたスタイルシート |
| `@libraz/formulon-cell/styles/{tokens,paper,ink,contrast,toolbar}.css` | 個別に読み込む場合のスタイルシート |

## 主な API

| API | 説明 |
|-----|------|
| `Spreadsheet.mount(host, opts)` | DOM 要素に表計算をマウントする |
| `WorkbookHandle.createDefault()` | WASM エンジンを使うワークブックを作る |
| `workbook.withBatchedRecalc(fn)` | 複数セルへの書き込みを 1 回の再計算にまとめる |
| `instance.i18n.setLocale(loc)` / `instance.setTheme(theme)` | 実行時にロケールやテーマを切り替える |
| `instance.setPolicy` / `setViewportOptions` / `setContextMenu` / `setUi` / `setToolbar` | 埋め込み設定を実行時に変更する |
| `instance.applyChanges(changes)` | ユーザー向けポリシーを通さない、ホストからの更新 |
| `instance.on(event, handler)` | `cellChange`、`selectionChange`、`changeBatch`、`recalc` など |

オプションと API の一覧は [ドキュメントサイト](https://formulon.libraz.net/ja/cell/) にあります。

## ドキュメント

- [デモ](https://formulon.libraz.net/ja/cell/demo)
- [インストール](https://formulon.libraz.net/ja/cell/install)
- [設定オプション](https://formulon.libraz.net/ja/cell/options)
- [埋め込みと利用例](https://formulon.libraz.net/ja/cell/embedding)
- [モーダルと全画面表示](https://formulon.libraz.net/ja/cell/modals)
- [React / Vue](https://formulon.libraz.net/ja/cell/frameworks)
- [テーマ](https://formulon.libraz.net/ja/cell/theming)
- [言語と表示文言](https://formulon.libraz.net/ja/cell/i18n)

## ライセンス

[Apache-2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
