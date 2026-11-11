# @libraz/formulon-cell

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — react](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![npm — vue](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-6-blue?logo=typescript)](https://www.typescriptlang.org/)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-blue)](https://formulon.libraz.net/ja/cell/)
[![Demo](https://img.shields.io/badge/demo-live-brightgreen)](https://formulon.libraz.net/ja/cell/demo)

**フレームワークなしで、Web ページに表計算を組み込めます。** DOM 要素にマウントすると、Canvas 描画のグリッドと、デスクトップ版 Excel に合わせた UI（数式バー、リボン、シートタブ、コンテキストメニュー）が表示され、計算は [formulon](https://github.com/libraz/formulon) の WASM エンジンが受け持ちます。機能は UI プロファイルと拡張ファクトリで組み合わせ、ロケールとテーマは再マウントせずに切り替えられます。

[デモ](https://formulon.libraz.net/ja/cell/demo) でブラウザ上ですぐに試せます。

このパッケージは Vanilla TypeScript / DOM のコアです。[`@libraz/formulon-cell-react`](https://www.npmjs.com/package/@libraz/formulon-cell-react) と [`@libraz/formulon-cell-vue`](https://www.npmjs.com/package/@libraz/formulon-cell-vue) は、同じマウント処理を包む薄いアダプタです。

> **ベータ版（1.0 未満）です。** UI はデスクトップ版 Excel に合わせて作っており、機能ごとに突き合わせています。公開 API はマイナーバージョンの間でも変わることがあり、まだ突き合わせていない挙動は異なる場合があります。違いを見つけたら [issue](https://github.com/libraz/formulon-cell/issues) で知らせてください。

## インストール

```sh
npm install @libraz/formulon-cell
```

標準の WASM エンジンは単一スレッド版で、COOP/COEP ヘッダは要りません。エンジンの初期化に失敗すると `Workbook.createDefault()` は reject します。インメモリのスタブは `engine: 'stub'` を渡したときだけ使われます。バンドラの設定は [プロジェクトの README](https://github.com/libraz/formulon-cell/blob/main/README_ja.md#バンドラの設定) を参照してください。

## クイックスタート

```ts
import { mount, Workbook } from '@libraz/formulon-cell';
import { ribbon } from '@libraz/formulon-cell/toolbar';
import '@libraz/formulon-cell/styles.css';

const wb = await Workbook.createDefault();
const sheet = await mount(document.getElementById('sheet')!, {
  workbook: wb,
  ui: { profile: 'full', features: { localeSwitcher: true } },
  locale: 'ja',
  toolbar: ribbon(), // リボンも同時にマウントし、sheet.toolbar で参照できる
});

sheet.update({ locale: 'en' });
sheet.update({ theme: 'ink' }); // グリッドとリボンのテーマをまとめて切り替え
```

`host: { workbookName: () => currentFilename }` を指定すると、ウォッチウィンドウに現在のブック名を表示します。React / Vueコンポーネントも同じオプションを受け取ります。

言語選択UIは初期状態では非表示です。`ui.features.localeSwitcher: true` で表示し、`sheet.update({ ui: { features: { localeSwitcher: false } } })` で動的に非表示へ戻せます（`ui` は丸ごと置き換わります）。`update({ locale })` は表示辞書と計算プロファイルを同時に切り替えます。切り替えが拒否された場合は `FormulonError` を投げ、現在の言語を維持します。ダイアログ、検索・置換、コメント編集、フィルターメニューを閉じてから切り替えてください。拒否時は入力途中の内容を保持します。

## 埋め込み

`ui` で何を表示するか、`policy` でユーザーが何を変更できるか、`viewport` でどのセルを表示・選択できるかを決めます。どれも `update({ ui, policy, viewport })` で実行時に変更できます。

```ts
import { fixedFormPolicy, mount, viewerPolicy } from '@libraz/formulon-cell';

const form = await mount(host, {
  workbook: wb,
  ui: { profile: 'embedded', features: { formulaBar: true } }, // グリッドと数式バーだけ
  policy: fixedFormPolicy(['B2:B11']), // どのシートでも B2:B11 だけ入力可能
  viewport: { mode: 'fixed', range: 'A1:D12', tabNavigation: 'editable' },
});

form.update({ policy: viewerPolicy() }); // 送信後は選択とコピーだけ
```

ポリシーは 27 種類の操作を個別に許可・拒否し、編集できるセルは A1 形式の範囲（`'B2:B11'` は全シート、`'Input!B2:B11'` はそのシートだけ）か述語で指定します。エディタ、クリップボード、フィルハンドル、コンテキストメニュー、リボン、インスタンス API のどの入力経路にも効きます。`restrict` フックは拒否を追加できますが、許可を足すことはできません。policy を設定すると、操作の分類がないリボンコマンドは無効のままになります。制限付きのシートは、リボンを出さずに表示するか、ホスト側で用意した操作 UI と組み合わせてください。詳しくは [埋め込みと利用例](https://formulon.libraz.net/ja/cell/embedding) を参照してください。

## UI プロファイル

`ui.profile` で機能の出発点を選び（既定は `'full'`）、`ui.features` で個別の機能を切り替えます。

| プロファイル | 内容 |
|--------|------|
| `embedded`  | グリッドだけを表示し、ほかの機能、リボン、印刷はすべてオフ |
| `minimal`   | 数式バー、ステータスバー、基本キーマップ |
| `standard`  | + クイック分析、コンテキストメニュー、検索と置換、クリップボード、書式のコピー/貼り付け |
| `full`      | + セルの書式設定、形式を選択して貼り付け、条件付き書式、名前の定義、ハイパーリンク、ピボットテーブルの作成、データの入力規則、オートコンプリート、コメントのホバー表示 |

## サブパスエクスポート

| import パス | 内容 |
|---|---|
| `@libraz/formulon-cell` | `mount`、`Workbook`、ポリシー、公開型 |
| `@libraz/formulon-cell/toolbar` | `ribbon()` とリボンのモデル |
| `@libraz/formulon-cell/charts` | グラフのレイアウト用ヘルパー |
| `@libraz/formulon-cell/extensions` | 拡張ファクトリ一式 |
| `@libraz/formulon-cell/i18n/{en,ja,de,fr,zh,ko,th}` | 7言語のロケール辞書 |
| `@libraz/formulon-cell/styles.css` | グリッド、ダイアログ、オーバーレイ、リボン、全テーマをまとめたスタイルシート |
| `@libraz/formulon-cell/styles/{tokens,paper,ink,contrast,toolbar}.css` | 個別に読み込む場合のスタイルシート |

## 主な API

| API | 説明 |
|-----|------|
| `mount(host, options)` | DOM 要素に表計算をマウントする |
| `Workbook.createDefault()` / `Workbook.load(bytes)` | WASM エンジンを使うワークブックを作る、または開く |
| `workbook.batch(fn)` | 複数セルへの書き込みを 1 回の再計算にまとめる |
| `instance.update(patch)` | `locale`、`theme`、`ui`、`policy`、`viewport`、`contextMenu`、`toolbar` などのオプションを実行時に変更する |
| `instance.applyChanges(changes, { history })` | ポリシーとシート保護で判定する、全部か無しかのセル一括更新 |
| `instance.on(event, handler)` | `change`、`cell`、`workbook`、`recalc`、`error`、`update` |

オプションと API の一覧は [ドキュメントサイト](https://formulon.libraz.net/ja/cell/) にあります。

CSV出力にはロケールと区切り文字を指定できます。例: `wb.sheet(0).exportCsv({ locale: 'de', delimiter: ';' })`。取り込みでは `wb.sheet(0).importCsv(text, { at: 'A1', locale: 'de', delimiter: ';' })` に同じ設定を渡すと、数値と真偽値を同じ規則で解析します。出力の既定値は、ワークブックのロケールとカンマ区切りです。取り込みは `locale` が未指定の場合、現在のワークブックのプロファイルを使います。

数式の編集と外部クリップボードからの入力には、現在のプロファイルの構文を使います。たとえば独語の `=SUMME(1,5;2)` は `=SUM(1.5,2)` として保存されます。コードから呼ぶ `Sheet.setFormula()` には正規構文を渡してください。ローカライズされた文字列は、先に `workbook.canonicalizeFormula()` で変換できます。表示用の変換は `workbook.localizeFormula()`、区切り文字や真偽値などの定義取得は `workbook.getLocaleFacts()` を使います。

校閲 > スペル チェックは、デスクトップ版と同じダイアログでシートを 1 語ずつ確認します。組み込みのチェッカーは英語のよくある誤字の短い一覧だけを持ち、ほかの言語の単語はすべて通します。実際の辞書で確認するには `spelling: { check(word, locale), suggest(word, locale), addWord?(word, locale) }` を渡してください。ユーザーが追加した単語は `addWord` に渡されるので、ホスト側で保存できます。ライブラリ自体はセッションを超えて何も保存しません。

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
