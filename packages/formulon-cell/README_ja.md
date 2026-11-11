# @libraz/formulon-cell

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — react](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![npm — vue](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-6-blue?logo=typescript)](https://www.typescriptlang.org/)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-blue)](https://formulon.libraz.net/ja/cell/)
[![Demo](https://img.shields.io/badge/demo-live-brightgreen)](https://formulon.libraz.net/ja/cell/demo)

**フレームワークなしで、Web ページに表計算を組み込めます。** DOM 要素にマウントすると、Canvas 描画のグリッドと、デスクトップ版 Excel に合わせた UI（数式バー、リボン、シートタブ、コンテキストメニュー）が表示され、計算は [formulon](https://github.com/libraz/formulon) の WASM エンジンが受け持ちます。機能はプリセットと拡張ファクトリで組み合わせ、ロケールとテーマは再マウントせずに切り替えられます。

[デモ](https://formulon.libraz.net/ja/cell/demo) でブラウザ上ですぐに試せます。

このパッケージは Vanilla TypeScript / DOM のコアです。[`@libraz/formulon-cell-react`](https://www.npmjs.com/package/@libraz/formulon-cell-react) と [`@libraz/formulon-cell-vue`](https://www.npmjs.com/package/@libraz/formulon-cell-vue) は、同じマウント処理を包む薄いアダプタです。

> **ベータ版（1.0 未満）です。** UI はデスクトップ版 Excel に合わせて作っており、機能ごとに突き合わせています。公開 API はマイナーバージョンの間でも変わることがあり、まだ突き合わせていない挙動は異なる場合があります。違いを見つけたら [issue](https://github.com/libraz/formulon-cell/issues) で知らせてください。

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
  features: { ...presets.full(), localeSwitcher: true },
  locale: 'ja',
  toolbar: true, // リボンも同時にマウントし、sheet.toolbar で参照できる
});

sheet.setLocale('en');
sheet.setTheme('ink'); // グリッドとリボンのテーマをまとめて切り替え
```

`getWorkbookName: () => currentFilename` を指定すると、ウォッチウィンドウに現在のブック名を表示します。React / Vueコンポーネントも同じコールバックを受け取ります。

言語選択UIは初期状態では非表示です。`features.localeSwitcher: true` で表示し、`sheet.setFeatures({ localeSwitcher: false })` で動的に非表示へ戻せます。`sheet.setLocale()` は表示辞書と計算プロファイルを同時に切り替えます。切り替えが拒否された場合は `false` を返し、現在の言語を維持します。ダイアログ、検索・置換、コメント編集、フィルターメニューを閉じてから切り替えてください。拒否時は入力途中の内容を保持します。

## 埋め込み

`ui` で何を表示するか、`policy` でユーザーが何を変更できるか、`viewport` でどのセルを表示・選択できるかを決めます。どれも実行時に変更できます（`setUi`、`setPolicy`、`setViewportOptions`）。

```ts
import { fixedFormPolicy, viewerPolicy } from '@libraz/formulon-cell';

const form = await Spreadsheet.mount(host, {
  workbook: wb,
  ui: { profile: 'embedded', features: { formulaBar: true } }, // グリッドと数式バーだけ
  policy: fixedFormPolicy(['B2:B11']), // どのシートでも B2:B11 だけ入力可能
  viewport: { range: 'A1:D12', tabNavigation: 'editable' },
});

form.setPolicy(viewerPolicy()); // 送信後は選択とコピーだけ
```

ポリシーは 28 種類の操作を個別に許可・拒否し、編集できるセルは A1 形式の範囲（`'B2:B11'` は全シート、`'Input!B2:B11'` はそのシートだけ）か述語で指定します。エディタ、クリップボード、フィルハンドル、コンテキストメニュー、リボン、インスタンス API のどの入力経路にも効きます。`restrict` フックは拒否を追加できますが、許可を足すことはできません。policy を設定すると、操作の分類がないリボンコマンドは無効のままになります。制限付きのシートは、リボンを出さずに表示するか、ホスト側で用意した操作 UI と組み合わせてください。詳しくは [埋め込みと利用例](https://formulon.libraz.net/ja/cell/embedding) を参照してください。

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
| `@libraz/formulon-cell/i18n/{en,ja,de,fr,zh,ko,th}` | 7言語のロケール辞書 |
| `@libraz/formulon-cell/styles.css` | グリッド、ダイアログ、オーバーレイ、リボン、全テーマをまとめたスタイルシート |
| `@libraz/formulon-cell/styles/{tokens,paper,ink,contrast,toolbar}.css` | 個別に読み込む場合のスタイルシート |

## 主な API

| API | 説明 |
|-----|------|
| `Spreadsheet.mount(host, opts)` | DOM 要素に表計算をマウントする |
| `WorkbookHandle.createDefault()` | WASM エンジンを使うワークブックを作る |
| `workbook.withBatchedRecalc(fn)` | 複数セルへの書き込みを 1 回の再計算にまとめる |
| `instance.setLocale(loc)` / `instance.setTheme(theme)` | 実行時にロケールやテーマを切り替える |
| `instance.setPolicy` / `setViewportOptions` / `setContextMenu` / `setUi` / `setToolbar` | 埋め込み設定を実行時に変更する |
| `instance.applyChanges(changes)` | ポリシーとシート保護で判定する、全部か無しかのセル一括更新 |
| `instance.on(event, handler)` | `cellChange`、`selectionChange`、`changeBatch`、`recalc` など |

オプションと API の一覧は [ドキュメントサイト](https://formulon.libraz.net/ja/cell/) にあります。

CSV出力にはロケールと区切り文字を指定できます。例: `exportCSV(state, { locale: 'de', delimiter: ';' })`。取り込みでは `importCSV(store, workbook, text, anchor, history, settings)` の第6引数へ同じ設定を渡すと、数値と真偽値を同じ規則で解析します。出力の既定値は英語の正規表記とカンマ区切りです。取り込みは `settings.locale` が未指定の場合、現在のワークブックのプロファイルを使います。

数式の編集と外部クリップボードからの入力には、現在のプロファイルの構文を使います。たとえば独語の `=SUMME(1,5;2)` は `=SUM(1.5,2)` として保存されます。コードから呼ぶ `workbook.setFormula()` には正規構文を渡してください。ローカライズされた文字列は、先に `workbook.canonicalizeLocalizedFormula()` で変換できます。表示用の変換は `workbook.localizeFormulaText()`、区切り文字や真偽値などの定義取得は `workbook.localeFacts()` を使います。

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
