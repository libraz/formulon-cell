# @libraz/formulon-cell

[![npm version](https://img.shields.io/npm/v/@libraz/formulon-cell.svg)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![license](https://img.shields.io/npm/l/@libraz/formulon-cell.svg)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![bundle size](https://img.shields.io/bundlephobia/minzip/@libraz/formulon-cell)](https://bundlephobia.com/package/@libraz/formulon-cell)

[formulon](https://github.com/libraz/formulon) WASM 計算エンジン向けの
スプレッドシート UI。デスクトップ表計算ソフト風の UI 表層、
Canvas 描画によるグリッド、拡張ベースの機能構成、実行時ロケール切替を
提供します。

## インストール

```sh
npm install @libraz/formulon-cell zustand
```

`zustand` はピア依存として公開しています。WASM エンジンは
[crossOriginIsolated](https://developer.mozilla.org/docs/Web/API/crossOriginIsolated)
コンテキスト (`COOP: same-origin` + `COEP: require-corp`) を必要とします。
ヘッダが無い環境では formulon-cell はインメモリのスタブエンジンに
フォールバックし、再計算と xlsx の読み書きは無効化されます。

Vite / webpack / esbuild の設定は
[バンドラ統合](https://github.com/libraz/formulon-cell/blob/main/README_ja.md#バンドラ統合)
を参照してください。

## クイックスタート

```ts
import { Spreadsheet, WorkbookHandle, presets } from '@libraz/formulon-cell';
import '@libraz/formulon-cell/styles.css';

const host = document.getElementById('sheet')!;
const wb = await WorkbookHandle.createDefault();
const sheet = await Spreadsheet.mount(host, {
  workbook: wb,
  features: presets.full(),
  locale: 'ja',
});

sheet.i18n.setLocale('en');     // 実行時にロケールを切り替え
sheet.setTheme('ink');           // ダークテーマへ切り替え
```

## プリセット

| プリセット | 含まれる機能 |
|----------|------------|
| `presets.minimal()`  | 数式バー、ステータスバー、基本キーマップ |
| `presets.standard()` | + ビューツールバー、クイック分析、コンテキストメニュー、検索／置換、クリップボード、書式コピー |
| `presets.full()`     | + 書式ダイアログ、形式を選択して貼り付け、条件付き書式、名前付き範囲、ハイパーリンクダイアログ、ピボットテーブル作成、入力規則、オートコンプリート、ホバーコメント |

## サブパスエクスポート

| インポートパス | 説明 |
|---|---|
| `@libraz/formulon-cell` | コア: `Spreadsheet`、`WorkbookHandle`、`presets`、拡張ファクトリ |
| `@libraz/formulon-cell/extensions` | 拡張ファクトリ一式（再エクスポート） |
| `@libraz/formulon-cell/extensions/*` | 個別の拡張 (`statusBar`、`findReplace`、`contextMenu` など) |
| `@libraz/formulon-cell/i18n/ja` | 日本語ロケール辞書 |
| `@libraz/formulon-cell/i18n/en` | 英語ロケール辞書 |
| `@libraz/formulon-cell/styles.css` | デフォルトスタイル束 |
| `@libraz/formulon-cell/styles/paper.css` | paper (ライト) テーマ |
| `@libraz/formulon-cell/styles/ink.css` | ink (ダーク) テーマ |
| `@libraz/formulon-cell/styles/contrast.css` | ハイコントラストテーマ |
| `@libraz/formulon-cell/styles/tokens.css` | テーマトークンのみ |

## 主な API

| API | 説明 |
|-----|------|
| `Spreadsheet.mount(host, opts)` | スプレッドシート UI を DOM 要素にマウント |
| `WorkbookHandle.createDefault()` | WASM エンジン（またはスタブ）でワークブックを生成 |
| `isUsingStub()` | スタブエンジンが使われているかを判定 |
| `presets.{minimal,standard,full}()` | 内蔵プリセット |
| `instance.i18n.setLocale(loc)` | 再マウント不要でロケールを切り替え |
| `instance.setTheme(theme)` | 実行時にテーマを切り替え |
| `createSessionChart(store, range, options)` | セッションの縦棒／折れ線チャートを作成 |
| `saveSheetView` / `activateSheetView` | セッション内のシートビュー管理 |
| `listDefinedNames` / `upsertDefinedName` | ヘッドレスな名前マネージャー API |

完全な API リファレンスは
[プロジェクト README](https://github.com/libraz/formulon-cell/blob/main/README_ja.md)
を参照してください。

## フレームワークコンポーネント

| パッケージ | 説明 |
|---------|------|
| [`@libraz/formulon-cell-react`](https://www.npmjs.com/package/@libraz/formulon-cell-react) | `<Spreadsheet>` React コンポーネント + フック |
| [`@libraz/formulon-cell-vue`](https://www.npmjs.com/package/@libraz/formulon-cell-vue) | `<Spreadsheet>` Vue コンポーネント + コンポーザブル |

## ライセンス

[Apache License 2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
