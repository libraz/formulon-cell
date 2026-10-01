# @libraz/formulon-cell

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — react](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![npm — vue](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-6-blue?logo=typescript)](https://www.typescriptlang.org/)

**フレームワーク非依存のまま、Web ページの中に表計算を組み込みます。** DOM 要素に
マウントすると、Canvas 描画のグリッドとデスクトップ表計算ソフト風の UI 表層
（数式バー、リボン、シートタブ、コンテキストメニュー）が手に入ります。その下では
[formulon](https://github.com/libraz/formulon) の WASM 計算エンジンが、メイン
スレッドの外で数式を評価します。機能はプリセットと拡張ファクトリで組み立てられ、
ロケールは再マウントせず実行時に切り替えられます。

本パッケージは Vanilla TypeScript / DOM のコアです。フレームワーク版は
[`@libraz/formulon-cell-react`](https://www.npmjs.com/package/@libraz/formulon-cell-react)
と [`@libraz/formulon-cell-vue`](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
にあり、どちらも同じマウント呼び出しに対する薄いアダプタです。

> **Excel 互換性について。** `formulon-cell` は、実際のブラウザ上で
> [**formulon**](https://github.com/libraz/formulon) を結合試験しながら、
> Excel 互換の表計算操作を目指して開発しています。選択、編集、数式入力、
> 再計算、ファイルの読み書きといった基本的なワークブック操作は使えます。
> ただし、詳細なコントロールの挙動、ダイアログ、キーボード操作、
> アクセシビリティなどの UI/UX は、まだ Excel と同じ振る舞いを保証して
> おらず、不具合が残る可能性もあります。現時点では、Excel をそのまま
> 置き換える完成済みのエンドユーザー向け表計算ソフトとして案内しないでください。

## インストール

```sh
npm install @libraz/formulon-cell zustand
```

`zustand` はピア依存です。formulon 0.12.0の標準WASMは単一スレッド版で、COOP/COEPヘッダを必要としません。起動に失敗すると `WorkbookHandle.createDefault()` はエラーを返します。スタブエンジンは、テストや明示的なデモ向けに `preferStub: true` を渡した場合だけ使います。

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
  toolbar: true,                 // 同じ呼び出しでリボンも → sheet.toolbar
});

sheet.i18n.setLocale('en');     // 実行時にロケールを切り替え
sheet.setTheme('ink');           // ダークテーマ — グリッドとツールバーが同時に切り替わる
```

## ホスト統合

ブラウザの API だけでは、デスクトップ表計算ソフトの統合ポイントをすべて
再現できません。ホスト側は `MountOptions` からその能力を差し込めるため、
リボンや Backstage の共通挙動はコアに置いたまま拡張できます。

```ts
const sheet = await Spreadsheet.mount(host, {
  workbook: wb,
  captureScreenClip: async () => ({
    src: await nativeCaptureRegionAsDataUrl(),
    alt: '画面の領域切り取り',
  }),
  printerProfiles: [
    {
      id: 'office-printer',
      name: 'Office Printer',
      paperSize: 'A4',
      orientation: 'portrait',
      printableBounds: { top: 0.16, right: 0.16, bottom: 0.16, left: 0.16 },
    },
  ],
  refreshPrinterProfiles: () => nativeListPrinterProfiles(),
  uploadStatus: 'saving',
  macroRecording: false,
});
```

`captureScreenClip` は「挿入 > スクリーンショット > 画面の領域切り取り」を
担います。省略した場合、このコマンドはネイティブの領域切り取りがホスト提供で
あることを通知します。プリンタープロファイルは、ページ設定と印刷プレビューに
物理プリンターの最小余白を与えるもので、ブラウザだけのホストでは省略できます。
`uploadStatus` は `saved`・`saving`・`error`・`null`、`macroRecording` は
`true`・`false`・`null` を受け取ります。

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
| `@libraz/formulon-cell/styles.css` | 全部入り束: グリッド・ダイアログ・オーバーレイ・リボンツールバー・3 テーマすべて。多くの利用者はこの 1 本だけで足りる |
| `@libraz/formulon-cell/styles/toolbar.css` | リボン用スタイルのみ — 全部入り束を使わない細粒度構成向け |
| `@libraz/formulon-cell/styles/paper.css` | paper (ライト) テーマのみ |
| `@libraz/formulon-cell/styles/ink.css` | ink (ダーク) テーマのみ |
| `@libraz/formulon-cell/styles/contrast.css` | ハイコントラストテーマのみ |
| `@libraz/formulon-cell/styles/tokens.css` | テーマトークンのみ |

## 主な API

| API | 説明 |
|-----|------|
| `Spreadsheet.mount(host, opts)` | スプレッドシート UI を DOM 要素にマウント |
| `WorkbookHandle.createDefault()` | WASM エンジンでワークブックを生成 |
| `isUsingStub()` | 明示的なスタブエンジンが使われているかを判定 |
| `presets.{minimal,standard,full}()` | 内蔵プリセット |
| `instance.i18n.setLocale(loc)` | 再マウント不要でロケールを切り替え |
| `instance.setTheme(theme)` | 実行時にテーマを切り替え |
| `createSessionChart(store, range, options)` | セッションの縦棒／折れ線チャートを作成 |
| `saveSheetView` / `activateSheetView` | セッション内のシートビュー管理 |
| `listDefinedNames` / `upsertDefinedName` | ヘッドレスな名前マネージャー API |
| `ribbonActivationEntries` / `ribbonSurfaceCommandIds` | ホスト側の監査とラッパー間の一致確認に使うリボンコマンドの共有マニフェスト |
| `attachRangePickerButton` | Excel 風ダイアログが共通で使う範囲選択コントロール |
| `appendConditionalApplyFormatControls` / `conditionalStyleOptions` | 条件付き書式ルール UI の共通ヘルパー |
| `showReport` / `reportDialogLabels` | ホスト依存の互換性レポート用ダイアログとラベル対応表 |
| `projectDisabledReason` / `projectDisabledState` | aria 属性・title・dataset・コントロール状態へ無効／読み取り専用の理由を射影する共通処理 |

完全な API リファレンスは
[プロジェクト README](https://github.com/libraz/formulon-cell/blob/main/README_ja.md)
を参照してください。

## 併せて使えるパッケージ

| パッケージ | 説明 |
|---------|------|
| [`@libraz/formulon-cell-react`](https://www.npmjs.com/package/@libraz/formulon-cell-react) | `<Spreadsheet>` React コンポーネント + フック + `SpreadsheetToolbar` リボン |
| [`@libraz/formulon-cell-vue`](https://www.npmjs.com/package/@libraz/formulon-cell-vue) | `<Spreadsheet>` Vue コンポーネント + コンポーザブル + `SpreadsheetToolbar` リボン |

React:

```tsx
import { SpreadsheetToolbar, type RibbonTab } from '@libraz/formulon-cell-react';
import '@libraz/formulon-cell-react/toolbar.css';
```

Vue:

```vue
<script setup lang="ts">
import { type RibbonTab } from '@libraz/formulon-cell-vue';
import SpreadsheetToolbar from '@libraz/formulon-cell-vue/toolbar.vue';
import '@libraz/formulon-cell-vue/toolbar.css';
</script>
```

どちらのラッパーも、コアの `Spreadsheet.mountToolbar` に対する薄いアダプタです。
リボンの DOM、メニューファクトリ、アクティベーションモデル、動的ドロップダウンの
ディスパッチャは `@libraz/formulon-cell` が持ちます。

## ライセンス

[Apache-2.0](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
