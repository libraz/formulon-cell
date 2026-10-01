# formulon-cell

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-cell/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-cell/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-cell?label=%40libraz%2Fformulon-cell)](https://www.npmjs.com/package/@libraz/formulon-cell)
[![npm — react](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=react)](https://www.npmjs.com/package/@libraz/formulon-cell-react)
[![npm — vue](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=vue)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)
[![codecov](https://codecov.io/gh/libraz/formulon-cell/branch/main/graph/badge.svg)](https://codecov.io/gh/libraz/formulon-cell)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-cell/blob/main/LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-6-blue?logo=typescript)](https://www.typescriptlang.org/)

**formulon-cell は、Web ページの中で動く表計算そのものを提供します。** DOM 要素・
React コンポーネント・Vue コンポーネントのいずれにマウントしても、Canvas 描画の
グリッドとデスクトップ表計算ソフト風の UI 表層（数式バー、リボン、シートタブ、
コンテキストメニュー）が手に入ります。その下では
[formulon](https://github.com/libraz/formulon) の WASM 計算エンジンが、メイン
スレッドの外で数式を評価します。コアはフレームワーク非依存の TypeScript で、
React 版と Vue 版は同じマウント呼び出しに対する薄いアダプタです。

**こんなときに使えます**

- **データグリッドではなく表計算を組み込む** — 選択、編集、数式入力、再計算、ファイルの読み書きが最初のマウントから動きます。
- **計算を本物のエンジンに任せる** — 数式は UI に後付けした JavaScript の数式パーサーではなく、WebAssembly へコンパイルした C++ エンジンが処理します。
- **使う機能だけを配る** — プリセットと拡張ファクトリでマウント対象を決められます。UI 表層をすべて外し、キャンバスとストアだけを使うこともできます。
- **実行時に言語を切り替える** — `ja` と `en` を同梱し、再マウントせずその場でロケールを差し替えられます。

> **Excel 互換性について。** `formulon-cell` は、実際のブラウザ上で
> [**formulon**](https://github.com/libraz/formulon) を結合試験しながら、
> Excel 互換の表計算操作を目指して開発しています。選択、編集、数式入力、
> 再計算、ファイルの読み書きといった基本的なワークブック操作は使えます。
> ただし、詳細なコントロールの挙動、ダイアログ、キーボード操作、
> アクセシビリティなどの UI/UX は、まだ Excel と同じ振る舞いを保証して
> おらず、不具合が残る可能性もあります。現時点では、Excel をそのまま
> 置き換える完成済みのエンドユーザー向け表計算ソフトとして案内しないでください。

## パッケージ

| パッケージ | npm | 概要 |
|---------|-----|------|
| [`@libraz/formulon-cell`](./packages/formulon-cell)             | [![npm](https://img.shields.io/npm/v/@libraz/formulon-cell?label=)](https://www.npmjs.com/package/@libraz/formulon-cell)             | Vanilla TS / DOM コア |
| [`@libraz/formulon-cell-react`](./packages/formulon-cell-react) | [![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-react?label=)](https://www.npmjs.com/package/@libraz/formulon-cell-react) | React 18+ コンポーネント・フック・リボンツールバー |
| [`@libraz/formulon-cell-vue`](./packages/formulon-cell-vue)     | [![npm](https://img.shields.io/npm/v/@libraz/formulon-cell-vue?label=)](https://www.npmjs.com/package/@libraz/formulon-cell-vue)     | Vue 3 コンポーネント・コンポーザブル・リボンツールバー |

## インストール

```sh
npm install @libraz/formulon-cell zustand
# または yarn / pnpm
```

`zustand` はピア依存として公開しています。UI 表層が購読しているストアに、
利用者側からも同じインスタンスでアクセスできるようにするためです。

formulon 0.12.0 の標準WASMは単一スレッド版で、COOP/COEPヘッダや `SharedArrayBuffer` を必要としません。起動に失敗した場合、`WorkbookHandle.createDefault()` はエラーを返します。インメモリのスタブエンジンは、テストや明示的なデモ向けに `preferStub: true` を渡した場合だけ使います。

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

## バンドラ統合

formulon-cell は `@libraz/formulon` の標準の単一スレッドWASMを利用します。以下の設定でWASMアセットを解決してください。

**1. ワーカー設定はスレッド版を直接使う場合に必要。** アプリで `@libraz/formulon/threads` を直接使う場合、Viteでは `worker: { format: 'es' }` を指定してください。formulon-cell の標準ロードではワーカーを起動しません。

**2. トップレベル await と動的な Node モジュール読み込みには es2022
ターゲットが必要。** エンジンファクトリはトップレベル await と条件付きの
`await import('node:...')` を利用します。メイン側・ワーカー側の双方で
ビルドターゲットを `es2022` 以上に引き上げてください。

```ts
// vite.config.ts
export default defineConfig({
  build: { target: 'es2022' },
});
```

**3. 依存関係の事前バンドル対象からエンジンを除外する。** formulon-cell は
`@libraz/formulon` を経由してロードし、その Emscripten ラッパーが
WASM アセットの解決を担当します。両パッケージを事前バンドルの
対象から外し、アセット解決はアプリ側のバンドラに委ねてください。

```ts
// vite.config.ts
export default defineConfig({
  optimizeDeps: { exclude: ['@libraz/formulon-cell', '@libraz/formulon'] },
});
```

**4. 標準WASMはクロスオリジン分離を必要としない。** `@libraz/formulon/threads` を直接使う場合だけ、`Cross-Origin-Opener-Policy: same-origin` と `Cross-Origin-Embedder-Policy: require-corp` ヘッダが必要です。スタブは `preferStub: true` を渡したテストや明示的なデモ専用です。数式評価・再計算・xlsxの読み書きが不完全なため、通常の実行には実WASMを使ってください。

```ts
import { WorkbookHandle, isUsingStub } from '@libraz/formulon-cell';

const wb = await WorkbookHandle.createDefault();
if (isUsingStub()) {
  console.warn('formulon-cell: 明示的にスタブエンジンを使用中');
}
```

## できること

- **デスクトップ表計算ソフト風の UI 表層** を標準装備（数式バー、
  ステータスバー、コンテキストメニュー、シートタブ、リボン）。
- **Canvas 描画によるグリッド** とテーマトークン。`paper`（ライト）と
  `ink`（ダーク）を同梱しており、ドキュメント化された CSS 変数で独自テーマ
  も作成可能。
- **拡張ベース API**: ビルトイン機能は機能フラグで制御。差し替え可能な
  パーツ（検索／置換、書式ダイアログ、形式を選択して貼り付け、
  ハイパーリンクダイアログ、ホバーコメント、
  クイック分析、ピボットテーブル作成 など）は拡張ファクトリとして提供。
- **実行時 i18n** — 再マウントせずにロケールを切り替え可能。`ja` と `en`
  を同梱しており、実行時にロケールを追加登録することもできます。
- **ヘッドレスモード** — キャンバスとストアのみを利用し、UI 表層を
  独自に実装することもできます。

### プリセット

| プリセット | 含まれる機能 |
|----------|------------|
| `presets.minimal()`  | 数式バー、ステータスバー、基本キーマップ |
| `presets.standard()` | + クイック分析、セッションチャートオーバーレイ、ワークブックオブジェクトインスペクター、コンテキストメニュー、検索／置換、クリップボード、書式コピー、ホイールスクロール |
| `presets.full()`     | + 書式ダイアログ、形式を選択して貼り付け、条件付き書式、反復計算設定、ジャンプ — セル選択、ページ設定、名前付き範囲、ハイパーリンクダイアログ、ピボットテーブル作成、入力規則、オートコンプリート、ホバーコメント、表計算キーマップ |

### i18n

```ts
import { Spreadsheet } from '@libraz/formulon-cell';

const sheet = await Spreadsheet.mount(host, { locale: 'en' });

// 実行時にロケールを切り替え — すべてのラベルがその場で更新される
sheet.i18n.setLocale('ja');

// 辞書をフォークせず、一部のキーだけ上書きする
sheet.i18n.extend('ja', { contextMenu: { copy: 'コピーする' } });

// 新しいロケールを登録する
import fr from './fr.js';
sheet.i18n.register('fr', fr);
sheet.i18n.setLocale('fr');
```

## デモアプリ

| アプリ | 起動 | 内容 |
|-----|-----|------|
| `apps/react-demo`  | `yarn dev` / `yarn dev:react` | React コンポーネント `<Spreadsheet>` と同じ機能面 |
| `apps/vue-demo`    | `yarn dev:vue`    | Vue コンポーネント `<Spreadsheet>` と同じ機能面 |

## フレームワーク向けリボンツールバー

React 版・Vue 版のパッケージは、コアの `Spreadsheet.mountToolbar` に対する
薄いアダプタを公開しています。リボンの DOM、メニューファクトリ、
アクティベーションモデル、動的ドロップダウンのディスパッチャはいずれも
`@libraz/formulon-cell` に集約されているため、フレームワーク側のラッパーが
別個のメニュー実装を抱えることはありません。コンポーネントと一緒に
ツールバー用 CSS も読み込んでください。

ホスト側でリボンを監査・拡張する場合は、コマンド一覧を React や Vue で
組み直さず、コアが公開する共有マニフェスト（`ribbonActivationEntries`、
`ribbonSurfaceCommandIds`、`DYNAMIC_RIBBON_DROPDOWN_HANDLER_ATTRS`）を
読み込んでください。`attachRangePickerButton`、
`appendConditionalApplyFormatControls`、`conditionalStyleOptions`、
`showReport`、`reportDialogLabels`、`projectDisabledReason`、
`projectDisabledState` といったダイアログ用のヘルパーもコアから公開しており、
内蔵の Excel 365 風ダイアログやレポートと同じ挙動を保ちたいホスト UI が
そのまま利用できます。

```tsx
import { SpreadsheetToolbar, type RibbonTab } from '@libraz/formulon-cell-react';
import '@libraz/formulon-cell-react/toolbar.css';
```

```vue
<script setup lang="ts">
import { type RibbonTab } from '@libraz/formulon-cell-vue';
import SpreadsheetToolbar from '@libraz/formulon-cell-vue/toolbar.vue';
import '@libraz/formulon-cell-vue/toolbar.css';
</script>
```

## 含まないもの（Non-goals）

formulon-cell は UI レイヤであり、計算エンジンでもアプリケーションでも
ありません。数式評価・再計算・xlsx の解析は
[formulon](https://github.com/libraz/formulon) の担当で、このリポジトリで
作り直すことはありません。サーバー、永続化層、共同編集のトランスポート、
認証もいずれも含まず、これらはホスト側が持ちます。また、静かに機能を
落とすこともしません。クロスオリジン分離が無い環境では、インメモリの
スタブへフォールバックせずマウント自体を失敗させます。React 版・Vue 版は
薄いアダプタのままとし、メニュー・コマンド・ダイアログはコアに置くことで、
3 つのホストの実装が互いにずれないようにしています。

## リリース

タグベースのリリースフローは [`docs/releasing.md`](./docs/releasing.md)
を参照してください。

## ライセンス

[Apache-2.0](LICENSE)
