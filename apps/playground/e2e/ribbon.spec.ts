import { expect, type Page, test } from '@playwright/test';

const ribbonTabs = [
  {
    id: 'home',
    label: 'Home',
    commands: [
      'paste',
      'cut',
      'copy',
      'formatPainter',
      'fontFamily',
      'fontSize',
      'fontGrow',
      'fontShrink',
      'font-row-2',
      'bold',
      'italic',
      'underline',
      'strike',
      'borders',
      'borderPreset',
      'borderStyle',
      'borderColor',
      'moreBorders',
      'drawBorder',
      'drawBorderGrid',
      'eraseBorder',
      'fontColor',
      'fillColor',
      'top',
      'middle',
      'bottomAlign',
      'textOrientation',
      'wrap',
      'alignment-row-2',
      'alignL',
      'alignC',
      'alignR',
      'indentDecrease',
      'indentIncrease',
      'merge',
      'numberFormat',
      'number-row-2',
      'currency',
      'percent',
      'comma',
      'decDown',
      'decUp',
      'conditional',
      'formatTableHome',
      'cellStyles',
      'insertRows',
      'deleteRows',
      'formatCellsHome',
      'autosum',
      'fillHome',
      'clearFormat',
      'sortFilterHome',
      'findHome',
    ],
  },
  {
    id: 'insert',
    label: 'Insert',
    commands: [
      'pivotTableInsert',
      'formatTableInsert',
      'namedRangesInsert',
      'removeDupesInsert',
      'pictureInsert',
      'shapesInsert',
      'screenshotInsert',
      'chartInsert',
      'hyperlinkInsert',
      'linksInsert',
      'commentInsert',
      'symbolInsert',
    ],
  },
  { id: 'draw', label: 'Draw', commands: ['drawPen', 'drawGrid', 'drawErase'] },
  {
    id: 'pageLayout',
    label: 'Page Layout',
    commands: [
      'pageTheme',
      'marginsPreset',
      'orientationPreset',
      'paperSizePreset',
      'pageSetupAdvanced',
      'printArea',
      'pageBreaks',
      'sheetBackground',
      'printTitles',
      'scaleWidth',
      'scaleHeight',
      'scalePercent',
      'pageLayoutGridlinesView',
      'pageLayoutGridlinesPrint',
      'pageLayoutHeadingsView',
      'pageLayoutHeadingsPrint',
      'printPageLayout',
    ],
  },
  {
    id: 'formulas',
    label: 'Formulas',
    commands: [
      'fx',
      'autosumFormula',
      'sum',
      'avg',
      'ifFormula',
      'xlookupFormula',
      'concatFormula',
      'todayFormula',
      'pmtFormula',
      'roundFormula',
      'namedRanges',
      'precedents',
      'dependents',
      'clearArrows',
      'errorChecking',
      'showFormulasFormula',
      'evaluateFormula',
      'recalcNow',
      'calcOptions',
      'watch',
    ],
  },
  {
    id: 'data',
    label: 'Data',
    commands: [
      'filter',
      'sortAsc',
      'sortDesc',
      'sortData',
      'textToColumns',
      'removeDupes',
      'dataValidation',
      'linksData',
      'outlineGroup',
      'outlineUngroup',
      'outlineShowDetail',
      'outlineHideDetail',
    ],
  },
  {
    id: 'review',
    label: 'Review',
    commands: [
      'spellingReview',
      'translateReview',
      'newCommentReview',
      'deleteCommentReview',
      'previousCommentReview',
      'nextCommentReview',
      'findReview',
      'protectReview',
      'protectWorkbookReview',
      'protectionReview',
      'accessibility',
    ],
  },
  {
    id: 'view',
    label: 'View',
    commands: [
      'viewNormal',
      'viewPageLayout',
      'viewPageBreakPreview',
      'watchView',
      'sheetViewSelect',
      'sheetViewSave',
      'sheetViewDelete',
      'workbookObjectsView',
      'viewGridlines',
      'viewHeadings',
      'viewFormulas',
      'viewFormulaBar',
      'viewR1C1',
      'freeze',
      'windowVisibility',
      'zoomDialog',
      'zoomSelection',
      'zoom75',
      'zoom100',
      'zoom125',
      'protect',
    ],
  },
  { id: 'automate', label: 'Automate', commands: ['script', 'recordActions', 'allScripts'] },
  { id: 'acrobat', label: 'Acrobat', commands: ['addIn', 'pdf'] },
] as const;

async function mount(page: Page, url = '/'): Promise<void> {
  await page.goto(url);
  await page.waitForSelector('.fc-host', { state: 'attached', timeout: 30_000 });
  await page.waitForFunction(
    () => {
      const host = document.querySelector('.fc-host') as HTMLElement | null;
      const state = host?.dataset.fcEngineState;
      return state === 'ready' || state === 'ready-stub';
    },
    { timeout: 30_000 },
  );
}

test('R00: Japanese conditional-format menu uses localized multi-level labels', async ({
  page,
}) => {
  await mount(page, '/?locale=ja');

  await page.getByRole('tab', { name: 'ホーム', exact: true }).click();
  await page.locator('[data-ribbon-command="conditional"]').click();

  const menu = page.locator('#menu-conditional');
  await expect(menu).toBeVisible();
  await expect(menu).toContainText('セルの強調表示ルール');
  await expect(menu).toContainText('上位/下位ルール');
  await expect(menu).toContainText('データ バー');
  await expect(menu).toContainText('カラー スケール');
  await expect(menu).toContainText('アイコン セット');
  await expect(menu).toContainText('ルールのクリア');
  await expect(menu).toContainText('ルールの管理...');

  await page.locator('[data-cf-submenu="dataBar"]').hover();
  await expect(
    page.locator('.app__submenu--cf-dataBar [data-cf-action="data-blue"]'),
  ).toHaveAttribute('aria-label', '塗りつぶし (グラデーション)、青のデータ バー');

  await page.locator('[data-cf-submenu="colorScale"]').hover();
  await expect(
    page.locator('.app__submenu--cf-colorScale [data-cf-action="scale-gyr"]'),
  ).toHaveAttribute('title', '緑 - 黄 - 赤のカラー スケール');

  await page.locator('[data-cf-submenu="iconSet"]').hover();
  const iconSetPanel = page.locator('.app__submenu--cf-iconSet');
  await expect(iconSetPanel).toBeVisible();
  await expect(iconSetPanel).toContainText('方向');
  await expect(iconSetPanel).toContainText('図形');
  await expect(iconSetPanel).toContainText('インジケーター');
  await expect(iconSetPanel).toContainText('評価');
  await expect(iconSetPanel.locator('[data-cf-action="icons-arrows3"]').first()).toHaveAttribute(
    'aria-label',
    '3 方向矢印',
  );
  await expect(iconSetPanel.locator('[data-cf-action="icons-symbols3"]').first()).toHaveAttribute(
    'aria-label',
    '3 記号',
  );
  await expect(iconSetPanel.locator('[data-cf-action="icons-boxes5"]').first()).toHaveAttribute(
    'aria-label',
    '5 ボックス',
  );

  await page.locator('#menu-conditional > [data-cf-action="new-rule"]').click();
  const newRuleDialog = page.getByRole('dialog', { name: '新しい書式ルール' });
  await expect(newRuleDialog).toBeVisible();
  await expect(newRuleDialog).toContainText('スタイル');
  await expect(newRuleDialog).toContainText('クラシック');
  await expect(newRuleDialog).toContainText('対象範囲');
  await expect(newRuleDialog).toContainText('種類');
  await expect(newRuleDialog).toContainText('セル値');
  await expect(newRuleDialog).toContainText('濃い赤の文字、明るい赤の背景');
  await expect(newRuleDialog.getByRole('button', { name: 'OK' })).toBeVisible();
  await newRuleDialog.getByRole('button', { name: 'キャンセル' }).click();
});

test('R00a: Japanese defined-name menu uses localized labels', async ({ page }) => {
  await mount(page, '/?locale=ja');

  await page.getByRole('tab', { name: '数式', exact: true }).click();
  await page.locator('[data-ribbon-command="namedRanges"]').click();

  const menu = page.locator('#menu-defined-names');
  await expect(menu).toBeVisible();
  await expect(menu).toContainText('名前の定義...');
  await expect(menu).toContainText('名前の管理...');
  await expect(menu).toContainText('上端行から作成');
  await expect(menu).toContainText('左端列から作成');
  await expect(menu).toContainText('数式で使用');
  await expect(menu).toContainText('定義された名前はありません');
});

test('R00-data-ja: Japanese Data remove-duplicates dialog uses localized labels', async ({
  page,
}) => {
  await mount(page, '/?locale=ja');

  await selectRangeAndSetValues(page, { r0: 0, c0: 0, r1: 2, c1: 1 }, [
    { row: 0, col: 0, value: '品目' },
    { row: 0, col: 1, value: '数量' },
    { row: 1, col: 0, value: 'ペン' },
    { row: 1, col: 1, value: 2 },
    { row: 2, col: 0, value: 'ペン' },
    { row: 2, col: 1, value: 2 },
  ]);

  await page.getByRole('tab', { name: 'データ', exact: true }).click();
  await page.locator('[data-ribbon-command="removeDupes"]').click();

  const dialog = page.getByRole('dialog', { name: '重複の削除' });
  await expect(dialog).toBeVisible();
  await expect(dialog).toContainText('列');
  await expect(dialog).toContainText('すべて選択');
  await expect(dialog).toContainText('すべて選択解除');
  await expect(dialog).toContainText('先頭行をデータの見出しとして使用する');
  await dialog.getByRole('button', { name: 'キャンセル' }).click();
});

test('R00ab: Japanese Home editing menus use localized labels', async ({ page }) => {
  await mount(page, '/?locale=ja');

  await page.getByRole('tab', { name: 'ホーム', exact: true }).click();

  await page.locator('[data-ribbon-command="autosum"] .demo__rb-split-chevron').click();
  const autosumMenu = page.locator('#menu-autosum-home');
  await expect(autosumMenu).toBeVisible();
  await expect(autosumMenu).toContainText('合計');
  await expect(autosumMenu).toContainText('平均');
  await expect(autosumMenu).toContainText('数値の個数');
  await expect(autosumMenu).toContainText('その他の関数...');
  await page.keyboard.press('Escape');

  await page.locator('[data-ribbon-command="fillHome"]').click();
  const fillMenu = page.locator('#menu-fill');
  await expect(fillMenu).toBeVisible();
  await expect(fillMenu).toContainText('下方向へコピー');
  await expect(fillMenu).toContainText('右方向へコピー');
  await expect(fillMenu).toContainText('連続データの作成...');
  await expect(fillMenu).toContainText('週日単位');
  await page.keyboard.press('Escape');

  await page.locator('[data-ribbon-command="clearFormat"]').click();
  const clearMenu = page.locator('#menu-clear');
  await expect(clearMenu).toBeVisible();
  await expect(clearMenu).toContainText('すべてクリア');
  await expect(clearMenu).toContainText('数式と値のクリア');
  await expect(clearMenu).toContainText('ハイパーリンクの削除');
  await expect(clearMenu).toContainText('条件付き書式のクリア');
  await page.keyboard.press('Escape');

  await page.locator('[data-ribbon-command="sortFilterHome"]').click();
  const sortMenu = page.locator('#menu-sort-home');
  await expect(sortMenu).toBeVisible();
  await expect(sortMenu).toContainText('昇順で並べ替え');
  await expect(sortMenu).toContainText('降順で並べ替え');
  await expect(sortMenu).toContainText('選択したセルの値でフィルター');
  await expect(sortMenu).toContainText('詳細設定...');
  await page.keyboard.press('Escape');

  await page.locator('[data-ribbon-command="findHome"]').click();
  const findMenu = page.locator('#menu-find-select');
  await expect(findMenu).toBeVisible();
  await expect(findMenu).toContainText('検索...');
  await expect(findMenu).toContainText('置換...');
  await expect(findMenu).toContainText('条件を選択してジャンプ...');
  await expect(findMenu).toContainText('データの入力規則');
});

test('R00aa: Japanese formula-auditing split menus use localized labels', async ({ page }) => {
  await mount(page, '/?locale=ja');

  await page.getByRole('tab', { name: '数式', exact: true }).click();
  await page.locator('[data-ribbon-command="errorChecking"]').click();
  await expect(page.locator('#menu-error-checking')).toContainText('エラー チェック...');
  await expect(page.locator('#menu-error-checking')).toContainText('エラー トレース');
  await expect(page.locator('#menu-error-checking')).toContainText('エラーを無視する');

  await page.locator('[data-ribbon-command="clearArrows"]').click();
  await expect(page.locator('#menu-clear-arrows')).toContainText('矢印の削除');
  await expect(page.locator('#menu-clear-arrows')).toContainText('参照元の矢印の削除');
  await expect(page.locator('#menu-clear-arrows')).toContainText('参照先の矢印の削除');
});

test('R00ab: Japanese Watch Window menu uses localized labels', async ({ page }) => {
  await mount(page, '/?locale=ja');

  await page.getByRole('tab', { name: '数式', exact: true }).click();
  await page.locator('[data-ribbon-command="watch"]').click();
  const menu = page.locator('#menu-watch-formulas');
  await expect(menu).toBeVisible();
  await expect(menu).toContainText('ウォッチ ウィンドウ');
  await expect(menu).toContainText('ウォッチ式の追加');
  await expect(menu).toContainText('ウォッチ式の削除');
  await expect(menu).toContainText('すべて削除');
});

test('R00ac: Japanese review comment delete menu uses localized labels', async ({ page }) => {
  await mount(page, '/?locale=ja');

  await page.getByRole('tab', { name: '校閲', exact: true }).click();
  await page.locator('[data-ribbon-command="deleteCommentReview"]').click();
  const menu = page.locator('#menu-review-comments');
  await expect(menu).toBeVisible();
  await expect(menu).toContainText('コメントの削除');
  await expect(menu).toContainText('すべてのコメントとメモを削除');
});

test('R00ad: Japanese Protect menu uses localized labels', async ({ page }) => {
  await mount(page, '/?locale=ja');

  await page.getByRole('tab', { name: '校閲', exact: true }).click();
  await page.locator('[data-ribbon-command="protectReview"]').click();
  const menu = page.locator('#menu-protect-review');
  await expect(menu).toBeVisible();
  await expect(menu).toContainText('シートの保護...');
  await expect(menu).toContainText('シート保護の解除...');
  await expect(menu).toContainText('セルのロック');
  await expect(menu).toContainText('セルのロック解除');
  await expect(menu).toContainText('ブックの保護...');
  await expect(menu).toContainText('ブック保護の解除...');
  await expect(menu).toContainText('範囲の編集を許可...');
  await expect(menu).toContainText('編集許可範囲のクリア');
});

test('R00b: Japanese review reports use localized labels and selected translate text', async ({
  page,
}) => {
  await mount(page, '/?locale=ja');
  await selectCellAndSetText(page, 44, 2, 'teh  teh');

  await page.getByRole('tab', { name: '校閲', exact: true }).click();
  await page.locator('[data-ribbon-command="spellingReview"]').click();
  await expect(page.getByRole('dialog', { name: 'スペル チェック' })).toBeVisible();
  await expect(page.getByRole('dialog')).toContainText('警告');
  await expect(page.getByRole('dialog')).toContainText('スペルミスの可能性');
  await page.keyboard.press('Escape');

  await page.locator('[data-ribbon-command="translateReview"]').click();
  await expect(page.getByRole('dialog', { name: '翻訳' })).toBeVisible();
  await expect(page.getByRole('dialog')).toContainText('翻訳対象テキスト');
  await expect(page.getByRole('dialog')).toContainText('teh  teh');
  await page.keyboard.press('Escape');

  await page.locator('[data-ribbon-command="accessibility"]').click();
  await expect(page.getByRole('dialog', { name: 'アクセシビリティ' })).toBeVisible();
  await expect(page.getByRole('dialog')).toContainText('問題は見つかりませんでした。');
  await page.keyboard.press('Escape');
});

test('R00c: Japanese Automate and Add-ins dialogs use localized copy', async ({ page }) => {
  await mount(page, '/?locale=ja');

  await page.getByRole('tab', { name: '自動化', exact: true }).click();
  await expect(page.locator('[data-ribbon-command="recordActions"]')).toHaveText(
    /アクションの記録/,
  );
  await expect(page.locator('[data-ribbon-command="allScripts"]')).toHaveText(/すべてのスクリプト/);
  await page.locator('[data-ribbon-command="script"]').click();
  await expect(page.locator('#menu-script')).toContainText('大文字に変換');
  await page.locator('#menu-script [data-script-action="custom"]').click();
  await expect(page.getByRole('dialog', { name: 'スクリプト' })).toBeVisible();
  await expect(page.getByRole('textbox', { name: 'コマンド' })).toHaveAttribute(
    'placeholder',
    'スクリプト コマンドを入力してください: uppercase, lowercase, trim, clear',
  );
  await page.getByRole('textbox', { name: 'コマンド' }).fill('bad-command');
  await page.getByRole('button', { name: '実行' }).click();
  await expect(page.getByRole('alert')).toContainText(
    'uppercase、lowercase、trim、clear のいずれかを入力してください。',
  );
  await page.keyboard.press('Escape');

  await page.getByRole('tab', { name: 'Acrobat', exact: true }).click();
  await page.locator('[data-ribbon-command="addIn"]').click();
  await expect(page.locator('#menu-add-ins')).toContainText('アドインを取得...');
  await expect(page.locator('#menu-add-ins')).toContainText('個人用アドイン');
  await page.locator('#menu-add-ins [data-add-in-action="my"]').click();
  await expect(page.getByRole('dialog', { name: 'アドイン' })).toBeVisible();
  await expect(page.getByRole('dialog')).toContainText('組み込みアドイン');
  await expect(page.getByRole('dialog')).toContainText('外部アドイン');
  await page.keyboard.press('Escape');

  await page.locator('[data-ribbon-command="pdf"]').click();
  await expect(page.locator('#menu-pdf')).toContainText('PDF を作成してリンクを共有');
  await page.locator('#menu-pdf [data-pdf-action="share"]').click();
  await expect(page.getByRole('dialog', { name: 'PDF' })).toContainText(
    'PDF の書き出し準備ができました。',
  );
  await page.keyboard.press('Escape');
});

test('R00d: Japanese Page Layout background menu uses localized labels', async ({ page }) => {
  await mount(page, '/?locale=ja');

  await page.getByRole('tab', { name: 'ページ レイアウト', exact: true }).click();
  await selectRangeAndSetValues(page, { r0: 1, c0: 1, r1: 2, c1: 2 }, []);
  await page.locator('[data-ribbon-command="printArea"]').click();
  const printAreaMenu = page.locator('#menu-print-area');
  await expect(printAreaMenu).toBeVisible();
  await expect(printAreaMenu).toContainText('印刷範囲の設定');
  await expect(printAreaMenu).toContainText('印刷範囲のクリア');
  await page.locator('#menu-print-area [data-print-area-action="set"]').click();
  await expect(page.getByRole('alertdialog', { name: '印刷範囲' })).toContainText(
    '印刷範囲を B2:C3 に設定しました。',
  );
  await page.getByRole('button', { name: 'OK', exact: true }).click();

  await page.locator('[data-ribbon-command="printTitles"]').click();
  const printTitlesMenu = page.locator('#menu-print-titles');
  await expect(printTitlesMenu).toBeVisible();
  await expect(printTitlesMenu).toContainText('タイトル行の設定');
  await expect(printTitlesMenu).toContainText('タイトル列の設定');
  await expect(printTitlesMenu).toContainText('印刷タイトルのクリア');
  await page.keyboard.press('Escape');

  await page.locator('[data-ribbon-command="sheetBackground"]').click();

  const menu = page.locator('#menu-sheet-background');
  await expect(menu).toBeVisible();
  await expect(menu).toContainText('背景の選択...');
  await expect(menu).toContainText('背景の削除');

  await page.locator('#menu-sheet-background [data-sheet-background-action="set"]').click();
  const dialog = page.getByRole('dialog', { name: '背景の選択...' });
  await expect(dialog).toBeVisible();
  await expect(dialog).toContainText('背景画像のURL');
  await dialog.getByRole('button', { name: 'キャンセル' }).click();

  await page.locator('[data-ribbon-select="scaleWidth"] .demo__rb-dd__btn').click();
  await expect(page.locator('[data-ribbon-select="scaleWidth"] [data-value="0"]')).toHaveText(
    '自動',
  );
  await expect(page.locator('[data-ribbon-select="scaleWidth"] [data-value="custom"]')).toHaveText(
    'ユーザー設定...',
  );
  await page.locator('[data-ribbon-select="scaleWidth"] [data-value="custom"]').click();
  const widthDialog = page.getByRole('dialog', { name: '横' });
  await expect(widthDialog).toBeVisible();
  await expect(widthDialog.getByRole('textbox', { name: 'ページ数 (1-99)' })).toBeVisible();
  await widthDialog.getByRole('button', { name: 'キャンセル' }).click();

  await page.locator('[data-ribbon-select="scaleHeight"] .demo__rb-dd__btn').click();
  await page.locator('[data-ribbon-select="scaleHeight"] [data-value="custom"]').click();
  const heightDialog = page.getByRole('dialog', { name: '縦' });
  await expect(heightDialog).toBeVisible();
  await expect(heightDialog.getByRole('textbox', { name: 'ページ数 (1-99)' })).toBeVisible();
  await heightDialog.getByRole('button', { name: 'キャンセル' }).click();

  await page.locator('[data-ribbon-select="scalePercent"] .demo__rb-dd__btn').click();
  await page.locator('[data-ribbon-select="scalePercent"] [data-value="custom"]').click();
  const scaleDialog = page.getByRole('dialog', { name: '拡大縮小' });
  await expect(scaleDialog).toBeVisible();
  await expect(scaleDialog.getByRole('textbox', { name: '倍率 (10-400)' })).toBeVisible();
  await scaleDialog.getByRole('button', { name: 'キャンセル' }).click();
});

test('R00e: Japanese Insert Symbol menu uses localized category labels', async ({ page }) => {
  await mount(page, '/?locale=ja');

  await page.getByRole('tab', { name: '挿入', exact: true }).click();
  await page.locator('[data-ribbon-command="symbolInsert"]').click();

  const menu = page.locator('#menu-symbol');
  await expect(menu).toBeVisible();
  await expect(menu).toContainText('数学記号');
  await expect(menu).toContainText('ギリシャ文字');
  await expect(menu).toContainText('通貨記号');
  await expect(menu).toContainText('法務記号');
  await expect(menu).toContainText('その他の記号...');
});

async function closeDialog(page: Page): Promise<void> {
  await page
    .getByRole('button', { name: /^(Cancel|Close)$/ })
    .last()
    .click();
}

type ActiveCellFormat = {
  bold?: boolean;
  italic?: boolean;
  strike?: boolean;
  fontFamily?: string;
  fontSize?: number;
  align?: string;
  vAlign?: string;
  wrap?: boolean;
  indent?: number;
  rotation?: number;
  color?: string;
  fill?: string;
  underline?: boolean;
  borders?: Record<string, unknown>;
  hyperlink?: string;
  comment?: string;
  locked?: boolean;
  numFmt?: {
    kind?: string;
    decimals?: number;
    symbol?: string;
    thousands?: boolean;
  };
};

type ValidationSummary =
  | {
      kind?: string;
      source?: string[] | { ref: string };
      op?: string;
      a?: number;
      b?: number;
      formula?: string;
      allowBlank?: boolean;
      errorStyle?: string;
      promptTitle?: string;
      promptMessage?: string;
      errorTitle?: string;
      errorMessage?: string;
    }
  | null
  | undefined;

type ConditionalRuleSummary = {
  kind: string;
  color?: string;
  stops?: string[];
  icons?: string;
  period?: string;
  mode?: string;
  n?: number;
  percent?: boolean;
  op?: string;
  a?: number;
  b?: number;
  text?: string;
  formula?: string;
};

type CommentSummary = {
  addr: CellAddr;
  text: string;
};

type TraceSummary = {
  kind: string;
  from: CellAddr;
  to: CellAddr;
};

type CellAddr = {
  sheet: number;
  row: number;
  col: number;
};

async function readActiveCellFormat(page: Page): Promise<ActiveCellFormat | null | undefined> {
  return page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            getState: () => {
              selection: { active: { sheet: number; row: number; col: number } };
              format: { formats: Map<string, ActiveCellFormat> };
            };
          };
        }
      | undefined;
    const state = inst?.store.getState();
    const active = state?.selection.active;
    return active ? state?.format.formats.get(`${active.sheet}:${active.row}:${active.col}`) : null;
  });
}

async function readActiveCellLocked(page: Page): Promise<boolean | null> {
  return page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            getState: () => {
              selection: { active: { sheet: number; row: number; col: number } };
              format: { formats: Map<string, { locked?: boolean }> };
            };
          };
        }
      | undefined;
    const state = inst?.store.getState();
    const active = state?.selection.active;
    if (!active) return null;
    return state?.format.formats.get(`${active.sheet}:${active.row}:${active.col}`)?.locked ?? null;
  });
}

async function patchActiveCellFormat(page: Page, patch: ActiveCellFormat): Promise<void> {
  await page.evaluate((patch) => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            setState: (fn: (state: unknown) => unknown) => void;
            getState: () => {
              selection: { active: { sheet: number; row: number; col: number } };
              format: { formats: Map<string, ActiveCellFormat> };
            };
          };
        }
      | undefined;
    const state = inst?.store.getState();
    const active = state?.selection.active;
    if (!inst || !state || !active) return;
    const key = `${active.sheet}:${active.row}:${active.col}`;
    inst.store.setState((raw) => {
      const s = raw as typeof state;
      const formats = new Map(s.format.formats);
      formats.set(key, { ...(formats.get(key) ?? {}), ...patch });
      return { ...s, format: { ...s.format, formats } };
    });
  }, patch);
}

async function readActiveValidation(page: Page): Promise<ValidationSummary> {
  return page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            getState: () => {
              selection: { active: { sheet: number; row: number; col: number } };
              format: { formats: Map<string, { validation?: ValidationSummary }> };
            };
          };
        }
      | undefined;
    const state = inst?.store.getState();
    const active = state?.selection.active;
    return active
      ? (state?.format.formats.get(`${active.sheet}:${active.row}:${active.col}`)?.validation ??
          null)
      : null;
  });
}

async function readValidationCircles(page: Page): Promise<string[]> {
  return page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            getState: () => {
              errorIndicators: { validationCircles: Set<string> };
            };
          };
        }
      | undefined;
    return [...(inst?.store.getState().errorIndicators.validationCircles ?? new Set<string>())];
  });
}

async function selectCellAndSetText(
  page: Page,
  row: number,
  col: number,
  text: string,
): Promise<void> {
  await page.evaluate(
    ({ row, col, text }) => {
      const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
        | {
            workbook: {
              setText: (addr: CellAddr, value: string) => void;
              recalc: () => void;
            };
            store: {
              setState: (
                fn: (state: {
                  selection: {
                    active: CellAddr;
                    anchor: CellAddr;
                    range: { sheet: number; r0: number; c0: number; r1: number; c1: number };
                    extraRanges?: unknown[];
                  };
                }) => unknown,
              ) => void;
            };
          }
        | undefined;
      if (!inst) return;
      const active = { sheet: 0, row, col };
      inst.workbook.setText(active, text);
      inst.workbook.recalc();
      inst.store.setState((state) => ({
        ...state,
        selection: {
          ...state.selection,
          active,
          anchor: active,
          range: { sheet: 0, r0: row, c0: col, r1: row, c1: col },
          extraRanges: [],
        },
      }));
    },
    { row, col, text },
  );
}

async function setCommentDirect(page: Page, row: number, col: number, text: string): Promise<void> {
  await page.evaluate(
    ({ row, col, text }) => {
      const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
        | {
            store: {
              setState: (
                fn: (state: { format: { formats: Map<string, { comment?: string }> } }) => unknown,
              ) => void;
            };
          }
        | undefined;
      if (!inst) return;
      inst.store.setState((state) => {
        const key = `0:${row}:${col}`;
        const formats = new Map(state.format.formats);
        formats.set(key, { ...(formats.get(key) ?? {}), comment: text });
        return { ...state, format: { formats } };
      });
    },
    { row, col, text },
  );
}

async function readCommentSummaries(page: Page): Promise<CommentSummary[]> {
  return page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            getState: () => {
              data: { sheetIndex: number };
              format: { formats: Map<string, { comment?: string }> };
            };
          };
        }
      | undefined;
    const state = inst?.store.getState();
    if (!state) return [];
    const out: CommentSummary[] = [];
    for (const [key, format] of state.format.formats) {
      if (!format.comment) continue;
      const [sheet, row, col] = key.split(':').map((part) => Number(part));
      if (sheet !== state.data.sheetIndex) continue;
      out.push({ addr: { sheet, row, col }, text: format.comment });
    }
    return out.sort((a, b) => a.addr.row - b.addr.row || a.addr.col - b.addr.col);
  });
}

async function readProtectionSummary(page: Page): Promise<{
  protected: boolean;
  password?: string;
  workbookStructureProtected?: boolean;
  workbookPassword?: string;
  allowedEditRanges?: {
    title: string;
    range: { sheet: number; r0: number; c0: number; r1: number; c1: number };
  }[];
}> {
  return page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            getState: () => {
              data: { sheetIndex: number };
              protection: {
                protectedSheets: Map<number, { password?: string }>;
                workbookStructure?: { password?: string };
                allowedEditRanges: {
                  title: string;
                  range: { sheet: number; r0: number; c0: number; r1: number; c1: number };
                }[];
              };
            };
          };
        }
      | undefined;
    const state = inst?.store.getState();
    const sheet = state?.data.sheetIndex ?? 0;
    const entry = state?.protection.protectedSheets.get(sheet);
    return {
      protected: !!entry,
      password: entry?.password,
      workbookStructureProtected: !!state?.protection.workbookStructure,
      workbookPassword: state?.protection.workbookStructure?.password,
      allowedEditRanges: state?.protection.allowedEditRanges.map((entry) => ({
        title: entry.title,
        range: entry.range,
      })),
    };
  });
}

async function readSheetCount(page: Page): Promise<number> {
  return page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | { workbook: { sheetCount: number } }
      | undefined;
    return inst?.workbook.sheetCount ?? 0;
  });
}

async function isCellWritableDirect(page: Page, row: number, col: number): Promise<boolean> {
  return page.evaluate(
    ({ row, col }) => {
      const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
        | {
            store: {
              getState: () => {
                protection: {
                  protectedSheets: Map<number, { password?: string }>;
                  allowedEditRanges: {
                    range: { sheet: number; r0: number; c0: number; r1: number; c1: number };
                  }[];
                };
                format: { formats: Map<string, { locked?: boolean }> };
              };
            };
          }
        | undefined;
      const state = inst?.store.getState();
      if (!state) return false;
      const addr = { sheet: 0, row, col };
      if (!state.protection.protectedSheets.has(0)) return true;
      const allowed = state.protection.allowedEditRanges.some(
        (entry) =>
          entry.range.sheet === addr.sheet &&
          addr.row >= entry.range.r0 &&
          addr.row <= entry.range.r1 &&
          addr.col >= entry.range.c0 &&
          addr.col <= entry.range.c1,
      );
      if (allowed) return true;
      return state.format.formats.get(`${addr.sheet}:${addr.row}:${addr.col}`)?.locked === false;
    },
    { row, col },
  );
}

async function selectRangeAndSetValues(
  page: Page,
  range: { r0: number; c0: number; r1: number; c1: number },
  values: Array<{ row: number; col: number; value: string | number }>,
): Promise<void> {
  await page.evaluate(
    ({ range, values }) => {
      const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
        | {
            workbook: {
              setText: (addr: CellAddr, value: string) => void;
              setNumber: (addr: CellAddr, value: number) => void;
              cells: (sheet: number) => Iterable<{
                addr: CellAddr;
                value: unknown;
                formula: string | null;
              }>;
              recalc: () => void;
            };
            store: {
              setState: (
                fn: (state: {
                  data: { cells: Map<string, unknown> };
                  selection: {
                    active: CellAddr;
                    anchor: CellAddr;
                    range: { sheet: number; r0: number; c0: number; r1: number; c1: number };
                    extraRanges?: unknown[];
                  };
                }) => unknown,
              ) => void;
            };
          }
        | undefined;
      if (!inst) return;
      for (const entry of values) {
        const addr = { sheet: 0, row: entry.row, col: entry.col };
        if (typeof entry.value === 'number') inst.workbook.setNumber(addr, entry.value);
        else inst.workbook.setText(addr, entry.value);
      }
      inst.workbook.recalc();
      const active = { sheet: 0, row: range.r0, col: range.c0 };
      const cells = new Map<string, { value: unknown; formula: string | null }>();
      for (const cell of inst.workbook.cells(0)) {
        cells.set(`${cell.addr.sheet}:${cell.addr.row}:${cell.addr.col}`, {
          value: cell.value,
          formula: cell.formula,
        });
      }
      inst.store.setState((state) => ({
        ...state,
        data: { ...state.data, cells },
        selection: {
          ...state.selection,
          active,
          anchor: active,
          range: { sheet: 0, ...range },
          extraRanges: [],
        },
      }));
    },
    { range, values },
  );
}

async function selectRangeAndSetFormulas(
  page: Page,
  range: { r0: number; c0: number; r1: number; c1: number },
  formulas: Array<{ row: number; col: number; formula: string }>,
): Promise<void> {
  await page.evaluate(
    ({ range, formulas }) => {
      const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
        | {
            workbook: {
              setFormula: (addr: CellAddr, formula: string) => void;
              recalc: () => void;
            };
            store: {
              setState: (
                fn: (state: {
                  selection: {
                    active: CellAddr;
                    anchor: CellAddr;
                    range: { sheet: number; r0: number; c0: number; r1: number; c1: number };
                    extraRanges?: unknown[];
                  };
                }) => unknown,
              ) => void;
            };
          }
        | undefined;
      if (!inst) return;
      for (const entry of formulas) {
        inst.workbook.setFormula({ sheet: 0, row: entry.row, col: entry.col }, entry.formula);
      }
      inst.workbook.recalc();
      const active = { sheet: 0, row: range.r0, col: range.c0 };
      inst.store.setState((state) => ({
        ...state,
        selection: {
          ...state.selection,
          active,
          anchor: active,
          range: { sheet: 0, ...range },
          extraRanges: [],
        },
      }));
    },
    { range, formulas },
  );
}

async function readCellText(page: Page, row: number, col: number): Promise<string | undefined> {
  return page.evaluate(
    ({ row, col }) => {
      const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
        | {
            workbook: {
              getValue: (addr: CellAddr) => { kind: string; value?: string };
            };
          }
        | undefined;
      const value = inst?.workbook.getValue({ sheet: 0, row, col });
      return value?.kind === 'text' ? value.value : undefined;
    },
    { row, col },
  );
}

async function readCellSummary(
  page: Page,
  row: number,
  col: number,
): Promise<{ kind?: string; value?: string | number; formula?: string | null }> {
  return page.evaluate(
    ({ row, col }) => {
      const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
        | {
            workbook: {
              getValue: (addr: CellAddr) => { kind: string; value?: string | number };
              cellFormula: (addr: CellAddr) => string | null;
            };
          }
        | undefined;
      const addr = { sheet: 0, row, col };
      const value = inst?.workbook.getValue(addr);
      return {
        kind: value?.kind,
        value: value?.value,
        formula: inst?.workbook.cellFormula(addr) ?? null,
      };
    },
    { row, col },
  );
}

async function readCellsGroupState(page: Page): Promise<{
  sheetCount: number;
  activeSheet: number;
  hiddenRows: number[];
  hiddenCols: number[];
  hiddenSheets: number[];
  sheetTabColors: Array<[number, string]>;
}> {
  return page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          workbook: { sheetCount: number };
          store: {
            getState: () => {
              data: { sheetIndex: number };
              layout: {
                hiddenRows: Set<number>;
                hiddenCols: Set<number>;
                hiddenSheets: Set<number>;
                sheetTabColors: Map<number, string>;
              };
            };
          };
        }
      | undefined;
    const state = inst?.store.getState();
    return {
      sheetCount: inst?.workbook.sheetCount ?? 0,
      activeSheet: state?.data.sheetIndex ?? -1,
      hiddenRows: [...(state?.layout.hiddenRows ?? new Set<number>())].sort((a, b) => a - b),
      hiddenCols: [...(state?.layout.hiddenCols ?? new Set<number>())].sort((a, b) => a - b),
      hiddenSheets: [...(state?.layout.hiddenSheets ?? new Set<number>())].sort((a, b) => a - b),
      sheetTabColors: [...(state?.layout.sheetTabColors ?? new Map<number, string>()).entries()],
    };
  });
}

async function undoViaInstance(page: Page): Promise<boolean> {
  return page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | { undo: () => boolean }
      | undefined;
    return inst?.undo() ?? false;
  });
}

async function redoViaInstance(page: Page): Promise<boolean> {
  return page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | { redo: () => boolean }
      | undefined;
    return inst?.redo() ?? false;
  });
}

async function readDefinedNames(page: Page): Promise<Array<{ name: string; formula: string }>> {
  return page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | { workbook: { definedNames: () => Iterable<{ name: string; formula: string }> } }
      | undefined;
    return inst ? Array.from(inst.workbook.definedNames()) : [];
  });
}

async function readWatchAddresses(page: Page): Promise<CellAddr[]> {
  return page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | { store: { getState: () => { watch: { watches: CellAddr[] } } } }
      | undefined;
    return inst?.store.getState().watch.watches ?? [];
  });
}

async function readIgnoredErrorKeys(page: Page): Promise<string[]> {
  return page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | { store: { getState: () => { errorIndicators: { ignoredErrors: Set<string> } } } }
      | undefined;
    return [...(inst?.store.getState().errorIndicators.ignoredErrors ?? new Set<string>())];
  });
}

async function readSelectionSummary(page: Page): Promise<{
  active: CellAddr;
  range: { sheet: number; r0: number; c0: number; r1: number; c1: number };
}> {
  return page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            getState: () => {
              selection: {
                active: CellAddr;
                range: { sheet: number; r0: number; c0: number; r1: number; c1: number };
              };
            };
          };
        }
      | undefined;
    const selection = inst?.store.getState().selection;
    return {
      active: selection?.active ?? { sheet: 0, row: -1, col: -1 },
      range: selection?.range ?? { sheet: 0, r0: -1, c0: -1, r1: -1, c1: -1 },
    };
  });
}

async function readInsertObjectSummary(page: Page): Promise<{
  tables: Array<{
    id?: string;
    source?: string;
    range?: unknown;
    style?: string;
    showHeader?: boolean;
  }>;
  charts: Array<{ id?: string; kind?: string; source?: unknown }>;
  pivots: Array<{ sheetIndex?: number; top?: number; left?: number; fields?: string[] }>;
}> {
  return page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          workbook: {
            getPivotTables: () => Array<{
              sheetIndex?: number;
              top?: number;
              left?: number;
              fields?: string[];
            }>;
          };
          store: {
            getState: () => {
              tables: {
                tables: Array<{
                  id?: string;
                  source?: string;
                  range?: unknown;
                  style?: string;
                  showHeader?: boolean;
                }>;
              };
              charts: { charts: Array<{ id?: string; kind?: string; source?: unknown }> };
            };
          };
        }
      | undefined;
    const state = inst?.store.getState();
    return {
      tables:
        state?.tables.tables.map((table) => ({
          id: table.id,
          source: table.source,
          range: table.range,
          style: table.style,
          showHeader: table.showHeader,
        })) ?? [],
      charts:
        state?.charts.charts.map((chart) => ({
          id: chart.id,
          kind: chart.kind,
          source: chart.source,
        })) ?? [],
      pivots:
        inst?.workbook.getPivotTables().map((pivot) => ({
          sheetIndex: pivot.sheetIndex,
          top: pivot.top,
          left: pivot.left,
          fields: pivot.fields,
        })) ?? [],
    };
  });
}

async function readPageLayoutSummary(page: Page): Promise<{
  setup: {
    orientation?: string;
    paperSize?: string;
    margins?: { top?: number; right?: number; bottom?: number; left?: number };
    printArea?: string;
    printTitleRows?: string;
    printTitleCols?: string;
    fitWidth?: number;
    fitHeight?: number;
    scale?: number;
    manualPageBreakRows?: number[];
    manualPageBreakCols?: number[];
    showGridlines?: boolean;
    showHeadings?: boolean;
  };
  ui: {
    showGridLines?: boolean;
    showHeaders?: boolean;
    background?: string;
  };
}> {
  return page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            getState: () => {
              data: { sheetIndex: number };
              pageSetup: {
                setupBySheet: Map<
                  number,
                  {
                    orientation?: string;
                    paperSize?: string;
                    margins?: { top?: number; right?: number; bottom?: number; left?: number };
                    printArea?: string;
                    printTitleRows?: string;
                    printTitleCols?: string;
                    fitWidth?: number;
                    fitHeight?: number;
                    scale?: number;
                    manualPageBreakRows?: number[];
                    manualPageBreakCols?: number[];
                    showGridlines?: boolean;
                    showHeadings?: boolean;
                  }
                >;
              };
              ui: {
                showGridLines?: boolean;
                showHeaders?: boolean;
                sheetBackgroundImages: Map<number, string>;
              };
            };
          };
        }
      | undefined;
    const state = inst?.store.getState();
    const sheet = state?.data.sheetIndex ?? 0;
    return {
      setup: state?.pageSetup.setupBySheet.get(sheet) ?? {},
      ui: {
        showGridLines: state?.ui.showGridLines,
        showHeaders: state?.ui.showHeaders,
        background: state?.ui.sheetBackgroundImages.get(sheet),
      },
    };
  });
}

async function readFilterSummary(page: Page): Promise<{
  filterRange: { sheet: number; r0: number; c0: number; r1: number; c1: number } | null;
  filterCriteria: Array<{
    range: { sheet: number; r0: number; c0: number; r1: number; c1: number };
    byCol: number;
    hiddenValues: string[];
  }>;
  hiddenRows: number[];
}> {
  return page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            getState: () => {
              ui: {
                filterRange: {
                  sheet: number;
                  r0: number;
                  c0: number;
                  r1: number;
                  c1: number;
                } | null;
                filterCriteria: Array<{
                  range: { sheet: number; r0: number; c0: number; r1: number; c1: number };
                  byCol: number;
                  hiddenValues: string[];
                }>;
              };
              layout: { hiddenRows: Set<number> };
            };
          };
        }
      | undefined;
    const state = inst?.store.getState();
    return {
      filterRange: state?.ui.filterRange ?? null,
      filterCriteria:
        state?.ui.filterCriteria.map((criteria) => ({
          range: criteria.range,
          byCol: criteria.byCol,
          hiddenValues: [...criteria.hiddenValues],
        })) ?? [],
      hiddenRows: state ? [...state.layout.hiddenRows].sort((a, b) => a - b) : [],
    };
  });
}

async function readViewSummary(page: Page): Promise<{
  workbookView?: string;
  showGridLines?: boolean;
  showHeaders?: boolean;
  showFormulas?: boolean;
  r1c1?: boolean;
  zoom?: number;
  freezeRows?: number;
  freezeCols?: number;
  sheetViews?: { id: string; name: string; sheet: number }[];
  activeSheetViewId?: string | null;
  formulaBarAttached: boolean;
}> {
  return page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            getState: () => {
              ui: {
                workbookView?: string;
                showGridLines?: boolean;
                showHeaders?: boolean;
                showFormulas?: boolean;
                r1c1?: boolean;
              };
              viewport: { zoom?: number };
              layout: { freezeRows: number; freezeCols: number };
              sheetViews: {
                views: { id: string; name: string; sheet: number }[];
                activeViewId: string | null;
              };
            };
          };
        }
      | undefined;
    const state = inst?.store.getState();
    return {
      workbookView: state?.ui.workbookView,
      showGridLines: state?.ui.showGridLines,
      showHeaders: state?.ui.showHeaders,
      showFormulas: state?.ui.showFormulas,
      r1c1: state?.ui.r1c1,
      zoom: state?.viewport.zoom,
      freezeRows: state?.layout.freezeRows,
      freezeCols: state?.layout.freezeCols,
      sheetViews: state?.sheetViews.views,
      activeSheetViewId: state?.sheetViews.activeViewId,
      formulaBarAttached: !!document.querySelector('.fc-host__formulabar'),
    };
  });
}

async function readTraceSummaries(page: Page): Promise<TraceSummary[]> {
  return page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            getState: () => {
              traces: { items: TraceSummary[] };
            };
          };
        }
      | undefined;
    return inst?.store.getState().traces.items ?? [];
  });
}

async function readLayoutSummary(page: Page): Promise<{
  hiddenRows: number[];
  hiddenCols: number[];
  rowHeights: [number, number][];
  colWidths: [number, number][];
  outlineRows: [number, number][];
  outlineCols: [number, number][];
  outlineRowGutter: number;
  outlineColGutter: number;
}> {
  return page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            getState: () => {
              layout: {
                hiddenRows: Set<number>;
                hiddenCols: Set<number>;
                rowHeights: Map<number, number>;
                colWidths: Map<number, number>;
                outlineRows: Map<number, number>;
                outlineCols: Map<number, number>;
                outlineRowGutter: number;
                outlineColGutter: number;
              };
            };
          };
        }
      | undefined;
    const layout = inst?.store.getState().layout;
    return {
      hiddenRows: layout ? [...layout.hiddenRows].sort((a, b) => a - b) : [],
      hiddenCols: layout ? [...layout.hiddenCols].sort((a, b) => a - b) : [],
      rowHeights: layout ? [...layout.rowHeights.entries()].sort(([a], [b]) => a - b) : [],
      colWidths: layout ? [...layout.colWidths.entries()].sort(([a], [b]) => a - b) : [],
      outlineRows: layout ? [...layout.outlineRows.entries()].sort(([a], [b]) => a - b) : [],
      outlineCols: layout ? [...layout.outlineCols.entries()].sort(([a], [b]) => a - b) : [],
      outlineRowGutter: layout?.outlineRowGutter ?? 0,
      outlineColGutter: layout?.outlineColGutter ?? 0,
    };
  });
}

async function readConditionalRuleSummaries(page: Page): Promise<ConditionalRuleSummary[]> {
  return page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            getState: () => {
              conditional: {
                rules: Array<{
                  kind: string;
                  color?: string;
                  stops?: string[];
                  icons?: string;
                  period?: string;
                  mode?: string;
                  n?: number;
                  percent?: boolean;
                  op?: string;
                  a?: number;
                  b?: number;
                  text?: string;
                  formula?: string;
                }>;
              };
            };
          };
        }
      | undefined;
    return (
      inst?.store.getState().conditional.rules.map((rule) => ({
        kind: rule.kind,
        color: rule.color,
        stops: rule.stops,
        icons: rule.icons,
        period: rule.period,
        mode: rule.mode,
        n: rule.n,
        percent: rule.percent,
        op: rule.op,
        a: rule.a,
        b: rule.b,
        text: rule.text,
        formula: rule.formula,
      })) ?? []
    );
  });
}

test('R01: ribbon tabs switch visible panels and render expected commands', async ({ page }) => {
  await mount(page, '/?locale=en');

  for (const tab of ribbonTabs) {
    await page.getByRole('tab', { name: tab.label, exact: true }).click();

    const visiblePanel = page.locator('.demo__ribbon:not([hidden])');
    await expect(visiblePanel).toHaveCount(1);
    await expect(visiblePanel).toHaveAttribute('data-ribbon-panel', tab.id);
    await expect
      .poll(() =>
        visiblePanel
          .locator('[data-ribbon-command]')
          .evaluateAll((nodes) => nodes.map((node) => node.getAttribute('data-ribbon-command'))),
      )
      .toEqual([...tab.commands]);
  }

  await page.getByRole('tab', { name: 'Home', exact: true }).click();
  await expect(page.locator('[data-ribbon-command="formatCellsHome"]')).toHaveAttribute(
    'aria-keyshortcuts',
    /Control\+1/,
  );
  await expect(page.locator('[data-ribbon-command="findHome"]')).toHaveAttribute(
    'aria-keyshortcuts',
    /Control\+F/,
  );

  await page.getByRole('tab', { name: 'Formulas', exact: true }).click();
  await expect(page.locator('[data-ribbon-command="namedRanges"]')).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+F3',
  );
  await expect(page.locator('[data-ribbon-command="fx"]')).toHaveAttribute(
    'aria-keyshortcuts',
    'Shift+F3',
  );
});

test('R02: Home font controls render and apply formatting', async ({ page }) => {
  await mount(page, '/?locale=en');

  await expect(page.locator('[data-ribbon-select="fontFamily"] .demo__rb-dd__value')).toHaveText(
    'Aptos',
  );
  await expect(page.locator('[data-ribbon-select="fontSize"] .demo__rb-dd__value')).toHaveText(
    '11',
  );
  await expect(page.locator('[data-ribbon-command="borders"]')).toBeVisible();
  await expect(page.locator('[data-ribbon-command="fontColor"] .demo__rb-color__swatch')).toHaveCSS(
    'background-color',
    'rgb(32, 31, 30)',
  );
  await expect(page.locator('[data-ribbon-command="fillColor"] .demo__rb-color__swatch')).toHaveCSS(
    'background-color',
    'rgb(255, 255, 255)',
  );

  await page.locator('[data-ribbon-command="bold"]').click();
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ bold: true });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(async () => (await readActiveCellFormat(page))?.bold).toBeUndefined();
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ bold: true });

  await page.locator('[data-ribbon-command="italic"]').click();
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ italic: true });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(async () => (await readActiveCellFormat(page))?.italic).toBeUndefined();
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ italic: true });

  await page.locator('[data-ribbon-command="underline"]').click();
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ underline: true });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(async () => (await readActiveCellFormat(page))?.underline).toBeUndefined();
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ underline: true });

  await page.locator('[data-ribbon-command="strike"]').click();
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ strike: true });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(async () => (await readActiveCellFormat(page))?.strike).toBeUndefined();
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ strike: true });

  await page.locator('[data-ribbon-select="fontFamily"] .demo__rb-dd__btn').click();
  await expect(page.locator('[data-ribbon-select="fontFamily"] .demo__rb-dd__list')).toBeVisible();
  await page.locator('[data-ribbon-select="fontFamily"] [data-value="Arial"]').click();
  await expect(page.locator('[data-ribbon-select="fontFamily"] .demo__rb-dd__value')).toHaveText(
    'Arial',
  );
  expect(await undoViaInstance(page)).toBe(true);
  await expect(page.locator('[data-ribbon-select="fontFamily"] .demo__rb-dd__value')).toHaveText(
    'Aptos',
  );
  expect(await redoViaInstance(page)).toBe(true);
  await expect(page.locator('[data-ribbon-select="fontFamily"] .demo__rb-dd__value')).toHaveText(
    'Arial',
  );

  await page.locator('[data-ribbon-select="fontSize"] .demo__rb-dd__btn').click();
  await page.locator('[data-ribbon-select="fontSize"] [data-value="14"]').click();
  await expect(page.locator('[data-ribbon-select="fontSize"] .demo__rb-dd__value')).toHaveText(
    '14',
  );
  expect(await undoViaInstance(page)).toBe(true);
  await expect(page.locator('[data-ribbon-select="fontSize"] .demo__rb-dd__value')).toHaveText(
    '11',
  );
  expect(await redoViaInstance(page)).toBe(true);
  await expect(page.locator('[data-ribbon-select="fontSize"] .demo__rb-dd__value')).toHaveText(
    '14',
  );

  await page.locator('[data-ribbon-command="fontGrow"]').click();
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ fontSize: 15 });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ fontSize: 14 });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ fontSize: 15 });

  await page.locator('[data-ribbon-command="fontShrink"]').click();
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ fontSize: 14 });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ fontSize: 15 });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ fontSize: 14 });

  await page.locator('[data-ribbon-command="fontColor"] .demo__rb-color__btn').click();
  await page.locator('.demo__color-flyout [data-color="#c00000"]').click();
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ color: '#c00000' });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(async () => (await readActiveCellFormat(page))?.color).toBeUndefined();
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ color: '#c00000' });

  await page.locator('[data-ribbon-command="fillColor"] .demo__rb-color__btn').click();
  await page.locator('.demo__color-flyout [data-color="#ffff00"]').click();
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ fill: '#ffff00' });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(async () => (await readActiveCellFormat(page))?.fill).toBeUndefined();
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ fill: '#ffff00' });

  const activeFormat = await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            getState: () => {
              selection: { active: { sheet: number; row: number; col: number } };
              format: { formats: Map<string, { fontFamily?: string; fontSize?: number }> };
            };
          };
        }
      | undefined;
    const state = inst?.store.getState();
    const active = state?.selection.active;
    return active ? state?.format.formats.get(`${active.sheet}:${active.row}:${active.col}`) : null;
  });

  expect(activeFormat).toMatchObject({
    bold: true,
    italic: true,
    underline: true,
    strike: true,
    fontFamily: 'Arial',
    fontSize: 14,
    color: '#c00000',
    fill: '#ffff00',
  });
});

test('R02-number: Home number controls apply Excel-style formats and undo', async ({ page }) => {
  await mount(page, '/?locale=en&fixture=empty');
  await selectRangeAndSetValues(page, { r0: 0, c0: 0, r1: 0, c1: 0 }, [
    { row: 0, col: 0, value: 1234.567 },
  ]);

  await page.locator('[data-ribbon-select="numberFormat"] .demo__rb-dd__btn').click();
  await page.locator('[data-ribbon-select="numberFormat"] [data-value="currency"]').click();
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({
      numFmt: { kind: 'currency', decimals: 2, symbol: '$' },
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(async () => (await readActiveCellFormat(page))?.numFmt).toBeUndefined();
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({
      numFmt: { kind: 'currency', decimals: 2, symbol: '$' },
    });
  await expect(page.locator('[data-ribbon-select="numberFormat"] .demo__rb-dd__value')).toHaveText(
    'Currency',
  );

  await page.locator('[data-ribbon-command="decDown"]').click();
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({
      numFmt: { kind: 'currency', decimals: 1, symbol: '$' },
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({
      numFmt: { kind: 'currency', decimals: 2, symbol: '$' },
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({
      numFmt: { kind: 'currency', decimals: 1, symbol: '$' },
    });
  await page.locator('[data-ribbon-command="decUp"]').click();
  await page.locator('[data-ribbon-command="decUp"]').click();
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({
      numFmt: { kind: 'currency', decimals: 3, symbol: '$' },
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({
      numFmt: { kind: 'currency', decimals: 2, symbol: '$' },
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({
      numFmt: { kind: 'currency', decimals: 3, symbol: '$' },
    });

  await page.locator('[data-ribbon-command="percent"]').click();
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({
      numFmt: { kind: 'percent', decimals: 0 },
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({
      numFmt: { kind: 'currency', decimals: 3, symbol: '$' },
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({
      numFmt: { kind: 'percent', decimals: 0 },
    });
  await page.locator('[data-ribbon-command="decUp"]').click();
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({
      numFmt: { kind: 'percent', decimals: 1 },
    });

  await page.locator('[data-ribbon-command="comma"]').click();
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({
      numFmt: { kind: 'fixed', decimals: 2, thousands: true },
    });

  const undoComma = await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          undo: () => boolean;
          store: {
            getState: () => {
              selection: { active: CellAddr };
              format: { formats: Map<string, ActiveCellFormat> };
            };
          };
        }
      | undefined;
    const ok = inst?.undo() ?? false;
    const state = inst?.store.getState();
    const active = state?.selection.active;
    return {
      ok,
      format: active
        ? state?.format.formats.get(`${active.sheet}:${active.row}:${active.col}`)
        : null,
    };
  });
  expect(undoComma).toMatchObject({
    ok: true,
    format: { numFmt: { kind: 'percent', decimals: 1 } },
  });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({
      numFmt: { kind: 'fixed', decimals: 2, thousands: true },
    });

  await page.locator('[data-ribbon-select="numberFormat"] .demo__rb-dd__btn').click();
  await page.locator('[data-ribbon-select="numberFormat"] [data-value="more"]').click();
  await expect(page.getByRole('dialog', { name: 'Format Cells' })).toBeVisible();
  await closeDialog(page);
});

test('R02a: Home Paste preserves internal multi-cell payloads and undoes as one step', async ({
  page,
}) => {
  await mount(page, '/?locale=en&fixture=empty');

  await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          workbook: {
            setNumber: (addr: CellAddr, value: number) => void;
            setText: (addr: CellAddr, value: string) => void;
            recalc: () => void;
          };
          store: {
            setState: (
              fn: (state: {
                data: { cells: Map<string, unknown> };
                selection: {
                  active: CellAddr;
                  anchor: CellAddr;
                  range: { sheet: number; r0: number; c0: number; r1: number; c1: number };
                  extraRanges?: unknown[];
                };
                format: { formats: Map<string, Record<string, unknown>> };
              }) => unknown,
            ) => void;
          };
        }
      | undefined;
    if (!inst) return;
    inst.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 11);
    inst.workbook.setText({ sheet: 0, row: 0, col: 1 }, 'north');
    inst.workbook.setNumber({ sheet: 0, row: 1, col: 0 }, 22);
    inst.workbook.setText({ sheet: 0, row: 1, col: 1 }, 'south');
    inst.workbook.recalc();
    inst.store.setState((state) => {
      const cells = new Map(state.data.cells);
      cells.set('0:0:0', { value: { kind: 'number', value: 11 }, formula: null });
      cells.set('0:0:1', { value: { kind: 'text', value: 'north' }, formula: null });
      cells.set('0:1:0', { value: { kind: 'number', value: 22 }, formula: null });
      cells.set('0:1:1', { value: { kind: 'text', value: 'south' }, formula: null });
      const formats = new Map(state.format.formats);
      formats.set('0:0:0', { bold: true, fill: '#fff2cc' });
      formats.set('0:1:1', { italic: true, color: '#c00000' });
      return {
        ...state,
        data: { ...state.data, cells },
        format: { formats },
        selection: {
          ...state.selection,
          active: { sheet: 0, row: 0, col: 0 },
          anchor: { sheet: 0, row: 0, col: 0 },
          range: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
          extraRanges: [],
        },
      };
    });
  });

  await page.locator('button[data-ribbon-command="copy"]').click();
  await page.evaluate(() => navigator.clipboard.writeText('11\tnorth\r\n22\tsouth'));
  await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            setState: (
              fn: (state: {
                selection: {
                  active: CellAddr;
                  anchor: CellAddr;
                  range: { sheet: number; r0: number; c0: number; r1: number; c1: number };
                  extraRanges?: unknown[];
                };
              }) => unknown,
            ) => void;
          };
        }
      | undefined;
    if (!inst) return;
    const active = { sheet: 0, row: 3, col: 3 };
    inst.store.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        active,
        anchor: active,
        range: { sheet: 0, r0: 3, c0: 3, r1: 3, c1: 3 },
        extraRanges: [],
      },
    }));
  });

  await page.locator('button[data-ribbon-command="paste"]').click();
  await expect
    .poll(() =>
      page.evaluate(() => {
        const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
          | {
              workbook: {
                getValue: (addr: CellAddr) => { kind: string; value?: string | number };
              };
            }
          | undefined;
        return inst?.workbook.getValue({ sheet: 0, row: 3, col: 3 }) ?? null;
      }),
    )
    .toEqual({ kind: 'number', value: 11 });

  const pasted = await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          workbook: {
            getValue: (addr: CellAddr) => { kind: string; value?: string | number };
          };
          store: {
            getState: () => {
              selection: {
                range: { sheet: number; r0: number; c0: number; r1: number; c1: number };
              };
              format: { formats: Map<string, Record<string, unknown>> };
            };
          };
        }
      | undefined;
    const state = inst?.store.getState();
    return inst && state
      ? {
          values: [
            inst.workbook.getValue({ sheet: 0, row: 3, col: 3 }),
            inst.workbook.getValue({ sheet: 0, row: 3, col: 4 }),
            inst.workbook.getValue({ sheet: 0, row: 4, col: 3 }),
            inst.workbook.getValue({ sheet: 0, row: 4, col: 4 }),
          ],
          range: state.selection.range,
          formats: {
            topLeft: state.format.formats.get('0:3:3') ?? null,
            bottomRight: state.format.formats.get('0:4:4') ?? null,
          },
        }
      : null;
  });
  expect(pasted).toMatchObject({
    values: [
      { kind: 'number', value: 11 },
      { kind: 'text', value: 'north' },
      { kind: 'number', value: 22 },
      { kind: 'text', value: 'south' },
    ],
    range: { sheet: 0, r0: 3, c0: 3, r1: 4, c1: 4 },
    formats: {
      topLeft: { bold: true, fill: '#fff2cc' },
      bottomRight: { italic: true, color: '#c00000' },
    },
  });

  const isMac = await page.evaluate(() => navigator.platform.toLowerCase().includes('mac'));
  await page.keyboard.press(`${isMac ? 'Meta' : 'Control'}+Z`);

  const undone = await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          workbook: {
            recalc: () => void;
            getValue: (addr: CellAddr) => { kind: string; value?: string | number };
          };
          store: {
            getState: () => {
              format: { formats: Map<string, Record<string, unknown>> };
            };
          };
        }
      | undefined;
    inst?.workbook.recalc();
    const state = inst?.store.getState();
    return inst && state
      ? {
          values: [
            inst.workbook.getValue({ sheet: 0, row: 3, col: 3 }),
            inst.workbook.getValue({ sheet: 0, row: 3, col: 4 }),
            inst.workbook.getValue({ sheet: 0, row: 4, col: 3 }),
            inst.workbook.getValue({ sheet: 0, row: 4, col: 4 }),
          ],
          formats: {
            topLeft: state.format.formats.get('0:3:3') ?? null,
            bottomRight: state.format.formats.get('0:4:4') ?? null,
          },
        }
      : null;
  });
  expect(undone).toEqual({
    values: [{ kind: 'blank' }, { kind: 'blank' }, { kind: 'blank' }, { kind: 'blank' }],
    formats: { topLeft: null, bottomRight: null },
  });
});

test('R02aa: Home Cut moves values and formats with undoable source cleanup', async ({ page }) => {
  await mount(page, '/?locale=en&fixture=empty');

  await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          workbook: {
            setText: (addr: CellAddr, value: string) => void;
            recalc: () => void;
          };
          store: {
            setState: (
              fn: (state: {
                data: { cells: Map<string, unknown> };
                selection: {
                  active: CellAddr;
                  anchor: CellAddr;
                  range: { sheet: number; r0: number; c0: number; r1: number; c1: number };
                  extraRanges?: unknown[];
                };
                format: { formats: Map<string, Record<string, unknown>> };
              }) => unknown,
            ) => void;
          };
        }
      | undefined;
    if (!inst) return;
    inst.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'move me');
    inst.workbook.recalc();
    inst.store.setState((state) => {
      const cells = new Map(state.data.cells);
      cells.set('0:0:0', { value: { kind: 'text', value: 'move me' }, formula: null });
      const formats = new Map(state.format.formats);
      formats.set('0:0:0', { bold: true, fill: '#fff2cc' });
      return {
        ...state,
        data: { ...state.data, cells },
        format: { formats },
        selection: {
          ...state.selection,
          active: { sheet: 0, row: 0, col: 0 },
          anchor: { sheet: 0, row: 0, col: 0 },
          range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
          extraRanges: [],
        },
      };
    });
  });

  await page.locator('button[data-ribbon-command="cut"]').click();
  const afterCut = await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          workbook: {
            recalc: () => void;
            getValue: (addr: CellAddr) => { kind: string; value?: string | number };
          };
          store: { getState: () => { format: { formats: Map<string, Record<string, unknown>> } } };
        }
      | undefined;
    inst?.workbook.recalc();
    const state = inst?.store.getState();
    return inst && state
      ? {
          value: inst.workbook.getValue({ sheet: 0, row: 0, col: 0 }),
          format: state.format.formats.get('0:0:0') ?? null,
        }
      : null;
  });
  expect(afterCut).toEqual({ value: { kind: 'blank' }, format: null });

  await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            setState: (
              fn: (state: {
                selection: {
                  active: CellAddr;
                  anchor: CellAddr;
                  range: { sheet: number; r0: number; c0: number; r1: number; c1: number };
                  extraRanges?: unknown[];
                };
              }) => unknown,
            ) => void;
          };
        }
      | undefined;
    if (!inst) return;
    const active = { sheet: 0, row: 2, col: 2 };
    inst.store.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        active,
        anchor: active,
        range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
        extraRanges: [],
      },
    }));
  });
  await page.locator('button[data-ribbon-command="paste"]').click();
  await expect
    .poll(() =>
      page.evaluate(() => {
        const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
          | {
              workbook: {
                getValue: (addr: CellAddr) => { kind: string; value?: string | number };
              };
            }
          | undefined;
        return inst?.workbook.getValue({ sheet: 0, row: 2, col: 2 }) ?? null;
      }),
    )
    .toEqual({ kind: 'text', value: 'move me' });

  const afterPaste = await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: { getState: () => { format: { formats: Map<string, Record<string, unknown>> } } };
        }
      | undefined;
    const state = inst?.store.getState();
    return {
      source: state?.format.formats.get('0:0:0') ?? null,
      dest: state?.format.formats.get('0:2:2') ?? null,
    };
  });
  expect(afterPaste).toMatchObject({
    source: null,
    dest: { bold: true, fill: '#fff2cc' },
  });

  const isMac = await page.evaluate(() => navigator.platform.toLowerCase().includes('mac'));
  await page.keyboard.press(`${isMac ? 'Meta' : 'Control'}+Z`);
  await page.keyboard.press(`${isMac ? 'Meta' : 'Control'}+Z`);

  const restored = await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          workbook: {
            recalc: () => void;
            getValue: (addr: CellAddr) => { kind: string; value?: string | number };
          };
          store: { getState: () => { format: { formats: Map<string, Record<string, unknown>> } } };
        }
      | undefined;
    inst?.workbook.recalc();
    const state = inst?.store.getState();
    return inst && state
      ? {
          sourceValue: inst.workbook.getValue({ sheet: 0, row: 0, col: 0 }),
          sourceFormat: state.format.formats.get('0:0:0') ?? null,
          destValue: inst.workbook.getValue({ sheet: 0, row: 2, col: 2 }),
          destFormat: state.format.formats.get('0:2:2') ?? null,
        }
      : null;
  });
  expect(restored).toMatchObject({
    sourceValue: { kind: 'text', value: 'move me' },
    sourceFormat: { bold: true, fill: '#fff2cc' },
    destValue: { kind: 'blank' },
    destFormat: null,
  });
});

test('R02b: Home alignment controls apply cell formatting', async ({ page }) => {
  await mount(page, '/?locale=en');

  await page.getByRole('tab', { name: 'Home', exact: true }).click();

  await page.locator('[data-ribbon-command="alignC"]').click();
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ align: 'center' });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(async () => (await readActiveCellFormat(page))?.align).toBeUndefined();
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ align: 'center' });

  await page.locator('[data-ribbon-command="alignR"]').click();
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ align: 'right' });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ align: 'center' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ align: 'right' });

  await page.locator('[data-ribbon-command="top"]').click();
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ vAlign: 'top' });

  await page.locator('[data-ribbon-command="middle"]').click();
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ vAlign: 'middle' });

  await page.locator('[data-ribbon-command="bottomAlign"]').click();
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ vAlign: 'bottom' });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ vAlign: 'middle' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ vAlign: 'bottom' });

  await page.locator('[data-ribbon-command="wrap"]').click();
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ wrap: true });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(async () => (await readActiveCellFormat(page))?.wrap).toBeUndefined();
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ wrap: true });

  await page.locator('[data-ribbon-command="textOrientation"]').click();
  await page.locator('#menu-text-orientation [data-text-orientation="ccw"]').click();
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ rotation: 45 });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(async () => (await readActiveCellFormat(page))?.rotation).toBeUndefined();
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ rotation: 45 });

  await page.locator('[data-ribbon-command="indentIncrease"]').click();
  await page.locator('[data-ribbon-command="indentIncrease"]').click();
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ indent: 2 });

  await page.locator('[data-ribbon-command="indentDecrease"]').click();
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ indent: 1 });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ indent: 2 });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ indent: 1 });
});

test('R02c: Home border authoring commands arm draw modes and open border dialog', async ({
  page,
}) => {
  await mount(page, '/?locale=en');

  await page.getByRole('tab', { name: 'Home', exact: true }).click();

  await page.locator('[data-ribbon-command="drawBorder"]').click();
  await expect
    .poll(() =>
      page.evaluate(() => {
        const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
          | { borderDraw?: { getMode: () => string | null } }
          | undefined;
        return inst?.borderDraw?.getMode() ?? null;
      }),
    )
    .toBe('draw');

  await page.locator('[data-ribbon-command="drawBorderGrid"]').click();
  await expect
    .poll(() =>
      page.evaluate(() => {
        const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
          | { borderDraw?: { getMode: () => string | null } }
          | undefined;
        return inst?.borderDraw?.getMode() ?? null;
      }),
    )
    .toBe('grid');

  await page.locator('[data-ribbon-command="eraseBorder"]').click();
  await expect
    .poll(() =>
      page.evaluate(() => {
        const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
          | { borderDraw?: { getMode: () => string | null } }
          | undefined;
        return inst?.borderDraw?.getMode() ?? null;
      }),
    )
    .toBe('erase');

  await page.locator('[data-ribbon-command="eraseBorder"]').click();
  await expect
    .poll(() =>
      page.evaluate(() => {
        const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
          | { borderDraw?: { getMode: () => string | null } }
          | undefined;
        return inst?.borderDraw?.getMode() ?? null;
      }),
    )
    .toBeNull();

  await page.locator('[data-ribbon-command="borders"]').click();
  await page.locator('#menu-borders [data-border-preset="bottom"]').click();
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({ borders: { bottom: { style: 'thin' } } });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(async () => (await readActiveCellFormat(page))?.borders).toBeUndefined();
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({ borders: { bottom: { style: 'thin' } } });

  await page.locator('[data-ribbon-command="moreBorders"]').click();
  const formatDialog = page.getByRole('dialog', { name: 'Format Cells' });
  await expect(formatDialog).toBeVisible();
  await expect(formatDialog).toContainText('Border');
  await closeDialog(page);
});

test('R02d: Conditional-format flyouts create and clear preset rules', async ({ page }) => {
  await mount(page, '/?locale=en');

  await page.getByRole('tab', { name: 'Home', exact: true }).click();

  await page.locator('[data-ribbon-command="conditional"]').click();
  await page.locator('[data-cf-submenu="dataBar"]').hover();
  await page.locator('.app__submenu--cf-dataBar [data-cf-action="data-solid-green"]').click();
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([{ kind: 'data-bar', color: '#70ad47' }]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readConditionalRuleSummaries(page)).toEqual([]);
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([{ kind: 'data-bar', color: '#70ad47' }]);

  await page.locator('[data-ribbon-command="conditional"]').click();
  await page.locator('[data-cf-submenu="colorScale"]').hover();
  await page.locator('.app__submenu--cf-colorScale [data-cf-action="scale-ryg"]').click();
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([
      { kind: 'data-bar', color: '#70ad47' },
      { kind: 'color-scale', stops: ['#f8696b', '#ffeb84', '#63be7b'] },
    ]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([{ kind: 'data-bar', color: '#70ad47' }]);
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([
      { kind: 'data-bar', color: '#70ad47' },
      { kind: 'color-scale', stops: ['#f8696b', '#ffeb84', '#63be7b'] },
    ]);

  await page.locator('[data-ribbon-command="conditional"]').click();
  await page.locator('[data-cf-submenu="iconSet"]').hover();
  await page.locator('.app__submenu--cf-iconSet [data-cf-action="icons-traffic3"]').first().click();
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([
      { kind: 'data-bar', color: '#70ad47' },
      { kind: 'color-scale', stops: ['#f8696b', '#ffeb84', '#63be7b'] },
      { kind: 'icon-set', icons: 'traffic3' },
    ]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([
      { kind: 'data-bar', color: '#70ad47' },
      { kind: 'color-scale', stops: ['#f8696b', '#ffeb84', '#63be7b'] },
    ]);
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([
      { kind: 'data-bar', color: '#70ad47' },
      { kind: 'color-scale', stops: ['#f8696b', '#ffeb84', '#63be7b'] },
      { kind: 'icon-set', icons: 'traffic3' },
    ]);

  await page.locator('[data-ribbon-command="conditional"]').click();
  await page.locator('[data-cf-submenu="iconSet"]').hover();
  await page.locator('.app__submenu--cf-iconSet [data-cf-action="icons-symbols3"]').click();
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([
      { kind: 'data-bar', color: '#70ad47' },
      { kind: 'color-scale', stops: ['#f8696b', '#ffeb84', '#63be7b'] },
      { kind: 'icon-set', icons: 'traffic3' },
      { kind: 'icon-set', icons: 'symbols3' },
    ]);

  await page.locator('[data-ribbon-command="conditional"]').click();
  await page.locator('[data-cf-submenu="highlight"]').hover();
  await page.locator('.app__submenu--cf-highlight [data-cf-action="date-occurring"]').click();
  const dateDialog = page.getByRole('dialog', { name: 'A Date Occurring...' });
  await expect(dateDialog).toBeVisible();
  await dateDialog.getByRole('radio', { name: 'Yesterday' }).check();
  await dateDialog.getByRole('button', { name: 'OK' }).click();
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([
      { kind: 'data-bar', color: '#70ad47' },
      { kind: 'color-scale', stops: ['#f8696b', '#ffeb84', '#63be7b'] },
      { kind: 'icon-set', icons: 'traffic3' },
      { kind: 'icon-set', icons: 'symbols3' },
      { kind: 'date-occurring', period: 'yesterday' },
    ]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([
      { kind: 'data-bar', color: '#70ad47' },
      { kind: 'color-scale', stops: ['#f8696b', '#ffeb84', '#63be7b'] },
      { kind: 'icon-set', icons: 'traffic3' },
      { kind: 'icon-set', icons: 'symbols3' },
    ]);
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([
      { kind: 'data-bar', color: '#70ad47' },
      { kind: 'color-scale', stops: ['#f8696b', '#ffeb84', '#63be7b'] },
      { kind: 'icon-set', icons: 'traffic3' },
      { kind: 'icon-set', icons: 'symbols3' },
      { kind: 'date-occurring', period: 'yesterday' },
    ]);

  await page.locator('[data-ribbon-command="conditional"]').click();
  await page.locator('[data-cf-submenu="topBottom"]').hover();
  await page.locator('.app__submenu--cf-topBottom [data-cf-action="top10"]').click();
  const top10Dialog = page.getByRole('dialog', { name: 'Top 10 Items' });
  await expect(top10Dialog).toBeVisible();
  await expect(top10Dialog.getByRole('spinbutton', { name: 'Value' })).toHaveValue('10');
  await top10Dialog.getByRole('button', { name: 'OK' }).click();
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([
      { kind: 'data-bar', color: '#70ad47' },
      { kind: 'color-scale', stops: ['#f8696b', '#ffeb84', '#63be7b'] },
      { kind: 'icon-set', icons: 'traffic3' },
      { kind: 'icon-set', icons: 'symbols3' },
      { kind: 'date-occurring', period: 'yesterday' },
      { kind: 'top-bottom', mode: 'top', n: 10, percent: false },
    ]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([
      { kind: 'data-bar', color: '#70ad47' },
      { kind: 'color-scale', stops: ['#f8696b', '#ffeb84', '#63be7b'] },
      { kind: 'icon-set', icons: 'traffic3' },
      { kind: 'icon-set', icons: 'symbols3' },
      { kind: 'date-occurring', period: 'yesterday' },
    ]);
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([
      { kind: 'data-bar', color: '#70ad47' },
      { kind: 'color-scale', stops: ['#f8696b', '#ffeb84', '#63be7b'] },
      { kind: 'icon-set', icons: 'traffic3' },
      { kind: 'icon-set', icons: 'symbols3' },
      { kind: 'date-occurring', period: 'yesterday' },
      { kind: 'top-bottom', mode: 'top', n: 10, percent: false },
    ]);

  await page.locator('[data-ribbon-command="conditional"]').click();
  await page.locator('[data-cf-submenu="topBottom"]').hover();
  await page.locator('.app__submenu--cf-topBottom [data-cf-action="above-avg"]').click();
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([
      { kind: 'data-bar', color: '#70ad47' },
      { kind: 'color-scale', stops: ['#f8696b', '#ffeb84', '#63be7b'] },
      { kind: 'icon-set', icons: 'traffic3' },
      { kind: 'icon-set', icons: 'symbols3' },
      { kind: 'date-occurring', period: 'yesterday' },
      { kind: 'top-bottom', mode: 'top', n: 10, percent: false },
      { kind: 'average', mode: 'above' },
    ]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([
      { kind: 'data-bar', color: '#70ad47' },
      { kind: 'color-scale', stops: ['#f8696b', '#ffeb84', '#63be7b'] },
      { kind: 'icon-set', icons: 'traffic3' },
      { kind: 'icon-set', icons: 'symbols3' },
      { kind: 'date-occurring', period: 'yesterday' },
      { kind: 'top-bottom', mode: 'top', n: 10, percent: false },
    ]);
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([
      { kind: 'data-bar', color: '#70ad47' },
      { kind: 'color-scale', stops: ['#f8696b', '#ffeb84', '#63be7b'] },
      { kind: 'icon-set', icons: 'traffic3' },
      { kind: 'icon-set', icons: 'symbols3' },
      { kind: 'date-occurring', period: 'yesterday' },
      { kind: 'top-bottom', mode: 'top', n: 10, percent: false },
      { kind: 'average', mode: 'above' },
    ]);

  await page.locator('[data-ribbon-command="conditional"]').click();
  await page.locator('#menu-conditional > [data-cf-action="new-rule"]').click();
  const newRuleDialog = page.getByRole('dialog', { name: 'New Formatting Rule' });
  await expect(newRuleDialog).toBeVisible();
  await expect(newRuleDialog.getByRole('button', { name: 'OK' })).toBeVisible();
  await expect(newRuleDialog.getByRole('button', { name: 'Cancel' })).toBeVisible();
  await expect(newRuleDialog).not.toContainText('No rules defined yet');
  await newRuleDialog.getByRole('button', { name: 'OK' }).click();
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([
      { kind: 'data-bar', color: '#70ad47' },
      { kind: 'color-scale', stops: ['#f8696b', '#ffeb84', '#63be7b'] },
      { kind: 'icon-set', icons: 'traffic3' },
      { kind: 'icon-set', icons: 'symbols3' },
      { kind: 'date-occurring', period: 'yesterday' },
      { kind: 'top-bottom', mode: 'top', n: 10, percent: false },
      { kind: 'average', mode: 'above' },
      { kind: 'cell-value' },
    ]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([
      { kind: 'data-bar', color: '#70ad47' },
      { kind: 'color-scale', stops: ['#f8696b', '#ffeb84', '#63be7b'] },
      { kind: 'icon-set', icons: 'traffic3' },
      { kind: 'icon-set', icons: 'symbols3' },
      { kind: 'date-occurring', period: 'yesterday' },
      { kind: 'top-bottom', mode: 'top', n: 10, percent: false },
      { kind: 'average', mode: 'above' },
    ]);
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([
      { kind: 'data-bar', color: '#70ad47' },
      { kind: 'color-scale', stops: ['#f8696b', '#ffeb84', '#63be7b'] },
      { kind: 'icon-set', icons: 'traffic3' },
      { kind: 'icon-set', icons: 'symbols3' },
      { kind: 'date-occurring', period: 'yesterday' },
      { kind: 'top-bottom', mode: 'top', n: 10, percent: false },
      { kind: 'average', mode: 'above' },
      { kind: 'cell-value' },
    ]);

  await page.locator('[data-ribbon-command="conditional"]').click();
  await page.locator('[data-cf-submenu="dataBar"]').hover();
  await page.locator('.app__submenu--cf-dataBar [data-cf-action="new-rule"]').click();
  await expect(newRuleDialog).toBeVisible();
  await expect(newRuleDialog.locator('.fc-conddlg__sub:not([hidden])')).toContainText('Bar color');
  await newRuleDialog.getByRole('button', { name: 'Cancel' }).click();

  await page.locator('[data-ribbon-command="conditional"]').click();
  await page.locator('#menu-conditional [data-cf-action="manage"]').click();
  const rulesDialog = page.getByRole('dialog', { name: 'Conditional Formatting — Manage Rules' });
  await expect(rulesDialog).toBeVisible();
  await expect(rulesDialog).toContainText('Data Bar');
  await expect(rulesDialog).toContainText('A1');
  await rulesDialog.getByRole('button', { name: 'Move Down' }).first().click();
  await expect
    .poll(async () => (await readConditionalRuleSummaries(page))[0])
    .toMatchObject({
      kind: 'color-scale',
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(async () => (await readConditionalRuleSummaries(page))[0])
    .toMatchObject({
      kind: 'data-bar',
      color: '#70ad47',
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(async () => (await readConditionalRuleSummaries(page))[0])
    .toMatchObject({
      kind: 'color-scale',
    });
  await rulesDialog.locator('.fc-cfrulesdlg__move-up:not(:disabled)').first().click();
  await expect
    .poll(async () => (await readConditionalRuleSummaries(page))[0])
    .toMatchObject({
      kind: 'data-bar',
      color: '#70ad47',
    });
  await rulesDialog.getByRole('button', { name: 'New Rule...' }).click();
  await expect(rulesDialog).toBeHidden();
  await expect(newRuleDialog).toBeVisible();
  await newRuleDialog.getByRole('button', { name: 'Cancel' }).click();

  await page.locator('[data-ribbon-command="conditional"]').click();
  await page.locator('#menu-conditional [data-cf-action="manage"]').click();
  await expect(rulesDialog).toBeVisible();
  await rulesDialog.getByRole('button', { name: 'Duplicate Rule' }).first().click();
  await expect.poll(() => readConditionalRuleSummaries(page)).toHaveLength(9);
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readConditionalRuleSummaries(page)).toHaveLength(8);
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readConditionalRuleSummaries(page)).toHaveLength(9);
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readConditionalRuleSummaries(page)).toHaveLength(8);
  await rulesDialog.getByRole('button', { name: 'Remove' }).first().click();
  await expect.poll(() => readConditionalRuleSummaries(page)).toHaveLength(7);
  await closeDialog(page);
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readConditionalRuleSummaries(page)).toHaveLength(8);
  await expect
    .poll(async () => (await readConditionalRuleSummaries(page))[0])
    .toMatchObject({
      kind: 'data-bar',
      color: '#70ad47',
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readConditionalRuleSummaries(page)).toHaveLength(7);
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readConditionalRuleSummaries(page)).toHaveLength(8);

  await page.locator('[data-ribbon-command="conditional"]').click();
  await page.locator('[data-cf-submenu="clear"]').hover();
  await page.locator('.app__submenu--cf-clear [data-cf-action="clear-selection"]').click();
  await expect.poll(() => readConditionalRuleSummaries(page)).toEqual([]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readConditionalRuleSummaries(page)).toHaveLength(8);
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readConditionalRuleSummaries(page)).toEqual([]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readConditionalRuleSummaries(page)).toHaveLength(8);

  await page.locator('[data-ribbon-command="conditional"]').click();
  await page.locator('[data-cf-submenu="clear"]').hover();
  await page.locator('.app__submenu--cf-clear [data-cf-action="clear-sheet"]').click();
  await expect.poll(() => readConditionalRuleSummaries(page)).toEqual([]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readConditionalRuleSummaries(page)).toHaveLength(8);
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readConditionalRuleSummaries(page)).toEqual([]);
});

test('R02d-highlight: Conditional-format highlight prompts create Excel-style predicate rules', async ({
  page,
}) => {
  await mount(page, '/?locale=en&fixture=empty');

  const fillNumberPrompt = async (value: string): Promise<void> => {
    const dialog = page.locator('.app__dlg');
    await expect(dialog).toBeVisible();
    await dialog.locator('input[type="number"]').fill(value);
    await dialog.getByRole('button', { name: 'OK' }).click();
  };
  const fillTextPrompt = async (value: string): Promise<void> => {
    const dialog = page.locator('.app__dlg');
    await expect(dialog).toBeVisible();
    await dialog.locator('input[type="text"]').fill(value);
    await dialog.getByRole('button', { name: 'OK' }).click();
  };
  const openConditionalSubmenu = async (submenu: string): Promise<void> => {
    await page.locator('[data-ribbon-command="conditional"]').click();
    await page.locator(`[data-cf-submenu="${submenu}"]`).hover();
  };

  await page.getByRole('tab', { name: 'Home', exact: true }).click();
  await selectRangeAndSetValues(page, { r0: 0, c0: 0, r1: 3, c1: 0 }, [
    { row: 0, col: 0, value: 4 },
    { row: 1, col: 0, value: 8 },
    { row: 2, col: 0, value: 12 },
    { row: 3, col: 0, value: 16 },
  ]);

  await openConditionalSubmenu('highlight');
  await page.locator('.app__submenu--cf-highlight [data-cf-action="cell-gt"]').click();
  await fillNumberPrompt('10');
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([{ kind: 'cell-value', op: '>', a: 10 }]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readConditionalRuleSummaries(page)).toEqual([]);
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([{ kind: 'cell-value', op: '>', a: 10 }]);

  await openConditionalSubmenu('highlight');
  await page.locator('.app__submenu--cf-highlight [data-cf-action="cell-between"]').click();
  await fillNumberPrompt('3');
  await fillNumberPrompt('7');
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([
      { kind: 'cell-value', op: '>', a: 10 },
      { kind: 'cell-value', op: 'between', a: 3, b: 7 },
    ]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([{ kind: 'cell-value', op: '>', a: 10 }]);
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([
      { kind: 'cell-value', op: '>', a: 10 },
      { kind: 'cell-value', op: 'between', a: 3, b: 7 },
    ]);

  await openConditionalSubmenu('highlight');
  await page.locator('.app__submenu--cf-highlight [data-cf-action="text-contains"]').click();
  await fillTextPrompt('ink');
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([
      { kind: 'cell-value', op: '>', a: 10 },
      { kind: 'cell-value', op: 'between', a: 3, b: 7 },
      { kind: 'text-contains', text: 'ink' },
    ]);

  await openConditionalSubmenu('topBottom');
  await page.locator('.app__submenu--cf-topBottom [data-cf-action="bottom10-percent"]').click();
  await fillNumberPrompt('25');
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([
      { kind: 'cell-value', op: '>', a: 10 },
      { kind: 'cell-value', op: 'between', a: 3, b: 7 },
      { kind: 'text-contains', text: 'ink' },
      { kind: 'top-bottom', mode: 'bottom', n: 25, percent: true },
    ]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([
      { kind: 'cell-value', op: '>', a: 10 },
      { kind: 'cell-value', op: 'between', a: 3, b: 7 },
      { kind: 'text-contains', text: 'ink' },
    ]);
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([
      { kind: 'cell-value', op: '>', a: 10 },
      { kind: 'cell-value', op: 'between', a: 3, b: 7 },
      { kind: 'text-contains', text: 'ink' },
      { kind: 'top-bottom', mode: 'bottom', n: 25, percent: true },
    ]);
});

test('R02d-new-rule: Conditional-format New Rule creates formula-based rules', async ({ page }) => {
  await mount(page, '/?locale=en&fixture=empty');

  await page.getByRole('tab', { name: 'Home', exact: true }).click();
  await selectRangeAndSetValues(page, { r0: 0, c0: 0, r1: 3, c1: 0 }, [
    { row: 0, col: 0, value: 4 },
    { row: 1, col: 0, value: 8 },
    { row: 2, col: 0, value: 12 },
    { row: 3, col: 0, value: 16 },
  ]);

  await page.locator('[data-ribbon-command="conditional"]').click();
  await page.locator('#menu-conditional > [data-cf-action="new-rule"]').click();
  const dialog = page.getByRole('dialog', { name: 'New Formatting Rule' });
  await expect(dialog).toBeVisible();

  await dialog.locator('.fc-conddlg__form select').nth(1).selectOption('formula');
  const formulaInput = dialog.locator('.fc-conddlg__sub:not([hidden]) input[type="text"]');
  await expect(formulaInput).toHaveCount(1);
  await formulaInput.fill('=A1>10');
  await dialog.getByRole('button', { name: 'OK' }).click();

  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([{ kind: 'formula', formula: '=A1>10' }]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readConditionalRuleSummaries(page)).toEqual([]);
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readConditionalRuleSummaries(page))
    .toMatchObject([{ kind: 'formula', formula: '=A1>10' }]);
});

test('R02e: Home cell insert, delete, and format menus mutate sheet state', async ({ page }) => {
  await mount(page, '/?locale=en');

  await page.getByRole('tab', { name: 'Home', exact: true }).click();

  await selectCellAndSetText(page, 0, 0, 'row-anchor');
  await page.locator('[data-ribbon-command="insertRows"]').click();
  await page.locator('#menu-insert-cells [data-cell-insert="rows"]').click();
  await expect.poll(() => readCellText(page, 1, 0)).toBe('row-anchor');

  await page.locator('[data-ribbon-command="deleteRows"]').click();
  await page.locator('#menu-delete-cells [data-cell-delete="rows"]').click();
  await expect.poll(() => readCellText(page, 0, 0)).toBe('row-anchor');

  await selectCellAndSetText(page, 0, 0, 'col-anchor');
  await page.locator('[data-ribbon-command="insertRows"]').click();
  await page.locator('#menu-insert-cells [data-cell-insert="cols"]').click();
  await expect.poll(() => readCellText(page, 0, 1)).toBe('col-anchor');

  await page.locator('[data-ribbon-command="deleteRows"]').click();
  await page.locator('#menu-delete-cells [data-cell-delete="cols"]').click();
  await expect.poll(() => readCellText(page, 0, 0)).toBe('col-anchor');

  await selectCellAndSetText(page, 0, 0, 'cell-shift-down');
  await page.locator('[data-ribbon-command="insertRows"]').click();
  await page.locator('#menu-insert-cells [data-cell-insert="cells"]').click();
  await expect(page.getByRole('dialog', { name: 'Insert Cells' })).toBeVisible();
  await page.getByRole('button', { name: 'OK', exact: true }).click();
  await expect.poll(() => readCellText(page, 1, 0)).toBe('cell-shift-down');
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 0, 0)).toBe('cell-shift-down');
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 1, 0)).toBe('cell-shift-down');

  await selectCellAndSetText(page, 0, 0, 'delete-up-anchor');
  await page.locator('[data-ribbon-command="deleteRows"]').click();
  await page.locator('#menu-delete-cells [data-cell-delete="cells"]').click();
  await expect(page.getByRole('dialog', { name: 'Delete Cells' })).toBeVisible();
  await page.getByRole('button', { name: 'OK', exact: true }).click();
  await expect.poll(() => readCellText(page, 0, 0)).toBe('cell-shift-down');
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 0, 0)).toBe('delete-up-anchor');
  await expect.poll(() => readCellText(page, 1, 0)).toBe('cell-shift-down');
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 0, 0)).toBe('cell-shift-down');

  await selectCellAndSetText(page, 0, 0, 'cell-shift-right');
  await page.locator('[data-ribbon-command="insertRows"]').click();
  await page.locator('#menu-insert-cells [data-cell-insert="cells"]').click();
  await page.getByRole('radio', { name: 'Shift cells right' }).check();
  await page.getByRole('button', { name: 'OK', exact: true }).click();
  await expect.poll(() => readCellText(page, 0, 1)).toBe('cell-shift-right');
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 0, 0)).toBe('cell-shift-right');
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 0, 1)).toBe('cell-shift-right');

  await selectCellAndSetText(page, 0, 0, 'delete-left-anchor');
  await page.locator('[data-ribbon-command="deleteRows"]').click();
  await page.locator('#menu-delete-cells [data-cell-delete="cells"]').click();
  await page.getByRole('radio', { name: 'Shift cells left' }).check();
  await page.getByRole('button', { name: 'OK', exact: true }).click();
  await expect.poll(() => readCellText(page, 0, 0)).toBe('cell-shift-right');
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 0, 0)).toBe('delete-left-anchor');
  await expect.poll(() => readCellText(page, 0, 1)).toBe('cell-shift-right');
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 0, 0)).toBe('cell-shift-right');

  await selectCellAndSetText(page, 2, 2, 'visibility');
  await page.locator('[data-ribbon-command="formatCellsHome"]').click();
  await page.locator('#menu-format-cells [data-cell-format="hide-rows"]').click();
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenRows: [2] });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenRows: [] });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenRows: [2] });

  await page.locator('[data-ribbon-command="formatCellsHome"]').click();
  await page.locator('#menu-format-cells [data-cell-format="show-rows"]').click();
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenRows: [] });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenRows: [2] });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenRows: [] });

  await page.locator('[data-ribbon-command="formatCellsHome"]').click();
  await page.locator('#menu-format-cells [data-cell-format="hide-cols"]').click();
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenCols: [2] });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenCols: [] });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenCols: [2] });

  await page.locator('[data-ribbon-command="formatCellsHome"]').click();
  await page.locator('#menu-format-cells [data-cell-format="show-cols"]').click();
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenCols: [] });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenCols: [2] });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenCols: [] });

  await selectCellAndSetText(page, 5, 5, 'manual-size');
  await page.locator('[data-ribbon-command="formatCellsHome"]').click();
  await page.locator('#menu-format-cells [data-cell-format="row-height"]').click();
  await expect(page.getByRole('dialog', { name: 'Row Height' })).toBeVisible();
  await page.getByRole('spinbutton', { name: 'Height (px)' }).fill('44');
  await page.getByRole('button', { name: 'OK', exact: true }).click();
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({
      rowHeights: expect.arrayContaining([[5, 44]]),
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ rowHeights: [] });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({
      rowHeights: expect.arrayContaining([[5, 44]]),
    });

  await page.locator('[data-ribbon-command="formatCellsHome"]').click();
  await page.locator('#menu-format-cells [data-cell-format="col-width"]').click();
  await expect(page.getByRole('dialog', { name: 'Column Width' })).toBeVisible();
  await page.getByRole('spinbutton', { name: 'Width (px)' }).fill('123');
  await page.getByRole('button', { name: 'OK', exact: true }).click();
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({
      colWidths: expect.arrayContaining([[5, 123]]),
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ colWidths: [] });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({
      colWidths: expect.arrayContaining([[5, 123]]),
    });

  await selectCellAndSetText(page, 3, 3, 'one\ntwo\nthree');
  await page.locator('[data-ribbon-command="formatCellsHome"]').click();
  await page.locator('#menu-format-cells [data-cell-format="row-autofit"]').click();
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({
      rowHeights: expect.arrayContaining([[3, expect.any(Number)]]),
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(async () => (await readLayoutSummary(page)).rowHeights.some(([row]) => row === 3))
    .toBe(false);
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({
      rowHeights: expect.arrayContaining([[3, expect.any(Number)]]),
    });

  await selectCellAndSetText(page, 4, 4, 'a long value that should drive column autofit');
  await page.locator('[data-ribbon-command="formatCellsHome"]').click();
  await page.locator('#menu-format-cells [data-cell-format="col-autofit"]').click();
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({
      colWidths: expect.arrayContaining([[4, expect.any(Number)]]),
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(async () => (await readLayoutSummary(page)).colWidths.some(([col]) => col === 4))
    .toBe(false);
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({
      colWidths: expect.arrayContaining([[4, expect.any(Number)]]),
    });

  await selectRangeAndSetValues(page, { r0: 8, c0: 8, r1: 8, c1: 8 }, []);
  await page.locator('[data-ribbon-command="formatCellsHome"]').click();
  await page.locator('#menu-format-cells [data-cell-format="unlock-cell"]').click();
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ locked: false });
  await expect(page.locator('#status-metric')).toHaveText('Unlocked selected cell(s)');
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(async () => (await readActiveCellFormat(page))?.locked).toBeUndefined();
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ locked: false });
  await page.locator('[data-ribbon-command="formatCellsHome"]').click();
  await page.locator('#menu-format-cells [data-cell-format="lock-cell"]').click();
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ locked: true });
  await expect(page.locator('#status-metric')).toHaveText('Locked selected cell(s)');
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ locked: false });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ locked: true });

  await page.locator('[data-ribbon-command="formatCellsHome"]').click();
  await page.locator('#menu-format-cells [data-cell-format="protect-sheet"]').click();
  const protectDialog = page.getByRole('dialog', { name: 'Protect Sheet' });
  await expect(protectDialog).toBeVisible();
  await protectDialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect.poll(() => readProtectionSummary(page)).toMatchObject({ protected: true });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readProtectionSummary(page)).toMatchObject({ protected: false });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readProtectionSummary(page)).toMatchObject({ protected: true });

  await page.locator('[data-ribbon-command="formatCellsHome"]').click();
  await page.locator('#menu-format-cells [data-cell-format="protect-sheet"]').click();
  await expect.poll(() => readProtectionSummary(page)).toMatchObject({ protected: false });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readProtectionSummary(page)).toMatchObject({ protected: true });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readProtectionSummary(page)).toMatchObject({ protected: false });
});

test('R02f: Home editing menus apply fill, clear, autosum, sort, and find actions', async ({
  page,
}) => {
  await mount(page, '/?locale=en');

  await page.getByRole('tab', { name: 'Home', exact: true }).click();

  await selectRangeAndSetValues(page, { r0: 30, c0: 0, r1: 32, c1: 0 }, [
    { row: 30, col: 0, value: 'fill-source' },
  ]);
  await page.locator('[data-ribbon-command="fillHome"]').click();
  await page.locator('#menu-fill [data-fill="down"]').click();
  await expect.poll(() => readCellText(page, 31, 0)).toBe('fill-source');
  await expect.poll(() => readCellText(page, 32, 0)).toBe('fill-source');
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 31, 0)).toMatchObject({ kind: 'blank' });
  await expect.poll(() => readCellSummary(page, 32, 0)).toMatchObject({ kind: 'blank' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 31, 0)).toBe('fill-source');
  await expect.poll(() => readCellText(page, 32, 0)).toBe('fill-source');

  await selectRangeAndSetValues(page, { r0: 42, c0: 14, r1: 44, c1: 14 }, [
    { row: 42, col: 14, value: 45296 },
  ]);
  await page.locator('[data-ribbon-command="fillHome"]').click();
  await expect(page.locator('#menu-fill')).toContainText('Weekdays');
  await page.locator('#menu-fill [data-fill="weekdays"]').click();
  await expect
    .poll(() => readCellSummary(page, 43, 14))
    .toMatchObject({
      kind: 'number',
      value: 45299,
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 43, 14)).toMatchObject({ kind: 'blank' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCellSummary(page, 43, 14))
    .toMatchObject({
      kind: 'number',
      value: 45299,
    });

  await selectRangeAndSetValues(page, { r0: 34, c0: 14, r1: 34, c1: 16 }, [
    { row: 34, col: 14, value: 'right-source' },
  ]);
  await page.locator('[data-ribbon-command="fillHome"]').click();
  await page.locator('#menu-fill [data-fill="right"]').click();
  await expect.poll(() => readCellText(page, 34, 15)).toBe('right-source');
  await expect.poll(() => readCellText(page, 34, 16)).toBe('right-source');
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 34, 15)).toMatchObject({ kind: 'blank' });
  await expect.poll(() => readCellSummary(page, 34, 16)).toMatchObject({ kind: 'blank' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 34, 15)).toBe('right-source');
  await expect.poll(() => readCellText(page, 34, 16)).toBe('right-source');

  await selectRangeAndSetValues(page, { r0: 35, c0: 14, r1: 37, c1: 14 }, [
    { row: 37, col: 14, value: 'up-source' },
  ]);
  await page.locator('[data-ribbon-command="fillHome"]').click();
  await page.locator('#menu-fill [data-fill="up"]').click();
  await expect.poll(() => readCellText(page, 35, 14)).toBe('up-source');
  await expect.poll(() => readCellText(page, 36, 14)).toBe('up-source');
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 35, 14)).toMatchObject({ kind: 'blank' });
  await expect.poll(() => readCellSummary(page, 36, 14)).toMatchObject({ kind: 'blank' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 35, 14)).toBe('up-source');
  await expect.poll(() => readCellText(page, 36, 14)).toBe('up-source');

  await selectRangeAndSetValues(page, { r0: 38, c0: 14, r1: 38, c1: 16 }, [
    { row: 38, col: 16, value: 'left-source' },
  ]);
  await page.locator('[data-ribbon-command="fillHome"]').click();
  await page.locator('#menu-fill [data-fill="left"]').click();
  await expect.poll(() => readCellText(page, 38, 14)).toBe('left-source');
  await expect.poll(() => readCellText(page, 38, 15)).toBe('left-source');
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 38, 14)).toMatchObject({ kind: 'blank' });
  await expect.poll(() => readCellSummary(page, 38, 15)).toMatchObject({ kind: 'blank' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 38, 14)).toBe('left-source');
  await expect.poll(() => readCellText(page, 38, 15)).toBe('left-source');

  await selectRangeAndSetValues(page, { r0: 40, c0: 14, r1: 40, c1: 15 }, [
    { row: 40, col: 14, value: 45322 },
  ]);
  await page.locator('[data-ribbon-command="fillHome"]').click();
  await page.locator('#menu-fill [data-fill="months"]').click();
  await expect
    .poll(() => readCellSummary(page, 40, 15))
    .toMatchObject({
      kind: 'number',
      value: 45351,
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 40, 15)).toMatchObject({ kind: 'blank' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCellSummary(page, 40, 15))
    .toMatchObject({
      kind: 'number',
      value: 45351,
    });

  await selectRangeAndSetValues(page, { r0: 41, c0: 14, r1: 41, c1: 15 }, [
    { row: 41, col: 14, value: 45351 },
  ]);
  await page.locator('[data-ribbon-command="fillHome"]').click();
  await page.locator('#menu-fill [data-fill="years"]').click();
  await expect
    .poll(() => readCellSummary(page, 41, 15))
    .toMatchObject({
      kind: 'number',
      value: 45716,
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 41, 15)).toMatchObject({ kind: 'blank' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCellSummary(page, 41, 15))
    .toMatchObject({
      kind: 'number',
      value: 45716,
    });

  await selectRangeAndSetValues(page, { r0: 45, c0: 14, r1: 45, c1: 16 }, [
    { row: 45, col: 14, value: 'series-copy' },
  ]);
  await page.locator('[data-ribbon-command="fillHome"]').click();
  await page.locator('#menu-fill [data-fill="series"]').click();
  const seriesDialog = page.getByRole('dialog', { name: 'Series' });
  await expect(seriesDialog).toBeVisible();
  await seriesDialog.getByRole('radio', { name: 'Rows' }).check();
  await seriesDialog.getByRole('radio', { name: 'Copy' }).check();
  await seriesDialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect.poll(() => readCellText(page, 45, 15)).toBe('series-copy');
  await expect.poll(() => readCellText(page, 45, 16)).toBe('series-copy');
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 45, 15)).toMatchObject({ kind: 'blank' });
  await expect.poll(() => readCellSummary(page, 45, 16)).toMatchObject({ kind: 'blank' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 45, 15)).toBe('series-copy');
  await expect.poll(() => readCellText(page, 45, 16)).toBe('series-copy');

  await selectRangeAndSetValues(page, { r0: 4, c0: 0, r1: 4, c1: 0 }, [
    { row: 4, col: 0, value: 'clear-me' },
  ]);
  await page.locator('[data-ribbon-command="clearFormat"]').click();
  await expect(page.locator('#menu-clear')).toContainText('Remove Hyperlinks');
  await page.locator('#menu-clear [data-clear="contents"]').click();
  await expect.poll(() => readCellSummary(page, 4, 0)).toMatchObject({ kind: 'blank' });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 4, 0)).toBe('clear-me');
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 4, 0)).toMatchObject({ kind: 'blank' });

  await selectRangeAndSetValues(page, { r0: 4, c0: 4, r1: 4, c1: 4 }, [
    { row: 4, col: 4, value: 'metadata' },
  ]);
  await patchActiveCellFormat(page, {
    fill: '#fff2cc',
    color: '#c00000',
    underline: true,
    hyperlink: 'https://example.com/metadata',
    locked: false,
  });
  await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            getState: () => {
              selection: { active: CellAddr };
              format: {
                formats: Map<string, ActiveCellFormat & { validation?: ValidationSummary }>;
              };
            };
            setState: (
              updater: (state: {
                selection: { active: CellAddr };
                format: {
                  formats: Map<string, ActiveCellFormat & { validation?: ValidationSummary }>;
                };
              }) => unknown,
            ) => void;
          };
        }
      | undefined;
    const state = inst?.store.getState();
    const active = state?.selection.active;
    if (!inst || !state || !active) return;
    const key = `${active.sheet}:${active.row}:${active.col}`;
    inst.store.setState((raw) => {
      const formats = new Map(raw.format.formats);
      formats.set(key, {
        ...(formats.get(key) ?? {}),
        validation: { kind: 'list', source: ['Open', 'Closed'] },
      });
      return { ...raw, format: { ...raw.format, formats } };
    });
  });
  await page.locator('[data-ribbon-command="clearFormat"]').click();
  await page.locator('#menu-clear [data-clear="formats"]').click();
  await expect
    .poll(async () => {
      const format = await readActiveCellFormat(page);
      const validation = await readActiveValidation(page);
      return {
        fill: format?.fill,
        color: format?.color,
        underline: format?.underline,
        hyperlink: format?.hyperlink,
        locked: format?.locked,
        validation,
      };
    })
    .toEqual({
      fill: undefined,
      color: undefined,
      underline: undefined,
      hyperlink: 'https://example.com/metadata',
      locked: false,
      validation: { kind: 'list', source: ['Open', 'Closed'] },
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({
      fill: '#fff2cc',
      color: '#c00000',
      underline: true,
      hyperlink: 'https://example.com/metadata',
      locked: false,
    });
  await expect
    .poll(() => readActiveValidation(page))
    .toEqual({
      kind: 'list',
      source: ['Open', 'Closed'],
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(async () => (await readActiveCellFormat(page))?.fill).toBeUndefined();
  await expect
    .poll(async () => (await readActiveCellFormat(page))?.hyperlink)
    .toBe('https://example.com/metadata');
  await expect
    .poll(() => readActiveValidation(page))
    .toEqual({
      kind: 'list',
      source: ['Open', 'Closed'],
    });

  await selectCellAndSetText(page, 4, 5, 'clear-link-only');
  await patchActiveCellFormat(page, {
    hyperlink: 'https://example.com/clear-only',
    color: '#0563c1',
    underline: true,
  });
  await page.locator('[data-ribbon-command="clearFormat"]').click();
  await page.locator('#menu-clear [data-clear="hyperlinks"]').click();
  await expect
    .poll(async () => {
      const format = await readActiveCellFormat(page);
      return [format?.hyperlink, format?.color, format?.underline];
    })
    .toEqual([undefined, '#0563c1', true]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(async () => {
      const format = await readActiveCellFormat(page);
      return [format?.hyperlink, format?.color, format?.underline];
    })
    .toEqual(['https://example.com/clear-only', '#0563c1', true]);
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(async () => {
      const format = await readActiveCellFormat(page);
      return [format?.hyperlink, format?.color, format?.underline];
    })
    .toEqual([undefined, '#0563c1', true]);

  await selectCellAndSetText(page, 4, 1, 'link');
  await patchActiveCellFormat(page, {
    hyperlink: 'https://example.com',
    color: '#0563c1',
    underline: true,
  });
  await page.locator('[data-ribbon-command="clearFormat"]').click();
  await page.locator('#menu-clear [data-clear="remove-hyperlinks"]').click();
  await expect
    .poll(async () => {
      const format = await readActiveCellFormat(page);
      return [format?.hyperlink, format?.color, format?.underline];
    })
    .toEqual([undefined, undefined, undefined]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(async () => {
      const format = await readActiveCellFormat(page);
      return [format?.hyperlink, format?.color, format?.underline];
    })
    .toEqual(['https://example.com', '#0563c1', true]);
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(async () => {
      const format = await readActiveCellFormat(page);
      return [format?.hyperlink, format?.color, format?.underline];
    })
    .toEqual([undefined, undefined, undefined]);

  await selectRangeAndSetValues(page, { r0: 4, c0: 6, r1: 4, c1: 6 }, [
    { row: 4, col: 6, value: 'commented' },
  ]);
  await patchActiveCellFormat(page, { fill: '#e2f0d9', color: '#375623' });
  await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            getState: () => {
              selection: { active: CellAddr };
              format: { formats: Map<string, ActiveCellFormat> };
            };
            setState: (
              updater: (state: {
                selection: { active: CellAddr };
                format: { formats: Map<string, ActiveCellFormat> };
              }) => unknown,
            ) => void;
          };
          workbook: {
            capabilities?: { comments?: boolean };
            setCommentEntry?: (
              sheet: number,
              row: number,
              col: number,
              author: string,
              text: string,
            ) => boolean;
          };
        }
      | undefined;
    const state = inst?.store.getState();
    const active = state?.selection.active;
    if (!inst || !state || !active) return;
    const key = `${active.sheet}:${active.row}:${active.col}`;
    inst.store.setState((raw) => {
      const formats = new Map(raw.format.formats);
      formats.set(key, { ...(formats.get(key) ?? {}), comment: 'keep the format' });
      return { ...raw, format: { ...raw.format, formats } };
    });
    inst.workbook.setCommentEntry?.(active.sheet, active.row, active.col, '', 'keep the format');
  });
  await expect
    .poll(() => readCommentSummaries(page))
    .toEqual([{ addr: { sheet: 0, row: 4, col: 6 }, text: 'keep the format' }]);
  await page.locator('[data-ribbon-command="clearFormat"]').click();
  await page.locator('#menu-clear [data-clear="comments"]').click();
  await expect.poll(() => readCommentSummaries(page)).toEqual([]);
  await expect.poll(() => readCellText(page, 4, 6)).toBe('commented');
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({
      fill: '#e2f0d9',
      color: '#375623',
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCommentSummaries(page))
    .toEqual([{ addr: { sheet: 0, row: 4, col: 6 }, text: 'keep the format' }]);
  await expect
    .poll(() =>
      page.evaluate(() => {
        const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
          | {
              workbook: {
                capabilities?: { comments?: boolean };
                getComment?: (sheet: number, row: number, col: number) => { text: string } | null;
              };
            }
          | undefined;
        if (!inst?.workbook.capabilities?.comments) return 'unsupported';
        return inst.workbook.getComment?.(0, 4, 6)?.text ?? null;
      }),
    )
    .toBe('keep the format');
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCommentSummaries(page)).toEqual([]);

  await selectRangeAndSetValues(page, { r0: 4, c0: 2, r1: 4, c1: 2 }, [
    { row: 4, col: 2, value: 12 },
  ]);
  await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            setState: (updater: (state: { conditional: { rules: unknown[] } }) => unknown) => void;
          };
        }
      | undefined;
    inst?.store.setState((state) => ({
      ...state,
      conditional: {
        rules: [
          ...state.conditional.rules,
          {
            kind: 'cell-value',
            range: { sheet: 0, r0: 4, c0: 2, r1: 4, c1: 2 },
            op: '>',
            a: 10,
            apply: { fill: '#fff2cc' },
          },
        ],
      },
    }));
  });
  await expect.poll(() => readConditionalRuleSummaries(page)).toHaveLength(1);
  await page.locator('[data-ribbon-command="clearFormat"]').click();
  await page.locator('#menu-clear [data-clear="conditional"]').click();
  await expect.poll(() => readConditionalRuleSummaries(page)).toEqual([]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readConditionalRuleSummaries(page)).toHaveLength(1);
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readConditionalRuleSummaries(page)).toEqual([]);

  await selectRangeAndSetValues(page, { r0: 4, c0: 3, r1: 4, c1: 3 }, [
    { row: 4, col: 3, value: 'clear-all' },
  ]);
  await patchActiveCellFormat(page, { fill: '#fff2cc', color: '#c00000' });
  await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            setState: (updater: (state: { conditional: { rules: unknown[] } }) => unknown) => void;
          };
        }
      | undefined;
    inst?.store.setState((state) => ({
      ...state,
      conditional: {
        rules: [
          ...state.conditional.rules,
          {
            kind: 'cell-value',
            range: { sheet: 0, r0: 4, c0: 3, r1: 4, c1: 3 },
            op: '=',
            a: 'clear-all',
            apply: { fill: '#f4cccc' },
          },
        ],
      },
    }));
  });
  await expect.poll(() => readCellText(page, 4, 3)).toBe('clear-all');
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({
      fill: '#fff2cc',
      color: '#c00000',
    });
  await expect.poll(() => readConditionalRuleSummaries(page)).toHaveLength(1);
  await page.locator('[data-ribbon-command="clearFormat"]').click();
  await page.locator('#menu-clear [data-clear="all"]').click();
  await expect.poll(() => readCellSummary(page, 4, 3)).toMatchObject({ kind: 'blank' });
  await expect.poll(async () => (await readActiveCellFormat(page))?.fill).toBeUndefined();
  await expect.poll(() => readConditionalRuleSummaries(page)).toEqual([]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 4, 3)).toBe('clear-all');
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({
      fill: '#fff2cc',
      color: '#c00000',
    });
  await expect.poll(() => readConditionalRuleSummaries(page)).toHaveLength(1);
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 4, 3)).toMatchObject({ kind: 'blank' });
  await expect.poll(async () => (await readActiveCellFormat(page))?.fill).toBeUndefined();
  await expect.poll(() => readConditionalRuleSummaries(page)).toEqual([]);

  await selectRangeAndSetValues(page, { r0: 20, c0: 10, r1: 22, c1: 10 }, [
    { row: 20, col: 10, value: 2 },
    { row: 21, col: 10, value: 3 },
  ]);
  await page.locator('[data-ribbon-command="autosum"]').click();
  await expect
    .poll(() => readCellSummary(page, 22, 10))
    .toMatchObject({
      kind: 'number',
      value: 5,
    });
  await expect
    .poll(() => readCellSummary(page, 22, 10))
    .toMatchObject({
      formula: '=SUM(K21:K22)',
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 22, 10)).toMatchObject({ kind: 'blank' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCellSummary(page, 22, 10))
    .toMatchObject({ kind: 'number', value: 5, formula: '=SUM(K21:K22)' });

  await selectRangeAndSetValues(page, { r0: 27, c0: 10, r1: 29, c1: 11 }, [
    { row: 27, col: 10, value: 4 },
    { row: 28, col: 10, value: 6 },
    { row: 27, col: 11, value: 8 },
    { row: 28, col: 11, value: 12 },
  ]);
  await page.locator('[data-ribbon-command="autosum"]').click();
  await expect
    .poll(() => readCellSummary(page, 29, 10))
    .toMatchObject({ kind: 'number', value: 10, formula: '=SUM(K28:K29)' });
  await expect
    .poll(() => readCellSummary(page, 29, 11))
    .toMatchObject({ kind: 'number', value: 20, formula: '=SUM(L28:L29)' });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 29, 10)).toMatchObject({ kind: 'blank' });
  await expect.poll(() => readCellSummary(page, 29, 11)).toMatchObject({ kind: 'blank' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCellSummary(page, 29, 10))
    .toMatchObject({ kind: 'number', value: 10, formula: '=SUM(K28:K29)' });
  await expect
    .poll(() => readCellSummary(page, 29, 11))
    .toMatchObject({ kind: 'number', value: 20, formula: '=SUM(L28:L29)' });

  await selectRangeAndSetValues(page, { r0: 20, c0: 11, r1: 22, c1: 11 }, [
    { row: 20, col: 11, value: 2 },
    { row: 21, col: 11, value: 3 },
  ]);
  await page.locator('[data-ribbon-command="autosum"] .demo__rb-split-chevron').click();
  await expect(page.locator('#menu-autosum-home')).toContainText('More Functions...');
  await page.locator('#menu-autosum-home [data-autosum-fn="COUNT"]').click();
  await expect
    .poll(() => readCellSummary(page, 22, 11))
    .toMatchObject({
      kind: 'number',
      value: 2,
      formula: '=COUNT(L21:L22)',
    });

  await selectRangeAndSetValues(page, { r0: 20, c0: 12, r1: 22, c1: 12 }, [
    { row: 20, col: 12, value: 10 },
    { row: 21, col: 12, value: 20 },
  ]);
  await page.locator('[data-ribbon-command="autosum"] .demo__rb-split-chevron').click();
  await page.locator('#menu-autosum-home [data-autosum-fn="AVERAGE"]').click();
  await expect
    .poll(() => readCellSummary(page, 22, 12))
    .toMatchObject({ kind: 'number', value: 15, formula: '=AVERAGE(M21:M22)' });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 22, 12)).toMatchObject({ kind: 'blank' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCellSummary(page, 22, 12))
    .toMatchObject({ kind: 'number', value: 15, formula: '=AVERAGE(M21:M22)' });

  await selectRangeAndSetValues(page, { r0: 24, c0: 12, r1: 26, c1: 12 }, [
    { row: 24, col: 12, value: 7 },
    { row: 25, col: 12, value: 11 },
  ]);
  await page.locator('[data-ribbon-command="autosum"] .demo__rb-split-chevron').click();
  await page.locator('#menu-autosum-home [data-autosum-fn="MAX"]').click();
  await expect
    .poll(() => readCellSummary(page, 26, 12))
    .toMatchObject({ kind: 'number', value: 11, formula: '=MAX(M25:M26)' });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 26, 12)).toMatchObject({ kind: 'blank' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCellSummary(page, 26, 12))
    .toMatchObject({ kind: 'number', value: 11, formula: '=MAX(M25:M26)' });

  await selectRangeAndSetValues(page, { r0: 28, c0: 12, r1: 30, c1: 12 }, [
    { row: 28, col: 12, value: 7 },
    { row: 29, col: 12, value: 11 },
  ]);
  await page.locator('[data-ribbon-command="autosum"] .demo__rb-split-chevron').click();
  await page.locator('#menu-autosum-home [data-autosum-fn="MIN"]').click();
  await expect
    .poll(() => readCellSummary(page, 30, 12))
    .toMatchObject({ kind: 'number', value: 7, formula: '=MIN(M29:M30)' });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 30, 12)).toMatchObject({ kind: 'blank' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCellSummary(page, 30, 12))
    .toMatchObject({ kind: 'number', value: 7, formula: '=MIN(M29:M30)' });

  await page.locator('[data-ribbon-command="autosum"] .demo__rb-split-chevron').click();
  await page.locator('#menu-autosum-home [data-autosum-fn="MORE"]').click();
  await expect(page.getByRole('dialog', { name: 'Function Arguments' })).toBeVisible();
  await closeDialog(page);

  await selectRangeAndSetValues(page, { r0: 24, c0: 10, r1: 26, c1: 10 }, [
    { row: 24, col: 10, value: 'c' },
    { row: 25, col: 10, value: 'a' },
    { row: 26, col: 10, value: 'b' },
  ]);
  await page.locator('[data-ribbon-command="sortFilterHome"]').click();
  await page.locator('#menu-sort-home [data-sort="asc"]').click();
  await expect.poll(() => readCellText(page, 24, 10)).toBe('a');
  await expect.poll(() => readCellText(page, 25, 10)).toBe('b');
  await expect.poll(() => readCellText(page, 26, 10)).toBe('c');
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 24, 10)).toBe('c');
  await expect.poll(() => readCellText(page, 25, 10)).toBe('a');
  await expect.poll(() => readCellText(page, 26, 10)).toBe('b');
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 24, 10)).toBe('a');
  await expect.poll(() => readCellText(page, 25, 10)).toBe('b');
  await expect.poll(() => readCellText(page, 26, 10)).toBe('c');

  await selectRangeAndSetValues(page, { r0: 48, c0: 14, r1: 52, c1: 16 }, [
    { row: 48, col: 14, value: 'Group' },
    { row: 48, col: 15, value: 'Rank' },
    { row: 48, col: 16, value: 'Name' },
    { row: 49, col: 14, value: 'B' },
    { row: 49, col: 15, value: 1 },
    { row: 49, col: 16, value: 'b-low' },
    { row: 50, col: 14, value: 'A' },
    { row: 50, col: 15, value: 2 },
    { row: 50, col: 16, value: 'a-high' },
    { row: 51, col: 14, value: 'B' },
    { row: 51, col: 15, value: 2 },
    { row: 51, col: 16, value: 'b-high' },
    { row: 52, col: 14, value: 'A' },
    { row: 52, col: 15, value: 1 },
    { row: 52, col: 16, value: 'a-low' },
  ]);
  await page.locator('[data-ribbon-command="sortFilterHome"]').click();
  await page.locator('#menu-sort-home [data-sort="custom"]').click();
  const homeCustomSort = page.getByRole('dialog', { name: 'Custom Sort...' });
  await expect(homeCustomSort).toBeVisible();
  await homeCustomSort.getByLabel('Sort by').selectOption('14');
  await homeCustomSort.locator('select').nth(1).selectOption('15');
  await homeCustomSort.locator('select').nth(3).selectOption('desc');
  await expect(homeCustomSort.getByRole('checkbox', { name: 'My data has headers' })).toBeChecked();
  await homeCustomSort.getByRole('button', { name: 'OK' }).click();
  await expect.poll(() => readCellText(page, 49, 16)).toBe('a-high');
  await expect.poll(() => readCellText(page, 50, 16)).toBe('a-low');
  await expect.poll(() => readCellText(page, 51, 16)).toBe('b-high');
  await expect.poll(() => readCellText(page, 52, 16)).toBe('b-low');
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 49, 16)).toBe('b-low');
  await expect.poll(() => readCellText(page, 50, 16)).toBe('a-high');
  await expect.poll(() => readCellText(page, 51, 16)).toBe('b-high');
  await expect.poll(() => readCellText(page, 52, 16)).toBe('a-low');
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 49, 16)).toBe('a-high');
  await expect.poll(() => readCellText(page, 50, 16)).toBe('a-low');
  await expect.poll(() => readCellText(page, 51, 16)).toBe('b-high');
  await expect.poll(() => readCellText(page, 52, 16)).toBe('b-low');

  await selectRangeAndSetValues(page, { r0: 32, c0: 10, r1: 32, c1: 10 }, [
    { row: 31, col: 10, value: 'Category' },
    { row: 32, col: 10, value: 'paper' },
    { row: 33, col: 10, value: 'ink' },
    { row: 34, col: 10, value: 'paper' },
  ]);
  await page.locator('[data-ribbon-command="sortFilterHome"]').click();
  await expect(page.locator('#menu-sort-home')).toContainText("Filter by Selected Cell's Value");
  await page.locator('#menu-sort-home [data-sort="filter-by-value"]').click();
  await expect
    .poll(() => readFilterSummary(page))
    .toMatchObject({
      filterRange: { sheet: 0, r0: 31, c0: 10, r1: 34, c1: 10 },
      hiddenRows: [33],
    });
  await page.locator('[data-ribbon-command="sortFilterHome"]').click();
  await page.locator('#menu-sort-home [data-sort="filter-clear"]').click();
  await expect.poll(() => readFilterSummary(page)).toMatchObject({ hiddenRows: [] });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readFilterSummary(page))
    .toMatchObject({
      filterRange: { sheet: 0, r0: 31, c0: 10, r1: 34, c1: 10 },
      hiddenRows: [33],
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readFilterSummary(page)).toMatchObject({ hiddenRows: [] });

  await selectRangeAndSetValues(page, { r0: 32, c0: 10, r1: 32, c1: 10 }, []);
  await page.locator('[data-ribbon-command="sortFilterHome"]').click();
  await page.locator('#menu-sort-home [data-sort="filter-by-value"]').click();
  await expect
    .poll(() => readFilterSummary(page))
    .toMatchObject({
      filterRange: { sheet: 0, r0: 31, c0: 10, r1: 34, c1: 10 },
      hiddenRows: [33],
    });
  await selectRangeAndSetValues(page, { r0: 33, c0: 10, r1: 33, c1: 10 }, [
    { row: 33, col: 10, value: 'paper' },
  ]);
  await page.locator('[data-ribbon-command="sortFilterHome"]').click();
  await page.locator('#menu-sort-home [data-sort="filter-reapply"]').click();
  await expect
    .poll(() => readFilterSummary(page))
    .toMatchObject({
      filterRange: { sheet: 0, r0: 31, c0: 10, r1: 34, c1: 10 },
      filterCriteria: [
        {
          range: { sheet: 0, r0: 31, c0: 10, r1: 34, c1: 10 },
          byCol: 10,
          hiddenValues: ['ink'],
        },
      ],
      hiddenRows: [],
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readFilterSummary(page))
    .toMatchObject({
      filterRange: { sheet: 0, r0: 31, c0: 10, r1: 34, c1: 10 },
      hiddenRows: [33],
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readFilterSummary(page)).toMatchObject({ hiddenRows: [] });
  await page.locator('[data-ribbon-command="sortFilterHome"]').click();
  await page.locator('#menu-sort-home [data-sort="filter-clear"]').click();
  await expect
    .poll(() => readFilterSummary(page))
    .toMatchObject({
      filterRange: null,
      filterCriteria: [],
      hiddenRows: [],
    });

  await selectRangeAndSetValues(page, { r0: 35, c0: 10, r1: 37, c1: 10 }, [
    { row: 35, col: 10, value: 'Name' },
    { row: 36, col: 10, value: 'alpha' },
    { row: 37, col: 10, value: 'beta' },
  ]);
  await page.locator('[data-ribbon-command="sortFilterHome"]').click();
  await page.locator('#menu-sort-home [data-sort="filter"]').click();
  await expect
    .poll(() => readFilterSummary(page))
    .toMatchObject({
      filterRange: { sheet: 0, r0: 35, c0: 10, r1: 37, c1: 10 },
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readFilterSummary(page)).toMatchObject({ filterRange: null });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readFilterSummary(page))
    .toMatchObject({
      filterRange: { sheet: 0, r0: 35, c0: 10, r1: 37, c1: 10 },
    });
  await page.locator('[data-ribbon-command="sortFilterHome"]').click();
  await page.locator('#menu-sort-home [data-sort="filter"]').click();
  await expect
    .poll(() => readFilterSummary(page))
    .toMatchObject({
      filterRange: null,
      filterCriteria: [],
      hiddenRows: [],
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readFilterSummary(page))
    .toMatchObject({
      filterRange: { sheet: 0, r0: 35, c0: 10, r1: 37, c1: 10 },
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readFilterSummary(page)).toMatchObject({ filterRange: null });

  await selectRangeAndSetValues(page, { r0: 30, c0: 10, r1: 30, c1: 10 }, [
    { row: 30, col: 10, value: 10 },
  ]);
  await page.locator('[data-ribbon-command="conditional"]').click();
  await page.locator('[data-cf-submenu="dataBar"]').hover();
  await page.locator('.app__submenu--cf-dataBar [data-cf-action="data-blue"]').click();
  await page.locator('[data-ribbon-command="findHome"]').click();
  await page.locator('#menu-find-select [data-find-select="conditional-format"]').click();
  await expect
    .poll(() => readSelectionSummary(page))
    .toMatchObject({
      active: { sheet: 0, row: 30, col: 10 },
      range: { sheet: 0, r0: 30, c0: 10, r1: 30, c1: 10 },
    });
});

test('R02fa: Find & Select locates formulas, constants, comments, and data validation cells', async ({
  page,
}) => {
  await mount(page, '/?locale=en&fixture=empty');

  await page.getByRole('tab', { name: 'Home', exact: true }).click();
  await page.locator('[data-ribbon-command="findHome"]').click();
  await page.locator('#menu-find-select [data-find-select="go-to"]').click();
  const normalGoTo = page.getByRole('dialog', { name: 'Go To' });
  await expect(normalGoTo).toBeVisible();
  await normalGoTo.getByLabel('Reference').fill('B2:D4');
  await normalGoTo.getByRole('button', { name: 'OK', exact: true }).click();
  await expect
    .poll(() => readSelectionSummary(page))
    .toMatchObject({
      active: { sheet: 0, row: 1, col: 1 },
      range: { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 3 },
    });

  await selectCellAndSetText(page, 2, 1, 'replace target');
  await page.locator('[data-ribbon-command="findHome"]').click();
  await page.locator('#menu-find-select [data-find-select="replace"]').click();
  const replaceDialog = page.getByRole('dialog', { name: 'Find and Replace' });
  await expect(replaceDialog).toBeVisible();
  await expect(replaceDialog.locator('.fc-find__tab[aria-selected="true"]')).toHaveText('Replace');
  await replaceDialog.getByLabel('Find what:').fill('replace target');
  await replaceDialog.getByLabel('Replace with:').fill('replaced target');
  await replaceDialog.getByRole('button', { name: 'Replace', exact: true }).click();
  await expect.poll(() => readCellText(page, 2, 1)).toBe('replaced target');
  await closeDialog(page);
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 2, 1)).toBe('replace target');
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 2, 1)).toBe('replaced target');

  await page.locator('[data-ribbon-command="findHome"]').click();
  await expect(page.locator('#menu-find-select')).toContainText('Go To Special...');
  await expect(page.locator('#menu-find-select')).toContainText('Data Validation');
  await page.locator('#menu-find-select [data-find-select="go-to-special"]').click();
  await expect(page.locator('.fc-goto')).toBeVisible();
  await closeDialog(page);

  await selectRangeAndSetValues(page, { r0: 0, c0: 0, r1: 0, c1: 1 }, [
    { row: 0, col: 0, value: 10 },
    { row: 0, col: 1, value: 20 },
  ]);
  await selectRangeAndSetFormulas(page, { r0: 0, c0: 2, r1: 0, c1: 2 }, [
    { row: 0, col: 2, formula: '=A1+B1' },
  ]);
  await selectCellAndSetText(page, 2, 1, 'Needs validation');
  await setCommentDirect(page, 1, 3, 'First note');
  await setCommentDirect(page, 3, 4, 'Second note');
  await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            setState: (
              fn: (state: {
                format: { formats: Map<string, { validation?: unknown }> };
              }) => unknown,
            ) => void;
          };
        }
      | undefined;
    inst?.store.setState((state) => {
      const formats = new Map(state.format.formats);
      formats.set('0:2:1', {
        ...(formats.get('0:2:1') ?? {}),
        validation: { kind: 'list', source: ['Open', 'Closed'] },
      });
      return { ...state, format: { formats } };
    });
  });

  await page.getByRole('tab', { name: 'Home', exact: true }).click();
  await page.locator('[data-ribbon-command="findHome"]').click();
  await page.locator('#menu-find-select [data-find-select="formulas"]').click();
  await expect
    .poll(() => readSelectionSummary(page))
    .toMatchObject({
      active: { sheet: 0, row: 0, col: 2 },
      range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
    });

  await page.locator('[data-ribbon-command="findHome"]').click();
  await page.locator('#menu-find-select [data-find-select="constants"]').click();
  await expect
    .poll(() => readSelectionSummary(page))
    .toMatchObject({
      active: { sheet: 0, row: 0, col: 0 },
      range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
    });

  await page.locator('[data-ribbon-command="findHome"]').click();
  await page.locator('#menu-find-select [data-find-select="data-validation"]').click();
  await expect
    .poll(() => readSelectionSummary(page))
    .toMatchObject({
      active: { sheet: 0, row: 2, col: 1 },
      range: { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 },
    });

  await page.locator('[data-ribbon-command="findHome"]').click();
  await page.locator('#menu-find-select [data-find-select="comments"]').click();
  await expect
    .poll(() => readSelectionSummary(page))
    .toMatchObject({
      active: { sheet: 0, row: 1, col: 3 },
      range: { sheet: 0, r0: 1, c0: 3, r1: 3, c1: 4 },
    });

  await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            setState: (
              updater: (state: {
                selection: {
                  active: CellAddr;
                  anchor: CellAddr;
                  range: { sheet: number; r0: number; c0: number; r1: number; c1: number };
                  extraRanges?: unknown[];
                };
              }) => unknown,
            ) => void;
          };
        }
      | undefined;
    const active = { sheet: 0, row: 0, col: 0 };
    inst?.store.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        active,
        anchor: active,
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
        extraRanges: [],
      },
    }));
  });
  await page.locator('[data-ribbon-command="findHome"]').click();
  await page.locator('#menu-find-select [data-find-select="go-to-special"]').click();
  const goToDialog = page.getByRole('dialog', { name: 'Go To Special' });
  await expect(goToDialog).toBeVisible();
  await expect(goToDialog.getByRole('radio', { name: 'Current selection' })).toBeChecked();
  await goToDialog.getByRole('radio', { name: 'Numbers' }).check();
  await goToDialog.getByRole('button', { name: 'OK' }).click();
  await expect
    .poll(() => readSelectionSummary(page))
    .toMatchObject({
      active: { sheet: 0, row: 0, col: 0 },
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
    });
});

test('R02g: Insert tab commands create objects, clean data, and open authoring dialogs', async ({
  page,
}) => {
  await mount(page, '/?locale=en');

  await page.getByRole('tab', { name: 'Insert', exact: true }).click();

  await selectRangeAndSetValues(page, { r0: 40, c0: 0, r1: 42, c1: 1 }, [
    { row: 40, col: 0, value: 'Item' },
    { row: 40, col: 1, value: 'Qty' },
    { row: 41, col: 0, value: 'Pen' },
    { row: 41, col: 1, value: 2 },
    { row: 42, col: 0, value: 'Ink' },
    { row: 42, col: 1, value: 3 },
  ]);

  await page.locator('[data-ribbon-command="formatTableInsert"]').click();
  await page.locator('#menu-table-style-insert [data-table-style="medium"]').first().click();
  await page
    .getByRole('dialog', { name: 'Format as Table' })
    .getByRole('button', { name: 'OK' })
    .click();
  await expect
    .poll(() => readInsertObjectSummary(page))
    .toMatchObject({
      tables: [expect.objectContaining({ source: 'session', style: 'medium' })],
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(async () => (await readInsertObjectSummary(page)).tables).toEqual([]);
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readInsertObjectSummary(page))
    .toMatchObject({
      tables: [expect.objectContaining({ source: 'session', style: 'medium' })],
    });

  await page.locator('[data-ribbon-command="chartInsert"]').click();
  await page.locator('#menu-chart-insert [data-chart-insert="line"]').click();
  await expect
    .poll(() => readInsertObjectSummary(page))
    .toMatchObject({
      charts: [expect.objectContaining({ kind: 'line' })],
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(async () => (await readInsertObjectSummary(page)).charts).toEqual([]);
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readInsertObjectSummary(page))
    .toMatchObject({
      charts: [expect.objectContaining({ kind: 'line' })],
    });
  await page.locator('[data-ribbon-command="chartInsert"]').click();
  await page.locator('#menu-chart-insert [data-chart-insert="pie"]').click();
  await page.locator('[data-ribbon-command="chartInsert"]').click();
  await page.locator('#menu-chart-insert [data-chart-insert="scatter"]').click();
  await expect
    .poll(async () => {
      const summary = await readInsertObjectSummary(page);
      return summary.charts.map((chart) => chart.kind);
    })
    .toEqual(expect.arrayContaining(['line', 'pie', 'scatter']));

  await selectRangeAndSetValues(page, { r0: 47, c0: 0, r1: 49, c1: 0 }, [
    { row: 47, col: 0, value: 3 },
    { row: 48, col: 0, value: 5 },
    { row: 49, col: 0, value: 2 },
  ]);
  await page.locator('[data-ribbon-command="chartInsert"]').click();
  await page.locator('#menu-chart-insert [data-chart-insert="recommended"]').click();
  const recommendedCharts = page.getByRole('dialog', { name: 'Recommended Charts' });
  await expect(recommendedCharts).toBeVisible();
  await expect(recommendedCharts.getByText('Recommended Charts: Bar')).toBeVisible();
  await recommendedCharts.getByRole('button', { name: 'OK', exact: true }).click();
  await expect
    .poll(async () => {
      const summary = await readInsertObjectSummary(page);
      return summary.charts.map((chart) => chart.kind);
    })
    .toEqual(expect.arrayContaining(['bar']));
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(async () => {
      const summary = await readInsertObjectSummary(page);
      return summary.charts.map((chart) => chart.kind);
    })
    .not.toContain('bar');
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(async () => {
      const summary = await readInsertObjectSummary(page);
      return summary.charts.map((chart) => chart.kind);
    })
    .toEqual(expect.arrayContaining(['bar']));

  await page.locator('[data-ribbon-command="pictureInsert"]').click();
  await expect(page.locator('#menu-picture-insert')).toBeVisible();
  await page.locator('#menu-picture-insert [data-picture-insert="online"]').click();
  await expect(page.getByRole('dialog', { name: 'Online Pictures...' })).toBeVisible();
  await page
    .getByRole('textbox', { name: 'Image URL' })
    .fill('data:image/svg+xml,%3Csvg xmlns=%22http://www.w3.org/2000/svg%22/%3E');
  await page.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(page.locator('.app-illustration[data-illustration-type="image"]')).toBeVisible();
  expect(await undoViaInstance(page)).toBe(true);
  await expect(page.locator('.app-illustration[data-illustration-type="image"]')).toHaveCount(0);
  expect(await redoViaInstance(page)).toBe(true);
  await expect(page.locator('.app-illustration[data-illustration-type="image"]')).toBeVisible();

  await page.locator('[data-ribbon-command="shapesInsert"]').click();
  await expect(page.locator('#menu-shapes-insert')).toBeVisible();
  await page.locator('#menu-shapes-insert [data-shape-insert="rounded-rectangle"]').click();
  const insertedShape = page.locator('.app-illustration[data-shape="rounded-rectangle"]');
  await expect(insertedShape).toBeVisible();
  const shapeLeftBefore = await insertedShape.evaluate((el) =>
    parseFloat((el as HTMLElement).style.left),
  );
  const shapeBox = await insertedShape.boundingBox();
  expect(shapeBox).not.toBeNull();
  await page.mouse.move(shapeBox!.x + 20, shapeBox!.y + 20);
  await page.mouse.down();
  await page.mouse.move(shapeBox!.x + 55, shapeBox!.y + 45);
  await page.mouse.up();
  await expect
    .poll(() => insertedShape.evaluate((el) => parseFloat((el as HTMLElement).style.left)))
    .toBeGreaterThan(shapeLeftBefore);
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => insertedShape.evaluate((el) => parseFloat((el as HTMLElement).style.left)))
    .toBe(shapeLeftBefore);
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => insertedShape.evaluate((el) => parseFloat((el as HTMLElement).style.left)))
    .toBeGreaterThan(shapeLeftBefore);
  const shapeWidthBefore = await insertedShape.evaluate((el) =>
    parseFloat((el as HTMLElement).style.width),
  );
  await insertedShape.focus();
  await page.keyboard.press('Alt+ArrowRight');
  await expect
    .poll(() => insertedShape.evaluate((el) => parseFloat((el as HTMLElement).style.width)))
    .toBeGreaterThan(shapeWidthBefore);
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => insertedShape.evaluate((el) => parseFloat((el as HTMLElement).style.width)))
    .toBe(shapeWidthBefore);
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => insertedShape.evaluate((el) => parseFloat((el as HTMLElement).style.width)))
    .toBeGreaterThan(shapeWidthBefore);

  await page.locator('[data-ribbon-command="screenshotInsert"]').click();
  await expect(page.locator('#menu-screenshot-insert')).toBeVisible();
  await page.locator('#menu-screenshot-insert [data-screenshot-insert="current-view"]').click();
  await expect(
    page.locator('.app-illustration[data-illustration-type="screenshot"]'),
  ).toBeVisible();
  expect(await undoViaInstance(page)).toBe(true);
  await expect(page.locator('.app-illustration[data-illustration-type="screenshot"]')).toHaveCount(
    0,
  );
  expect(await redoViaInstance(page)).toBe(true);
  await expect(
    page.locator('.app-illustration[data-illustration-type="screenshot"]'),
  ).toBeVisible();

  await selectRangeAndSetValues(page, { r0: 44, c0: 0, r1: 46, c1: 1 }, [
    { row: 44, col: 0, value: 'A' },
    { row: 44, col: 1, value: 1 },
    { row: 45, col: 0, value: 'A' },
    { row: 45, col: 1, value: 1 },
    { row: 46, col: 0, value: 'B' },
    { row: 46, col: 1, value: 2 },
  ]);
  await page.locator('[data-ribbon-command="removeDupesInsert"]').click();
  await expect(page.getByRole('dialog', { name: 'Remove Duplicates' })).toBeVisible();
  await expect(page.getByRole('checkbox', { name: 'A', exact: true })).toBeChecked();
  await page.getByRole('button', { name: 'OK', exact: true }).click();
  await expect.poll(() => readCellText(page, 45, 0)).toBe('B');
  await expect.poll(() => readCellSummary(page, 46, 0)).toMatchObject({ kind: 'blank' });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 45, 0)).toBe('A');
  await expect.poll(() => readCellSummary(page, 45, 1)).toMatchObject({ kind: 'number', value: 1 });
  await expect.poll(() => readCellText(page, 46, 0)).toBe('B');
  await expect.poll(() => readCellSummary(page, 46, 1)).toMatchObject({ kind: 'number', value: 2 });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 45, 0)).toBe('B');
  await expect.poll(() => readCellSummary(page, 46, 0)).toMatchObject({ kind: 'blank' });

  await selectCellAndSetText(page, 48, 0, 'symbol');
  await page.locator('[data-ribbon-command="symbolInsert"]').click();
  await page.locator('#menu-symbol [data-symbol="Ω"]').click();
  await expect.poll(() => readCellText(page, 48, 0)).toBe('symbolΩ');
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 48, 0)).toBe('symbol');
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 48, 0)).toBe('symbolΩ');
  await page.locator('[data-ribbon-command="symbolInsert"]').click();
  await page.locator('#menu-symbol [data-symbol-action="more"]').click();
  await expect(page.getByRole('dialog', { name: 'More Symbols...' })).toBeVisible();
  await page.getByRole('textbox', { name: 'Character or Unicode text' }).fill('✓');
  await page.getByRole('button', { name: 'OK', exact: true }).click();
  await expect.poll(() => readCellText(page, 48, 0)).toBe('symbolΩ✓');
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 48, 0)).toBe('symbolΩ');
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 48, 0)).toBe('symbolΩ✓');

  await selectCellAndSetText(page, 49, 0, 'locked');
  await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | { setSheetProtected: (protectedState: boolean, password?: string) => void }
      | undefined;
    inst?.setSheetProtected(true);
  });
  await page.locator('[data-ribbon-command="symbolInsert"]').click();
  await page.locator('#menu-symbol [data-symbol="±"]').click();
  await expect(page.getByRole('alertdialog', { name: 'Symbol' })).toBeVisible();
  await page
    .getByRole('alertdialog', { name: 'Symbol' })
    .getByRole('button', { name: 'OK', exact: true })
    .click();
  await expect.poll(() => readCellText(page, 49, 0)).toBe('locked');
  await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | { setSheetProtected: (protectedState: boolean, password?: string) => void }
      | undefined;
    inst?.setSheetProtected(false);
  });

  await selectRangeAndSetValues(page, { r0: 50, c0: 0, r1: 53, c1: 2 }, [
    { row: 50, col: 0, value: 'Region' },
    { row: 50, col: 1, value: 'Product' },
    { row: 50, col: 2, value: 'Sales' },
    { row: 51, col: 0, value: 'East' },
    { row: 51, col: 1, value: 'Pen' },
    { row: 51, col: 2, value: 3 },
    { row: 52, col: 0, value: 'West' },
    { row: 52, col: 1, value: 'Ink' },
    { row: 52, col: 2, value: 4 },
    { row: 53, col: 0, value: 'East' },
    { row: 53, col: 1, value: 'Ink' },
    { row: 53, col: 2, value: 2 },
  ]);
  await page.locator('[data-ribbon-command="pivotTableInsert"]').click();
  await expect(page.locator('#menu-pivot-table')).toBeVisible();
  await expect(page.locator('#menu-pivot-table')).toContainText('Recommended PivotTables');
  await page.locator('#menu-pivot-table [data-pivot-table-action="recommended"]').click();
  const recommendedPivots = page.getByRole('dialog', { name: 'Recommended PivotTables' });
  await expect(recommendedPivots).toBeVisible();
  await expect(recommendedPivots.getByText('Region by Product - Sum of Sales')).toBeVisible();
  await recommendedPivots.getByRole('button', { name: 'OK', exact: true }).click();
  await expect
    .poll(async () => (await readInsertObjectSummary(page)).pivots.length)
    .toBeGreaterThan(0);

  await page.locator('[data-ribbon-command="pivotTableInsert"]').click();
  await page.locator('#menu-pivot-table [data-pivot-table-action="existing-sheet"]').click();
  await expect(page.locator('.fc-pivotdlg')).toBeVisible();
  await closeDialog(page);

  await selectRangeAndSetValues(page, { r0: 54, c0: 0, r1: 56, c1: 1 }, [
    { row: 54, col: 0, value: 'Net Sales' },
    { row: 54, col: 1, value: 'Tax Rate' },
    { row: 55, col: 0, value: 10 },
    { row: 55, col: 1, value: 0.08 },
    { row: 56, col: 0, value: 20 },
    { row: 56, col: 1, value: 0.1 },
  ]);
  await page.locator('[data-ribbon-command="namedRangesInsert"]').click();
  await expect(page.locator('#menu-defined-names-insert')).toBeVisible();
  await page
    .locator('#menu-defined-names-insert [data-defined-name-action="create-top-row"]')
    .click();
  await expect
    .poll(() => readDefinedNames(page))
    .toEqual(
      expect.arrayContaining([
        { name: 'Net_Sales', formula: '=$A$56:$A$57' },
        { name: 'Tax_Rate', formula: '=$B$56:$B$57' },
      ]),
    );
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readDefinedNames(page)).toEqual([]);
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readDefinedNames(page))
    .toEqual(
      expect.arrayContaining([
        { name: 'Net_Sales', formula: '=$A$56:$A$57' },
        { name: 'Tax_Rate', formula: '=$B$56:$B$57' },
      ]),
    );

  await selectRangeAndSetValues(page, { r0: 58, c0: 0, r1: 58, c1: 0 }, []);
  await page.locator('[data-ribbon-command="namedRangesInsert"]').click();
  await expect(page.locator('#menu-defined-names-insert')).toContainText('Net_Sales');
  await page
    .locator('#menu-defined-names-insert [data-defined-name-action="insert:Net_Sales"]')
    .click();
  await expect
    .poll(() => readCellSummary(page, 58, 0))
    .toMatchObject({
      formula: '=Net_Sales',
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 58, 0)).toMatchObject({ kind: 'blank' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCellSummary(page, 58, 0))
    .toMatchObject({
      formula: '=Net_Sales',
    });

  await page.locator('[data-ribbon-command="namedRangesInsert"]').click();
  await page.locator('#menu-defined-names-insert [data-defined-name-action="manager"]').click();
  await expect(page.locator('.fc-namedlg')).toBeVisible();
  await closeDialog(page);

  await page.locator('[data-ribbon-command="hyperlinkInsert"]').click();
  await expect(page.locator('.fc-hldlg')).toBeVisible();
  await page.getByLabel('URL').fill('https://example.com');
  await page.getByRole('button', { name: 'OK', exact: true }).click();
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({
      hyperlink: 'https://example.com',
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(async () => (await readActiveCellFormat(page))?.hyperlink).toBeUndefined();
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({
      hyperlink: 'https://example.com',
    });

  await page.locator('[data-ribbon-command="linksInsert"]').click();
  await expect(page.locator('#menu-links-insert')).toBeVisible();
  await page.locator('#menu-links-insert [data-link-action="clear"]').click();
  await expect
    .poll(() => readActiveCellFormat(page))
    .not.toMatchObject({
      hyperlink: 'https://example.com',
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({
      hyperlink: 'https://example.com',
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(async () => (await readActiveCellFormat(page))?.hyperlink).toBeUndefined();

  await page.locator('[data-ribbon-command="linksInsert"]').click();
  await page.locator('#menu-links-insert [data-link-action="external"]').click();
  await expect(page.locator('.fc-extlinkdlg')).toBeVisible();
  await closeDialog(page);

  await page.locator('[data-ribbon-command="commentInsert"]').click();
  await expect(page.locator('.fc-cmtnote')).toBeVisible();
  await page.locator('.fc-cmtnote').getByRole('button', { name: 'Cancel', exact: true }).click();
});

test('R02h: Page Layout ribbon commands update print and display setup', async ({ page }) => {
  await mount(page, '/?locale=en');

  await page.getByRole('tab', { name: 'Page Layout', exact: true }).click();

  await page.locator('[data-ribbon-command="pageTheme"]').click();
  await expect(page.locator('#menu-page-theme')).toContainText('Office Light');
  await page.locator('#menu-page-theme [data-page-theme-action="dark"]').click();
  await expect
    .poll(() =>
      page.evaluate(() => ({
        shell: document.documentElement.dataset.theme,
        core: document.querySelector<HTMLElement>('.fc-host')?.dataset.fcTheme,
      })),
    )
    .toEqual({ shell: 'dark', core: 'ink' });
  await page.locator('[data-ribbon-command="pageTheme"]').click();
  await page.locator('#menu-page-theme [data-page-theme-action="contrast"]').click();
  await expect
    .poll(() =>
      page.evaluate(() => ({
        shell: document.documentElement.dataset.theme,
        core: document.querySelector<HTMLElement>('.fc-host')?.dataset.fcTheme,
      })),
    )
    .toEqual({ shell: 'contrast', core: 'contrast' });
  await page.locator('[data-ribbon-command="pageTheme"]').click();
  await page.locator('#menu-page-theme [data-page-theme-action="light"]').click();
  await expect
    .poll(() =>
      page.evaluate(() => ({
        shell: document.documentElement.dataset.theme,
        core: document.querySelector<HTMLElement>('.fc-host')?.dataset.fcTheme,
      })),
    )
    .toEqual({ shell: 'light', core: 'paper' });

  await page.locator('[data-ribbon-select="marginsPreset"] .demo__rb-dd__btn').click();
  await page.locator('[data-ribbon-select="marginsPreset"] [data-value="narrow"]').click();
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { margins: { right: 0.25, left: 0.25 } },
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(async () => (await readPageLayoutSummary(page)).setup.margins ?? null)
    .toBeNull();
  await page.locator('[data-ribbon-select="marginsPreset"] .demo__rb-dd__btn').click();
  await page.locator('[data-ribbon-select="marginsPreset"] [data-value="narrow"]').click();
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { margins: { right: 0.25, left: 0.25 } },
    });

  await page.locator('[data-ribbon-select="orientationPreset"] .demo__rb-dd__btn').click();
  await page.locator('[data-ribbon-select="orientationPreset"] [data-value="landscape"]').click();
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { orientation: 'landscape' },
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(async () => (await readPageLayoutSummary(page)).setup.orientation ?? null)
    .toBe('portrait');
  await page.locator('[data-ribbon-select="orientationPreset"] .demo__rb-dd__btn').click();
  await page.locator('[data-ribbon-select="orientationPreset"] [data-value="landscape"]').click();
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { orientation: 'landscape' },
    });

  await page.locator('[data-ribbon-select="paperSizePreset"] .demo__rb-dd__btn').click();
  await page.locator('[data-ribbon-select="paperSizePreset"] [data-value="letter"]').click();
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { paperSize: 'letter' },
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(async () => (await readPageLayoutSummary(page)).setup.paperSize ?? null)
    .toBe('A4');
  await page.locator('[data-ribbon-select="paperSizePreset"] .demo__rb-dd__btn').click();
  await page.locator('[data-ribbon-select="paperSizePreset"] [data-value="letter"]').click();
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { paperSize: 'letter' },
    });

  await page.locator('[data-ribbon-command="pageSetupAdvanced"]').click();
  const pageSetup = page.getByRole('dialog', { name: 'Page Setup' });
  await expect(pageSetup).toBeVisible();
  await pageSetup.locator('select[aria-label="Orientation"]').selectOption('portrait');
  await pageSetup.locator('select[aria-label="Paper size"]').selectOption('A5');
  await pageSetup.getByLabel('Adjust to').check();
  await pageSetup.getByRole('spinbutton', { name: 'Scale', exact: true }).fill('90');

  await pageSetup.locator('[data-pgsetup-tab="margins"][role="tab"]').click();
  await pageSetup.getByRole('spinbutton', { name: 'Top', exact: true }).fill('0.4');
  await pageSetup.getByRole('spinbutton', { name: 'Right', exact: true }).fill('0.6');
  await pageSetup.getByRole('spinbutton', { name: 'Bottom', exact: true }).fill('0.8');
  await pageSetup.getByRole('spinbutton', { name: 'Left', exact: true }).fill('0.5');
  await pageSetup.getByLabel('Horizontally').check();

  await pageSetup.locator('[data-pgsetup-tab="sheet"][role="tab"]').click();
  await pageSetup.getByLabel('Print area').fill('A1:B4');
  await pageSetup.getByLabel('Print title rows').fill('1:1');
  await pageSetup.getByLabel('Print title columns').fill('A:A');
  await pageSetup.getByLabel('Print gridlines').check();
  await pageSetup.getByLabel('Print headings').check();
  await pageSetup.getByRole('button', { name: 'OK', exact: true }).click();
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: {
        orientation: 'portrait',
        paperSize: 'A5',
        margins: { top: 0.4, right: 0.6, bottom: 0.8, left: 0.5 },
        printArea: 'A1:B4',
        printTitleRows: '1:1',
        printTitleCols: 'A:A',
        scale: 0.9,
        fitWidth: 0,
        fitHeight: 0,
        showGridlines: true,
        showHeadings: true,
      },
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { orientation: 'landscape', paperSize: 'letter' },
    });
  await expect
    .poll(async () => (await readPageLayoutSummary(page)).setup.printArea ?? null)
    .toBeNull();
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: {
        orientation: 'portrait',
        paperSize: 'A5',
        printArea: 'A1:B4',
        printTitleRows: '1:1',
        printTitleCols: 'A:A',
        scale: 0.9,
        showGridlines: true,
        showHeadings: true,
      },
    });

  await selectRangeAndSetValues(page, { r0: 2, c0: 1, r1: 4, c1: 3 }, []);
  await page.locator('[data-ribbon-command="printArea"]').click();
  await page.locator('#menu-print-area [data-print-area-action="set"]').click();
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { printArea: 'B3:D5' },
    });
  await page.getByRole('button', { name: 'OK', exact: true }).click();
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(async () => (await readPageLayoutSummary(page)).setup.printArea ?? null)
    .toBe('A1:B4');
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { printArea: 'B3:D5' },
    });

  await selectRangeAndSetValues(page, { r0: 6, c0: 0, r1: 6, c1: 0 }, []);
  await page.locator('[data-ribbon-command="pageBreaks"]').click();
  await page.locator('#menu-page-breaks [data-page-break-action="insert-row"]').click();
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { manualPageBreakRows: [6] },
    });
  await selectRangeAndSetValues(page, { r0: 0, c0: 4, r1: 0, c1: 4 }, []);
  await page.locator('[data-ribbon-command="pageBreaks"]').click();
  await page.locator('#menu-page-breaks [data-page-break-action="insert-col"]').click();
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { manualPageBreakRows: [6], manualPageBreakCols: [4] },
    });
  await page.locator('[data-ribbon-command="pageBreaks"]').click();
  await page.locator('#menu-page-breaks [data-page-break-action="remove-col"]').click();
  await expect
    .poll(async () => (await readPageLayoutSummary(page)).setup.manualPageBreakRows ?? [])
    .toEqual([6]);
  await expect
    .poll(async () => (await readPageLayoutSummary(page)).setup.manualPageBreakCols ?? [])
    .toEqual([]);
  await page.locator('[data-ribbon-command="pageBreaks"]').click();
  await page.locator('#menu-page-breaks [data-page-break-action="reset-all"]').click();
  await expect
    .poll(async () => (await readPageLayoutSummary(page)).setup.manualPageBreakRows ?? [])
    .toEqual([]);
  await expect
    .poll(async () => (await readPageLayoutSummary(page)).setup.manualPageBreakCols ?? [])
    .toEqual([]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(async () => (await readPageLayoutSummary(page)).setup.manualPageBreakRows ?? [])
    .toEqual([6]);
  await expect
    .poll(async () => (await readPageLayoutSummary(page)).setup.manualPageBreakCols ?? [])
    .toEqual([]);
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(async () => (await readPageLayoutSummary(page)).setup.manualPageBreakRows ?? [])
    .toEqual([]);
  await expect
    .poll(async () => (await readPageLayoutSummary(page)).setup.manualPageBreakCols ?? [])
    .toEqual([]);

  await page.locator('[data-ribbon-command="sheetBackground"]').click();
  await expect(page.locator('#menu-sheet-background')).toBeVisible();
  await page.locator('#menu-sheet-background [data-sheet-background-action="set"]').click();
  await expect(page.getByRole('dialog', { name: 'Choose Background...' })).toBeVisible();
  await page.getByRole('button', { name: 'OK', exact: true }).click();
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      ui: { background: expect.stringContaining('linear-gradient') },
    });
  await page.locator('[data-ribbon-command="sheetBackground"]').click();
  await page.locator('#menu-sheet-background [data-sheet-background-action="clear"]').click();
  await expect
    .poll(async () => (await readPageLayoutSummary(page)).ui.background ?? null)
    .toBeNull();
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(async () => (await readPageLayoutSummary(page)).ui.background ?? '')
    .toContain('linear-gradient');
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(async () => (await readPageLayoutSummary(page)).ui.background ?? null)
    .toBeNull();

  await selectRangeAndSetValues(page, { r0: 0, c0: 1, r1: 1, c1: 3 }, []);
  await page.locator('[data-ribbon-command="printTitles"]').click();
  await page.locator('#menu-print-titles [data-print-titles-action="rows"]').click();
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { printTitleRows: '1:2' },
    });
  await page.locator('[data-ribbon-command="printTitles"]').click();
  await page.locator('#menu-print-titles [data-print-titles-action="cols"]').click();
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { printTitleRows: '1:2', printTitleCols: 'B:D' },
    });
  await page.locator('[data-ribbon-command="printTitles"]').click();
  await page.locator('#menu-print-titles [data-print-titles-action="clear"]').click();
  await expect
    .poll(async () => (await readPageLayoutSummary(page)).setup.printTitleRows ?? null)
    .toBeNull();
  await expect
    .poll(async () => (await readPageLayoutSummary(page)).setup.printTitleCols ?? null)
    .toBeNull();
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { printTitleRows: '1:2', printTitleCols: 'B:D' },
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(async () => (await readPageLayoutSummary(page)).setup.printTitleRows ?? null)
    .toBeNull();
  await expect
    .poll(async () => (await readPageLayoutSummary(page)).setup.printTitleCols ?? null)
    .toBeNull();

  await page.locator('[data-ribbon-select="scaleWidth"] .demo__rb-dd__btn').click();
  await page.locator('[data-ribbon-select="scaleWidth"] [data-value="2"]').click();
  await page.locator('[data-ribbon-select="scaleHeight"] .demo__rb-dd__btn').click();
  await page.locator('[data-ribbon-select="scaleHeight"] [data-value="1"]').click();
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { fitWidth: 2, fitHeight: 1 },
    });
  await page.locator('[data-ribbon-select="scaleWidth"] .demo__rb-dd__btn').click();
  await page.locator('[data-ribbon-select="scaleWidth"] [data-value="custom"]').click();
  await expect(page.getByRole('dialog', { name: 'Width' })).toBeVisible();
  await page.getByRole('textbox', { name: 'Number of pages (1-99)' }).fill('5');
  await page.getByRole('button', { name: 'OK', exact: true }).click();
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { fitWidth: 5, fitHeight: 1 },
    });
  await expect(page.locator('[data-ribbon-select="scaleWidth"] .demo__rb-dd__value')).toHaveText(
    '5 pages',
  );
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { fitWidth: 2, fitHeight: 1 },
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { fitWidth: 5, fitHeight: 1 },
    });

  await page.locator('[data-ribbon-select="scaleHeight"] .demo__rb-dd__btn').click();
  await page.locator('[data-ribbon-select="scaleHeight"] [data-value="custom"]').click();
  await expect(page.getByRole('dialog', { name: 'Height' })).toBeVisible();
  await page.getByRole('textbox', { name: 'Number of pages (1-99)' }).fill('3');
  await page.getByRole('button', { name: 'OK', exact: true }).click();
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { fitWidth: 5, fitHeight: 3 },
    });
  await expect(page.locator('[data-ribbon-select="scaleHeight"] .demo__rb-dd__value')).toHaveText(
    '3 pages',
  );
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { fitWidth: 5, fitHeight: 1 },
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { fitWidth: 5, fitHeight: 3 },
    });

  await page.locator('[data-ribbon-select="scalePercent"] .demo__rb-dd__btn').click();
  await page.locator('[data-ribbon-select="scalePercent"] [data-value="125"]').click();
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { fitWidth: 0, fitHeight: 0, scale: 1.25 },
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { fitWidth: 5, fitHeight: 3 },
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { fitWidth: 0, fitHeight: 0, scale: 1.25 },
    });
  await page.locator('[data-ribbon-select="scalePercent"] .demo__rb-dd__btn').click();
  await page.locator('[data-ribbon-select="scalePercent"] [data-value="custom"]').click();
  await expect(page.getByRole('dialog', { name: 'Scale' })).toBeVisible();
  await page.getByRole('textbox', { name: 'Scale percentage (10-400)' }).fill('175');
  await page.getByRole('button', { name: 'OK', exact: true }).click();
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { fitWidth: 0, fitHeight: 0, scale: 1.75 },
    });
  await expect(page.locator('[data-ribbon-select="scalePercent"] .demo__rb-dd__value')).toHaveText(
    '175%',
  );
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { fitWidth: 0, fitHeight: 0, scale: 1.25 },
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { fitWidth: 0, fitHeight: 0, scale: 1.75 },
    });

  await page.locator('[data-ribbon-command="pageLayoutGridlinesView"]').click();
  await page.locator('[data-ribbon-command="pageLayoutHeadingsView"]').click();
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      ui: { showGridLines: false, showHeaders: false },
    });

  await page.locator('[data-ribbon-command="pageLayoutGridlinesPrint"]').click();
  await page.locator('[data-ribbon-command="pageLayoutHeadingsPrint"]').click();
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { showGridlines: false, showHeadings: false },
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { showGridlines: false, showHeadings: true },
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { showGridlines: true, showHeadings: true },
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { showGridlines: false, showHeadings: true },
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readPageLayoutSummary(page))
    .toMatchObject({
      setup: { showGridlines: false, showHeadings: false },
    });
});

test('R02i: Formulas ribbon commands evaluate, show formulas, and check errors', async ({
  page,
}) => {
  await mount(page, '/?locale=en');

  await selectRangeAndSetFormulas(page, { r0: 0, c0: 0, r1: 1, c1: 1 }, [
    { row: 0, col: 0, formula: '=1+2' },
    { row: 1, col: 1, formula: '=A1/0' },
  ]);
  await expect
    .poll(() => readCellSummary(page, 1, 1))
    .toMatchObject({
      formula: '=A1/0',
    });

  await page.getByRole('tab', { name: 'Formulas', exact: true }).click();
  await page.locator('[data-ribbon-command="showFormulasFormula"]').click();
  await expect
    .poll(() =>
      page.evaluate(() => {
        const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
          | { store: { getState: () => { ui: { showFormulas?: boolean } } } }
          | undefined;
        return inst?.store.getState().ui.showFormulas ?? false;
      }),
    )
    .toBe(true);
  await expect(page.locator('[data-ribbon-command="showFormulasFormula"]')).toHaveAttribute(
    'aria-pressed',
    'true',
  );

  await page.locator('[data-ribbon-command="errorChecking"]').click();
  await expect(page.locator('#menu-error-checking')).toBeVisible();
  await page.locator('#menu-error-checking [data-formula-audit-action="error-checking"]').click();
  await expect
    .poll(() => readSelectionSummary(page))
    .toMatchObject({
      active: { sheet: 0, row: 1, col: 1 },
    });

  await page.locator('[data-ribbon-command="errorChecking"]').click();
  await page.locator('#menu-error-checking [data-formula-audit-action="trace-error"]').click();
  await expect
    .poll(() => readTraceSummaries(page))
    .toContainEqual({
      kind: 'precedent',
      from: { sheet: 0, row: 0, col: 0 },
      to: { sheet: 0, row: 1, col: 1 },
    });

  await page.locator('[data-ribbon-command="errorChecking"]').click();
  await page.locator('#menu-error-checking [data-formula-audit-action="ignore-error"]').click();
  await expect.poll(() => readIgnoredErrorKeys(page)).toEqual(['0:1:1']);
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readIgnoredErrorKeys(page)).toEqual([]);
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readIgnoredErrorKeys(page)).toEqual(['0:1:1']);

  await page.locator('[data-ribbon-command="clearArrows"]').click();
  await expect(page.locator('#menu-clear-arrows')).toBeVisible();
  await page.locator('#menu-clear-arrows [data-formula-audit-action="clear-precedents"]').click();
  await expect.poll(() => readTraceSummaries(page)).toEqual([]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readTraceSummaries(page))
    .toContainEqual({
      kind: 'precedent',
      from: { sheet: 0, row: 0, col: 0 },
      to: { sheet: 0, row: 1, col: 1 },
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readTraceSummaries(page)).toEqual([]);

  await page.locator('[data-ribbon-command="evaluateFormula"]').click();
  await expect(page.locator('.fc-evaldlg')).toBeVisible();
  await expect(page.locator('.fc-evaldlg')).toContainText('B2');
  await expect(page.locator('.fc-evaldlg')).toContainText('=A1/0');
  await page.locator('.fc-evaldlg__btn', { hasText: 'Evaluate' }).click();
  await expect(page.locator('.fc-evaldlg__box--evaluation')).toContainText('=3/0');
  await closeDialog(page);

  await selectRangeAndSetValues(page, { r0: 6, c0: 0, r1: 6, c1: 1 }, [
    { row: 6, col: 0, value: 10 },
    { row: 6, col: 1, value: 20 },
  ]);
  await selectRangeAndSetFormulas(page, { r0: 6, c0: 2, r1: 6, c1: 2 }, [
    { row: 6, col: 2, formula: '=A7+B7' },
  ]);
  await page.locator('[data-ribbon-command="precedents"]').click();
  await expect
    .poll(() => readTraceSummaries(page))
    .toEqual([
      {
        kind: 'precedent',
        from: { sheet: 0, row: 6, col: 0 },
        to: { sheet: 0, row: 6, col: 2 },
      },
      {
        kind: 'precedent',
        from: { sheet: 0, row: 6, col: 1 },
        to: { sheet: 0, row: 6, col: 2 },
      },
    ]);

  await selectRangeAndSetValues(page, { r0: 6, c0: 0, r1: 6, c1: 0 }, []);
  await page.locator('[data-ribbon-command="dependents"]').click();
  await expect
    .poll(() => readTraceSummaries(page))
    .toContainEqual({
      kind: 'dependent',
      from: { sheet: 0, row: 6, col: 0 },
      to: { sheet: 0, row: 6, col: 2 },
    });

  await page.locator('[data-ribbon-command="clearArrows"]').click();
  await page.locator('#menu-clear-arrows [data-formula-audit-action="clear-dependents"]').click();
  await expect
    .poll(() => readTraceSummaries(page))
    .not.toContainEqual({
      kind: 'dependent',
      from: { sheet: 0, row: 6, col: 0 },
      to: { sheet: 0, row: 6, col: 2 },
    });
  await expect
    .poll(() => readTraceSummaries(page))
    .toContainEqual({
      kind: 'precedent',
      from: { sheet: 0, row: 6, col: 0 },
      to: { sheet: 0, row: 6, col: 2 },
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readTraceSummaries(page))
    .toContainEqual({
      kind: 'dependent',
      from: { sheet: 0, row: 6, col: 0 },
      to: { sheet: 0, row: 6, col: 2 },
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readTraceSummaries(page))
    .not.toContainEqual({
      kind: 'dependent',
      from: { sheet: 0, row: 6, col: 0 },
      to: { sheet: 0, row: 6, col: 2 },
    });

  await page.locator('[data-ribbon-command="clearArrows"]').click();
  await page.locator('#menu-clear-arrows [data-formula-audit-action="clear-all"]').click();
  await expect.poll(() => readTraceSummaries(page)).toEqual([]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readTraceSummaries(page)).not.toEqual([]);
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readTraceSummaries(page)).toEqual([]);
});

test('R02i-functions: Formulas function library buttons seed arguments and insert formulas', async ({
  page,
}) => {
  await mount(page, '/?locale=en&fixture=empty');

  await selectRangeAndSetValues(page, { r0: 0, c0: 0, r1: 0, c1: 1 }, [
    { row: 0, col: 0, value: 10 },
    { row: 0, col: 1, value: 20 },
  ]);
  await selectRangeAndSetValues(page, { r0: 0, c0: 2, r1: 0, c1: 2 }, []);
  await page.getByRole('tab', { name: 'Formulas', exact: true }).click();

  await selectRangeAndSetValues(page, { r0: 2, c0: 0, r1: 4, c1: 0 }, [
    { row: 2, col: 0, value: 10 },
    { row: 3, col: 0, value: 20 },
  ]);
  await page.locator('[data-ribbon-command="autosumFormula"] .demo__rb-split-chevron').click();
  await page.locator('#menu-autosum-formulas [data-autosum-fn="AVERAGE"]').click();
  await expect
    .poll(() => readCellSummary(page, 4, 0))
    .toMatchObject({ kind: 'number', value: 15, formula: '=AVERAGE(A3:A4)' });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 4, 0)).toMatchObject({ kind: 'blank' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCellSummary(page, 4, 0))
    .toMatchObject({ kind: 'number', value: 15, formula: '=AVERAGE(A3:A4)' });

  await selectRangeAndSetValues(page, { r0: 2, c0: 1, r1: 4, c1: 1 }, [
    { row: 2, col: 1, value: 7 },
    { row: 3, col: 1, value: 12 },
  ]);
  await page.locator('[data-ribbon-command="autosumFormula"] .demo__rb-split-chevron').click();
  await page.locator('#menu-autosum-formulas [data-autosum-fn="MAX"]').click();
  await expect
    .poll(() => readCellSummary(page, 4, 1))
    .toMatchObject({ kind: 'number', value: 12, formula: '=MAX(B3:B4)' });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 4, 1)).toMatchObject({ kind: 'blank' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCellSummary(page, 4, 1))
    .toMatchObject({ kind: 'number', value: 12, formula: '=MAX(B3:B4)' });

  await selectRangeAndSetValues(page, { r0: 2, c0: 2, r1: 4, c1: 2 }, [
    { row: 2, col: 2, value: 7 },
    { row: 3, col: 2, value: 12 },
  ]);
  await page.locator('[data-ribbon-command="autosumFormula"] .demo__rb-split-chevron').click();
  await page.locator('#menu-autosum-formulas [data-autosum-fn="MIN"]').click();
  await expect
    .poll(() => readCellSummary(page, 4, 2))
    .toMatchObject({ kind: 'number', value: 7, formula: '=MIN(C3:C4)' });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 4, 2)).toMatchObject({ kind: 'blank' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCellSummary(page, 4, 2))
    .toMatchObject({ kind: 'number', value: 7, formula: '=MIN(C3:C4)' });

  await selectRangeAndSetValues(page, { r0: 0, c0: 2, r1: 0, c1: 2 }, []);

  const seededButtons: Array<[string, string]> = [
    ['sum', 'SUM(number1, [number2], ...)'],
    ['avg', 'AVERAGE(number1, [number2], ...)'],
    ['ifFormula', 'IF(logical_test, value_if_true, [value_if_false])'],
    [
      'xlookupFormula',
      'XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])',
    ],
    ['concatFormula', 'CONCAT(text1, [text2], ...)'],
    ['todayFormula', 'TODAY()'],
    ['pmtFormula', 'PMT(rate, nper, pv, [fv], [type])'],
    ['roundFormula', 'ROUND(number, num_digits)'],
  ];

  for (const [command, signature] of seededButtons) {
    await page.locator(`[data-ribbon-command="${command}"]`).click();
    await expect(page.locator('.fc-fxdialog')).toBeVisible();
    await expect(page.locator('.fc-fxdialog__args-name')).toHaveText(signature);
    await closeDialog(page);
  }

  await page.locator('[data-ribbon-command="sum"]').click();
  await page.locator('.fc-fxdialog__arg-input').first().fill('A1:B1');
  await page.locator('.fc-fxdialog .fc-fmtdlg__btn--primary').click();
  await expect.poll(() => readCellSummary(page, 0, 2)).toMatchObject({ formula: '=SUM(A1:B1)' });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 0, 2)).toMatchObject({ kind: 'blank' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 0, 2)).toMatchObject({ formula: '=SUM(A1:B1)' });

  await selectRangeAndSetValues(page, { r0: 0, c0: 3, r1: 0, c1: 3 }, []);
  await page.locator('[data-ribbon-command="ifFormula"]').click();
  await page.locator('.fc-fxdialog__arg-input').nth(0).fill('A1>5');
  await page.locator('.fc-fxdialog__arg-input').nth(1).fill('"yes"');
  await page.locator('.fc-fxdialog__arg-input').nth(2).fill('"no"');
  await page.locator('.fc-fxdialog .fc-fmtdlg__btn--primary').click();
  await expect
    .poll(() => readCellSummary(page, 0, 3))
    .toMatchObject({ formula: '=IF(A1>5, "yes", "no")' });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 0, 3)).toMatchObject({ kind: 'blank' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCellSummary(page, 0, 3))
    .toMatchObject({ formula: '=IF(A1>5, "yes", "no")' });

  await selectRangeAndSetValues(page, { r0: 0, c0: 4, r1: 0, c1: 4 }, []);
  await page.locator('[data-ribbon-command="roundFormula"]').click();
  await page.locator('.fc-fxdialog__arg-input').nth(0).fill('1.234');
  await page.locator('.fc-fxdialog__arg-input').nth(1).fill('2');
  await page.locator('.fc-fxdialog .fc-fmtdlg__btn--primary').click();
  await expect
    .poll(() => readCellSummary(page, 0, 4))
    .toMatchObject({ formula: '=ROUND(1.234, 2)' });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 0, 4)).toMatchObject({ kind: 'blank' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCellSummary(page, 0, 4))
    .toMatchObject({ formula: '=ROUND(1.234, 2)' });

  await selectRangeAndSetValues(page, { r0: 0, c0: 5, r1: 0, c1: 5 }, []);
  await page.locator('[data-ribbon-command="todayFormula"]').click();
  await expect(page.locator('.fc-fxdialog__arg-input')).toHaveCount(0);
  await page.locator('.fc-fxdialog .fc-fmtdlg__btn--primary').click();
  await expect.poll(() => readCellSummary(page, 0, 5)).toMatchObject({ formula: '=TODAY()' });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 0, 5)).toMatchObject({ kind: 'blank' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 0, 5)).toMatchObject({ formula: '=TODAY()' });
});

test('R02j: Data Text to Columns menu supports presets and custom delimiters', async ({ page }) => {
  await mount(page, '/?locale=en');

  await page.getByRole('tab', { name: 'Data', exact: true }).click();

  await selectCellAndSetText(page, 52, 0, 'red;green;blue');
  await page.locator('[data-ribbon-command="textToColumns"]').click();
  await expect(page.locator('#menu-text-to-columns')).toBeVisible();
  await page.locator('#menu-text-to-columns [data-text-to-columns-delimiter=";"]').click();
  await expect.poll(() => readCellText(page, 52, 0)).toBe('red');
  await expect.poll(() => readCellText(page, 52, 1)).toBe('green');
  await expect.poll(() => readCellText(page, 52, 2)).toBe('blue');
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 52, 0)).toBe('red;green;blue');
  await expect.poll(() => readCellSummary(page, 52, 1)).toMatchObject({ kind: 'blank' });
  await expect.poll(() => readCellSummary(page, 52, 2)).toMatchObject({ kind: 'blank' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 52, 0)).toBe('red');
  await expect.poll(() => readCellText(page, 52, 1)).toBe('green');
  await expect.poll(() => readCellText(page, 52, 2)).toBe('blue');

  await selectCellAndSetText(page, 53, 0, 'north|south|west');
  await page.locator('[data-ribbon-command="textToColumns"]').click();
  await page.locator('#menu-text-to-columns [data-text-to-columns-delimiter="custom"]').click();
  await expect(page.getByRole('dialog', { name: 'Convert Text to Columns' })).toBeVisible();
  await page.locator('.app__dlg__input').fill('|');
  await page.getByRole('button', { name: 'OK', exact: true }).click();
  await expect.poll(() => readCellText(page, 53, 0)).toBe('north');
  await expect.poll(() => readCellText(page, 53, 1)).toBe('south');
  await expect.poll(() => readCellText(page, 53, 2)).toBe('west');
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 53, 0)).toBe('north|south|west');
  await expect.poll(() => readCellSummary(page, 53, 1)).toMatchObject({ kind: 'blank' });
  await expect.poll(() => readCellSummary(page, 53, 2)).toMatchObject({ kind: 'blank' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 53, 0)).toBe('north');
  await expect.poll(() => readCellText(page, 53, 1)).toBe('south');
  await expect.poll(() => readCellText(page, 53, 2)).toBe('west');
});

test('R02ja: Japanese Text to Columns localizes custom delimiter dialog and status', async ({
  page,
}) => {
  await mount(page, '/?locale=ja');

  await page.getByRole('tab', { name: 'データ', exact: true }).click();
  await selectCellAndSetText(page, 54, 0, '東|西|南');
  await page.locator('[data-ribbon-command="textToColumns"]').click();
  await page.locator('#menu-text-to-columns [data-text-to-columns-delimiter="custom"]').click();
  await expect(page.getByRole('dialog', { name: '区切り位置指定ウィザード' })).toBeVisible();
  await page.getByRole('textbox', { name: '区切り文字' }).fill('|');
  await page.getByRole('button', { name: 'OK', exact: true }).click();
  await expect.poll(() => readCellText(page, 54, 0)).toBe('東');
  await expect.poll(() => readCellText(page, 54, 1)).toBe('西');
  await expect.poll(() => readCellText(page, 54, 2)).toBe('南');
  await expect(page.locator('#status-metric')).toContainText('テキストを 3 列に分割しました');

  await selectCellAndSetText(page, 55, 0, '区切りなし');
  await page.locator('[data-ribbon-command="textToColumns"]').click();
  await page.locator('#menu-text-to-columns [data-text-to-columns-delimiter=";"]').click();
  await expect(page.locator('#status-metric')).toContainText(
    '区切り文字を含むテキストが見つかりません',
  );
});

test('R02k: Data filter menu infers current region, applies values, and clears filter', async ({
  page,
}) => {
  await mount(page, '/?locale=en');

  await selectRangeAndSetValues(page, { r0: 75, c0: 0, r1: 77, c1: 1 }, [
    { row: 75, col: 0, value: 'Bronze' },
    { row: 75, col: 1, value: 1 },
    { row: 76, col: 0, value: 'Gold' },
    { row: 76, col: 1, value: 3 },
    { row: 77, col: 0, value: 'Silver' },
    { row: 77, col: 1, value: 2 },
  ]);
  await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            setState: (
              fn: (state: {
                selection: {
                  active: CellAddr;
                  anchor: CellAddr;
                  range: { sheet: number; r0: number; c0: number; r1: number; c1: number };
                  extraRanges?: unknown[];
                };
              }) => unknown,
            ) => void;
          };
        }
      | undefined;
    const active = { sheet: 0, row: 75, col: 1 };
    inst?.store.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        active,
        anchor: active,
        range: { sheet: 0, r0: 75, c0: 0, r1: 77, c1: 1 },
        extraRanges: [],
      },
    }));
  });
  await page.getByRole('tab', { name: 'Data', exact: true }).click();
  await page.locator('[data-ribbon-command="sortDesc"]').click();
  await expect.poll(() => readCellText(page, 75, 0)).toBe('Gold');
  await expect.poll(() => readCellText(page, 76, 0)).toBe('Silver');
  await expect.poll(() => readCellText(page, 77, 0)).toBe('Bronze');
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 75, 0)).toBe('Bronze');
  await expect.poll(() => readCellText(page, 76, 0)).toBe('Gold');
  await expect.poll(() => readCellText(page, 77, 0)).toBe('Silver');
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 75, 0)).toBe('Gold');
  await expect.poll(() => readCellText(page, 76, 0)).toBe('Silver');
  await expect.poll(() => readCellText(page, 77, 0)).toBe('Bronze');

  await selectRangeAndSetValues(page, { r0: 80, c0: 0, r1: 83, c1: 1 }, [
    { row: 80, col: 0, value: 'Item' },
    { row: 80, col: 1, value: 'Qty' },
    { row: 81, col: 0, value: 'Gamma' },
    { row: 81, col: 1, value: 2 },
    { row: 82, col: 0, value: 'Alpha' },
    { row: 82, col: 1, value: 3 },
    { row: 83, col: 0, value: 'Beta' },
    { row: 83, col: 1, value: 1 },
  ]);
  await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            setState: (
              fn: (state: {
                selection: {
                  active: CellAddr;
                  anchor: CellAddr;
                  range: { sheet: number; r0: number; c0: number; r1: number; c1: number };
                  extraRanges?: unknown[];
                };
              }) => unknown,
            ) => void;
          };
        }
      | undefined;
    const active = { sheet: 0, row: 82, col: 1 };
    inst?.store.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        active,
        anchor: active,
        range: { sheet: 0, r0: 82, c0: 1, r1: 82, c1: 1 },
        extraRanges: [],
      },
    }));
  });
  await page.locator('[data-ribbon-command="sortAsc"]').click();
  await expect.poll(() => readCellText(page, 80, 0)).toBe('Item');
  await expect.poll(() => readCellText(page, 81, 0)).toBe('Beta');
  await expect.poll(() => readCellText(page, 82, 0)).toBe('Gamma');
  await expect.poll(() => readCellText(page, 83, 0)).toBe('Alpha');
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 81, 0)).toBe('Gamma');
  await expect.poll(() => readCellText(page, 82, 0)).toBe('Alpha');
  await expect.poll(() => readCellText(page, 83, 0)).toBe('Beta');
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 81, 0)).toBe('Beta');
  await expect.poll(() => readCellText(page, 82, 0)).toBe('Gamma');
  await expect.poll(() => readCellText(page, 83, 0)).toBe('Alpha');

  await page.locator('[data-ribbon-command="sortData"]').click();
  const customSort = page.getByRole('dialog', { name: 'Custom Sort...' });
  await expect(customSort).toBeVisible();
  await customSort.getByLabel('Sort by').selectOption('0');
  await customSort.locator('select').nth(2).selectOption('desc');
  await customSort.getByRole('button', { name: 'OK' }).click();
  await expect.poll(() => readCellText(page, 80, 0)).toBe('Item');
  await expect.poll(() => readCellText(page, 81, 0)).toBe('Gamma');
  await expect.poll(() => readCellText(page, 82, 0)).toBe('Beta');
  await expect.poll(() => readCellText(page, 83, 0)).toBe('Alpha');
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 81, 0)).toBe('Beta');
  await expect.poll(() => readCellText(page, 82, 0)).toBe('Gamma');
  await expect.poll(() => readCellText(page, 83, 0)).toBe('Alpha');
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 81, 0)).toBe('Gamma');
  await expect.poll(() => readCellText(page, 82, 0)).toBe('Beta');
  await expect.poll(() => readCellText(page, 83, 0)).toBe('Alpha');

  await selectRangeAndSetValues(page, { r0: 86, c0: 0, r1: 88, c1: 1 }, [
    { row: 86, col: 0, value: 'Item' },
    { row: 86, col: 1, value: 'Qty' },
    { row: 87, col: 0, value: 'Pen' },
    { row: 87, col: 1, value: 2 },
    { row: 88, col: 0, value: 'Pen' },
    { row: 88, col: 1, value: 2 },
  ]);
  await page.locator('[data-ribbon-command="removeDupes"]').click();
  await expect(page.getByRole('dialog', { name: 'Remove Duplicates' })).toBeVisible();
  await expect(page.getByRole('checkbox', { name: /Item|A/ }).first()).toBeChecked();
  await expect(page.getByRole('checkbox', { name: /Qty|B/ }).first()).toBeChecked();
  await page.getByRole('button', { name: 'OK', exact: true }).click();
  await expect.poll(() => readCellText(page, 87, 0)).toBe('Pen');
  await expect.poll(() => readCellSummary(page, 88, 0)).toMatchObject({ kind: 'blank' });
  await expect.poll(() => readCellSummary(page, 88, 1)).toMatchObject({ kind: 'blank' });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 87, 0)).toBe('Pen');
  await expect.poll(() => readCellSummary(page, 87, 1)).toMatchObject({ kind: 'number', value: 2 });
  await expect.poll(() => readCellText(page, 88, 0)).toBe('Pen');
  await expect.poll(() => readCellSummary(page, 88, 1)).toMatchObject({ kind: 'number', value: 2 });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 87, 0)).toBe('Pen');
  await expect.poll(() => readCellSummary(page, 88, 0)).toMatchObject({ kind: 'blank' });
  await expect.poll(() => readCellSummary(page, 88, 1)).toMatchObject({ kind: 'blank' });

  await selectRangeAndSetValues(page, { r0: 70, c0: 0, r1: 73, c1: 1 }, [
    { row: 70, col: 0, value: 'Region' },
    { row: 70, col: 1, value: 'Qty' },
    { row: 71, col: 0, value: 'East' },
    { row: 71, col: 1, value: 3 },
    { row: 72, col: 0, value: 'West' },
    { row: 72, col: 1, value: 1 },
    { row: 73, col: 0, value: 'East' },
    { row: 73, col: 1, value: 2 },
  ]);
  await selectRangeAndSetValues(page, { r0: 71, c0: 0, r1: 71, c1: 0 }, []);

  await page.getByRole('tab', { name: 'Data', exact: true }).click();
  await page.locator('[data-ribbon-command="filter"]').click();
  await page.locator('#menu-sort [data-sort="filter"]').click();

  await expect(page.locator('.fc-filter-dropdown')).toBeVisible();
  await expect
    .poll(() => readFilterSummary(page))
    .toMatchObject({
      filterRange: { sheet: 0, r0: 70, c0: 0, r1: 73, c1: 1 },
    });

  await page
    .locator('.fc-filter-dropdown__row')
    .filter({ hasText: 'West' })
    .locator('input')
    .uncheck();
  await page.locator('.fc-filter-dropdown__apply').click();
  await expect
    .poll(() => readFilterSummary(page))
    .toMatchObject({
      hiddenRows: [72],
      filterCriteria: [
        {
          range: { sheet: 0, r0: 70, c0: 0, r1: 73, c1: 1 },
          byCol: 0,
          hiddenValues: ['West'],
        },
      ],
    });

  await page.locator('[data-ribbon-command="filter"]').click();
  await page.locator('#menu-sort [data-sort="filter-clear"]').click();
  await expect
    .poll(() => readFilterSummary(page))
    .toMatchObject({ filterRange: null, filterCriteria: [], hiddenRows: [] });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readFilterSummary(page))
    .toMatchObject({
      hiddenRows: [72],
      filterRange: { sheet: 0, r0: 70, c0: 0, r1: 73, c1: 1 },
      filterCriteria: [
        {
          range: { sheet: 0, r0: 70, c0: 0, r1: 73, c1: 1 },
          byCol: 0,
          hiddenValues: ['West'],
        },
      ],
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readFilterSummary(page))
    .toMatchObject({ filterRange: null, filterCriteria: [], hiddenRows: [] });

  await selectRangeAndSetValues(page, { r0: 90, c0: 0, r1: 93, c1: 1 }, [
    { row: 90, col: 0, value: 'Region' },
    { row: 90, col: 1, value: 'Qty' },
    { row: 91, col: 0, value: 'East' },
    { row: 91, col: 1, value: 3 },
    { row: 92, col: 0, value: 'West' },
    { row: 92, col: 1, value: 1 },
    { row: 93, col: 0, value: 'East' },
    { row: 93, col: 1, value: 2 },
    { row: 95, col: 0, value: 'Region' },
    { row: 96, col: 0, value: 'East' },
  ]);
  await page.locator('[data-ribbon-command="filter"]').click();
  await page.locator('#menu-sort [data-sort="filter-advanced"]').click();
  const advancedFilter = page.getByRole('dialog', { name: 'Advanced Filter' });
  await expect(advancedFilter).toBeVisible();
  await advancedFilter.getByRole('textbox', { name: 'List range' }).fill('A91:B94');
  await advancedFilter.getByRole('textbox', { name: 'Criteria range' }).fill('A96:A97');
  await advancedFilter.getByRole('textbox', { name: 'Copy to (optional)' }).fill('D91');
  await advancedFilter.getByRole('checkbox', { name: 'Unique records only' }).check();
  await advancedFilter.getByRole('button', { name: 'OK', exact: true }).click();
  await expect.poll(() => readCellText(page, 90, 3)).toBe('Region');
  await expect.poll(() => readCellText(page, 90, 4)).toBe('Qty');
  await expect.poll(() => readCellText(page, 91, 3)).toBe('East');
  await expect.poll(() => readCellText(page, 92, 3)).toBe('East');
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 90, 3)).toMatchObject({ kind: 'blank' });
  await expect.poll(() => readCellSummary(page, 90, 4)).toMatchObject({ kind: 'blank' });
  await expect.poll(() => readCellSummary(page, 91, 3)).toMatchObject({ kind: 'blank' });
  await expect.poll(() => readCellSummary(page, 92, 3)).toMatchObject({ kind: 'blank' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 90, 3)).toBe('Region');
  await expect.poll(() => readCellText(page, 90, 4)).toBe('Qty');
  await expect.poll(() => readCellText(page, 91, 3)).toBe('East');
  await expect.poll(() => readCellText(page, 92, 3)).toBe('East');
});

test('R02l: Calculation Options menu persists workbook calc mode', async ({ page }) => {
  await mount(page, '/?locale=en');

  await page.getByRole('tab', { name: 'Formulas', exact: true }).click();
  await page.locator('[data-ribbon-command="calcOptions"]').click();
  await expect(page.locator('#menu-calc-options')).toBeVisible();
  await page.locator('#menu-calc-options [data-calc-option="manual"]').click();
  await expect
    .poll(() =>
      page.evaluate(() => {
        const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
          | { workbook: { calcMode: () => 0 | 1 | 2 | null } }
          | undefined;
        return inst?.workbook.calcMode() ?? null;
      }),
    )
    .toBe(1);

  await page.locator('[data-ribbon-command="calcOptions"]').click();
  await expect(page.locator('#menu-calc-options [data-calc-option="manual"]')).toHaveAttribute(
    'aria-checked',
    'true',
  );
  await page.locator('#menu-calc-options [data-calc-option="auto-no-table"]').click();
  await expect
    .poll(() =>
      page.evaluate(() => {
        const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
          | { workbook: { calcMode: () => 0 | 1 | 2 | null } }
          | undefined;
        return inst?.workbook.calcMode() ?? null;
      }),
    )
    .toBe(2);

  await page.locator('[data-ribbon-command="calcOptions"]').click();
  await page.locator('#menu-calc-options [data-calc-option="iterative"]').click();
  await expect(page.locator('.fc-iterdlg')).toBeVisible();
  await closeDialog(page);
});

test('R02m: Data Validation creates list rules and enforces allowed values', async ({ page }) => {
  await mount(page, '/?locale=en');

  await selectRangeAndSetValues(page, { r0: 5, c0: 0, r1: 5, c1: 0 }, []);
  await page.getByRole('tab', { name: 'Data', exact: true }).click();
  await page.locator('[data-ribbon-command="dataValidation"]').click();
  await expect(page.locator('#menu-data-validation')).toBeVisible();
  await expect(page.locator('#menu-data-validation')).toContainText('Data Validation...');
  await expect(page.locator('#menu-data-validation')).toContainText('Circle Invalid Data');
  await expect(page.locator('#menu-data-validation')).toContainText('Clear Validation Circles');
  await page.locator('#menu-data-validation [data-validation-action="settings"]').click();

  const formatDialog = page.getByRole('dialog', { name: 'Format Cells' });
  await expect(formatDialog).toBeVisible();
  await formatDialog.locator('select[aria-label="Kind"]').selectOption('list', { force: true });
  await formatDialog.getByRole('textbox', { name: 'Literal values' }).fill('Open\nClosed\nHold');
  await formatDialog.getByRole('textbox', { name: 'Input title' }).fill('Status');
  await formatDialog
    .getByRole('textbox', { name: 'Input message' })
    .fill('Choose a workflow state.');
  await formatDialog.getByRole('textbox', { name: 'Error title' }).fill('Invalid status');
  await formatDialog
    .getByRole('textbox', { name: 'Error message' })
    .fill('Use Open, Closed, or Hold.');
  await formatDialog.getByRole('button', { name: 'OK', exact: true }).click();

  await expect
    .poll(() => readActiveValidation(page))
    .toMatchObject({
      kind: 'list',
      source: ['Open', 'Closed', 'Hold'],
      promptTitle: 'Status',
      promptMessage: 'Choose a workflow state.',
      errorTitle: 'Invalid status',
      errorMessage: 'Use Open, Closed, or Hold.',
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveValidation(page)).toBeNull();
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readActiveValidation(page))
    .toMatchObject({
      kind: 'list',
      source: ['Open', 'Closed', 'Hold'],
      promptTitle: 'Status',
      promptMessage: 'Choose a workflow state.',
      errorTitle: 'Invalid status',
      errorMessage: 'Use Open, Closed, or Hold.',
    });

  await expect(page.locator('.fc-validation-prompt')).toContainText('Status');
  await expect(page.locator('.fc-validation-prompt')).toContainText('Choose a workflow state.');

  const beforeInvalid = await readCellSummary(page, 5, 0);
  await page.locator('.fc-host__formulabar-input').fill('Invalid');
  await page.locator('.fc-host__formulabar-input').press('Enter');
  await expect(page.getByRole('dialog', { name: 'Invalid status' })).toBeVisible();
  await expect(page.getByRole('dialog', { name: 'Invalid status' })).toContainText(
    'Use Open, Closed, or Hold.',
  );
  await expect.poll(() => readCellSummary(page, 5, 0)).toEqual(beforeInvalid);
  await page.getByRole('button', { name: 'OK', exact: true }).click();

  await page.locator('.fc-host__formulabar-input').fill('Closed');
  await page.locator('.fc-host__formulabar-input').press('Enter');
  await expect.poll(() => readCellText(page, 5, 0)).toBe('Closed');

  await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            setState: (
              updater: (state: {
                data: { cells: Map<string, { value: unknown; formula: unknown }> };
              }) => { data: { cells: Map<string, { value: unknown; formula: unknown }> } },
            ) => void;
          };
        }
      | undefined;
    inst?.store.setState((state) => {
      const cells = new Map(state.data.cells);
      cells.set('0:5:0', { value: { kind: 'text', value: 'Invalid' }, formula: null });
      return { ...state, data: { ...state.data, cells } };
    });
  });
  await page.locator('[data-ribbon-command="dataValidation"]').click();
  await page.locator('#menu-data-validation [data-validation-action="circle-invalid"]').click();
  await expect.poll(() => readValidationCircles(page)).toEqual(['0:5:0']);
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readValidationCircles(page)).toEqual([]);
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readValidationCircles(page)).toEqual(['0:5:0']);

  await page.locator('[data-ribbon-command="dataValidation"]').click();
  await page.locator('#menu-data-validation [data-validation-action="clear-circles"]').click();
  await expect.poll(() => readValidationCircles(page)).toEqual([]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readValidationCircles(page)).toEqual(['0:5:0']);
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readValidationCircles(page)).toEqual([]);

  await selectCellAndSetText(page, 5, 0, 'Invalid');
  await page.locator('[data-ribbon-command="dataValidation"]').click();
  await page.locator('#menu-data-validation [data-validation-action="clear-rules"]').click();
  await expect.poll(() => readActiveValidation(page)).toBeNull();
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readActiveValidation(page))
    .toMatchObject({
      kind: 'list',
      source: ['Open', 'Closed', 'Hold'],
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveValidation(page)).toBeNull();
});

test('R02m-whole: Data Validation creates whole-number bounds and blocks invalid input', async ({
  page,
}) => {
  await mount(page, '/?locale=en&fixture=empty');

  await selectRangeAndSetValues(page, { r0: 6, c0: 0, r1: 6, c1: 0 }, []);
  await page.getByRole('tab', { name: 'Data', exact: true }).click();
  await page.locator('[data-ribbon-command="dataValidation"]').click();
  await page.locator('#menu-data-validation [data-validation-action="settings"]').click();

  const dialog = page.getByRole('dialog', { name: 'Format Cells' });
  await expect(dialog).toBeVisible();
  await dialog.locator('select[aria-label="Kind"]').selectOption('whole', { force: true });
  await dialog.locator('select[aria-label="Condition"]').selectOption('between', { force: true });
  await dialog.getByRole('spinbutton', { name: 'Value', exact: true }).fill('1');
  await dialog.getByRole('spinbutton', { name: 'Upper value' }).fill('10');
  await dialog.getByRole('textbox', { name: 'Error title' }).fill('Quantity out of range');
  await dialog
    .getByRole('textbox', { name: 'Error message' })
    .fill('Use a whole number from 1 to 10.');
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();

  await expect
    .poll(() => readActiveValidation(page))
    .toMatchObject({
      kind: 'whole',
      op: 'between',
      a: 1,
      b: 10,
      errorTitle: 'Quantity out of range',
      errorMessage: 'Use a whole number from 1 to 10.',
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveValidation(page)).toBeNull();
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readActiveValidation(page))
    .toMatchObject({ kind: 'whole', op: 'between', a: 1, b: 10 });

  const beforeInvalid = await readCellSummary(page, 6, 0);
  await page.locator('.fc-host__formulabar-input').fill('12');
  await page.locator('.fc-host__formulabar-input').press('Enter');
  await expect(page.getByRole('dialog', { name: 'Quantity out of range' })).toBeVisible();
  await expect(page.getByRole('dialog', { name: 'Quantity out of range' })).toContainText(
    'Use a whole number from 1 to 10.',
  );
  await expect.poll(() => readCellSummary(page, 6, 0)).toEqual(beforeInvalid);
  await page.getByRole('button', { name: 'OK', exact: true }).click();

  await page.locator('.fc-host__formulabar-input').fill('7');
  await page.locator('.fc-host__formulabar-input').press('Enter');
  await expect
    .poll(() => readCellSummary(page, 6, 0))
    .toMatchObject({
      kind: 'number',
      value: 7,
    });
});

test('R02m-text-length: Data Validation enforces text-length rules', async ({ page }) => {
  await mount(page, '/?locale=en&fixture=empty');

  await selectRangeAndSetValues(page, { r0: 7, c0: 0, r1: 7, c1: 0 }, []);
  await page.getByRole('tab', { name: 'Data', exact: true }).click();
  await page.locator('[data-ribbon-command="dataValidation"]').click();
  await page.locator('#menu-data-validation [data-validation-action="settings"]').click();

  const dialog = page.getByRole('dialog', { name: 'Format Cells' });
  await expect(dialog).toBeVisible();
  await dialog.locator('select[aria-label="Kind"]').selectOption('textLength', { force: true });
  await dialog.locator('select[aria-label="Condition"]').selectOption('<=', { force: true });
  await dialog.getByRole('spinbutton', { name: 'Value', exact: true }).fill('5');
  await dialog.getByRole('textbox', { name: 'Error title' }).fill('Text too long');
  await dialog
    .getByRole('textbox', { name: 'Error message' })
    .fill('Use five characters or fewer.');
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();

  await expect
    .poll(() => readActiveValidation(page))
    .toMatchObject({
      kind: 'textLength',
      op: '<=',
      a: 5,
      errorTitle: 'Text too long',
      errorMessage: 'Use five characters or fewer.',
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveValidation(page)).toBeNull();
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readActiveValidation(page))
    .toMatchObject({ kind: 'textLength', op: '<=', a: 5 });

  const beforeInvalid = await readCellSummary(page, 7, 0);
  await page.locator('.fc-host__formulabar-input').fill('abcdef');
  await page.locator('.fc-host__formulabar-input').press('Enter');
  await expect(page.getByRole('dialog', { name: 'Text too long' })).toBeVisible();
  await expect(page.getByRole('dialog', { name: 'Text too long' })).toContainText(
    'Use five characters or fewer.',
  );
  await expect.poll(() => readCellSummary(page, 7, 0)).toEqual(beforeInvalid);
  await page.getByRole('button', { name: 'OK', exact: true }).click();

  await page.locator('.fc-host__formulabar-input').fill('abcde');
  await page.locator('.fc-host__formulabar-input').press('Enter');
  await expect.poll(() => readCellText(page, 7, 0)).toBe('abcde');
});

test('R02n: Data outline commands group, collapse, expand, and ungroup rows and columns', async ({
  page,
}) => {
  await mount(page, '/?locale=en');

  await page.getByRole('tab', { name: 'Data', exact: true }).click();

  await selectRangeAndSetValues(page, { r0: 10, c0: 0, r1: 12, c1: 0 }, []);
  await page.locator('[data-ribbon-command="outlineGroup"]').click();
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({
      outlineRows: [
        [10, 1],
        [11, 1],
        [12, 1],
      ],
      outlineRowGutter: 14,
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({ outlineRows: [], outlineRowGutter: 0 });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({
      outlineRows: [
        [10, 1],
        [11, 1],
        [12, 1],
      ],
      outlineRowGutter: 14,
    });

  await selectRangeAndSetValues(page, { r0: 11, c0: 0, r1: 11, c1: 0 }, []);
  await page.locator('[data-ribbon-command="outlineHideDetail"]').click();
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenRows: [10, 11, 12] });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenRows: [] });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenRows: [10, 11, 12] });

  await page.locator('[data-ribbon-command="outlineShowDetail"]').click();
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenRows: [] });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenRows: [10, 11, 12] });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenRows: [] });

  await selectRangeAndSetValues(page, { r0: 10, c0: 0, r1: 12, c1: 0 }, []);
  await page.locator('[data-ribbon-command="outlineUngroup"]').click();
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({ outlineRows: [], outlineRowGutter: 0 });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({
      outlineRows: [
        [10, 1],
        [11, 1],
        [12, 1],
      ],
      outlineRowGutter: 14,
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({ outlineRows: [], outlineRowGutter: 0 });

  await selectRangeAndSetValues(page, { r0: 20, c0: 0, r1: 24, c1: 0 }, []);
  await page.locator('[data-ribbon-command="outlineGroup"]').click();
  await selectRangeAndSetValues(page, { r0: 21, c0: 0, r1: 23, c1: 0 }, []);
  await page.locator('[data-ribbon-command="outlineGroup"]').click();
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({
      outlineRows: [
        [20, 1],
        [21, 2],
        [22, 2],
        [23, 2],
        [24, 1],
      ],
      outlineRowGutter: 28,
    });

  await page.locator('[data-ribbon-command="outlineUngroup"]').click();
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({
      outlineRows: [
        [20, 1],
        [21, 1],
        [22, 1],
        [23, 1],
        [24, 1],
      ],
      outlineRowGutter: 14,
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({
      outlineRows: [
        [20, 1],
        [21, 2],
        [22, 2],
        [23, 2],
        [24, 1],
      ],
      outlineRowGutter: 28,
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({
      outlineRows: [
        [20, 1],
        [21, 1],
        [22, 1],
        [23, 1],
        [24, 1],
      ],
      outlineRowGutter: 14,
    });

  await selectRangeAndSetValues(page, { r0: 0, c0: 2, r1: 0, c1: 4 }, []);
  await page.locator('[data-ribbon-command="outlineGroup"]').click();
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({
      outlineCols: [
        [2, 1],
        [3, 1],
        [4, 1],
      ],
      outlineColGutter: 14,
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({ outlineCols: [], outlineColGutter: 0 });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({
      outlineCols: [
        [2, 1],
        [3, 1],
        [4, 1],
      ],
      outlineColGutter: 14,
    });

  await selectRangeAndSetValues(page, { r0: 0, c0: 3, r1: 0, c1: 3 }, []);
  await page.locator('[data-ribbon-command="outlineHideDetail"]').click();
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenCols: [2, 3, 4] });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenCols: [] });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenCols: [2, 3, 4] });

  await page.locator('[data-ribbon-command="outlineShowDetail"]').click();
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenCols: [] });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenCols: [2, 3, 4] });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenCols: [] });

  await selectRangeAndSetValues(page, { r0: 0, c0: 2, r1: 0, c1: 4 }, []);
  await page.locator('[data-ribbon-command="outlineUngroup"]').click();
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({ outlineCols: [], outlineColGutter: 0 });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({
      outlineCols: [
        [2, 1],
        [3, 1],
        [4, 1],
      ],
      outlineColGutter: 14,
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({ outlineCols: [], outlineColGutter: 0 });
});

test('R02o: Review comment commands create, navigate, and delete notes', async ({ page }) => {
  await mount(page, '/?locale=en');

  await selectCellAndSetText(page, 60, 0, 'review anchor');
  await page.getByRole('tab', { name: 'Review', exact: true }).click();
  await page.locator('[data-ribbon-command="newCommentReview"]').click();
  const note = page.locator('.fc-cmtnote');
  await expect(note).toBeVisible();
  await note.locator('.fc-cmtnote__textarea').fill('First review note');
  await note.getByRole('button', { name: 'OK', exact: true }).click();
  await expect
    .poll(() => readCommentSummaries(page))
    .toEqual([{ addr: { sheet: 0, row: 60, col: 0 }, text: 'First review note' }]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCommentSummaries(page)).toEqual([]);
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCommentSummaries(page))
    .toEqual([{ addr: { sheet: 0, row: 60, col: 0 }, text: 'First review note' }]);

  await setCommentDirect(page, 61, 1, 'Second review note');
  await expect
    .poll(() => readCommentSummaries(page))
    .toEqual([
      { addr: { sheet: 0, row: 60, col: 0 }, text: 'First review note' },
      { addr: { sheet: 0, row: 61, col: 1 }, text: 'Second review note' },
    ]);

  await selectRangeAndSetValues(page, { r0: 0, c0: 0, r1: 0, c1: 0 }, []);
  await page.locator('[data-ribbon-command="nextCommentReview"]').click();
  await expect
    .poll(() => readSelectionSummary(page))
    .toMatchObject({
      active: { sheet: 0, row: 60, col: 0 },
    });
  await expect(page.locator('.fc-cmtnote')).toBeVisible();
  await expect(page.locator('.fc-cmtnote__textarea')).toHaveValue('First review note');
  await page.locator('.fc-cmtnote').getByRole('button', { name: 'Cancel' }).click();

  await page.locator('[data-ribbon-command="nextCommentReview"]').click();
  await expect
    .poll(() => readSelectionSummary(page))
    .toMatchObject({
      active: { sheet: 0, row: 61, col: 1 },
    });
  await expect(page.locator('.fc-cmtnote')).toBeVisible();
  await expect(page.locator('.fc-cmtnote__textarea')).toHaveValue('Second review note');
  await page.locator('.fc-cmtnote').getByRole('button', { name: 'Cancel' }).click();

  await page.locator('[data-ribbon-command="previousCommentReview"]').click();
  await expect
    .poll(() => readSelectionSummary(page))
    .toMatchObject({
      active: { sheet: 0, row: 60, col: 0 },
    });
  await expect(page.locator('.fc-cmtnote')).toBeVisible();
  await expect(page.locator('.fc-cmtnote__textarea')).toHaveValue('First review note');
  await page.locator('.fc-cmtnote').getByRole('button', { name: 'Cancel' }).click();

  await page.locator('[data-ribbon-command="deleteCommentReview"]').click();
  await expect(page.locator('#menu-review-comments')).toBeVisible();
  await page.locator('#menu-review-comments [data-comment-action="delete-active"]').click();
  await expect
    .poll(() => readCommentSummaries(page))
    .toEqual([{ addr: { sheet: 0, row: 61, col: 1 }, text: 'Second review note' }]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCommentSummaries(page))
    .toEqual([
      { addr: { sheet: 0, row: 60, col: 0 }, text: 'First review note' },
      { addr: { sheet: 0, row: 61, col: 1 }, text: 'Second review note' },
    ]);
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCommentSummaries(page))
    .toEqual([{ addr: { sheet: 0, row: 61, col: 1 }, text: 'Second review note' }]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCommentSummaries(page))
    .toEqual([
      { addr: { sheet: 0, row: 60, col: 0 }, text: 'First review note' },
      { addr: { sheet: 0, row: 61, col: 1 }, text: 'Second review note' },
    ]);

  await page.locator('[data-ribbon-command="deleteCommentReview"]').click();
  await page.locator('#menu-review-comments [data-comment-action="delete-all"]').click();
  await expect.poll(() => readCommentSummaries(page)).toEqual([]);
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCommentSummaries(page))
    .toEqual([
      { addr: { sheet: 0, row: 60, col: 0 }, text: 'First review note' },
      { addr: { sheet: 0, row: 61, col: 1 }, text: 'Second review note' },
    ]);
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCommentSummaries(page)).toEqual([]);
});

test('R02p: Review Protect Sheet command toggles password-protected state', async ({ page }) => {
  await mount(page, '/?locale=en');

  await page.getByRole('tab', { name: 'Review', exact: true }).click();
  await selectCellAndSetText(page, 70, 0, 'lock target');
  await page.locator('[data-ribbon-command="protectReview"]').click();
  await expect(page.locator('#menu-protect-review')).toBeVisible();
  await page.locator('#menu-protect-review [data-protect-action="unlock-cell"]').click();
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ locked: false });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellLocked(page)).toBeNull();
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellLocked(page)).toBe(false);

  await page.locator('[data-ribbon-command="protectReview"]').click();
  await page.locator('#menu-protect-review [data-protect-action="lock-cell"]').click();
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ locked: true });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellLocked(page)).toBe(false);
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readActiveCellLocked(page)).toBe(true);

  await page.locator('[data-ribbon-command="protectReview"]').click();
  await page.locator('#menu-protect-review [data-protect-action="allow-edit-ranges"]').click();
  const allowRangesDialog = page.getByRole('dialog', { name: 'Allow Users to Edit Ranges' });
  await expect(allowRangesDialog).toBeVisible();
  await expect(allowRangesDialog.locator('.app__dlg__input')).toHaveValue('A71');
  await allowRangesDialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect
    .poll(() => readProtectionSummary(page))
    .toMatchObject({
      allowedEditRanges: [{ title: 'A71', range: { sheet: 0, r0: 70, c0: 0, r1: 70, c1: 0 } }],
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readProtectionSummary(page)).toMatchObject({ allowedEditRanges: [] });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readProtectionSummary(page))
    .toMatchObject({
      allowedEditRanges: [{ title: 'A71', range: { sheet: 0, r0: 70, c0: 0, r1: 70, c1: 0 } }],
    });

  await page.locator('[data-ribbon-command="protectReview"]').click();
  await page.locator('#menu-protect-review [data-protect-action="protect-sheet"]').click();
  const protectDialog = page.getByRole('dialog', { name: 'Protect Sheet' });
  await expect(protectDialog).toBeVisible();
  await protectDialog.locator('.app__dlg__input').fill('secret');
  await protectDialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect
    .poll(() => readProtectionSummary(page))
    .toMatchObject({ protected: true, password: 'secret' });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readProtectionSummary(page)).toMatchObject({ protected: false });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readProtectionSummary(page))
    .toMatchObject({ protected: true, password: 'secret' });
  await expect.poll(() => isCellWritableDirect(page, 70, 0)).toBe(true);
  await expect.poll(() => isCellWritableDirect(page, 70, 1)).toBe(false);

  await page.locator('[data-ribbon-command="protectReview"]').click();
  await page.locator('#menu-protect-review [data-protect-action="unprotect-sheet"]').click();
  const wrongDialog = page.getByRole('dialog', { name: 'Unprotect Sheet' });
  await expect(wrongDialog).toBeVisible();
  await wrongDialog.locator('.app__dlg__input').fill('wrong');
  await wrongDialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(page.getByRole('alertdialog', { name: 'Unprotect Sheet' })).toBeVisible();
  await page
    .getByRole('alertdialog', { name: 'Unprotect Sheet' })
    .getByRole('button', { name: 'OK', exact: true })
    .click();
  await expect
    .poll(() => readProtectionSummary(page))
    .toMatchObject({ protected: true, password: 'secret' });

  await page.locator('[data-ribbon-command="protectReview"]').click();
  await page.locator('#menu-protect-review [data-protect-action="unprotect-sheet"]').click();
  const unprotectDialog = page.getByRole('dialog', { name: 'Unprotect Sheet' });
  await expect(unprotectDialog).toBeVisible();
  await unprotectDialog.locator('.app__dlg__input').fill('secret');
  await unprotectDialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect.poll(() => readProtectionSummary(page)).toMatchObject({ protected: false });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readProtectionSummary(page))
    .toMatchObject({ protected: true, password: 'secret' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readProtectionSummary(page)).toMatchObject({ protected: false });

  await page.locator('[data-ribbon-command="protectReview"]').click();
  await page
    .locator('#menu-protect-review [data-protect-action="clear-allowed-edit-ranges"]')
    .click();
  await expect.poll(() => readProtectionSummary(page)).toMatchObject({ allowedEditRanges: [] });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readProtectionSummary(page))
    .toMatchObject({
      allowedEditRanges: [{ title: 'A71', range: { sheet: 0, r0: 70, c0: 0, r1: 70, c1: 0 } }],
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readProtectionSummary(page)).toMatchObject({ allowedEditRanges: [] });

  await page.locator('[data-ribbon-command="protectionReview"]').click();
  const directAllowRangesDialog = page.getByRole('dialog', { name: 'Allow Users to Edit Ranges' });
  await expect(directAllowRangesDialog).toBeVisible();
  await expect(directAllowRangesDialog.locator('.app__dlg__input')).toHaveValue('A71');
  await directAllowRangesDialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect
    .poll(() => readProtectionSummary(page))
    .toMatchObject({
      allowedEditRanges: [{ title: 'A71', range: { sheet: 0, r0: 70, c0: 0, r1: 70, c1: 0 } }],
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readProtectionSummary(page)).toMatchObject({ allowedEditRanges: [] });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readProtectionSummary(page))
    .toMatchObject({
      allowedEditRanges: [{ title: 'A71', range: { sheet: 0, r0: 70, c0: 0, r1: 70, c1: 0 } }],
    });

  const beforeSheets = await readSheetCount(page);
  await page.locator('[data-ribbon-command="protectReview"]').click();
  await page.locator('#menu-protect-review [data-protect-action="protect-workbook"]').click();
  const protectWorkbookDialog = page.getByRole('dialog', { name: 'Protect Workbook' });
  await expect(protectWorkbookDialog).toBeVisible();
  await protectWorkbookDialog.locator('.app__dlg__input').fill('book');
  await protectWorkbookDialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect
    .poll(() => readProtectionSummary(page))
    .toMatchObject({
      workbookStructureProtected: true,
      workbookPassword: 'book',
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readProtectionSummary(page))
    .toMatchObject({ workbookStructureProtected: false });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readProtectionSummary(page))
    .toMatchObject({
      workbookStructureProtected: true,
      workbookPassword: 'book',
    });

  await page.locator('#btn-sheet-add').click();
  await expect.poll(() => readSheetCount(page)).toBe(beforeSheets);

  await page.locator('[data-ribbon-command="protectReview"]').click();
  await page.locator('#menu-protect-review [data-protect-action="unprotect-workbook"]').click();
  const unprotectWorkbookDialog = page.getByRole('dialog', { name: 'Unprotect Workbook' });
  await expect(unprotectWorkbookDialog).toBeVisible();
  await unprotectWorkbookDialog.locator('.app__dlg__input').fill('book');
  await unprotectWorkbookDialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect
    .poll(() => readProtectionSummary(page))
    .toMatchObject({
      workbookStructureProtected: false,
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readProtectionSummary(page))
    .toMatchObject({
      workbookStructureProtected: true,
      workbookPassword: 'book',
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readProtectionSummary(page))
    .toMatchObject({
      workbookStructureProtected: false,
    });

  await page.locator('#btn-sheet-add').click();
  await expect.poll(() => readSheetCount(page)).toBe(beforeSheets + 1);

  await page.locator('[data-ribbon-command="protectWorkbookReview"]').click();
  const directProtectWorkbookDialog = page.getByRole('dialog', { name: 'Protect Workbook' });
  await expect(directProtectWorkbookDialog).toBeVisible();
  await directProtectWorkbookDialog.locator('.app__dlg__input').fill('direct-book');
  await directProtectWorkbookDialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect
    .poll(() => readProtectionSummary(page))
    .toMatchObject({
      workbookStructureProtected: true,
      workbookPassword: 'direct-book',
    });
  await page.locator('#btn-sheet-add').click();
  await expect.poll(() => readSheetCount(page)).toBe(beforeSheets + 1);

  await page.locator('[data-ribbon-command="protectWorkbookReview"]').click();
  const directUnprotectWorkbookDialog = page.getByRole('dialog', { name: 'Unprotect Workbook' });
  await expect(directUnprotectWorkbookDialog).toBeVisible();
  await directUnprotectWorkbookDialog.locator('.app__dlg__input').fill('direct-book');
  await directUnprotectWorkbookDialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect
    .poll(() => readProtectionSummary(page))
    .toMatchObject({
      workbookStructureProtected: false,
    });
  await page.locator('#btn-sheet-add').click();
  await expect.poll(() => readSheetCount(page)).toBe(beforeSheets + 2);
});

test('R02q: View ribbon commands update workbook view, toggles, and zoom state', async ({
  page,
}) => {
  await mount(page, '/?locale=en');

  await page.getByRole('tab', { name: 'View', exact: true }).click();

  await expect(page.locator('[data-ribbon-command="viewNormal"]')).toHaveAttribute(
    'aria-pressed',
    'true',
  );

  await page.locator('[data-ribbon-command="viewPageLayout"]').click();
  await expect.poll(() => readViewSummary(page)).toMatchObject({ workbookView: 'pageLayout' });
  await expect(page.locator('[data-ribbon-command="viewPageLayout"]')).toHaveAttribute(
    'aria-pressed',
    'true',
  );

  await page.locator('[data-ribbon-command="viewPageBreakPreview"]').click();
  await expect
    .poll(() => readViewSummary(page))
    .toMatchObject({ workbookView: 'pageBreakPreview' });
  await expect(page.locator('[data-ribbon-command="viewPageBreakPreview"]')).toHaveAttribute(
    'aria-pressed',
    'true',
  );

  await page.locator('[data-ribbon-command="viewNormal"]').click();
  await expect.poll(() => readViewSummary(page)).toMatchObject({ workbookView: 'normal' });

  await page.locator('[data-ribbon-command="sheetViewSave"]').click();
  const saveViewDialog = page.getByRole('dialog', { name: 'Save' });
  await expect(saveViewDialog).toBeVisible();
  await saveViewDialog.locator('.app__dlg__input').fill('Ops view');
  await saveViewDialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect
    .poll(() => readViewSummary(page))
    .toMatchObject({
      sheetViews: [{ name: 'Ops view', sheet: 0 }],
    });
  const savedView = await readViewSummary(page);
  expect(savedView.activeSheetViewId).toBeTruthy();
  await expect(page.locator('[data-ribbon-command="sheetViewSelect"]')).toContainText('Ops view');
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readViewSummary(page))
    .toMatchObject({
      sheetViews: [],
      activeSheetViewId: null,
    });
  expect(await redoViaInstance(page)).toBe(true);
  const restoredSavedView = await readViewSummary(page);
  expect(restoredSavedView.activeSheetViewId).toBeTruthy();
  await expect
    .poll(() => readViewSummary(page))
    .toMatchObject({
      sheetViews: [{ name: 'Ops view', sheet: 0 }],
    });

  await page.locator('[data-ribbon-command="sheetViewSelect"] .demo__rb-dd__btn').click();
  await page.getByRole('option', { name: 'Current view' }).click();
  await expect.poll(() => readViewSummary(page)).toMatchObject({ activeSheetViewId: null });

  await page.locator('[data-ribbon-command="sheetViewSelect"] .demo__rb-dd__btn').click();
  await page.getByRole('option', { name: 'Ops view' }).click();
  await expect
    .poll(() => readViewSummary(page))
    .toMatchObject({
      activeSheetViewId: restoredSavedView.activeSheetViewId,
    });

  await page.locator('[data-ribbon-command="sheetViewDelete"]').click();
  await expect
    .poll(() => readViewSummary(page))
    .toMatchObject({
      sheetViews: [],
      activeSheetViewId: null,
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readViewSummary(page))
    .toMatchObject({
      sheetViews: [{ name: 'Ops view', sheet: 0 }],
      activeSheetViewId: restoredSavedView.activeSheetViewId,
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readViewSummary(page))
    .toMatchObject({
      sheetViews: [],
      activeSheetViewId: null,
    });

  await page.evaluate(() => {
    const range = { sheet: 0, r0: 90, c0: 0, r1: 93, c1: 1 };
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            setState: (
              fn: (state: {
                ui: {
                  filterRange: typeof range | null;
                  filterCriteria: Array<{
                    range: typeof range;
                    byCol: number;
                    hiddenValues: string[];
                  }>;
                };
                layout: {
                  freezeRows: number;
                  freezeCols: number;
                  hiddenRows: Set<number>;
                  hiddenCols: Set<number>;
                };
                sheetViews: { activeViewId: string | null };
              }) => unknown,
            ) => void;
          };
        }
      | undefined;
    inst?.store.setState((state) => ({
      ...state,
      ui: {
        ...state.ui,
        filterRange: range,
        filterCriteria: [{ range, byCol: 0, hiddenValues: ['West'] }],
      },
      layout: {
        ...state.layout,
        freezeRows: 2,
        freezeCols: 1,
        hiddenRows: new Set([92]),
        hiddenCols: new Set([5]),
      },
    }));
  });
  await page.locator('[data-ribbon-command="sheetViewSave"]').click();
  const statefulViewDialog = page.getByRole('dialog', { name: 'Save' });
  await expect(statefulViewDialog).toBeVisible();
  await statefulViewDialog.locator('.app__dlg__input').fill('Filtered view');
  await statefulViewDialog.getByRole('button', { name: 'OK', exact: true }).click();
  const statefulView = await readViewSummary(page);
  expect(statefulView.activeSheetViewId).toBeTruthy();

  await page.locator('[data-ribbon-command="sheetViewSelect"] .demo__rb-dd__btn').click();
  await page.getByRole('option', { name: 'Current view' }).click();
  await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            setState: (
              fn: (state: {
                ui: { filterRange: unknown; filterCriteria: unknown[] };
                layout: {
                  freezeRows: number;
                  freezeCols: number;
                  hiddenRows: Set<number>;
                  hiddenCols: Set<number>;
                };
                viewport: { rowStart: number; colStart: number };
              }) => unknown,
            ) => void;
          };
        }
      | undefined;
    inst?.store.setState((state) => ({
      ...state,
      ui: { ...state.ui, filterRange: null, filterCriteria: [] },
      layout: {
        ...state.layout,
        freezeRows: 0,
        freezeCols: 0,
        hiddenRows: new Set(),
        hiddenCols: new Set(),
      },
      viewport: { ...state.viewport, rowStart: 0, colStart: 0 },
    }));
  });
  await expect
    .poll(() => readFilterSummary(page))
    .toMatchObject({
      filterRange: null,
      filterCriteria: [],
      hiddenRows: [],
    });
  await expect
    .poll(() => readLayoutSummary(page))
    .toMatchObject({
      hiddenCols: [],
      outlineRows: [],
    });

  await page.locator('[data-ribbon-command="sheetViewSelect"] .demo__rb-dd__btn').click();
  await page.getByRole('option', { name: 'Filtered view' }).click();
  await expect
    .poll(() => readFilterSummary(page))
    .toMatchObject({
      filterRange: { sheet: 0, r0: 90, c0: 0, r1: 93, c1: 1 },
      filterCriteria: [
        {
          range: { sheet: 0, r0: 90, c0: 0, r1: 93, c1: 1 },
          byCol: 0,
          hiddenValues: ['West'],
        },
      ],
      hiddenRows: [92],
    });
  await expect
    .poll(() => readViewSummary(page))
    .toMatchObject({
      activeSheetViewId: statefulView.activeSheetViewId,
      freezeRows: 2,
      freezeCols: 1,
    });
  await expect.poll(() => readLayoutSummary(page)).toMatchObject({ hiddenCols: [5] });
  await page.locator('[data-ribbon-command="sheetViewSelect"] .demo__rb-dd__btn').click();
  await page.getByRole('option', { name: 'Current view' }).click();
  await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            setState: (
              fn: (state: {
                ui: { filterRange: unknown; filterCriteria: unknown[] };
                layout: {
                  freezeRows: number;
                  freezeCols: number;
                  hiddenRows: Set<number>;
                  hiddenCols: Set<number>;
                };
                viewport: { rowStart: number; colStart: number };
                sheetViews: { activeViewId: string | null };
              }) => unknown,
            ) => void;
          };
        }
      | undefined;
    inst?.store.setState((state) => ({
      ...state,
      ui: { ...state.ui, filterRange: null, filterCriteria: [] },
      layout: {
        ...state.layout,
        freezeRows: 0,
        freezeCols: 0,
        hiddenRows: new Set(),
        hiddenCols: new Set(),
      },
      viewport: { ...state.viewport, rowStart: 0, colStart: 0 },
      sheetViews: { ...state.sheetViews, activeViewId: null },
    }));
  });
  await expect
    .poll(() => readViewSummary(page))
    .toMatchObject({ activeSheetViewId: null, freezeRows: 0, freezeCols: 0 });

  await page.locator('[data-ribbon-command="workbookObjectsView"]').click();
  await expect(page.locator('.fc-objects')).toBeVisible();
  await page.locator('.fc-objects__close').click();

  await selectRangeAndSetValues(page, { r0: 5, c0: 5, r1: 5, c1: 5 }, []);
  await page.locator('[data-ribbon-command="viewGridlines"]').click();
  await page.locator('[data-ribbon-command="viewHeadings"]').click();
  await page.locator('[data-ribbon-command="viewFormulas"]').click();
  await page.locator('[data-ribbon-command="viewR1C1"]').click();
  await expect
    .poll(() => readViewSummary(page))
    .toMatchObject({
      showGridLines: false,
      showHeaders: false,
      showFormulas: true,
      r1c1: true,
    });
  for (const command of ['viewGridlines', 'viewHeadings']) {
    await expect(page.locator(`[data-ribbon-command="${command}"]`)).toHaveAttribute(
      'aria-pressed',
      'false',
    );
  }
  for (const command of ['viewFormulas', 'viewR1C1']) {
    await expect(page.locator(`[data-ribbon-command="${command}"]`)).toHaveAttribute(
      'aria-pressed',
      'true',
    );
  }
  await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          workbook: {
            setNumber: (addr: CellAddr, value: number) => void;
            setFormula: (addr: CellAddr, formula: string) => void;
            cells: (sheet: number) => Iterable<{
              addr: CellAddr;
              value: unknown;
              formula: string | null;
            }>;
            recalc: () => void;
          };
          store: {
            setState: (
              fn: (state: {
                data: { cells: Map<string, unknown> };
                selection: {
                  active: CellAddr;
                  anchor: CellAddr;
                  range: { sheet: number; r0: number; c0: number; r1: number; c1: number };
                  extraRanges?: unknown[];
                };
              }) => unknown,
            ) => void;
          };
        }
      | undefined;
    if (!inst) return;
    inst.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 3);
    inst.workbook.setNumber({ sheet: 0, row: 0, col: 1 }, 4);
    inst.workbook.setFormula({ sheet: 0, row: 3, col: 3 }, '=A1+B1');
    inst.workbook.recalc();
    const cells = new Map<string, { value: unknown; formula: string | null }>();
    for (const cell of inst.workbook.cells(0)) {
      cells.set(`${cell.addr.sheet}:${cell.addr.row}:${cell.addr.col}`, {
        value: cell.value,
        formula: cell.formula,
      });
    }
    const active = { sheet: 0, row: 3, col: 3 };
    inst.store.setState((state) => ({
      ...state,
      data: { ...state.data, cells },
      selection: {
        ...state.selection,
        active,
        anchor: active,
        range: { sheet: 0, r0: 3, c0: 3, r1: 3, c1: 3 },
        extraRanges: [],
      },
    }));
  });
  await expect(page.locator('.fc-host__formulabar-input')).toHaveValue('=R[-3]C[-3]+R[-3]C[-2]');
  await page.locator('.fc-host canvas').click({ position: { x: 10, y: 10 } });
  await expect
    .poll(() => readSelectionSummary(page))
    .toMatchObject({
      active: { sheet: 0, row: 0, col: 0 },
    });

  await expect.poll(() => readViewSummary(page)).toMatchObject({ formulaBarAttached: true });
  await page.locator('[data-ribbon-command="viewFormulaBar"]').click();
  await expect.poll(() => readViewSummary(page)).toMatchObject({ formulaBarAttached: false });
  await expect(page.locator('[data-ribbon-command="viewFormulaBar"]')).toHaveAttribute(
    'aria-pressed',
    'false',
  );

  await page.locator('[data-ribbon-command="freeze"]').click();
  await expect(page.locator('#menu-freeze')).toBeVisible();
  await page.locator('#menu-freeze [data-freeze="row"]').click();
  await expect.poll(() => readViewSummary(page)).toMatchObject({ freezeRows: 1, freezeCols: 0 });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readViewSummary(page)).toMatchObject({ freezeRows: 0, freezeCols: 0 });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readViewSummary(page)).toMatchObject({ freezeRows: 1, freezeCols: 0 });

  await page.locator('[data-ribbon-command="freeze"]').click();
  await page.locator('#menu-freeze [data-freeze="col"]').click();
  await expect.poll(() => readViewSummary(page)).toMatchObject({ freezeRows: 0, freezeCols: 1 });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readViewSummary(page)).toMatchObject({ freezeRows: 1, freezeCols: 0 });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readViewSummary(page)).toMatchObject({ freezeRows: 0, freezeCols: 1 });

  await selectRangeAndSetValues(page, { r0: 4, c0: 3, r1: 4, c1: 3 }, []);
  await page.locator('[data-ribbon-command="freeze"]').click();
  await page.locator('#menu-freeze [data-freeze="selection"]').click();
  await expect.poll(() => readViewSummary(page)).toMatchObject({ freezeRows: 4, freezeCols: 3 });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readViewSummary(page)).toMatchObject({ freezeRows: 0, freezeCols: 1 });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readViewSummary(page)).toMatchObject({ freezeRows: 4, freezeCols: 3 });

  await page.locator('[data-ribbon-command="freeze"]').click();
  await page.locator('#menu-freeze [data-freeze="off"]').click();
  await expect.poll(() => readViewSummary(page)).toMatchObject({ freezeRows: 0, freezeCols: 0 });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readViewSummary(page)).toMatchObject({ freezeRows: 4, freezeCols: 3 });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readViewSummary(page)).toMatchObject({ freezeRows: 0, freezeCols: 0 });

  await selectRangeAndSetValues(page, { r0: 0, c0: 0, r1: 39, c1: 19 }, []);
  await page.locator('[data-ribbon-command="zoomSelection"]').click();
  await expect.poll(() => readViewSummary(page)).toMatchObject({ zoom: 0.5 });

  await page.locator('[data-ribbon-command="zoom125"]').click();
  await expect.poll(() => readViewSummary(page)).toMatchObject({ zoom: 1.25 });
  await page.locator('[data-ribbon-command="zoomDialog"]').click();
  const zoomDialog = page.getByRole('dialog', { name: 'Zoom' });
  await expect(zoomDialog).toBeVisible();
  await expect(zoomDialog.locator('input[type="number"]')).toHaveValue('125');
  await zoomDialog.locator('input[type="number"]').fill('150');
  await zoomDialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect.poll(() => readViewSummary(page)).toMatchObject({ zoom: 1.5 });
  await page.locator('[data-ribbon-command="zoom100"]').click();
  await expect.poll(() => readViewSummary(page)).toMatchObject({ zoom: 1 });

  await page.locator('[data-ribbon-command="windowVisibility"]').click();
  const windowDialog = page.getByRole('dialog', { name: 'Format Cells' });
  await expect(windowDialog).toBeVisible();
  await windowDialog.locator('.fc-fmtdlg__footer').getByRole('button', { name: 'Cancel' }).click();
});

test('R02r: Home style commands create undoable tables and apply cell styles', async ({ page }) => {
  await mount(page, '/?locale=en');

  await page.getByRole('tab', { name: 'Home', exact: true }).click();

  await selectRangeAndSetValues(page, { r0: 80, c0: 0, r1: 82, c1: 1 }, [
    { row: 80, col: 0, value: 'Item' },
    { row: 80, col: 1, value: 'Qty' },
    { row: 81, col: 0, value: 'Paper' },
    { row: 81, col: 1, value: 4 },
    { row: 82, col: 0, value: 'Ink' },
    { row: 82, col: 1, value: 2 },
  ]);
  await page.locator('[data-ribbon-command="formatTableHome"]').click();
  await expect(page.locator('#menu-table-style-home')).toBeVisible();
  await page.locator('#menu-table-style-home [data-table-style="dark"]').first().click();
  const formatTableDialog = page.getByRole('dialog', { name: 'Format as Table' });
  await expect(formatTableDialog).toBeVisible();
  await expect(formatTableDialog.getByLabel('Where is the data for your table?')).toHaveValue(
    'A81:B83',
  );
  await expect(
    formatTableDialog.getByRole('checkbox', { name: 'My table has headers' }),
  ).toBeChecked();
  await formatTableDialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect
    .poll(() => readInsertObjectSummary(page))
    .toMatchObject({
      tables: [
        expect.objectContaining({
          source: 'session',
          range: { sheet: 0, r0: 80, c0: 0, r1: 82, c1: 1 },
          style: 'dark',
          showHeader: true,
        }),
      ],
    });

  const undoTable = await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          undo: () => boolean;
          store: {
            getState: () => { tables: { tables: unknown[] } };
          };
        }
      | undefined;
    return {
      ok: inst?.undo() ?? false,
      count: inst?.store.getState().tables.tables.length ?? -1,
    };
  });
  expect(undoTable).toEqual({ ok: true, count: 0 });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readInsertObjectSummary(page))
    .toMatchObject({
      tables: [
        expect.objectContaining({
          source: 'session',
          range: { sheet: 0, r0: 80, c0: 0, r1: 82, c1: 1 },
          style: 'dark',
          showHeader: true,
        }),
      ],
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(async () => (await readInsertObjectSummary(page)).tables).toHaveLength(0);

  await selectRangeAndSetValues(page, { r0: 80, c0: 3, r1: 81, c1: 4 }, [
    { row: 80, col: 3, value: 'No header item' },
    { row: 80, col: 4, value: 7 },
    { row: 81, col: 3, value: 'No header ink' },
    { row: 81, col: 4, value: 8 },
  ]);
  await page.locator('[data-ribbon-command="formatTableHome"]').click();
  await page.locator('#menu-table-style-home [data-table-style="light"]').first().click();
  const noHeaderDialog = page.getByRole('dialog', { name: 'Format as Table' });
  await noHeaderDialog.getByRole('checkbox', { name: 'My table has headers' }).uncheck();
  await noHeaderDialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect
    .poll(() => readInsertObjectSummary(page))
    .toMatchObject({
      tables: [
        expect.objectContaining({
          range: { sheet: 0, r0: 80, c0: 3, r1: 81, c1: 4 },
          style: 'light',
          showHeader: false,
        }),
      ],
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(async () => (await readInsertObjectSummary(page)).tables).toHaveLength(0);
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readInsertObjectSummary(page))
    .toMatchObject({
      tables: [
        expect.objectContaining({
          range: { sheet: 0, r0: 80, c0: 3, r1: 81, c1: 4 },
          style: 'light',
          showHeader: false,
        }),
      ],
    });

  await selectRangeAndSetValues(page, { r0: 84, c0: 0, r1: 84, c1: 0 }, []);
  await page.locator('[data-ribbon-command="cellStyles"]').click();
  const stylesMenu = page.locator('#menu-cell-styles-home');
  await expect(stylesMenu).toBeVisible();
  await expect(stylesMenu).toContainText('Good, Bad and Neutral');
  await expect(stylesMenu).toContainText('Data and Model');
  await expect(stylesMenu).toContainText('Titles and Headings');
  await expect(stylesMenu).toContainText('Themed Cell Styles');
  await expect(stylesMenu).toContainText('Number Format');
  await expect(stylesMenu.locator('[data-cell-style="checkCell"]')).toContainText('Check Cell');
  await expect(stylesMenu.locator('[data-cell-style="accent1_20"]')).toContainText('20% - Accent1');
  await stylesMenu.locator('[data-cell-style="good"]').click();
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({ color: '#006100', fill: '#c6efce' });
  expect(await undoViaInstance(page)).toBe(true);
  const undoneGoodStyle = await readActiveCellFormat(page);
  expect(undoneGoodStyle?.color).toBeUndefined();
  expect(undoneGoodStyle?.fill).toBeUndefined();
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({ color: '#006100', fill: '#c6efce' });

  await page.locator('[data-ribbon-command="cellStyles"]').click();
  await page.locator('#menu-cell-styles-home [data-cell-style="accent5_20"]').click();
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({ color: '#1f4e79', fill: '#ddebf7' });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({ color: '#006100', fill: '#c6efce' });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({ color: '#1f4e79', fill: '#ddebf7' });

  await page.locator('[data-ribbon-command="cellStyles"]').click();
  await page.locator('#menu-cell-styles-home [data-cell-style="normal"]').click();
  const normalFormat = await readActiveCellFormat(page);
  expect(normalFormat?.color).toBeUndefined();
  expect(normalFormat?.fill).toBeUndefined();
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readActiveCellFormat(page))
    .toMatchObject({ color: '#1f4e79', fill: '#ddebf7' });
  expect(await redoViaInstance(page)).toBe(true);
  const redoneNormalFormat = await readActiveCellFormat(page);
  expect(redoneNormalFormat?.color).toBeUndefined();
  expect(redoneNormalFormat?.fill).toBeUndefined();
});

test('R02s: Home Cells menus mutate workbook state and participate in undo where supported', async ({
  page,
}) => {
  await mount(page, '/?locale=en');
  await page.getByRole('tab', { name: 'Home', exact: true }).click();

  await selectRangeAndSetValues(page, { r0: 20, c0: 1, r1: 21, c1: 1 }, [
    { row: 20, col: 1, value: 'B21' },
    { row: 21, col: 1, value: 'B22' },
  ]);
  await page.locator('[data-ribbon-command="insertRows"]').click();
  await expect(page.locator('#menu-insert-cells')).toBeVisible();
  await page.locator('#menu-insert-cells [data-cell-insert="shift-down"]').click();
  await expect
    .poll(() => readCellSummary(page, 22, 1))
    .toMatchObject({
      kind: 'text',
      value: 'B21',
    });
  await expect.poll(() => readCellSummary(page, 20, 1)).toMatchObject({ kind: 'blank' });
  await expect.poll(() => readCellSummary(page, 21, 1)).toMatchObject({ kind: 'blank' });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCellSummary(page, 20, 1))
    .toMatchObject({
      kind: 'text',
      value: 'B21',
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCellSummary(page, 22, 1))
    .toMatchObject({
      kind: 'text',
      value: 'B21',
    });
  await expect.poll(() => readCellSummary(page, 20, 1)).toMatchObject({ kind: 'blank' });

  await selectRangeAndSetValues(page, { r0: 22, c0: 2, r1: 22, c1: 2 }, [
    { row: 22, col: 2, value: 'C23' },
    { row: 22, col: 3, value: 'D23' },
  ]);
  await page.locator('[data-ribbon-command="deleteRows"]').click();
  await expect(page.locator('#menu-delete-cells')).toBeVisible();
  await page.locator('#menu-delete-cells [data-cell-delete="shift-left"]').click();
  await expect
    .poll(() => readCellSummary(page, 22, 2))
    .toMatchObject({
      kind: 'text',
      value: 'D23',
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCellSummary(page, 22, 2))
    .toMatchObject({
      kind: 'text',
      value: 'C23',
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCellSummary(page, 22, 2))
    .toMatchObject({
      kind: 'text',
      value: 'D23',
    });

  await selectRangeAndSetValues(page, { r0: 24, c0: 0, r1: 24, c1: 0 }, []);
  await page.locator('[data-ribbon-command="formatCellsHome"]').click();
  await expect(page.locator('#menu-format-cells')).toBeVisible();
  await page.locator('#menu-format-cells [data-cell-format="hide-rows"]').click();
  await expect.poll(() => readCellsGroupState(page)).toMatchObject({ hiddenRows: [24] });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellsGroupState(page)).toMatchObject({ hiddenRows: [] });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellsGroupState(page)).toMatchObject({ hiddenRows: [24] });

  const beforeSheets = await readCellsGroupState(page);
  await page.locator('[data-ribbon-command="insertRows"]').click();
  await page.locator('#menu-insert-cells [data-cell-insert="sheet"]').click();
  await expect
    .poll(() => readCellsGroupState(page))
    .toMatchObject({
      sheetCount: beforeSheets.sheetCount + 1,
      activeSheet: beforeSheets.sheetCount,
    });
  await expect(page.locator('.app__tab')).toHaveCount(beforeSheets.sheetCount + 1);
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCellsGroupState(page))
    .toMatchObject({
      sheetCount: beforeSheets.sheetCount,
      activeSheet: beforeSheets.activeSheet,
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCellsGroupState(page))
    .toMatchObject({
      sheetCount: beforeSheets.sheetCount + 1,
      activeSheet: beforeSheets.sheetCount,
    });

  await page.locator('[data-ribbon-command="formatCellsHome"]').click();
  await page.locator('#menu-format-cells [data-cell-format="tab-color-red"]').click();
  await expect(page.locator('.app__tab[aria-selected="true"]')).toHaveAttribute(
    'data-sheet-tab-color',
    'true',
  );
  await expect
    .poll(() => readCellsGroupState(page))
    .toMatchObject({
      sheetTabColors: [[beforeSheets.sheetCount, '#c00000']],
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellsGroupState(page)).toMatchObject({ sheetTabColors: [] });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCellsGroupState(page))
    .toMatchObject({
      sheetTabColors: [[beforeSheets.sheetCount, '#c00000']],
    });

  await page.locator('[data-ribbon-command="formatCellsHome"]').click();
  await page.locator('#menu-format-cells [data-cell-format="hide-sheet"]').click();
  await expect
    .poll(() => readCellsGroupState(page))
    .toMatchObject({
      activeSheet: 0,
      hiddenSheets: [beforeSheets.sheetCount],
    });
  expect(await undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCellsGroupState(page))
    .toMatchObject({
      activeSheet: 0,
      hiddenSheets: [],
    });
  expect(await redoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readCellsGroupState(page))
    .toMatchObject({
      activeSheet: 0,
      hiddenSheets: [beforeSheets.sheetCount],
    });
  await expect(page.locator('.app__tab--unhide')).toBeVisible();
});

test('R03: routed ribbon commands open dialogs and mutate workbook state', async ({ page }) => {
  await mount(page, '/?locale=en');

  await page.getByRole('tab', { name: 'Page Layout', exact: true }).click();
  await page.locator('[data-ribbon-command="pageSetupAdvanced"]').click();
  await expect(page.locator('.fc-pgsetup')).toBeVisible();
  await closeDialog(page);
  await page.evaluate(() => {
    const w = window as Window & {
      __fcInst?: { print: (mode?: 'print' | 'pdf') => void };
      __ribbonPrintCalls?: string[];
    };
    const inst = w.__fcInst;
    w.__ribbonPrintCalls = [];
    if (inst) inst.print = (mode = 'print') => void w.__ribbonPrintCalls?.push(mode);
  });
  await page.locator('[data-ribbon-command="printPageLayout"]').click();
  await expect
    .poll(() =>
      page.evaluate(() => {
        const w = window as Window & { __ribbonPrintCalls?: string[] };
        return w.__ribbonPrintCalls?.join(',');
      }),
    )
    .toBe('print');

  await page.getByRole('tab', { name: 'Data', exact: true }).click();
  await page.locator('[data-ribbon-command="linksData"]').click();
  await expect(page.locator('#menu-links-data')).toBeVisible();
  await page.locator('#menu-links-data [data-link-action="external"]').click();
  await expect(page.locator('.fc-extlinkdlg')).toBeVisible();
  await closeDialog(page);

  await page.getByRole('tab', { name: 'Home', exact: true }).click();
  await page.locator('[data-ribbon-command="findHome"]').click();
  await page.locator('#menu-find-select [data-find-select="find"]').click();
  await expect(page.locator('.fc-find')).toBeVisible();
  await closeDialog(page);

  await page.locator('[data-ribbon-command="conditional"]').click();
  await page.locator('#menu-conditional [data-cf-action="manage"]').click();
  await expect(page.locator('.fc-cfrulesdlg')).toBeVisible();
  await closeDialog(page);

  await page.getByRole('tab', { name: 'Draw', exact: true }).click();
  await expect(page.locator('[data-ribbon-command="drawPen"]')).toBeEnabled();
  await expect(page.locator('[data-ribbon-command="drawGrid"]')).toBeEnabled();
  await expect(page.locator('[data-ribbon-command="drawErase"]')).toBeEnabled();
  await page.locator('[data-ribbon-command="drawGrid"]').click();
  await expect
    .poll(() =>
      page.evaluate(() => {
        const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
          | { borderDraw?: { getMode: () => string | null } }
          | undefined;
        return inst?.borderDraw?.getMode() ?? null;
      }),
    )
    .toBe('grid');
  await page.locator('[data-ribbon-command="drawPen"]').click();
  await expect(page.locator('[data-ribbon-command="drawPen"]')).toHaveAttribute(
    'aria-pressed',
    'true',
  );
  const gridBox = await page.locator('.fc-host__grid').boundingBox();
  expect(gridBox).not.toBeNull();
  await page.mouse.move(gridBox!.x + 120, gridBox!.y + 90);
  await page.mouse.down();
  await page.mouse.move(gridBox!.x + 165, gridBox!.y + 120);
  await page.mouse.move(gridBox!.x + 210, gridBox!.y + 95);
  await page.mouse.up();
  await expect(page.locator('.app-ink__stroke')).toHaveCount(1);
  await expect.poll(() => undoViaInstance(page)).toBe(true);
  await expect(page.locator('.app-ink__stroke')).toHaveCount(0);
  await expect.poll(() => redoViaInstance(page)).toBe(true);
  await expect(page.locator('.app-ink__stroke')).toHaveCount(1);
  await expect.poll(() => undoViaInstance(page)).toBe(true);
  await expect(page.locator('.app-ink__stroke')).toHaveCount(0);
  await page.mouse.move(gridBox!.x + 120, gridBox!.y + 90);
  await page.mouse.down();
  await page.mouse.move(gridBox!.x + 165, gridBox!.y + 120);
  await page.mouse.move(gridBox!.x + 210, gridBox!.y + 95);
  await page.mouse.up();
  await expect(page.locator('.app-ink__stroke')).toHaveCount(1);
  await page.locator('[data-ribbon-command="drawErase"]').click();
  await expect(page.locator('[data-ribbon-command="drawErase"]')).toHaveAttribute(
    'aria-pressed',
    'true',
  );
  await page.mouse.click(gridBox!.x + 165, gridBox!.y + 120);
  await expect(page.locator('.app-ink__stroke')).toHaveCount(0);
  await expect.poll(() => undoViaInstance(page)).toBe(true);
  await expect(page.locator('.app-ink__stroke')).toHaveCount(1);
  await expect.poll(() => redoViaInstance(page)).toBe(true);
  await expect(page.locator('.app-ink__stroke')).toHaveCount(0);
  await expect.poll(() => undoViaInstance(page)).toBe(true);
  await expect(page.locator('.app-ink__stroke')).toHaveCount(1);

  await page.getByRole('tab', { name: 'Insert', exact: true }).click();
  await page.locator('[data-ribbon-command="formatTableInsert"]').click();
  await page.locator('#menu-table-style-insert [data-table-style="medium"]').first().click();
  await page
    .getByRole('dialog', { name: 'Format as Table' })
    .getByRole('button', { name: 'OK' })
    .click();
  await page.locator('[data-ribbon-command="chartInsert"]').click();
  await page.locator('#menu-chart-insert [data-chart-insert="column"]').click();
  await expect(page.locator('.fc-chart')).toBeVisible();

  const objectCounts = await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            getState: () => {
              tables: { tables: unknown[] };
              charts: { charts: unknown[] };
            };
          };
        }
      | undefined;
    const state = inst?.store.getState();
    return {
      tables: state?.tables.tables.length ?? 0,
      charts: state?.charts.charts.length ?? 0,
    };
  });
  expect(objectCounts.tables).toBeGreaterThan(0);
  expect(objectCounts.charts).toBeGreaterThan(0);

  await page.getByRole('tab', { name: 'Formulas', exact: true }).click();
  await selectRangeAndSetValues(page, { r0: 90, c0: 0, r1: 90, c1: 0 }, [
    { row: 90, col: 0, value: 'watched' },
  ]);
  await page.locator('[data-ribbon-command="watch"]').click();
  await expect(page.locator('#menu-watch-formulas')).toBeVisible();
  await page.locator('#menu-watch-formulas [data-watch-action="add"]').click();
  await expect(page.locator('.fc-host__watchdock')).toBeVisible();
  await expect.poll(() => readWatchAddresses(page)).toEqual([{ sheet: 0, row: 90, col: 0 }]);
  await expect(page.locator('.fc-watch')).toContainText('A91');
  await expect.poll(() => undoViaInstance(page)).toBe(true);
  await expect.poll(() => readWatchAddresses(page)).toEqual([]);
  await expect.poll(() => redoViaInstance(page)).toBe(true);
  await expect.poll(() => readWatchAddresses(page)).toEqual([{ sheet: 0, row: 90, col: 0 }]);
  await expect.poll(() => undoViaInstance(page)).toBe(true);
  await expect.poll(() => readWatchAddresses(page)).toEqual([]);

  await page.locator('[data-ribbon-command="watch"]').click();
  await page.locator('#menu-watch-formulas [data-watch-action="add"]').click();
  await expect.poll(() => readWatchAddresses(page)).toEqual([{ sheet: 0, row: 90, col: 0 }]);

  await page.locator('[data-ribbon-command="watch"]').click();
  await page.locator('#menu-watch-formulas [data-watch-action="delete"]').click();
  await expect.poll(() => readWatchAddresses(page)).toEqual([]);
  await expect.poll(() => undoViaInstance(page)).toBe(true);
  await expect.poll(() => readWatchAddresses(page)).toEqual([{ sheet: 0, row: 90, col: 0 }]);
  await expect.poll(() => redoViaInstance(page)).toBe(true);
  await expect.poll(() => readWatchAddresses(page)).toEqual([]);
  await expect.poll(() => undoViaInstance(page)).toBe(true);
  await expect.poll(() => readWatchAddresses(page)).toEqual([{ sheet: 0, row: 90, col: 0 }]);

  await selectRangeAndSetValues(page, { r0: 91, c0: 0, r1: 91, c1: 1 }, [
    { row: 91, col: 0, value: 'left' },
    { row: 91, col: 1, value: 'right' },
  ]);
  await page.locator('[data-ribbon-command="watch"]').click();
  await page.locator('#menu-watch-formulas [data-watch-action="add"]').click();
  await expect
    .poll(() => readWatchAddresses(page))
    .toEqual([
      { sheet: 0, row: 90, col: 0 },
      { sheet: 0, row: 91, col: 0 },
      { sheet: 0, row: 91, col: 1 },
    ]);
  await page.locator('[data-ribbon-command="watch"]').click();
  await page.locator('#menu-watch-formulas [data-watch-action="delete-all"]').click();
  await expect.poll(() => readWatchAddresses(page)).toEqual([]);
  await expect.poll(() => undoViaInstance(page)).toBe(true);
  await expect
    .poll(() => readWatchAddresses(page))
    .toEqual([
      { sheet: 0, row: 90, col: 0 },
      { sheet: 0, row: 91, col: 0 },
      { sheet: 0, row: 91, col: 1 },
    ]);
  await expect.poll(() => redoViaInstance(page)).toBe(true);
  await expect.poll(() => readWatchAddresses(page)).toEqual([]);

  await page.locator('[data-ribbon-command="fx"]').click();
  await expect(page.locator('.fc-fxdialog')).toBeVisible();
  await closeDialog(page);

  await page.getByRole('tab', { name: 'Review', exact: true }).click();
  for (const command of ['spellingReview', 'translateReview', 'accessibility']) {
    await expect(page.locator(`[data-ribbon-command="${command}"]`)).toBeEnabled();
    await page.locator(`[data-ribbon-command="${command}"]`).click();
    await expect(page.getByRole('dialog')).toBeVisible();
    await closeDialog(page);
  }

  await page.getByRole('tab', { name: 'Automate', exact: true }).click();
  await expect(page.locator('[data-ribbon-command="script"]')).toBeEnabled();
  await expect(page.locator('[data-ribbon-command="recordActions"]')).toBeEnabled();
  await expect(page.locator('[data-ribbon-command="allScripts"]')).toBeEnabled();
  await page.locator('[data-ribbon-command="recordActions"]').click();
  await expect(page.getByRole('dialog')).toContainText('Recorded selected range action');
  await closeDialog(page);
  await page.locator('[data-ribbon-command="allScripts"]').click();
  await expect(page.getByRole('dialog')).toContainText('Built-in scripts');
  await expect(page.getByRole('dialog')).toContainText('Recent script runs');
  await closeDialog(page);
  await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          workbook: {
            setText: (addr: { sheet: number; row: number; col: number }, value: string) => void;
          };
          store: {
            getState: () => {
              selection: { active: { sheet: number; row: number; col: number } };
            };
            setState: (
              fn: (state: {
                selection: {
                  active: { sheet: number; row: number; col: number };
                  anchor: { sheet: number; row: number; col: number };
                  range: { sheet: number; r0: number; c0: number; r1: number; c1: number };
                };
              }) => unknown,
            ) => void;
          };
        }
      | undefined;
    const active = inst?.store.getState().selection.active;
    if (inst && active) {
      inst.workbook.setText(active, 'script value');
      inst.workbook.setText(
        { sheet: active.sheet, row: active.row, col: active.col + 1 },
        'second value',
      );
      inst.store.setState((state) => ({
        ...state,
        selection: {
          ...state.selection,
          active,
          anchor: active,
          range: {
            sheet: active.sheet,
            r0: active.row,
            c0: active.col,
            r1: active.row,
            c1: active.col + 1,
          },
        },
      }));
    }
  });
  await page.locator('[data-ribbon-command="script"]').click();
  await expect(page.locator('#menu-script')).toBeVisible();
  await page.locator('#menu-script [data-script-action="uppercase"]').click();
  const scriptResult = await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          workbook: {
            getValue: (addr: { sheet: number; row: number; col: number }) => {
              kind: string;
              value?: string;
            };
          };
          store: {
            getState: () => {
              selection: { active: { sheet: number; row: number; col: number } };
            };
          };
        }
      | undefined;
    const active = inst?.store.getState().selection.active;
    return inst && active
      ? [
          inst.workbook.getValue(active),
          inst.workbook.getValue({ sheet: active.sheet, row: active.row, col: active.col + 1 }),
        ]
      : null;
  });
  expect(scriptResult).toMatchObject([
    { kind: 'text', value: 'SCRIPT VALUE' },
    { kind: 'text', value: 'SECOND VALUE' },
  ]);
  const undoResult = await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          undo: () => boolean;
          workbook: {
            getValue: (addr: { sheet: number; row: number; col: number }) => {
              kind: string;
              value?: string;
            };
          };
          store: {
            getState: () => {
              selection: { active: { sheet: number; row: number; col: number } };
            };
          };
        }
      | undefined;
    const active = inst?.store.getState().selection.active;
    const ok = inst?.undo() ?? false;
    return inst && active
      ? {
          ok,
          values: [
            inst.workbook.getValue(active),
            inst.workbook.getValue({ sheet: active.sheet, row: active.row, col: active.col + 1 }),
          ],
        }
      : null;
  });
  expect(undoResult).toMatchObject({
    ok: true,
    values: [
      { kind: 'text', value: 'script value' },
      { kind: 'text', value: 'second value' },
    ],
  });
  const redoResult = await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          redo: () => boolean;
          workbook: {
            getValue: (addr: { sheet: number; row: number; col: number }) => {
              kind: string;
              value?: string;
            };
          };
          store: {
            getState: () => {
              selection: { active: { sheet: number; row: number; col: number } };
            };
          };
        }
      | undefined;
    const active = inst?.store.getState().selection.active;
    const ok = inst?.redo() ?? false;
    return inst && active
      ? {
          ok,
          values: [
            inst.workbook.getValue(active),
            inst.workbook.getValue({ sheet: active.sheet, row: active.row, col: active.col + 1 }),
          ],
        }
      : null;
  });
  expect(redoResult).toMatchObject({
    ok: true,
    values: [
      { kind: 'text', value: 'SCRIPT VALUE' },
      { kind: 'text', value: 'SECOND VALUE' },
    ],
  });

  await selectRangeAndSetValues(page, { r0: 94, c0: 0, r1: 94, c1: 1 }, [
    { row: 94, col: 0, value: 'MiXeD' },
    { row: 94, col: 1, value: 'SECOND' },
  ]);
  await page.locator('[data-ribbon-command="script"]').click();
  await page.locator('#menu-script [data-script-action="lowercase"]').click();
  await expect.poll(() => readCellText(page, 94, 0)).toBe('mixed');
  await expect.poll(() => readCellText(page, 94, 1)).toBe('second');
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 94, 0)).toBe('MiXeD');
  await expect.poll(() => readCellText(page, 94, 1)).toBe('SECOND');
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 94, 0)).toBe('mixed');
  await expect.poll(() => readCellText(page, 94, 1)).toBe('second');

  await selectRangeAndSetValues(page, { r0: 95, c0: 0, r1: 95, c1: 2 }, [
    { row: 95, col: 0, value: '  Alpha  ' },
    { row: 95, col: 1, value: 'Beta ' },
    { row: 95, col: 2, value: 9 },
  ]);
  await page.locator('[data-ribbon-command="script"]').click();
  await page.locator('#menu-script [data-script-action="custom"]').click();
  const scriptDialog = page.getByRole('dialog', { name: 'Script' });
  await expect(scriptDialog).toBeVisible();
  await scriptDialog.getByRole('textbox', { name: 'Command' }).fill('trim');
  await scriptDialog.getByRole('button', { name: 'Run' }).click();
  await expect.poll(() => readCellText(page, 95, 0)).toBe('Alpha');
  await expect.poll(() => readCellText(page, 95, 1)).toBe('Beta');
  await expect.poll(() => readCellSummary(page, 95, 2)).toMatchObject({ kind: 'number', value: 9 });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 95, 0)).toBe('  Alpha  ');
  await expect.poll(() => readCellText(page, 95, 1)).toBe('Beta ');
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 95, 0)).toBe('Alpha');
  await expect.poll(() => readCellText(page, 95, 1)).toBe('Beta');

  await selectRangeAndSetValues(page, { r0: 96, c0: 0, r1: 96, c1: 1 }, [
    { row: 96, col: 0, value: 'delete me' },
    { row: 96, col: 1, value: 123 },
  ]);
  await page.locator('[data-ribbon-command="script"]').click();
  await page.locator('#menu-script [data-script-action="clear"]').click();
  await expect.poll(() => readCellSummary(page, 96, 0)).toMatchObject({ kind: 'blank' });
  await expect.poll(() => readCellSummary(page, 96, 1)).toMatchObject({ kind: 'blank' });
  expect(await undoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellText(page, 96, 0)).toBe('delete me');
  await expect
    .poll(() => readCellSummary(page, 96, 1))
    .toMatchObject({ kind: 'number', value: 123 });
  expect(await redoViaInstance(page)).toBe(true);
  await expect.poll(() => readCellSummary(page, 96, 0)).toMatchObject({ kind: 'blank' });
  await expect.poll(() => readCellSummary(page, 96, 1)).toMatchObject({ kind: 'blank' });

  await page.getByRole('tab', { name: 'Acrobat', exact: true }).click();
  await expect(page.locator('[data-ribbon-command="addIn"]')).toBeEnabled();
  await page.locator('[data-ribbon-command="addIn"]').click();
  await expect(page.locator('#menu-add-ins')).toBeVisible();
  await page.locator('#menu-add-ins [data-add-in-action="get"]').click();
  await expect(page.getByRole('dialog')).toContainText('Office Add-ins store');
  await closeDialog(page);
  await page.locator('[data-ribbon-command="addIn"]').click();
  await page.locator('#menu-add-ins [data-add-in-action="manage"]').click();
  await expect(page.getByRole('dialog')).toContainText('Add-in management');
  await closeDialog(page);
  await page.locator('[data-ribbon-command="addIn"]').click();
  await page.locator('#menu-add-ins [data-add-in-action="my"]').click();
  await expect(page.getByRole('dialog', { name: 'My Add-ins' })).toContainText('Built-in add-ins');
  await expect(page.getByRole('dialog', { name: 'My Add-ins' })).toContainText('External add-ins');
  await closeDialog(page);
  await page.locator('[data-ribbon-command="pdf"]').click();
  await expect(page.locator('#menu-pdf')).toBeVisible();
  await page.locator('#menu-pdf [data-pdf-action="create"]').click();
  await expect
    .poll(() =>
      page.evaluate(() => {
        const w = window as Window & { __ribbonPrintCalls?: string[] };
        return w.__ribbonPrintCalls?.join(',');
      }),
    )
    .toBe('print,pdf');
  await page.locator('[data-ribbon-command="pdf"]').click();
  await page.locator('#menu-pdf [data-pdf-action="share"]').click();
  await expect(page.getByRole('dialog', { name: 'PDF' })).toContainText('PDF export is ready.');
  await closeDialog(page);
  await expect
    .poll(() =>
      page.evaluate(() => {
        const w = window as Window & { __ribbonPrintCalls?: string[] };
        return w.__ribbonPrintCalls?.join(',');
      }),
    )
    .toBe('print,pdf,pdf');
  await page.locator('[data-ribbon-command="pdf"]').click();
  await page.locator('#menu-pdf [data-pdf-action="preferences"]').click();
  const pdfPreferences = page.getByRole('dialog', { name: 'Page Setup' });
  await expect(pdfPreferences).toBeVisible();
  await expect(pdfPreferences.locator('[data-pgsetup-tab="page"][role="tab"]')).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await pdfPreferences.locator('[data-pgsetup-tab="sheet"][role="tab"]').click();
  await expect(pdfPreferences.getByLabel('Print area')).toBeVisible();
  await pdfPreferences.getByRole('button', { name: 'Cancel', exact: true }).click();
  await expect(pdfPreferences).toBeHidden();

  await page.getByRole('tab', { name: 'View', exact: true }).click();
  await page.locator('[data-ribbon-command="zoom125"]').click();
  const zoom = await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            getState: () => {
              viewport: { zoom: number };
            };
          };
        }
      | undefined;
    return inst?.store.getState().viewport.zoom;
  });
  expect(zoom).toBe(1.25);
});

test('R04: ribbon tabs support Excel-style arrow, Home, and End keyboard navigation', async ({
  page,
}) => {
  await mount(page, '/?locale=en');

  const home = page.getByRole('tab', { name: 'Home', exact: true });
  const insert = page.getByRole('tab', { name: 'Insert', exact: true });
  const acrobat = page.getByRole('tab', { name: 'Acrobat', exact: true });

  await home.focus();
  await expect(home).toBeFocused();
  await page.keyboard.press('ArrowRight');
  await expect(insert).toBeFocused();
  await expect(insert).toHaveAttribute('aria-selected', 'true');
  await expect(page.locator('.demo__ribbon:not([hidden])')).toHaveAttribute(
    'data-ribbon-panel',
    'insert',
  );

  await page.keyboard.press('End');
  await expect(acrobat).toBeFocused();
  await expect(acrobat).toHaveAttribute('aria-selected', 'true');

  await page.keyboard.press('Home');
  await expect(home).toBeFocused();
  await expect(home).toHaveAttribute('aria-selected', 'true');

  await page.keyboard.press('ArrowLeft');
  await expect(acrobat).toBeFocused();
  await expect(acrobat).toHaveAttribute('tabindex', '0');
  await expect(home).toHaveAttribute('tabindex', '-1');
});

test('R04b: ribbon supports Excel-style collapsed tabs-only mode', async ({ page }) => {
  await mount(page, '/?locale=en');

  const shell = page.locator('.app__ribbon-shell').first();
  const tabs = page.locator('.demo__ribbon-tabs').first();
  const home = page.getByRole('tab', { name: 'Home', exact: true });

  await expect(shell).not.toHaveClass(/demo__ribbon-shell--collapsed/);
  await expect(tabs).toHaveAttribute('data-ribbon-collapsed', 'false');
  await expect(page.getByRole('button', { name: 'Ribbon Display Options' })).toHaveAttribute(
    'aria-expanded',
    'false',
  );

  await page
    .locator('.fc-host')
    .first()
    .click({ position: { x: 240, y: 220 } });
  await page.keyboard.press('Control+F1');
  await expect(shell).toHaveClass(/demo__ribbon-shell--collapsed/);
  await expect(tabs).toHaveAttribute('data-ribbon-collapsed', 'true');
  await expect(page.locator('.demo__ribbon:not([hidden])')).not.toBeVisible();
  await page.getByRole('button', { name: 'Ribbon Display Options' }).click();
  await expect(page.getByRole('menuitemradio', { name: 'Show tabs only' })).toHaveAttribute(
    'aria-checked',
    'true',
  );
  await page
    .locator('.fc-host')
    .first()
    .click({ position: { x: 260, y: 260 } });
  await expect(page.getByRole('menuitemradio', { name: 'Show tabs only' })).toBeHidden();

  await page.getByRole('button', { name: 'Ribbon Display Options' }).click();
  await page.getByRole('menuitemradio', { name: 'Always show Ribbon' }).click();
  await expect(shell).not.toHaveClass(/demo__ribbon-shell--collapsed/);
  await expect(tabs).toHaveAttribute('data-ribbon-collapsed', 'false');

  await page.getByRole('button', { name: 'Ribbon Display Options' }).focus();
  await page.keyboard.press('ArrowDown');
  await expect(page.getByRole('menuitemradio', { name: 'Always show Ribbon' })).toBeFocused();
  await page.keyboard.press('End');
  await expect(page.getByRole('menuitemradio', { name: 'Show tabs only' })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(page.getByRole('menuitemradio', { name: 'Show tabs only' })).toBeHidden();

  await home.dblclick();
  await expect(shell).toHaveClass(/demo__ribbon-shell--collapsed/);
  await expect(tabs).toHaveAttribute('data-ribbon-collapsed', 'true');
});

test('R04c: Mac-style ribbon keeps File backstage tab hidden', async ({ page }) => {
  await mount(page, '/?locale=en');

  await expect(page.getByRole('tab', { name: 'File', exact: true })).toHaveCount(0);
  await expect(page.getByRole('tab', { name: 'Home', exact: true })).toHaveAttribute(
    'aria-selected',
    'true',
  );
});

test('R04d: title-bar quick access buttons route Save, Save As, Undo, Redo, and Home', async ({
  page,
}) => {
  await page.addInitScript(() => {
    const w = window as Window & { __titleDownloads?: string[] };
    w.__titleDownloads = [];
    HTMLAnchorElement.prototype.click = function click(this: HTMLAnchorElement): void {
      w.__titleDownloads?.push(this.download);
    };
  });
  await mount(page, '/?locale=en');

  const autosave = page.locator('.app__autosave-switch');
  await expect(autosave).toHaveAttribute('aria-pressed', 'false');
  await autosave.click();
  await expect(autosave).toHaveAttribute('aria-pressed', 'true');
  await expect(page.locator('#status-metric')).toHaveText('AutoSave is on');
  await autosave.click();
  await expect(autosave).toHaveAttribute('aria-pressed', 'false');
  await expect(page.locator('#status-metric')).toHaveText('AutoSave is off');

  await page.locator('.app__title [data-shell-i18n-label="save"]').click();
  await expect
    .poll(() =>
      page.evaluate(() => {
        const w = window as Window & { __titleDownloads?: string[] };
        return w.__titleDownloads ?? [];
      }),
    )
    .toEqual(['Book1.xlsx']);

  await page.locator('.app__title [data-shell-i18n-label="saveAs"]').click();
  const saveAsDialog = page.getByRole('dialog', { name: 'Save As' });
  await expect(saveAsDialog).toBeVisible();
  await saveAsDialog.locator('.app__dlg__input').fill('Ribbon QA');
  await saveAsDialog.getByRole('button', { name: 'Save', exact: true }).click();
  await expect(page.locator('#doc-name')).toHaveText('Ribbon QA');
  await expect
    .poll(() =>
      page.evaluate(() => {
        const w = window as Window & { __titleDownloads?: string[] };
        return w.__titleDownloads ?? [];
      }),
    )
    .toEqual(['Book1.xlsx', 'Ribbon QA.xlsx']);

  await page.locator('[data-ribbon-command="bold"]').click();
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ bold: true });
  await page.locator('.app__title [data-shell-i18n-label="undo"]').click();
  await expect.poll(async () => (await readActiveCellFormat(page))?.bold).toBeUndefined();
  await page.locator('.app__title [data-shell-i18n-label="redo"]').click();
  await expect.poll(() => readActiveCellFormat(page)).toMatchObject({ bold: true });

  await page.getByRole('tab', { name: 'Data', exact: true }).click();
  await page.locator('.app__title [data-shell-i18n-label="home"]').click();
  await expect(page.getByRole('tab', { name: 'Home', exact: true })).toHaveAttribute(
    'aria-selected',
    'true',
  );

  await page.locator('.app__title [data-shell-i18n-label="comments"]').click();
  await expect(page.getByRole('dialog', { name: 'New note' })).toBeVisible();
  await closeDialog(page);

  await page.locator('.app__title [data-shell-i18n-label="share"]').click();
  const shareDialog = page.getByRole('dialog', { name: 'Share' });
  await expect(shareDialog).toBeVisible();
  await expect(shareDialog).toContainText('This workbook is ready to share.');
  await closeDialog(page);

  await page.locator('.app__title [data-shell-i18n-label="more"]').click();
  const titleMoreMenu = page.locator('#menu-title-more');
  await expect(titleMoreMenu).toBeVisible();
  await expect(titleMoreMenu.getByRole('menuitem', { name: 'Save', exact: true })).toBeVisible();
  await expect(titleMoreMenu.getByRole('menuitem', { name: 'Save As', exact: true })).toBeVisible();
  await titleMoreMenu.getByRole('menuitem', { name: 'AutoSave', exact: true }).click();
  await expect(autosave).toHaveAttribute('aria-pressed', 'true');
  await page.locator('.app__title [data-shell-i18n-label="more"]').click();
  await titleMoreMenu.getByRole('menuitem', { name: 'Share', exact: true }).click();
  await expect(page.getByRole('dialog', { name: 'Share' })).toContainText(
    'This workbook is ready to share.',
  );
  await closeDialog(page);

  await selectRangeAndSetValues(page, { r0: 0, c0: 0, r1: 0, c1: 0 }, [
    { row: 10, col: 2, value: 'Search target' },
  ]);
  const titleSearch = page.locator('.app__search input[type="search"]');
  await page.keyboard.press('Meta+Control+U');
  await expect(titleSearch).toBeFocused();
  await titleSearch.fill('target');
  await titleSearch.press('Enter');
  const findDialog = page.getByRole('dialog', { name: 'Find and Replace' });
  await expect(findDialog).toBeVisible();
  await expect(findDialog.locator('input[type="text"]').first()).toHaveValue('target');
  await expect
    .poll(() => readSelectionSummary(page))
    .toMatchObject({
      active: { sheet: 0, row: 10, col: 2 },
    });
  await closeDialog(page);
  await titleSearch.fill('not-present');
  await titleSearch.press('Enter');
  await expect(findDialog).toBeVisible();
  await expect(page.locator('#status-metric')).toHaveText('No matches for "not-present"');
  await closeDialog(page);
});

test('R05: ribbon dropdowns support keyboard selection and Escape dismissal', async ({ page }) => {
  await mount(page, '/?locale=en');

  const fontButton = page.locator('[data-ribbon-select="fontFamily"] .demo__rb-dd__btn');
  await fontButton.focus();
  await page.keyboard.press('ArrowDown');

  const list = page.locator('[data-ribbon-select="fontFamily"] .demo__rb-dd__list');
  await expect(list).toBeVisible();
  await expect(fontButton).toHaveAttribute('aria-expanded', 'true');
  const options = page.locator('[data-ribbon-select="fontFamily"] [role="option"]');
  await expect(options.nth(0)).toBeFocused();
  await page.keyboard.press('ArrowDown');
  await expect(options.nth(1)).toBeFocused();
  const secondLabel = (await options.nth(1).locator('.demo__rb-dd__label').textContent()) ?? '';
  await page.keyboard.press('Enter');
  await expect(list).toBeHidden();
  await expect(fontButton).toBeFocused();
  await expect(page.locator('[data-ribbon-select="fontFamily"] .demo__rb-dd__value')).toHaveText(
    secondLabel,
  );

  await page.keyboard.press('ArrowDown');
  await expect(list).toBeVisible();
  await page.keyboard.press('End');
  await expect(options.last()).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(list).toBeHidden();
  await expect(fontButton).toBeFocused();
  await expect(fontButton).toHaveAttribute('aria-expanded', 'false');
});

test('R06: ribbon split menus support menu keyboard navigation and focus return', async ({
  page,
}) => {
  await mount(page, '/?locale=en');

  const paste = page.locator('button[data-ribbon-command="paste"]');
  const pasteMenu = page.locator('#menu-paste');
  await paste.locator('.demo__rb-split-chevron').click();
  await expect(pasteMenu).toBeVisible();
  await expect(pasteMenu.getByRole('menuitem', { name: 'Paste', exact: true })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(pasteMenu).toBeHidden();
  await expect(paste).toBeFocused();
  await paste.focus();
  await page.keyboard.press('ArrowDown');
  await expect(pasteMenu).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(pasteMenu).toBeHidden();

  const borders = page.locator('#btn-borders');
  const borderMenu = page.locator('#menu-borders');
  await borders.click();
  await expect(borderMenu).toBeVisible();
  // Excel-365 first preset row: 下罫線 / Bottom Border.
  await expect(
    borderMenu.getByRole('menuitem', { name: 'Bottom Border', exact: true }),
  ).toBeFocused();
  await page.keyboard.press('End');
  await expect(
    borderMenu.getByRole('menuitem', { name: 'More Borders...', exact: true }),
  ).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(borderMenu).toBeHidden();
  await expect(borders).toBeFocused();

  await page.getByRole('tab', { name: 'Data', exact: true }).click();
  const sort = page.locator('[data-ribbon-command="filter"]');
  const sortMenu = page.locator('#menu-sort');
  await sort.click();
  await expect(sortMenu).toBeVisible();
  await page.keyboard.press('ArrowDown');
  await expect(sortMenu.getByRole('menuitem', { name: 'Sort Z → A' })).toBeFocused();
  await page.keyboard.press('Home');
  await expect(sortMenu.getByRole('menuitem', { name: 'Sort A → Z' })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(sortMenu).toBeHidden();
  await expect(sort).toBeFocused();
});
