import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** I01 — `?locale=ja` boots the app in Japanese and keeps the demo chrome
 *  aligned with the Japanese Excel desktop baseline. */
export async function runLocaleBootScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await page.goto('/?locale=ja');
  await sp.waitForReady();
  await sp.expectNoStub();
  const jaToggle = page.getByRole('button', { name: 'JA', exact: true });
  if ((await jaToggle.count()) > 0) await jaToggle.first().click();

  await expect(page.getByRole('tab', { name: 'ホーム', exact: true })).toBeVisible();
  await expect(page.getByRole('tab', { name: '挿入', exact: true })).toBeVisible();
  await expect(page.getByRole('button', { name: '貼り付け', exact: true }).first()).toBeVisible();
  const searchBox = page.locator('.fc-tb__search input').first();
  await expect(searchBox).toHaveAttribute('aria-label', 'コマンドの検索');
  await searchBox.fill('貼り付け');
  const searchResults = page.locator('#demo-search-results');
  await expect(searchResults).toBeVisible();
  await expect(searchResults.getByRole('option', { name: /貼り付け/ }).first()).toBeVisible();
  await expect(searchResults).not.toContainText('Paste');
  await searchBox.fill('');
  await page.keyboard.press('Escape');
  await expect(searchResults).toBeHidden();
  const ribbonDisplayToggle = page.locator('[data-ribbon-toggle]').first();
  await expect(ribbonDisplayToggle).toHaveAttribute('aria-label', 'リボンの表示オプション');
  await ribbonDisplayToggle.click();
  const ribbonDisplayMenu = page.locator('.fc-tb__ribbon-display-menu');
  await expect(ribbonDisplayMenu).toBeVisible();
  for (const label of [
    'リボンを常に表示',
    '1 行のリボン',
    'タブのみ表示',
    'リボンを自動的に非表示',
  ]) {
    await expect(ribbonDisplayMenu.getByRole('menuitemradio', { name: label })).toBeVisible();
  }
  await expect(ribbonDisplayMenu).not.toContainText('Ribbon Display Options');
  await expect(ribbonDisplayMenu).not.toContainText('Show tabs only');
  await ribbonDisplayToggle.click();
  await expect(ribbonDisplayMenu).toHaveCount(0);

  const statusbar = page.locator('.fc-host__statusbar').first();
  await expect(statusbar).toBeVisible();
  await expect(statusbar).toContainText('準備完了');
  await statusbar.click({ button: 'right' });
  const statusChooser = page.locator('.fc-statusbar__chooser');
  await expect(statusChooser).toBeVisible();
  for (const label of [
    '集計表示',
    'ステータス バー項目',
    '平均',
    'データの個数',
    '合計',
    '表示選択ショートカット',
    'ズーム',
  ]) {
    await expect(statusChooser.getByText(label, { exact: true })).toBeVisible();
  }
  await expect(statusChooser).not.toContainText('Customize Status Bar');
  await expect(statusChooser).not.toContainText('Average');
  await page.keyboard.press('Escape');
  await expect(statusChooser).toBeHidden();

  const selectedSheetTab = page.locator('.fc-host__sheetbar-tab[aria-selected="true"]').first();
  await expect(selectedSheetTab).toBeVisible();
  await selectedSheetTab.click({ button: 'right' });
  const sheetMenu = page.locator('.fc-sheetmenu');
  await expect(sheetMenu).toBeVisible();
  for (const label of [
    '名前の変更',
    'シートの挿入',
    '左へ移動',
    '右へ移動',
    '削除',
    '非表示',
    'シートの再表示',
    'タブの色',
  ]) {
    await expect(sheetMenu.getByText(label, { exact: true }).first()).toBeVisible();
  }
  await expect(sheetMenu).not.toContainText('Rename');
  await expect(sheetMenu).not.toContainText('Insert Sheet');
  await expect(sheetMenu).not.toContainText('Tab Color');
  await page.keyboard.press('Escape');
  await expect(sheetMenu).toBeHidden();

  await sp.focusHost();
  await page.keyboard.type('seed');
  await page.keyboard.press('Enter');
  await page.keyboard.press('ArrowUp');
  await sp.shortcut('c');
  await page
    .locator('.fc-host')
    .first()
    .click({ button: 'right', position: { x: 200, y: 200 } });
  const contextMenu = page.locator('.fc-ctxmenu:not(.fc-ctxmenu__sub)');
  await expect(contextMenu).toBeVisible();
  await expect(contextMenu).toHaveAttribute('aria-label', 'コンテキスト メニュー');
  for (const label of [
    'コピー',
    '切り取り',
    '貼り付け',
    '形式を選択して貼り付け…',
    'クリア',
    'セルの書式設定…',
  ]) {
    await expect(contextMenu.getByText(label, { exact: true }).first()).toBeVisible();
  }
  await expect(contextMenu).not.toContainText('Copy');
  await expect(contextMenu).not.toContainText('Paste Special');
  await expect(contextMenu).not.toContainText('Format Cells');
  await contextMenu.locator('[data-fc-submenu="pasteSpecialMenu"]').hover();
  const pasteSpecialSubmenu = page.locator('.fc-ctxmenu__sub');
  await expect(pasteSpecialSubmenu).toBeVisible();
  await pasteSpecialSubmenu.locator('[data-fc-action="pasteSpecial"]').click();
  await expect(contextMenu).toBeHidden();
  const pasteSpecialDialog = page.getByRole('dialog', { name: '形式を選択して貼り付け' });
  await expect(pasteSpecialDialog).toBeVisible();
  for (const label of [
    '貼り付け',
    '演算',
    'すべて',
    '数式',
    '値',
    '書式',
    '空白セルを無視する',
    '行/列の入れ替え',
  ]) {
    await expect(pasteSpecialDialog.getByText(label, { exact: true }).first()).toBeVisible();
  }
  await expect(pasteSpecialDialog).not.toContainText('Paste Special');
  await expect(pasteSpecialDialog).not.toContainText('Operation');
  await expect(pasteSpecialDialog).not.toContainText('Skip blanks');
  await pasteSpecialDialog
    .locator('.fc-fmtdlg__footer')
    .getByText('キャンセル', { exact: true })
    .click();
  await expect(pasteSpecialDialog).toHaveCount(0);

  await page
    .locator('.fc-host')
    .first()
    .click({ button: 'right', position: { x: 200, y: 200 } });
  await expect(contextMenu).toBeVisible();
  await contextMenu.getByText('セルの書式設定…', { exact: true }).first().click();
  await expect(contextMenu).toBeHidden();
  const formatDialog = page.getByRole('dialog', { name: 'セルの書式設定' });
  await expect(formatDialog).toBeVisible();
  for (const label of [
    '表示形式',
    '配置',
    'フォント',
    '罫線',
    '塗りつぶし',
    'その他',
    '標準',
    '数値',
    '通貨',
    'パーセンテージ',
  ]) {
    await expect(formatDialog.getByText(label, { exact: true }).first()).toBeVisible();
  }
  await expect(formatDialog).not.toContainText('Format Cells');
  await expect(formatDialog).not.toContainText('Alignment');
  await expect(formatDialog).not.toContainText('Currency');
  await formatDialog.locator('.fc-fmtdlg__footer').getByText('キャンセル', { exact: true }).click();
  await expect(formatDialog).toHaveCount(0);

  await page.locator('[data-ribbon-tab="home"]').click();
  const homeRibbon = page.getByRole('toolbar', { name: 'ホーム リボン' });
  await expect(homeRibbon).toBeVisible();
  for (const name of [
    '貼り付け',
    '切り取り',
    'コピー',
    '書式のコピー',
    'フォント サイズの拡大',
    'フォント サイズの縮小',
    '太字 (Ctrl+B)',
    '斜体 (Ctrl+I)',
    '下線 (Ctrl+U)',
    '取り消し線',
    '罫線',
    '上揃え',
    '上下中央揃え',
    '下揃え',
    '左揃え',
    '中央揃え',
    '右揃え',
    '文字列の方向',
    '折り返して全体を表示',
    'インデントを減らす',
    'インデントを増やす',
    'セルの結合',
    '通貨',
    'パーセント',
    '桁区切りスタイル',
    '小数点以下の桁数を減らす',
    '小数点以下の桁数を増やす',
    '条件付き書式',
    'テーブルとして書式設定',
    'セル スタイル',
    '選択した行を挿入',
    '選択した行を削除',
    'セルの書式設定',
    'オートSUM (Σ)',
    'フィル',
    'クリア',
    '並べ替えとフィルター',
    '検索と選択 (Ctrl+F)',
  ]) {
    await expect(homeRibbon.getByRole('button', { name, exact: true }).first()).toBeVisible();
  }
  await expect(page.locator('[data-ribbon-command="paste"]').first()).toHaveAttribute(
    'data-ribbon-activation',
    'splitPrimary',
  );
  await expect(page.locator('[data-ribbon-command="conditional"]').first()).toHaveAttribute(
    'data-ribbon-activation',
    'gallery',
  );
  await expect(page.locator('[data-ribbon-command="formatTableHome"]').first()).toHaveAttribute(
    'data-ribbon-activation',
    'gallery',
  );
  await expect(page.locator('[data-ribbon-command="cellStyles"]').first()).toHaveAttribute(
    'data-ribbon-activation',
    'gallery',
  );
  await expect(homeRibbon).not.toContainText('Clipboard');
  await expect(homeRibbon).not.toContainText('Format Painter');
  await expect(homeRibbon).not.toContainText('Conditional Formatting');

  const conditionalButton = page.locator('[data-ribbon-command="conditional"]').first();
  await expect(conditionalButton).toBeVisible();
  await conditionalButton.click();
  const conditionalMenu = page.locator('#menu-conditional');
  await expect(conditionalMenu).toBeVisible();
  for (const label of [
    'セルの強調表示ルール',
    '上位/下位ルール',
    'データ バー',
    'カラー スケール',
    'アイコン セット',
    '新しいルール...',
    'ルールのクリア',
    'ルールの管理...',
  ]) {
    await expect(conditionalMenu.getByText(label, { exact: true }).first()).toBeVisible();
  }
  await expect(conditionalMenu).not.toContainText('Highlight Cells Rules');
  await expect(conditionalMenu).not.toContainText('Top/Bottom Rules');
  await expect(conditionalMenu).not.toContainText('Manage Rules');
  await conditionalMenu.locator('[data-cf-submenu="highlight"]').hover();
  const highlightMenu = page.locator('#menu-conditional-highlight');
  await expect(highlightMenu).toBeVisible();
  for (const label of [
    '指定の値より大きい...',
    '指定の値より小さい...',
    '指定の範囲内...',
    '指定の値に等しい...',
    '文字列...',
    '日付...',
    '重複する値...',
    '一意の値...',
    'その他のルール...',
  ]) {
    await expect(highlightMenu.getByText(label, { exact: true }).first()).toBeVisible();
  }
  await expect(highlightMenu).not.toContainText('Greater Than');
  await expect(highlightMenu).not.toContainText('Duplicate Values');
  await conditionalMenu.locator('[data-cf-submenu="topBottom"]').hover();
  const topBottomMenu = page.locator('#menu-conditional-topBottom');
  await expect(topBottomMenu).toBeVisible();
  for (const label of [
    '上位 10 項目',
    '下位 10 項目',
    '上位 10%',
    '下位 10%',
    '平均より上',
    '平均より下',
    'その他のルール...',
  ]) {
    await expect(topBottomMenu.getByText(label, { exact: true }).first()).toBeVisible();
  }
  await expect(topBottomMenu).not.toContainText('Top 10 Items');
  await expect(topBottomMenu).not.toContainText('Above Average');
  await conditionalMenu.locator('[data-cf-submenu="clear"]').hover();
  const clearRulesMenu = page.locator('#menu-conditional-clear');
  await expect(clearRulesMenu).toBeVisible();
  for (const label of ['選択したセルからルールをクリア', 'シート全体からルールをクリア']) {
    await expect(clearRulesMenu.getByText(label, { exact: true }).first()).toBeVisible();
  }
  await expect(clearRulesMenu).not.toContainText('Clear Rules from Selected Cells');
  await expect(clearRulesMenu).not.toContainText('Clear Rules from Entire Sheet');
  await conditionalMenu.locator('[data-cf-submenu="dataBar"]').hover();
  const dataBarMenu = page.locator('#menu-conditional-dataBar');
  await expect(dataBarMenu).toBeVisible();
  for (const label of ['塗りつぶし (グラデーション)', '塗りつぶし (単色)']) {
    await expect(dataBarMenu.getByText(label, { exact: true }).first()).toBeVisible();
  }
  for (const label of [
    '塗りつぶし (グラデーション)、青のデータ バー',
    '塗りつぶし (単色)、青のデータ バー',
  ]) {
    await expect(dataBarMenu.getByRole('menuitem', { name: label, exact: true })).toBeVisible();
  }
  await expect(dataBarMenu).not.toContainText('Gradient Fill');
  await conditionalMenu.locator('[data-cf-submenu="colorScale"]').hover();
  const colorScaleMenu = page.locator('#menu-conditional-colorScale');
  await expect(colorScaleMenu).toBeVisible();
  for (const label of [
    '緑 - 黄 - 赤のカラー スケール',
    '赤 - 黄 - 緑のカラー スケール',
    '青 - 白 - 赤のカラー スケール',
  ]) {
    await expect(colorScaleMenu.getByRole('menuitem', { name: label, exact: true })).toBeVisible();
  }
  await expect(colorScaleMenu).not.toContainText('Green - Yellow - Red Color Scale');
  await conditionalMenu.locator('[data-cf-submenu="iconSet"]').hover();
  const iconSetMenu = page.locator('#menu-conditional-iconSet');
  await expect(iconSetMenu).toBeVisible();
  for (const label of ['方向', '図形', 'インジケーター', '評価']) {
    await expect(iconSetMenu.getByText(label, { exact: true }).first()).toBeVisible();
  }
  for (const label of ['3 方向矢印', '3 色の信号', '3 フラグ', '5 評価']) {
    await expect(iconSetMenu.getByRole('menuitem', { name: label, exact: true })).toBeVisible();
  }
  await expect(iconSetMenu).not.toContainText('Directional');
  await expect(iconSetMenu).not.toContainText('3 Traffic Lights');
  await page.keyboard.press('Escape');
  await expect(conditionalMenu).toBeHidden();

  const findSelectButton = page.locator('[data-ribbon-command="findHome"]').first();
  await expect(findSelectButton).toBeVisible();
  await findSelectButton.click();
  const findSelectMenu = page.locator('#menu-find-select');
  await expect(findSelectMenu).toBeVisible();
  for (const label of [
    '検索...',
    '置換...',
    'ジャンプ...',
    '条件を選択してジャンプ...',
    '数式',
    '定数',
    '条件付き書式',
    'データの入力規則',
    'コメントとメモ',
  ]) {
    await expect(findSelectMenu.getByText(label, { exact: true }).first()).toBeVisible();
  }
  await expect(findSelectMenu).not.toContainText('Find...');
  await expect(findSelectMenu).not.toContainText('Go To Special');
  await findSelectMenu.locator('[data-find-select="find"]').click();
  const findDialog = page.locator('.fc-find');
  await expect(findDialog).toBeVisible();
  for (const label of [
    '検索と置換',
    '検索',
    '置換',
    '検索する文字列:',
    'オプション >>',
    'すべて検索',
    '前へ',
    '次へ',
    '閉じる',
  ]) {
    await expect(findDialog.getByText(label, { exact: true }).first()).toBeVisible();
  }
  await expect(findDialog).not.toContainText('Find and Replace');
  await expect(findDialog).not.toContainText('Find what:');
  await expect(findDialog).not.toContainText('Find All');
  await findDialog.getByRole('button', { name: 'オプション >>', exact: true }).click();
  for (const label of [
    'オプション <<',
    '検索場所:',
    '検索方向:',
    '検索対象:',
    '大文字/小文字を区別',
    'セル内容が完全に同一であるものを検索する',
    '書式...',
  ]) {
    await expect(findDialog.getByText(label, { exact: true }).first()).toBeVisible();
  }
  for (const label of ['シート', 'ブック']) {
    await expect(findDialog.locator('#fc-find-within')).toContainText(label);
  }
  for (const label of ['行', '列']) {
    await expect(findDialog.locator('#fc-find-search')).toContainText(label);
  }
  for (const label of ['数式', '値', 'コメント', 'メモ']) {
    await expect(findDialog.locator('#fc-find-look-in')).toContainText(label);
  }
  await expect(findDialog).not.toContainText('Within:');
  await expect(findDialog).not.toContainText('Search:');
  await expect(findDialog).not.toContainText('Look in:');
  await findDialog.getByRole('button', { name: 'オプション <<', exact: true }).click();
  await findDialog.locator('.fc-find__btn--close').click();
  await expect(findDialog).toBeHidden();

  await findSelectButton.click();
  await expect(findSelectMenu).toBeVisible();
  await findSelectMenu.locator('[data-find-select="replace"]').click();
  await expect(findDialog).toBeVisible();
  await expect(findDialog.locator('.fc-find__tab[aria-selected="true"]')).toHaveText('置換');
  for (const label of [
    '検索と置換',
    '検索する文字列:',
    '置換後の文字列:',
    'オプション >>',
    '置換',
    'すべて置換',
    '前へ',
    '次へ',
    '閉じる',
  ]) {
    await expect(findDialog.getByText(label, { exact: true }).first()).toBeVisible();
  }
  await expect(findDialog).not.toContainText('Replace with:');
  await expect(findDialog).not.toContainText('Replace all');
  await findDialog.locator('.fc-find__btn--close').click();
  await expect(findDialog).toBeHidden();

  await findSelectButton.click();
  await expect(findSelectMenu).toBeVisible();
  await findSelectMenu.locator('[data-find-select="go-to-special"]').click();
  const goToDialog = page.getByRole('dialog', { name: '選択オプション' });
  await expect(goToDialog).toBeVisible();
  for (const label of [
    '参照先',
    '範囲',
    'アクティブなシート',
    '現在の選択範囲',
    '種類',
    '空白セル',
    '空白以外のセル',
    '数式',
    '定数',
    '入力規則',
    '条件付き書式',
  ]) {
    await expect(goToDialog.getByText(label, { exact: true }).first()).toBeVisible();
  }
  await expect(goToDialog).not.toContainText('Go To Special');
  await expect(goToDialog).not.toContainText('Current selection');
  await expect(goToDialog).not.toContainText('Conditional formats');
  await goToDialog.getByRole('button', { name: 'キャンセル', exact: true }).click();
  await expect(goToDialog).toHaveCount(0);
  await page.keyboard.press('Escape');
  await expect(findSelectMenu).toBeHidden();

  const optionsButton = page.locator('.demo__account').getByRole('button', {
    name: 'オプション',
    exact: true,
  });
  await expect(optionsButton.first()).toBeVisible();
  await optionsButton.first().click();
  const optionsPanel = page.locator('.demo__panel');
  await expect(optionsPanel).toBeVisible();
  await expect(optionsPanel).toHaveAttribute('aria-label', 'オプション パネル');
  await expect(optionsPanel.getByRole('heading', { name: 'デモ表示' })).toBeVisible();
  await expect(optionsPanel.getByRole('heading', { name: 'プリセット' })).toBeVisible();
  await expect(optionsPanel.getByText('数式バー', { exact: true })).toBeVisible();
  await expect(optionsPanel.getByText('GREET("Workbook")', { exact: true })).toBeVisible();
  await expect(optionsPanel).not.toContainText('GREET("React")');
  await expect(optionsPanel).not.toContainText('GREET("Vue")');
  await optionsButton.first().click();
  await expect(optionsPanel).toBeHidden();

  await page.locator('[data-ribbon-tab="insert"]').click();
  const insertRibbon = page.getByRole('toolbar', { name: '挿入 リボン' });
  await expect(insertRibbon).toBeVisible();
  for (const name of [
    'ピボットテーブル',
    'テーブル',
    '画像',
    '図形',
    'スクリーンショット',
    'グラフ',
    'リンク (Ctrl+K)',
    'メモを挿入',
    '記号',
  ]) {
    await expect(insertRibbon.getByRole('button', { name, exact: true }).first()).toBeVisible();
  }
  await expect(insertRibbon).not.toContainText('Pictures');
  await expect(insertRibbon).not.toContainText('Shapes');
  await expect(insertRibbon).not.toContainText('Screenshot');
  const pivotTableButton = page.locator('[data-ribbon-command="pivotTableInsert"]').first();
  await expect(pivotTableButton).toBeVisible();
  await pivotTableButton.click();
  const pivotDialog = page.getByRole('dialog', { name: 'ピボットテーブルの作成' });
  await expect(pivotDialog).toBeVisible();
  for (const label of [
    '分析するデータを選択してください。',
    'テーブル/範囲',
    'ピボットテーブル レポートを配置する場所を選択してください。',
    '新規ワークシート',
    '既存のワークシート',
  ]) {
    await expect(pivotDialog.getByText(label, { exact: true }).first()).toBeVisible();
  }
  await expect(pivotDialog).not.toContainText('Create PivotTable');
  await expect(pivotDialog).not.toContainText('Table/Range');
  await pivotDialog.getByRole('button', { name: 'キャンセル', exact: true }).click();
  await expect(pivotDialog).toHaveCount(0);

  const tableButton = page.locator('[data-ribbon-command="formatTableInsert"]').first();
  await expect(tableButton).toBeVisible();
  await tableButton.click();
  const createTableDialog = page.getByRole('dialog', { name: 'テーブルの作成' });
  await expect(createTableDialog).toBeVisible();
  for (const label of [
    '表に変換するデータ範囲を指定してください。',
    '先頭行をテーブルの見出しとして使用する',
  ]) {
    await expect(createTableDialog.getByText(label, { exact: true }).first()).toBeVisible();
  }
  await expect(createTableDialog).not.toContainText('Create Table');
  await expect(createTableDialog).not.toContainText('My table has headers');
  await createTableDialog.getByRole('button', { name: 'キャンセル', exact: true }).click();
  await expect(createTableDialog).toHaveCount(0);

  await page.locator('[data-ribbon-tab="data"]').click();
  const dataRibbon = page.getByRole('toolbar', { name: 'データ リボン' });
  await expect(dataRibbon).toBeVisible();
  for (const name of [
    'フィルター',
    '昇順で並べ替え',
    '降順で並べ替え',
    'ユーザー設定の並べ替え...',
    '区切り位置',
    '重複の削除',
    'データの入力規則',
    'リンク',
    '選択した行または列をグループ化',
    '選択した行または列のグループ解除',
    'グループの詳細を表示',
    'グループの詳細を非表示',
  ]) {
    await expect(dataRibbon.getByRole('button', { name, exact: true }).first()).toBeVisible();
  }
  await expect(dataRibbon).not.toContainText('Sort & Filter');
  await expect(dataRibbon).not.toContainText('Remove Duplicates');
  await expect(dataRibbon).not.toContainText('Data Validation');
  const textToColumnsButton = page.locator('[data-ribbon-command="textToColumns"]').first();
  await expect(textToColumnsButton).toBeVisible();
  await textToColumnsButton.click();
  const textToColumnsDialog = page.getByRole('dialog', { name: '区切り位置指定ウィザード' });
  await expect(textToColumnsDialog).toBeVisible();
  for (const label of [
    '元のデータの形式',
    '区切り記号付き',
    '固定幅',
    '区切り文字',
    'タブ',
    'セミコロン',
    'カンマ',
    'スペース',
    'その他',
    '連続した区切り文字は 1 文字として扱う',
    'データのプレビュー',
  ]) {
    await expect(textToColumnsDialog.getByText(label, { exact: true }).first()).toBeVisible();
  }
  await expect(textToColumnsDialog).not.toContainText('Convert Text to Columns');
  await expect(textToColumnsDialog).not.toContainText('Original data type');
  await expect(textToColumnsDialog).not.toContainText('Delimiters');
  await textToColumnsDialog.getByRole('button', { name: 'キャンセル', exact: true }).click();
  await expect(textToColumnsDialog).toHaveCount(0);

  await page.locator('[data-ribbon-tab="formulas"]').click();
  const formulasRibbon = page.getByRole('toolbar', { name: '数式 リボン' });
  await expect(formulasRibbon).toBeVisible();
  for (const name of [
    '関数の挿入',
    'オートSUM (Σ)',
    '名前',
    '参照元',
    '参照先',
    '矢印の削除',
    'エラー チェック',
    '数式',
    '数式の検証',
    '再計算 (F9)',
    'オプション',
    'ウォッチ',
  ]) {
    await expect(formulasRibbon.getByRole('button', { name, exact: true }).first()).toBeVisible();
  }
  await expect(formulasRibbon).not.toContainText('Function Library');
  await expect(formulasRibbon).not.toContainText('Formula Auditing');
  await expect(formulasRibbon).not.toContainText('Calculation');
  const fxButton = page.locator('[data-ribbon-command="fx"]').first();
  await expect(fxButton).toBeVisible();
  await fxButton.click();
  const fxDialog = page.getByRole('dialog', { name: '関数の引数' });
  await expect(fxDialog).toBeVisible();
  await expect(fxDialog.getByPlaceholder('関数を検索…')).toBeVisible();
  for (const label of ['カテゴリを選択', '挿入', 'キャンセル']) {
    await expect(fxDialog.getByText(label, { exact: true }).first()).toBeVisible();
  }
  await expect(fxDialog).toContainText('数式の結果');
  for (const label of [
    'すべて',
    '最近使用した関数',
    '論理',
    '検索/行列',
    '文字列操作',
    '日付/時刻',
    '数学/三角',
    '財務',
  ]) {
    await expect(fxDialog.locator('.fc-fxdialog__category')).toContainText(label);
  }
  await expect(fxDialog).not.toContainText('Function Arguments');
  await expect(fxDialog).not.toContainText('Select a category');
  await expect(fxDialog).not.toContainText('Formula result');
  await fxDialog.getByRole('button', { name: 'キャンセル', exact: true }).click();
  await expect(fxDialog).toHaveCount(0);

  await page.locator('[data-ribbon-tab="pageLayout"]').click();
  const pageLayoutRibbon = page.getByRole('toolbar', { name: 'ページ レイアウト リボン' });
  await expect(pageLayoutRibbon).toBeVisible();
  for (const name of [
    'テーマ',
    '余白',
    '印刷の向き',
    '用紙サイズ',
    'ページ設定',
    '改ページ',
    '背景',
    '印刷タイトル',
    '横をページ数に合わせる',
    '縦をページ数に合わせる',
    '拡大縮小',
    '目盛線',
    '枠線を印刷',
    '見出し',
    '行列番号を印刷',
    '配置',
    'オブジェクトの選択と表示',
    '印刷',
  ]) {
    await expect(pageLayoutRibbon.getByRole('button', { name, exact: true }).first()).toBeVisible();
  }
  await expect(
    pageLayoutRibbon.getByRole('button', {
      name: '印刷範囲: 印刷範囲の設定/印刷範囲に追加/印刷範囲のクリア',
    }),
  ).toBeVisible();
  const ribbonOptionLabels = async (command: string): Promise<string[]> => {
    const raw = await page
      .locator(`[data-ribbon-command="${command}"]`)
      .first()
      .getAttribute('data-ribbon-options');
    return raw ? (JSON.parse(raw) as { label: string }[]).map((option) => option.label) : [];
  };
  for (const label of ['標準', '広い', '狭い', 'ユーザー設定']) {
    expect(await ribbonOptionLabels('marginsPreset')).toContain(label);
  }
  for (const label of ['縦', '横']) {
    expect(await ribbonOptionLabels('orientationPreset')).toContain(label);
  }
  for (const label of ['A4', 'A3', 'A5', 'レター', 'リーガル', 'タブロイド']) {
    expect(await ribbonOptionLabels('paperSizePreset')).toContain(label);
  }
  await expect(pageLayoutRibbon).not.toContainText('Page Setup');
  await expect(pageLayoutRibbon).not.toContainText('Print Area');
  await expect(pageLayoutRibbon).not.toContainText('Orientation');

  await page.locator('[data-ribbon-tab="view"]').click();
  const viewRibbon = page.getByRole('toolbar', { name: '表示 リボン' });
  await expect(viewRibbon).toBeVisible();
  for (const name of [
    '標準',
    'ページ レイアウト',
    '改ページ プレビュー',
    'ウォッチ',
    'ビュー',
    '保存',
    '削除',
    'オブジェクト',
    'ピボットテーブルのフィールド',
    '目盛線',
    '見出し',
    '数式',
    '数式バー',
    'R1C1参照形式',
    'ウィンドウ枠',
    '書式',
    'ズーム',
    '選択範囲に合わせる',
    'ズーム 75%',
    'ズーム 100%',
    'ズーム 125%',
    '保護',
  ]) {
    await expect(viewRibbon.getByRole('button', { name, exact: true }).first()).toBeVisible();
  }
  for (const label of ['現在の表示', 'ズーム...', '75%', '100%', '125%']) {
    await expect(viewRibbon.getByText(label, { exact: true }).first()).toBeVisible();
  }
  await expect(viewRibbon).not.toContainText('Page Break Preview');
  await expect(viewRibbon).not.toContainText('Formula Bar');
  await expect(viewRibbon).not.toContainText('Freeze');
  await page.locator('[data-ribbon-command="zoomDialog"]').first().click();
  const zoomDialog = page.getByRole('dialog', { name: 'ズーム' });
  await expect(zoomDialog).toBeVisible();
  await expect(zoomDialog.getByText('倍率', { exact: true }).first()).toBeVisible();
  await expect(zoomDialog).not.toContainText('Magnification');
  await zoomDialog.getByRole('button', { name: 'キャンセル', exact: true }).click();
  await expect(zoomDialog).toHaveCount(0);

  await page.locator('[data-ribbon-tab="file"]').click();
  const backstage = page.locator('.fc-tb__backstage[role="dialog"]').first();
  await expect(backstage).toBeVisible();
  await expect(backstage).toHaveAttribute('aria-label', 'ファイル');
  await expect(backstage.getByText('ブック · スプレッドシート レイアウト')).toBeVisible();
  for (const label of [
    '情報',
    '新規',
    '開く',
    '保存',
    '名前を付けて保存',
    '印刷',
    '共有',
    'エクスポート',
    'オプション',
    '閉じる',
  ]) {
    await expect(backstage.getByRole('button', { name: label, exact: true }).first()).toBeVisible();
  }
  await expect(backstage).not.toContainText('React');
  await expect(backstage).not.toContainText('Vue');
  await backstage.getByRole('button', { name: '印刷', exact: true }).first().click();
  await expect(backstage.getByRole('heading', { name: '印刷' })).toBeVisible();
  await expect(backstage.getByRole('button', { name: 'PDF にエクスポート' })).toBeVisible();
  const backstagePageSetup = backstage.getByRole('button', { name: 'ページ設定' });
  await expect(backstagePageSetup).toBeVisible();
  await expect(backstage.getByText('アクティブ シート', { exact: true })).toBeVisible();
  await expect(backstage.getByText('印刷の向き', { exact: true })).toBeVisible();
  await expect(backstage.getByText('縦方向', { exact: true })).toBeVisible();
  await expect(backstage.getByText('印刷範囲なし', { exact: true })).toBeVisible();
  await expect(backstage).not.toContainText('portrait');
  await expect(backstage).not.toContainText('landscape');
  await backstagePageSetup.click();
  const pageSetupDialog = page.getByRole('dialog', { name: 'ページ設定' });
  await expect(pageSetupDialog).toBeVisible();
  for (const label of [
    'ページ',
    '余白',
    'ヘッダー/フッター',
    'シート',
    '印刷の向き',
    'プリンター',
    '用紙サイズ',
  ]) {
    await expect(pageSetupDialog.getByText(label, { exact: true }).first()).toBeVisible();
  }
  await pageSetupDialog.getByRole('tab', { name: '余白', exact: true }).click();
  await expect(pageSetupDialog.getByText('ページ中央', { exact: true }).first()).toBeVisible();
  await expect(pageSetupDialog).not.toContainText('Page Setup');
  await expect(pageSetupDialog).not.toContainText('Orientation');
  await expect(pageSetupDialog).not.toContainText('Margins');
  await pageSetupDialog.getByRole('button', { name: 'キャンセル', exact: true }).click();
  await expect(pageSetupDialog).toHaveCount(0);
  await backstage.getByRole('button', { name: '閉じる', exact: true }).click();
  await expect(backstage).toBeHidden();

  await page.locator('[data-ribbon-tab="review"]').click();
  const reviewRibbon = page.getByRole('toolbar', { name: '校閲 リボン' });
  await expect(reviewRibbon).toBeVisible();
  for (const name of [
    'スペル チェック',
    'アクセシビリティ',
    '翻訳',
    'メモを挿入',
    'コメントまたはメモの削除',
    '前のコメントまたはメモ',
    '次のコメントまたはメモ',
    '検索 (Ctrl+F)',
    '保護',
    'ブックの保護...',
    '範囲の編集を許可...',
  ]) {
    await expect(reviewRibbon.getByRole('button', { name, exact: true }).first()).toBeVisible();
  }
  await expect(reviewRibbon).not.toContainText('Spelling');
  await expect(reviewRibbon).not.toContainText('Accessibility');
  await expect(reviewRibbon).not.toContainText('Translate');
  await expect(page.locator('[data-ribbon-command="deleteCommentReview"]').first()).toHaveAttribute(
    'data-ribbon-activation',
    'splitPrimary',
  );
  await expect(page.locator('[data-ribbon-command="deleteCommentReview"]').first()).toHaveAttribute(
    'data-ribbon-menu-id',
    'menu-review-comments',
  );
  await expect(page.locator('[data-ribbon-command="protectReview"]').first()).toHaveAttribute(
    'data-ribbon-activation',
    'splitPrimary',
  );
  await expect(page.locator('[data-ribbon-command="protectReview"]').first()).toHaveAttribute(
    'data-ribbon-menu-id',
    'menu-protect-review',
  );

  await page.locator('[data-ribbon-command="accessibility"]').first().click();
  const accessibilityDialog = page.getByRole('dialog', { name: 'アクセシビリティ チェック' });
  await expect(accessibilityDialog).toBeVisible();
  await expect(accessibilityDialog).toContainText(/問題は見つかりませんでした。|リボン コマンド/);
  await expect(accessibilityDialog).not.toContainText('Accessibility Check');
  await accessibilityDialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(accessibilityDialog).toHaveCount(0);

  await page.locator('[data-ribbon-command="translateReview"]').first().click();
  const translateDialog = page.getByRole('dialog', { name: '翻訳' });
  await expect(translateDialog).toBeVisible();
  await expect(translateDialog).toContainText('このデモには翻訳サービスが接続されていません。');
  await expect(translateDialog).not.toContainText('No translation service');
  await translateDialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(translateDialog).toHaveCount(0);

  await page.locator('[data-ribbon-command="newCommentReview"]').first().click();
  const commentDialog = page.getByRole('dialog', { name: 'メモを挿入' });
  await expect(commentDialog).toBeVisible();
  await expect(commentDialog.getByPlaceholder('メモを入力')).toBeVisible();
  await expect(commentDialog).not.toContainText('Insert Note');
  await commentDialog.getByRole('button', { name: 'キャンセル', exact: true }).click();
  await expect(commentDialog).toHaveCount(0);

  await page.evaluate(() => {
    const inst = (
      window as unknown as {
        __fcInst?: {
          workbook?: {
            setText?: (addr: { sheet: number; row: number; col: number }, value: string) => void;
          };
        };
      }
    ).__fcInst;
    if (!inst?.workbook?.setText) {
      throw new Error('window.__fcInst.workbook.setText is required');
    }
    inst.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'teh teh');
  });
  const spellingCommand = page.getByRole('button', { name: /スペル/ }).first();
  await expect(spellingCommand).toBeVisible();
  await spellingCommand.click();
  const dialog = page.getByRole('dialog', { name: 'スペル チェック' });
  await expect(dialog).toBeVisible();
  await expect(dialog).toContainText('同じ語が繰り返されています');
  await expect(dialog).toContainText('スペルミスの可能性');
  await expect(dialog).not.toContainText('Spelling Review');
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(dialog).toHaveCount(0);

  await page.locator('[data-ribbon-tab="help"]').click();
  const helpRibbon = page.getByRole('toolbar', { name: 'ヘルプ リボン' });
  await expect(helpRibbon).toBeVisible();
  await expect(helpRibbon.getByRole('button', { name: 'ヘルプ', exact: true })).toBeVisible();
  await expect(page.locator('[data-ribbon-command="helpSearch"]').first()).toHaveAttribute(
    'data-ribbon-activation',
    'disabled',
  );
  await expect(helpRibbon).not.toContainText('Help');

  const lang = await page.evaluate(() => document.documentElement.lang);
  expect(lang === 'ja' || lang === '').toBe(true);
  expect(await sp.isCrossOriginIsolated()).toBe(true);

  const enToggle = page.getByRole('button', { name: 'EN', exact: true });
  if ((await enToggle.count()) > 0) {
    await enToggle.first().click();
    await expect(page.getByRole('tab', { name: 'Home', exact: true })).toBeVisible();
    await expect(page.getByRole('tab', { name: 'Insert', exact: true })).toBeVisible();
    await expect(page.getByRole('button', { name: 'Paste', exact: true }).first()).toBeVisible();
    await expect(searchBox).toHaveAttribute('aria-label', 'Search commands');
  } else {
    await page.goto('/?locale=en');
    await sp.waitForReady();
    await expect(page.getByRole('tab', { name: 'Home', exact: true })).toBeVisible();
    await expect(page.getByRole('tab', { name: 'Insert', exact: true })).toBeVisible();
    await expect(page.getByRole('button', { name: 'Paste', exact: true }).first()).toBeVisible();
  }
}

/** I02 — `?theme=dark` boots the app in the `ink` core theme.
 *  Observable via the host's `data-fc-theme` attribute. */
export async function runThemeBootScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await page.goto('/?theme=dark');
  await sp.waitForReady();

  const themeAttr = await page.evaluate(() => {
    const host = document.querySelector('.fc-host') as HTMLElement | null;
    return host?.dataset.fcTheme ?? null;
  });
  expect(themeAttr === 'ink' || themeAttr === 'dark').toBe(true);
}
