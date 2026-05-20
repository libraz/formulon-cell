import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import {
  backstageMenuText,
  buildRibbonModel,
  conditionalMenuText,
  EXCEL365_STANDARD_RIBBON_TABS,
  fluentIconPaths,
  OPTIONAL_RIBBON_TABS,
  pageScaleMenuText,
  ribbonDisplayText,
  toolbarMenuText,
  viewToggleMenuText,
} from '../../../src/index.js';

type ReactRibbonControlKind = 'tool' | 'select' | 'color' | 'break';

interface RibbonControl {
  id: string;
  kind: ReactRibbonControlKind;
}

const reactToolbarSource = (name: string): string => {
  const sourcePath = [
    resolve(process.cwd(), `../formulon-cell-react/src/toolbar/${name}`),
    resolve(process.cwd(), `packages/formulon-cell-react/src/toolbar/${name}`),
  ].find((candidate) => existsSync(candidate));
  if (!sourcePath) throw new Error(`React toolbar source not found: ${name}`);
  return readFileSync(sourcePath, 'utf8');
};

const reactRibbonControls = (): RibbonControl[] => {
  const source = `${reactToolbarSource('groups.tsx')}\n${reactToolbarSource('add-in-groups.tsx')}`;
  const controls: (RibbonControl & { index: number })[] = [];
  const re = /\b(tool|select|optionSelect|color|rowBreak)(?:<[^>]+>)?\(\s*['"]([^'"]+)/g;
  for (const match of source.matchAll(re)) {
    const [, kind, id] = match;
    if (!kind || !id) continue;
    controls.push({
      id,
      index: match.index,
      kind:
        kind === 'rowBreak'
          ? 'break'
          : kind === 'optionSelect'
            ? 'select'
            : (kind as ReactRibbonControlKind),
    });
  }
  const mergeIndex = source.lastIndexOf('mergeMenu');
  if (mergeIndex >= 0) {
    const insertAt = controls.findIndex((control) => control.index > mergeIndex);
    controls.splice(insertAt >= 0 ? insertAt : controls.length, 0, {
      id: 'merge',
      index: mergeIndex,
      kind: 'select',
    });
  }
  const menuProps: Record<string, RibbonControl> = {
    autosumFormulaMenu: { id: 'autosumFormula', kind: 'tool' },
    autosumMenu: { id: 'autosum', kind: 'tool' },
    addInMenu: { id: 'addIn', kind: 'tool' },
    calcOptionsMenu: { id: 'calcOptions', kind: 'tool' },
    cellDeleteMenu: { id: 'deleteRows', kind: 'tool' },
    cellFormatMenu: { id: 'formatCellsHome', kind: 'tool' },
    cellInsertMenu: { id: 'insertRows', kind: 'tool' },
    cellStylesMenu: { id: 'cellStyles', kind: 'tool' },
    chartMenu: { id: 'chartInsert', kind: 'tool' },
    conditionalMenu: { id: 'conditional', kind: 'tool' },
    clearArrowsMenu: { id: 'clearArrows', kind: 'tool' },
    dataFilterMenu: { id: 'filter', kind: 'tool' },
    dataSortMenu: { id: 'sortData', kind: 'tool' },
    dataValidationMenu: { id: 'dataValidation', kind: 'tool' },
    definedNamesMenu: { id: 'namedRanges', kind: 'tool' },
    deleteCommentMenu: { id: 'deleteCommentReview', kind: 'tool' },
    errorCheckingMenu: { id: 'errorChecking', kind: 'tool' },
    freezeMenu: { id: 'freeze', kind: 'tool' },
    formatTableHomeMenu: { id: 'formatTableHome', kind: 'tool' },
    formatTableInsertMenu: { id: 'formatTableInsert', kind: 'tool' },
    clearMenu: { id: 'clearFormat', kind: 'tool' },
    fillMenu: { id: 'fillHome', kind: 'tool' },
    findMenu: { id: 'findHome', kind: 'tool' },
    pageBreaksMenu: { id: 'pageBreaks', kind: 'tool' },
    pdfMenu: { id: 'pdf', kind: 'tool' },
    pivotTableMenu: { id: 'pivotTableInsert', kind: 'tool' },
    pasteMenu: { id: 'paste', kind: 'tool' },
    pictureInsertMenu: { id: 'pictureInsert', kind: 'tool' },
    printAreaMenu: { id: 'printArea', kind: 'tool' },
    printTitlesMenu: { id: 'printTitles', kind: 'tool' },
    protectionMenu: { id: 'protectionReview', kind: 'tool' },
    sheetBackgroundMenu: { id: 'sheetBackground', kind: 'tool' },
    shapesInsertMenu: { id: 'shapesInsert', kind: 'tool' },
    screenshotInsertMenu: { id: 'screenshotInsert', kind: 'tool' },
    sortMenu: { id: 'sortFilterHome', kind: 'tool' },
    symbolMenu: { id: 'symbolInsert', kind: 'tool' },
    themeMenu: { id: 'pageTheme', kind: 'tool' },
    textOrientationMenu: { id: 'textOrientation', kind: 'tool' },
    textToColumnsMenu: { id: 'textToColumns', kind: 'tool' },
    functionLogicalMenu: { id: 'ifFormula', kind: 'tool' },
    functionLookupMenu: { id: 'xlookupFormula', kind: 'tool' },
    functionTextMenu: { id: 'concatFormula', kind: 'tool' },
    functionDateTimeMenu: { id: 'todayFormula', kind: 'tool' },
    functionFinancialMenu: { id: 'pmtFormula', kind: 'tool' },
    functionMathTrigMenu: { id: 'roundFormula', kind: 'tool' },
    hyperlinkMenu: { id: 'hyperlinkInsert', kind: 'tool' },
    outlineGroupMenu: { id: 'outlineGroup', kind: 'tool' },
    outlineUngroupMenu: { id: 'outlineUngroup', kind: 'tool' },
    watchMenu: { id: 'watch', kind: 'tool' },
    watchViewMenu: { id: 'watchView', kind: 'tool' },
    windowMenu: { id: 'windowVisibility', kind: 'tool' },
  };
  for (const [prop, control] of Object.entries(menuProps)) {
    const index = source.lastIndexOf(prop);
    if (index < 0) continue;
    const insertAt = controls.findIndex((entry) => entry.index > index);
    controls.splice(insertAt >= 0 ? insertAt : controls.length, 0, { ...control, index });
  }
  return controls.map(({ id, kind }) => ({ id, kind }));
};

const modelControls = (): RibbonControl[] =>
  buildRibbonModel('en').flatMap((tab) =>
    tab.groups.flatMap((group) =>
      group.commands.map((command) => ({
        id: command.id,
        kind:
          command.kind === 'select' || command.kind === 'color'
            ? command.kind
            : command.kind === 'break'
              ? 'break'
              : 'tool',
      })),
    ),
  );

describe('toolbar/ribbon-model', () => {
  it('keeps the shared ribbon command surface aligned with the React toolbar', () => {
    const byId = (a: RibbonControl, b: RibbonControl): number => a.id.localeCompare(b.id);

    expect([...modelControls()].sort(byId)).toEqual([...reactRibbonControls()].sort(byId));
  });

  it('can project the Excel 365 standard tab surface without optional add-in tabs', () => {
    const tabs = buildRibbonModel('en', { tabs: EXCEL365_STANDARD_RIBBON_TABS }).map(
      (tab) => tab.id,
    );

    expect(tabs).toEqual([
      'file',
      'home',
      'insert',
      'pageLayout',
      'formulas',
      'data',
      'review',
      'view',
      'help',
    ]);
    expect(tabs).not.toEqual(expect.arrayContaining([...OPTIONAL_RIBBON_TABS]));
  });

  it('keeps Excel 365 command placement for names, duplicates, and links', () => {
    const tabs = new Map(buildRibbonModel('en').map((tab) => [tab.id, tab]));
    const commandIds = (tab: string): string[] =>
      tabs
        .get(tab as never)
        ?.groups.flatMap((group) => group.commands.map((command) => command.id)) ?? [];

    expect(commandIds('insert')).not.toEqual(
      expect.arrayContaining(['namedRangesInsert', 'removeDupesInsert', 'linksInsert']),
    );
    expect(commandIds('formulas')).toEqual(expect.arrayContaining(['namedRanges']));
    expect(commandIds('data')).toEqual(expect.arrayContaining(['removeDupes', 'linksData']));
    expect(commandIds('insert')).toEqual(expect.arrayContaining(['hyperlinkInsert']));
    expect(commandIds('pageLayout')).toEqual(
      expect.arrayContaining(['arrangeObjectsPageLayout', 'selectionPanePageLayout']),
    );
    const commands = new Map(
      buildRibbonModel('en')
        .flatMap((tab) => tab.groups)
        .flatMap((group) => group.commands)
        .map((command) => [command.id, command]),
    );
    expect(commands.get('formatTableInsert')).toMatchObject({
      label: 'Table',
      title: 'Table',
    });
    expect(commands.get('formatTableHome')).toMatchObject({
      label: 'Format as Table',
      title: 'Format as Table',
    });
    expect(tabs.get('review')?.groups.map((group) => group.title)).toEqual([
      'Proofing',
      'Accessibility',
      'Language',
      'Comments',
      'Find',
      'Protection',
    ]);
  });

  it('has Fluent SVG paths for every icon used by the ribbon model', () => {
    const missing = buildRibbonModel('en')
      .flatMap((tab) => tab.groups)
      .flatMap((group) => group.commands)
      .filter((command) => command.icon && !fluentIconPaths(command.icon))
      .map((command) => `${command.id}:${command.icon}`);

    expect(missing).toEqual([]);
  });

  it('localizes ribbon model command titles for Japanese Office-like surfaces', () => {
    const commands = new Map(
      buildRibbonModel('ja')
        .flatMap((tab) => tab.groups)
        .flatMap((group) => group.commands)
        .map((command) => [command.id, command.title]),
    );

    expect(commands.get('numberFormat')).toBe('数値');
    expect(commands.get('bold')).toBe('太字 (⌘B)');
    expect(commands.get('borderPreset')).toBe('罫線パターン');
    expect(commands.get('borderStyle')).toBe('罫線のスタイル');
    const borderPreset = buildRibbonModel('ja')
      .flatMap((tab) => tab.groups)
      .flatMap((group) => group.commands)
      .find((command) => command.id === 'borderPreset');
    expect(borderPreset?.options?.map((option) => option.label)).toEqual(
      expect.arrayContaining([
        '内側',
        '内側横罫線',
        '内側縦罫線',
        '右下がり斜め罫線',
        '右上がり斜め罫線',
      ]),
    );
    expect(commands.get('moreBorders')).toBe('その他の罫線...');
    expect(commands.get('drawBorder')).toBe('罫線の作成');
    expect(commands.get('drawBorderGrid')).toBe('罫線グリッドの作成');
    expect(commands.get('eraseBorder')).toBe('罫線の削除');
    const borderStyle = buildRibbonModel('ja')
      .flatMap((tab) => tab.groups)
      .flatMap((group) => group.commands)
      .find((command) => command.id === 'borderStyle');
    expect(borderStyle?.options?.map((option) => option.label)).toEqual(
      expect.arrayContaining(['極細線', '中太破線', '一点鎖線', '中太一点鎖線', '二点鎖線']),
    );
    expect(commands.get('pageSetupAdvanced')).toBe('ページ設定');
    expect(commands.get('sum')).toBe('SUM の引数');
    expect(commands.get('sortAsc')).toBe('昇順で並べ替え');
    expect(commands.get('outlineGroup')).toBe('選択した行または列をグループ化');
    expect(commands.get('deleteCommentReview')).toBe('コメントまたはメモの削除');
    expect(commands.get('viewNormal')).toBe('標準');
    expect(commands.get('viewPageLayout')).toBe('ページ レイアウト');
    expect(commands.get('viewPageBreakPreview')).toBe('改ページ プレビュー');
    expect(commands.get('showFormulasFormula')).toBe('数式');
    expect(commands.get('errorChecking')).toBe('エラー チェック');
    expect(commands.get('evaluateFormula')).toBe('数式の検証');
    expect(commands.get('watchView')).toBe('ウォッチ');
    expect(commands.get('zoom100')).toBe('ズーム 100%');
    expect(commands.get('protect')).toBe('保護');
  });

  it('does not expose internal row-break ids as command titles', () => {
    const breaks = buildRibbonModel('ja')
      .flatMap((tab) => tab.groups)
      .flatMap((group) => group.commands)
      .filter((command) => command.kind === 'break');

    expect(breaks.map((command) => command.title)).toEqual(breaks.map(() => ''));
  });

  it('keeps Japanese ribbon command titles free of untranslated English phrases', () => {
    const allowed =
      /^(Acrobat|PDF|R1C1|fx|SUM|AVG|AVERAGE|IF|XLOOKUP|CONCAT|TODAY|PMT|ROUND|A-Z|Z-A|\$|%|,|\.0|\.00|B|I|U|S|\d+%)$/;
    const allowedInline =
      /^(?:オートSUM \(Σ\)|(?:SUM|AVERAGE|IF|XLOOKUP|CONCAT|TODAY|PMT|ROUND) の引数)$/;
    const untranslated = buildRibbonModel('ja')
      .flatMap((tab) => tab.groups)
      .flatMap((group) => group.commands)
      .filter((command) => /[A-Za-z]{3,}/.test(command.title))
      .filter((command) => !allowed.test(command.title) && !allowedInline.test(command.title))
      .map((command) => `${command.id}:${command.title}`);

    expect(untranslated).toEqual([]);
  });

  it('keeps shared ribbon menu labels localized for React and Vue wrappers', () => {
    expect(toolbarMenuText('ja').validationSettings).toBe('データの入力規則...');
    expect(toolbarMenuText('en').validationSettings).toBe('Data Validation...');
    expect(toolbarMenuText('ja').errorChecking).toBe('エラー チェック...');
    expect(toolbarMenuText('en').traceError).toBe('Trace Error');
    expect(toolbarMenuText('ja').sortCustom).toBe('ユーザー設定の並べ替え...');
    expect(toolbarMenuText('en').sortCustom).toBe('Custom Sort...');
    expect(toolbarMenuText('ja').symbolGreek).toBe('ギリシャ文字');
    expect(toolbarMenuText('en').symbolMore).toBe('More Symbols...');
    expect(toolbarMenuText('ja').autosumMoreFunctions).toBe('その他の関数...');
    expect(toolbarMenuText('en').autosumMoreFunctions).toBe('More Functions...');
    expect(toolbarMenuText('ja').tableStyleDark).toBe('濃色');
    expect(toolbarMenuText('en').tableStyleDark).toBe('Dark');
    expect(toolbarMenuText('ja').scriptCommandPrompt).toContain(
      'スクリプト コマンドを入力してください',
    );
    expect(toolbarMenuText('ja').autoFitRowHeight).toBe('行の高さの自動調整');
    expect(toolbarMenuText('en').autoFitColWidth).toBe('AutoFit Column Width');
    expect(toolbarMenuText('ja').orientationAngleCounterclockwise).toBe('左回りに回転');
    expect(toolbarMenuText('en').orientationHorizontalText).toBe('Horizontal Text');
    expect(toolbarMenuText('en').scriptCommandPrompt).toBe(
      'Script command: uppercase, lowercase, trim, clear',
    );
    expect(toolbarMenuText('ja').automationRunStatus).toBe(
      'スクリプト · {count} 個のセルを変更しました',
    );
    expect(toolbarMenuText('en').automationRunStatus).toBe('Script · {count} cell(s) changed');
    expect(conditionalMenuText('ja').iconSets).toBe('アイコン セット');
    expect(conditionalMenuText('en').iconSets).toBe('Icon Sets');
    expect(conditionalMenuText('ja').unique).toBe('一意の値...');
    expect(conditionalMenuText('ja').equal).toBe('指定の値に等しい...');
    expect(conditionalMenuText('en').textContains).toBe('Text that Contains...');
    expect(conditionalMenuText('ja').datePrompt).toBe(
      '日付条件: 昨日、今日、明日、過去 7 日間、先週、今週、来週、先月、今月、来月',
    );
    expect(conditionalMenuText('ja').datePrompt).not.toContain('last7');
    expect(conditionalMenuText('en').top10Percent).toBe('Top 10%');
    expect(conditionalMenuText('ja').dataBarSolidGreen).toBe('塗りつぶし (単色)、緑のデータ バー');
    expect(conditionalMenuText('en').dataBarSolidGreen).toBe('Solid Fill, Green Data Bar');
    expect(conditionalMenuText('ja').colorScaleGreenWhiteGreen).toBe(
      '緑 - 白 - 緑のカラー スケール',
    );
    expect(conditionalMenuText('en').colorScaleGreenWhiteGreen).toBe(
      'Green - White - Green Color Scale',
    );
    expect(conditionalMenuText('ja').iconTraffic3).toBe('3 色の信号');
    expect(conditionalMenuText('en').iconStars3).toBe('3 Stars');
    expect(ribbonDisplayText('ja').label).toBe('リボンの表示オプション');
    expect(ribbonDisplayText('en').collapsed).toBe('Show tabs only');
    expect(ribbonDisplayText('en').singleLine).toBe('Single Line Ribbon');
    expect(ribbonDisplayText('en').autoHide).toBe('Auto-hide Ribbon');
    expect(backstageMenuText('ja').back).toBe('戻る');
    expect(backstageMenuText('en').workbookInfo).toBe('Workbook Information');
    expect(pageScaleMenuText('ja').automatic).toBe('自動');
    expect(pageScaleMenuText('en').pages).toBe('pages');
    expect(pageScaleMenuText('ja').custom).toBe('ユーザー設定...');
    expect(pageScaleMenuText('en').invalidScale).toBe('Enter a scale from 10 to 400.');
    expect(viewToggleMenuText('ja').formulaBar).toBe('数式バー');
    expect(viewToggleMenuText('en').gridlines).toBe('Gridlines');
  });
});
