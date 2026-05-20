import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  backstageMenuText,
  buildRibbonModel,
  conditionalMenuText,
  EXCEL365_STANDARD_RIBBON_TABS,
  fluentIconPaths,
  HOME_MIXED_LAYOUT_GROUP_VARIANTS,
  HOME_STACKED_LAYOUT_GROUP_VARIANTS,
  HOME_TILE_LAYOUT_GROUP_VARIANTS,
  OPTIONAL_RIBBON_TABS,
  pageScaleMenuText,
  ribbonActivatableCommandIds,
  ribbonActivatableSurfaceCommandIds,
  ribbonCommandIds,
  ribbonCommands,
  ribbonDisplayText,
  ribbonSurfaceCommandIds,
  ribbonTabCommandIds,
  toolbarMenuText,
  viewToggleMenuText,
} from '../../../src/index.js';

describe('toolbar/ribbon-model', () => {
  const ribbonStylesDir = join(process.cwd(), 'src/styles/toolbar/ribbon');

  it('keeps the shared core ribbon command surface unique within each tab', () => {
    const keys = buildRibbonModel('en').flatMap((tab) =>
      tab.groups.flatMap((group) =>
        group.commands.map((command) => `${tab.id}:${command.id}:${command.kind}`),
      ),
    );
    const duplicates = keys.filter((key, index) => keys.indexOf(key) !== index);

    expect(duplicates).toEqual([]);
  });

  it('keeps command surfaces locale-independent', () => {
    expect(ribbonCommandIds('ja')).toEqual(ribbonCommandIds('en'));
    expect(ribbonActivatableCommandIds('ja')).toEqual(ribbonActivatableCommandIds('en'));
    expect(ribbonSurfaceCommandIds()).toEqual(ribbonCommandIds('en'));
    expect(ribbonActivatableSurfaceCommandIds()).toEqual(ribbonActivatableCommandIds('en'));
    expect(
      ribbonSurfaceCommandIds({ tabs: EXCEL365_STANDARD_RIBBON_TABS }),
    ).toEqual(ribbonCommandIds('en', { tabs: EXCEL365_STANDARD_RIBBON_TABS }));
    expect(
      ribbonActivatableSurfaceCommandIds({ tabs: EXCEL365_STANDARD_RIBBON_TABS }),
    ).toEqual(ribbonActivatableCommandIds('en', { tabs: EXCEL365_STANDARD_RIBBON_TABS }));
  });

  it('keeps the Excel 365 Home tab group and command placement explicit', () => {
    const home = buildRibbonModel('en').find((tab) => tab.id === 'home');

    expect(home?.groups.map((group) => group.variant)).toEqual([
      'clipboard',
      'font',
      'alignment',
      'number',
      'styles',
      'cells',
      'editing',
    ]);

    const commandsByGroup = new Map(
      home?.groups.map((group) => [
        group.variant,
        group.commands.map((command) => command.id),
      ]) ?? [],
    );

    expect(commandsByGroup.get('clipboard')).toEqual([
      'paste',
      'cut',
      'copy',
      'formatPainter',
    ]);
    expect(commandsByGroup.get('font')).toEqual([
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
      'fillColor',
      'fontColor',
    ]);
    expect(commandsByGroup.get('alignment')).toEqual([
      'top',
      'middle',
      'bottomAlign',
      'alignL',
      'alignC',
      'alignR',
      'alignment-row-2',
      'textOrientation',
      'wrap',
      'indentDecrease',
      'indentIncrease',
      'merge',
    ]);
    expect(commandsByGroup.get('number')).toEqual([
      'numberFormat',
      'number-row-2',
      'currency',
      'percent',
      'comma',
      'decDown',
      'decUp',
    ]);
    expect(commandsByGroup.get('styles')).toEqual([
      'conditional',
      'formatTableHome',
      'cellStyles',
    ]);
    expect(commandsByGroup.get('cells')).toEqual([
      'insertRows',
      'deleteRows',
      'formatCellsHome',
    ]);
    expect(commandsByGroup.get('editing')).toEqual([
      'autosum',
      'fillHome',
      'clearFormat',
      'sortFilterHome',
      'findHome',
    ]);
  });

  it('keeps Home wide-button groups on the shared tile layout CSS', () => {
    const groupsCss = readFileSync(join(ribbonStylesDir, 'groups.css'), 'utf8');

    for (const variant of HOME_TILE_LAYOUT_GROUP_VARIANTS) {
      expect(groupsCss).toContain(`.demo__ribbon-group--${variant} .demo__ribbon-tools`);
    }
    expect(groupsCss).toContain('.demo__ribbon-group--tiles .demo__rb--wide');
    expect(groupsCss).toContain('.demo__ribbon-group--tiles .demo__rb-icon');
    expect(groupsCss).toContain('flex-wrap: nowrap;');
    expect(groupsCss).toContain('height: 66px;');
    expect(groupsCss).toContain('.demo__ribbon-group--tiles .demo__rb .demo__rb-split-chevron');
    expect(groupsCss).toContain('position: absolute;');
    expect(groupsCss).toContain('bottom: 5px;');
  });

  it('loads group-specific ribbon CSS after generic button sizing', () => {
    for (const cssFile of ['../ribbon.css', '../../toolbar.css']) {
      const css = readFileSync(join(ribbonStylesDir, cssFile), 'utf8');
      const generic = css.indexOf('layout-and-buttons.css');
      const groups = css.indexOf('groups.css');
      expect(generic, cssFile).toBeGreaterThanOrEqual(0);
      expect(groups, cssFile).toBeGreaterThan(generic);
    }
  });

  it('keeps Home dense layout widths large enough to avoid hidden row wraps', () => {
    const groupsCss = readFileSync(join(ribbonStylesDir, 'groups.css'), 'utf8');
    const home = buildRibbonModel('en').find((tab) => tab.id === 'home');
    const tileWidth = 74;
    const tileGap = 4;
    const widthFor = (variant: string): number => {
      const match = new RegExp(
        String.raw`\.demo__ribbon-group--${variant} \.demo__ribbon-tools\s*\{\s*width:\s*(\d+)px;`,
        'm',
      ).exec(groupsCss);
      if (!match?.[1]) throw new Error(`missing ${variant} ribbon tool width`);
      return Number(match[1]);
    };

    for (const variant of HOME_TILE_LAYOUT_GROUP_VARIANTS) {
      const commandCount =
        home?.groups.find((group) => group.variant === variant)?.commands.length ?? 0;
      const requiredWidth = commandCount * tileWidth + Math.max(0, commandCount - 1) * tileGap;
      expect(widthFor(variant)).toBeGreaterThanOrEqual(requiredWidth);
    }
    expect(widthFor('cells')).toBeGreaterThanOrEqual(92);
    expect(widthFor('editing')).toBeGreaterThanOrEqual(238);
    expect(groupsCss).toContain('.demo__ribbon-group--stacked .demo__ribbon-tools');
    expect(groupsCss).toContain('.demo__ribbon-group--mixed .demo__ribbon-tools');
    expect(groupsCss).toContain('grid-template-rows: repeat(3, 22px);');
    expect(groupsCss).toContain(
      '.demo__ribbon-group--mixed .demo__rb--wide:not(.demo__rb--stacked) .demo__rb-split-chevron',
    );
    expect(groupsCss).toContain(
      '.demo__ribbon-group--stacked .demo__rb--stacked .demo__rb-split-chevron',
    );
    expect(groupsCss).toContain('margin: 0 0 0 auto;');
  });

  it('keeps Home tile-layout groups backed only by wide ribbon commands', () => {
    const home = buildRibbonModel('en').find((tab) => tab.id === 'home');
    const nonWideCommands =
      home?.groups
        .filter((group) =>
          HOME_TILE_LAYOUT_GROUP_VARIANTS.includes(
            group.variant as (typeof HOME_TILE_LAYOUT_GROUP_VARIANTS)[number],
          ),
        )
        .flatMap((group) =>
          group.commands
            .filter((command) => command.kind !== 'wide')
            .map((command) => `${group.variant}:${command.id}:${command.kind ?? 'button'}`),
        ) ?? [];

    expect(nonWideCommands).toEqual([]);
  });

  it('keeps Home stacked and mixed groups backed by shared layout metadata', () => {
    const home = buildRibbonModel('en').find((tab) => tab.id === 'home');
    const groups = new Map(home?.groups.map((group) => [group.variant, group]) ?? []);

    expect(HOME_STACKED_LAYOUT_GROUP_VARIANTS).toEqual(['cells']);
    expect(HOME_MIXED_LAYOUT_GROUP_VARIANTS).toEqual(['editing']);
    expect(groups.get('cells')?.commands.map((command) => command.layout)).toEqual([
      'stacked',
      'stacked',
      'stacked',
    ]);
    expect(groups.get('editing')?.commands.map((command) => command.layout ?? 'large')).toEqual([
      'stacked',
      'stacked',
      'stacked',
      'large',
      'large',
    ]);
    expect(
      [...(groups.get('cells')?.commands ?? []), ...(groups.get('editing')?.commands ?? [])].map(
        (command) => command.className ?? '',
      ),
    ).toEqual(['', '', '', '', '', '', '', '']);
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
    const commandIds = (tab: Parameters<typeof ribbonTabCommandIds>[1]): string[] =>
      ribbonTabCommandIds('en', tab);

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
      ribbonCommands('en').map((command) => [command.id, command]),
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
    const missing = ribbonCommands('en')
      .filter((command) => command.icon && !fluentIconPaths(command.icon))
      .map((command) => `${command.id}:${command.icon}`);

    expect(missing).toEqual([]);
  });

  it('localizes ribbon model command titles for Japanese Office-like surfaces', () => {
    const commands = new Map(
      ribbonCommands('ja').map((command) => [command.id, command.title]),
    );

    expect(commands.get('numberFormat')).toBe('数値');
    expect(commands.get('paste')).toBe('貼り付け');
    expect(commands.get('bold')).toBe('太字 (⌘B)');
    expect(commands.get('borders')).toBe('罫線');
    const homeFontCommands =
      buildRibbonModel('en')
        .find((tab) => tab.id === 'home')
        ?.groups.find((group) => group.variant === 'font')
        ?.commands.map((command) => command.id) ?? [];
    expect(homeFontCommands.slice(-3)).toEqual(['borders', 'fillColor', 'fontColor']);
    expect(commands.has('borderPreset')).toBe(false);
    expect(commands.has('borderStyle')).toBe(false);
    expect(commands.has('moreBorders')).toBe(false);
    expect(commands.has('drawBorder')).toBe(false);
    expect(commands.has('drawBorderGrid')).toBe(false);
    expect(commands.has('eraseBorder')).toBe(false);
    const modelCommands = ribbonCommands('en');
    const homeAlignmentCommands = buildRibbonModel('en')
      .find((tab) => tab.id === 'home')
      ?.groups.find((group) => group.variant === 'alignment')
      ?.commands.map((command) => command.id) ?? [];
    expect(homeAlignmentCommands).toEqual([
      'top',
      'middle',
      'bottomAlign',
      'alignL',
      'alignC',
      'alignR',
      'alignment-row-2',
      'textOrientation',
      'wrap',
      'indentDecrease',
      'indentIncrease',
      'merge',
    ]);
    expect(modelCommands.find((command) => command.id === 'wrap')?.kind).toBe('button');
    expect(modelCommands.find((command) => command.id === 'currency')?.kind).toBe('button');
    expect(modelCommands.find((command) => command.id === 'merge')?.kind).toBe('button');
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
    const untranslated = ribbonCommands('ja')
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
