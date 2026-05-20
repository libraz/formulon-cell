import type { Strings } from '../../i18n/strings.js';
import {
  type BuildRibbonModelOptions,
  buildRibbonModel,
  type RibbonTab,
  type ToolbarLang,
  toolbarText,
} from '../ribbon-model.js';

export type RibbonSearchItemKind = 'tab' | 'command' | 'help';

export interface RibbonSearchItem {
  id: string;
  kind: RibbonSearchItemKind;
  label: string;
  hint: string;
  tab: RibbonTab;
  commandId?: string;
  disabled?: boolean;
  disabledReason?: string;
  keywords: string;
}

export interface BuildRibbonSearchIndexOptions extends BuildRibbonModelOptions {
  includeDisabled?: boolean;
}

export interface RibbonSearchUsagePrior {
  /** Command-specific usage weight supplied by a host or persisted app state.
   *  Values are additive with the shared static prior and are intentionally
   *  capped so exact/prefix text matches still dominate. */
  commandBoosts?: Readonly<Record<string, number>>;
}

export interface QueryRibbonSearchIndexOptions {
  usagePrior?: RibbonSearchUsagePrior;
}

const normalizeSearchText = (value: string): string =>
  value.normalize('NFKC').toLowerCase().replace(/\s+/g, ' ').trim();

const searchWords = (value: string): string[] =>
  normalizeSearchText(value)
    .split(/[^a-z0-9\u3040-\u30ff\u3400-\u9fff]+/i)
    .filter(Boolean);

const COMMAND_SEARCH_ALIASES: Readonly<Record<string, string>> = {
  freeze: 'freeze panes lock panes split window 固定 ウィンドウ枠固定',
  pageBreaks: 'page break breaks 改ページ page breaks',
  printTitles: 'repeat rows repeat columns print title rows print title columns タイトル行',
  textToColumns: 'split columns split text delimiter delimiters csv 区切り 文字列 分割',
  removeDupes: 'dedupe deduplicate duplicates unique remove duplicate 重複 一意',
  formatTableHome: 'format as table table style テーブルとして書式設定',
  formatTableInsert: 'table insert table create table listobject テーブル 挿入 作成',
  formatAsTable: 'format as table table style テーブルとして書式設定',
  pivotTableInsert: 'pivot pivot table summarize cross tab ピボット 集計',
  pivotFieldListView:
    'pivot field list pivot fields pivottable fields field list pivot table fields ピボット フィールド リスト',
  selectionPanePageLayout:
    'selection pane arrange objects pictures shapes images drawings 選択 ウィンドウ オブジェクト 配置 図形 画像',
  arrangeObjectsPageLayout:
    'arrange bring forward bring to front send backward send to back front back order objects pictures shapes drawings 配置 前面 背面 最前面 最背面 図形 画像 オブジェクト',
  sortFilterHome: 'filter sort dropdown autofilter フィルター 並べ替え',
  filter: 'filter sort dropdown autofilter フィルター 並べ替え',
  conditional: 'conditional format highlight rules color scale data bars 条件付き書式',
  accessibility:
    'check accessibility accessible issues review warnings alt text アクセシビリティ チェック 確認',
};

/** Small, static popularity prior for Search/Tell me. This is not telemetry:
 *  it keeps common spreadsheet workflows above broad tab/group keyword matches
 *  when the textual score is otherwise close. Exact/prefix label matches still
 *  dominate the ranking. */
const COMMAND_SEARCH_BOOSTS: Readonly<Record<string, number>> = {
  open: 40,
  save: 38,
  print: 36,
  printPageLayout: 34,
  pivotTableInsert: 34,
  pivotFieldListView: 30,
  sortFilterHome: 32,
  filter: 32,
  sortAsc: 30,
  sortDesc: 30,
  conditional: 28,
  formatTableHome: 26,
  formatAsTable: 26,
  freeze: 24,
  textToColumns: 22,
  removeDupes: 22,
  printArea: 20,
  printTitles: 20,
  pageSetup: 18,
  findHome: 18,
  findReview: 18,
  accessibility: 18,
};

const usageBoostFor = (commandId: string, prior: RibbonSearchUsagePrior | undefined): number => {
  const raw = prior?.commandBoosts?.[commandId];
  if (typeof raw !== 'number' || !Number.isFinite(raw)) return 0;
  return Math.max(-80, Math.min(80, raw));
};

const commandLabel = (label: string, title: string): string => label.trim() || title.trim();

const helpSearchText = (
  input: Strings | ToolbarLang,
): { label: string; hint: string; keywords: string } => {
  const tr = toolbarText(input);
  return tr.tabs.help === 'ヘルプ'
    ? {
        label: 'ヘルプとトレーニング',
        hint: 'ヘルプ、サポート、使い方を検索',
        keywords: 'ヘルプ トレーニング サポート 使い方 help support training',
      }
    : {
        label: 'Help and training',
        hint: 'Search help, support, and how-to topics',
        keywords: 'help training support how to tell me search',
      };
};

export function buildRibbonSearchIndex(
  input: Strings | ToolbarLang,
  opts: BuildRibbonSearchIndexOptions = {},
): RibbonSearchItem[] {
  const items: RibbonSearchItem[] = [];
  const tr = toolbarText(input);
  const tabs = buildRibbonModel(input, opts);
  for (const tab of tabs) {
    items.push({
      id: `tab:${tab.id}`,
      kind: 'tab',
      label: tab.label,
      hint: tab.label,
      tab: tab.id,
      keywords: normalizeSearchText(`${tab.label} ${tab.id}`),
    });
    for (const group of tab.groups) {
      for (const command of group.commands) {
        if (command.kind === 'break') continue;
        if (command.disabled && opts.includeDisabled !== true) continue;
        const label = commandLabel(command.label, command.title);
        if (!label) continue;
        const hint = command.title && command.title !== label ? command.title : group.title;
        const keywords = normalizeSearchText(
          [
            label,
            hint,
            group.title,
            tab.label,
            tab.id,
            command.id,
            COMMAND_SEARCH_ALIASES[command.id],
            command.options?.map((option) => option.label).join(' '),
          ]
            .filter(Boolean)
            .join(' '),
        );
        items.push({
          id: `command:${command.id}`,
          kind: 'command',
          label,
          hint,
          tab: tab.id,
          commandId: command.id,
          disabled: command.disabled,
          disabledReason: command.disabled ? tr.disabled : undefined,
          keywords,
        });
      }
    }
  }
  if (tabs.some((tab) => tab.id === 'help')) {
    const help = helpSearchText(input);
    items.push({
      id: 'help:helpAndTraining',
      kind: 'help',
      label: help.label,
      hint: help.hint,
      tab: 'help',
      keywords: normalizeSearchText(`${help.label} ${help.hint} ${help.keywords}`),
    });
  }
  return items;
}

export function queryRibbonSearchIndex(
  items: readonly RibbonSearchItem[],
  query: string,
  limit = 8,
  opts: QueryRibbonSearchIndexOptions = {},
): RibbonSearchItem[] {
  const q = normalizeSearchText(query);
  if (!q) return items.slice(0, limit);
  const qTokens = searchWords(q);
  const scoreItem = (item: RibbonSearchItem): number | null => {
    const label = normalizeSearchText(item.label);
    const hint = normalizeSearchText(item.hint);
    const keywords = normalizeSearchText(item.keywords);
    const words = searchWords(`${item.label} ${item.hint} ${item.keywords}`);
    if (!keywords.includes(q) && !qTokens.every((token) => keywords.includes(token))) return null;

    let score = 0;
    if (label === q) score += 1200;
    else if (label.startsWith(q)) score += 950;
    else if (words.some((word) => word.startsWith(q))) score += 720;
    else if (label.includes(q)) score += 640;
    else if (hint === q) score += 560;
    else if (hint.startsWith(q)) score += 500;
    else if (keywords.includes(q)) score += 300;

    for (const token of qTokens) {
      if (label === token) score += 110;
      else if (searchWords(item.label).some((word) => word.startsWith(token))) score += 90;
      else if (label.includes(token)) score += 60;
      else if (searchWords(item.hint).some((word) => word.startsWith(token))) score += 35;
      else if (words.some((word) => word.startsWith(token))) score += 25;
      else if (keywords.includes(token)) score += 10;
    }

    if (item.kind === 'command') score += 20;
    if (item.kind === 'help') score += 10;
    if (item.kind === 'tab' && label === q) score += 200;
    else if (item.kind === 'tab' && label.startsWith(q)) score += 100;
    if (item.commandId) {
      score += COMMAND_SEARCH_BOOSTS[item.commandId] ?? 0;
      score += usageBoostFor(item.commandId, opts.usagePrior);
    }
    if (item.disabled) score -= 250;
    return score;
  };

  return items
    .map((item, index) => ({ item, index, score: scoreItem(item) }))
    .filter(
      (entry): entry is { item: RibbonSearchItem; index: number; score: number } =>
        entry.score !== null,
    )
    .sort((a, b) => b.score - a.score || a.index - b.index)
    .map((entry) => entry.item)
    .slice(0, limit);
}
