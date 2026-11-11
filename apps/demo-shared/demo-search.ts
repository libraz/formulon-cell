import type { SpreadsheetInstance, ThemeName } from '@libraz/formulon-cell';
import {
  buildRibbonSearchIndex,
  EXCEL365_STANDARD_RIBBON_TABS,
  type ExcelRibbonIconName,
  queryRibbonSearchIndex,
  type RibbonSearchItem,
  type RibbonSearchUsagePrior,
  type RibbonTab,
} from '@libraz/formulon-cell/toolbar';
import {
  openDemoDialog,
  runDemoCommand,
  toggleDemoSheetProtection,
  traceDemoPrecedents,
} from './demo-actions.js';
import { demoDictionaryLocale } from './demo-locale.js';
import type { DemoCommandStrings, DemoUiStrings } from './demo-strings.js';

export interface DemoCommandItem {
  readonly id: string;
  readonly label: string;
  readonly hint: string;
  readonly keywords?: string;
  readonly tab?: RibbonTab;
  readonly run: () => void;
}

export interface DemoSearchItem {
  readonly id: string;
  readonly label: string;
  readonly hint: string;
  readonly commandId?: string;
  readonly disabled?: boolean;
  readonly disabledReason?: string;
  readonly keywords?: string;
  readonly tab?: RibbonTab;
  readonly run: () => void;
}

export type DemoSearchUsagePrior = RibbonSearchUsagePrior;

const DEMO_SEARCH_USAGE_KEY = 'formulon-cell.demo.searchUsagePrior';

const isSearchUsagePrior = (value: unknown): value is DemoSearchUsagePrior => {
  if (!value || typeof value !== 'object') return false;
  const boosts = (value as { commandBoosts?: unknown }).commandBoosts;
  return (
    boosts === undefined ||
    (!!boosts &&
      typeof boosts === 'object' &&
      Object.values(boosts).every((entry) => typeof entry === 'number' && Number.isFinite(entry)))
  );
};

export const loadDemoSearchUsagePrior = (): DemoSearchUsagePrior => {
  try {
    const raw = globalThis.localStorage?.getItem(DEMO_SEARCH_USAGE_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw) as unknown;
    return isSearchUsagePrior(parsed) ? parsed : {};
  } catch {
    return {};
  }
};

export const saveDemoSearchUsagePrior = (prior: DemoSearchUsagePrior): void => {
  try {
    globalThis.localStorage?.setItem(DEMO_SEARCH_USAGE_KEY, JSON.stringify(prior));
  } catch {
    // Storage is optional in embedded demos.
  }
};

export const recordDemoSearchUsage = (
  prior: DemoSearchUsagePrior,
  item: DemoSearchItem,
): DemoSearchUsagePrior => {
  const commandId = item.commandId;
  if (!commandId) return prior;
  return {
    commandBoosts: {
      ...(prior.commandBoosts ?? {}),
      [commandId]: Math.min(100, (prior.commandBoosts?.[commandId] ?? 0) + 12),
    },
  };
};

export interface BuildDemoCommandsOptions {
  commandText: DemoCommandStrings;
  instance: SpreadsheetInstance | null;
  openWorkbook: () => void;
  saveWorkbook: () => void;
  setRibbonTab: (tab: RibbonTab) => void;
  togglePanel: () => void;
  setTheme: (theme: ThemeName) => void;
  setLocale: (locale: string) => void;
}

export const buildDemoCommands = ({
  commandText,
  instance,
  openWorkbook,
  saveWorkbook,
  setRibbonTab,
  togglePanel,
  setTheme,
  setLocale,
}: BuildDemoCommandsOptions): readonly DemoCommandItem[] => [
  {
    id: 'open',
    label: commandText.commands.open.label,
    hint: commandText.commands.open.hint,
    tab: 'file',
    run: openWorkbook,
  },
  {
    id: 'save',
    label: commandText.commands.save.label,
    hint: commandText.commands.save.hint,
    tab: 'file',
    run: saveWorkbook,
  },
  {
    id: 'page-setup',
    label: commandText.commands.pageSetup.label,
    hint: commandText.commands.pageSetup.hint,
    tab: 'file',
    run: () => openDemoDialog(instance, 'pageSetup'),
  },
  {
    id: 'print',
    label: commandText.commands.print.label,
    hint: commandText.commands.print.hint,
    tab: 'file',
    run: () => instance?.print('print'),
  },
  {
    id: 'format-cells',
    label: commandText.commands.formatCells.label,
    hint: commandText.commands.formatCells.hint,
    tab: 'home',
    run: () => openDemoDialog(instance, 'format'),
  },
  {
    id: 'conditional',
    label: commandText.commands.conditionalFormatting.label,
    hint: commandText.commands.conditionalFormatting.hint,
    tab: 'insert',
    run: () => openDemoDialog(instance, 'conditional'),
  },
  {
    id: 'cell-styles',
    label: commandText.commands.cellStyles.label,
    hint: commandText.commands.cellStyles.hint,
    tab: 'insert',
    run: () => openDemoDialog(instance, 'cellStylesGallery'),
  },
  {
    id: 'name-manager',
    label: commandText.commands.nameManager.label,
    hint: commandText.commands.nameManager.hint,
    tab: 'insert',
    run: () => openDemoDialog(instance, 'namedRange'),
  },
  {
    id: 'insert-function',
    label: commandText.commands.insertFunction.label,
    hint: commandText.commands.insertFunction.hint,
    tab: 'formulas',
    run: () => openDemoDialog(instance, 'functionArguments'),
  },
  {
    id: 'trace-precedents',
    label: commandText.commands.tracePrecedents.label,
    hint: commandText.commands.tracePrecedents.hint,
    tab: 'formulas',
    run: () => traceDemoPrecedents(instance),
  },
  {
    id: 'watch-window',
    label: commandText.commands.watchWindow.label,
    hint: commandText.commands.watchWindow.hint,
    tab: 'formulas',
    run: () => runDemoCommand(instance, 'watch'),
  },
  {
    id: 'filter',
    label: commandText.commands.filter.label,
    hint: commandText.commands.filter.hint,
    tab: 'data',
    run: () => setRibbonTab('data'),
  },
  {
    id: 'sort',
    label: commandText.commands.sort.label,
    hint: commandText.commands.sort.hint,
    tab: 'data',
    run: () => setRibbonTab('data'),
  },
  {
    id: 'freeze-panes',
    label: commandText.commands.freezePanes.label,
    hint: commandText.commands.freezePanes.hint,
    tab: 'view',
    run: () => setRibbonTab('view'),
  },
  {
    id: 'protect-sheet',
    label: commandText.commands.protectSheet.label,
    hint: commandText.commands.protectSheet.hint,
    tab: 'view',
    run: () => toggleDemoSheetProtection(instance),
  },
  {
    id: 'options-pane',
    label: commandText.commands.options.label,
    hint: commandText.commands.options.hint,
    run: togglePanel,
  },
  {
    id: 'theme-light',
    label: commandText.commands.lightTheme.label,
    hint: commandText.commands.lightTheme.hint,
    run: () => setTheme('paper'),
  },
  {
    id: 'theme-dark',
    label: commandText.commands.darkTheme.label,
    hint: commandText.commands.darkTheme.hint,
    run: () => setTheme('ink'),
  },
  {
    id: 'locale-ja',
    label: commandText.commands.japaneseLocale.label,
    hint: commandText.commands.japaneseLocale.hint,
    run: () => setLocale('ja'),
  },
  {
    id: 'locale-en',
    label: commandText.commands.englishLocale.label,
    hint: commandText.commands.englishLocale.hint,
    run: () => setLocale('en'),
  },
];

const ribbonSearchHint = (item: RibbonSearchItem): string => {
  const hint =
    item.kind === 'tab' || item.kind === 'help' ? item.hint : `${item.hint} · ${item.tab}`;
  return item.disabledReason ? `${hint} · ${item.disabledReason}` : hint;
};

export const buildDemoSearchItems = (
  commandItems: readonly DemoCommandItem[],
  locale: string,
  setRibbonTab: (tab: RibbonTab) => void,
  applyRibbonCommand?: (commandId: string) => boolean,
  ribbonTabs: readonly RibbonTab[] = EXCEL365_STANDARD_RIBBON_TABS,
): readonly DemoSearchItem[] => {
  const commandKeys = new Set(commandItems.flatMap((item) => [item.id, item.label]));
  const ribbonItems = buildRibbonSearchIndex(demoDictionaryLocale(locale), {
    includeDisabled: true,
    tabs: ribbonTabs,
  })
    .filter((item) => !commandKeys.has(item.commandId ?? '') && !commandKeys.has(item.label))
    .map(
      (item): DemoSearchItem => ({
        id: `ribbon:${item.id}`,
        label: item.label,
        hint: ribbonSearchHint(item),
        commandId: item.commandId,
        disabled: item.disabled,
        disabledReason: item.disabledReason,
        keywords: item.keywords,
        tab: item.tab,
        run: () => {
          if (item.commandId && applyRibbonCommand?.(item.commandId)) return;
          setRibbonTab(item.tab);
        },
      }),
    );
  return [...commandItems, ...ribbonItems];
};

export const queryDemoSearchItems = (
  items: readonly DemoSearchItem[],
  query: string,
  limit = 8,
  usagePrior: DemoSearchUsagePrior = {},
): readonly DemoSearchItem[] => {
  const q = query.trim().toLowerCase();
  if (!q) return items.slice(0, limit);
  const ribbonBacked = items.filter((item) => item.id.startsWith('ribbon:'));
  const demoBacked = items.filter((item) => !item.id.startsWith('ribbon:'));
  const demoMatches = demoBacked.filter((item) =>
    `${item.label} ${item.hint} ${item.keywords ?? ''}`.toLowerCase().includes(q),
  );
  const ribbonMatches = queryRibbonSearchIndex(
    ribbonBacked.map(
      (item): RibbonSearchItem => ({
        id: item.id,
        kind: 'command',
        label: item.label,
        hint: item.hint,
        tab: item.tab ?? 'home',
        commandId: item.commandId,
        disabled: item.disabled,
        disabledReason: item.disabledReason,
        keywords: `${item.label} ${item.hint} ${item.keywords ?? ''}`.toLowerCase(),
      }),
    ),
    query,
    limit,
    { usagePrior },
  );
  const byRibbonId = new Map(ribbonBacked.map((item) => [item.id, item]));
  const orderedRibbonMatches = ribbonMatches.flatMap((item) => {
    const match = byRibbonId.get(item.id);
    return match ? [match] : [];
  });
  return [...demoMatches, ...orderedRibbonMatches].slice(0, limit);
};

export const demoSearchOptionId = (index: number): string => `demo-search-option-${index}`;

export const nextDemoSearchIndex = (
  current: number,
  count: number,
  direction: 'first' | 'next' | 'previous',
): number => {
  if (count <= 0) return -1;
  if (direction === 'first') return Math.max(current, 0);
  if (direction === 'next') return (current + 1 + count) % count;
  return (current < 0 ? count - 1 : current - 1 + count) % count;
};

export type DemoSearchSuggestionId = keyof DemoUiStrings['searchSuggestions']['items'];

export interface DemoSearchSuggestion {
  readonly id: DemoSearchSuggestionId;
  readonly label: string;
  /** Ribbon glyph drawn before the label. */
  readonly glyph: ExcelRibbonIconName;
  /** The row leads to a submenu on the desktop. */
  readonly submenu?: boolean;
  readonly disabled?: boolean;
  readonly run: () => void;
}

export interface DemoSearchSuggestionSection {
  readonly id: 'recommended' | 'findReplace' | 'feedback';
  /** Null for the feedback row, which sits under a rule instead of a heading. */
  readonly heading: string | null;
  readonly items: readonly DemoSearchSuggestion[];
}

export interface DemoSearchSuggestionDeps {
  readonly instance: SpreadsheetInstance | null;
  readonly setRibbonTab: (tab: RibbonTab) => void;
  readonly applyCommand: (commandId: string) => boolean;
}

/** The Mac popup under an empty search field: Suggested Actions, Find and
 *  Replace, then the feedback row. */
export const buildDemoSearchSuggestions = (
  ui: DemoUiStrings,
  deps: DemoSearchSuggestionDeps,
): readonly DemoSearchSuggestionSection[] => {
  const t = ui.searchSuggestions;
  const item = (
    id: DemoSearchSuggestionId,
    glyph: ExcelRibbonIconName,
    run: () => void,
    extra: Pick<DemoSearchSuggestion, 'submenu' | 'disabled'> = {},
  ): DemoSearchSuggestion => ({ id, label: t.items[id], glyph, run, ...extra });
  return [
    {
      id: 'recommended',
      heading: t.recommended,
      items: [
        item('freeze', 'freeze', () => deps.setRibbonTab('view'), { submenu: true }),
        item('headerFooter', 'headerFooter', () =>
          openDemoDialog(deps.instance, 'pageSetup', { tab: 'headerFooter' }),
        ),
        item('insertSheetRows', 'insertRows', () => deps.applyCommand('insertRows')),
        item('removeDuplicates', 'removeDuplicates', () => deps.applyCommand('removeDupes')),
        item('macros', 'code', () => {}, { disabled: true }),
      ],
    },
    {
      id: 'findReplace',
      heading: t.findReplace,
      items: [
        item('find', 'find', () => openDemoDialog(deps.instance, 'findReplace', { tab: 'find' })),
        item('findReplace', 'replaceFind', () =>
          openDemoDialog(deps.instance, 'findReplace', { tab: 'replace' }),
        ),
      ],
    },
    {
      id: 'feedback',
      heading: null,
      items: [item('feedback', 'feedback', () => {}, { disabled: true })],
    },
  ];
};
