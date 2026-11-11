import type { ExcelRibbonIconName, RibbonTab, SpreadsheetInstance } from '@libraz/formulon-cell';
import type { DemoUiStrings } from './demo-strings.js';

/** Commands the Mac Quick Access Toolbar can show, in its customize-menu order. */
export const DEMO_QUICK_ACCESS_COMMANDS = [
  'autosave',
  'new',
  'home',
  'save',
  'saveAs',
  'sendWorkbook',
  'print',
  'spelling',
  'undo',
  'redo',
  'sortAsc',
  'sortDesc',
] as const;

export type DemoQuickAccessId = (typeof DEMO_QUICK_ACCESS_COMMANDS)[number];

/** The desktop's default set: AutoSave, Home, Save, Save As, Undo, Repeat. */
export const DEFAULT_DEMO_QUICK_ACCESS: readonly DemoQuickAccessId[] = [
  'autosave',
  'home',
  'save',
  'saveAs',
  'undo',
  'redo',
];

/** Ribbon glyph each toolbar button draws; AutoSave is a switch and draws none. */
export const DEMO_QUICK_ACCESS_GLYPHS: Readonly<
  Record<DemoQuickAccessId, ExcelRibbonIconName | null>
> = {
  autosave: null,
  new: 'page',
  home: 'home',
  save: 'save',
  saveAs: 'saveAs',
  sendWorkbook: 'share',
  print: 'print',
  spelling: 'spelling',
  undo: 'undo',
  redo: 'repeat',
  sortAsc: 'sortAsc',
  sortDesc: 'sortDesc',
};

export interface DemoQuickAccessMenuItem {
  readonly id: DemoQuickAccessId;
  readonly label: string;
  readonly checked: boolean;
}

/** Rows of the customize menu, each checked when its command is on the toolbar. */
export const demoQuickAccessMenu = (
  ui: DemoUiStrings,
  shown: readonly DemoQuickAccessId[],
): readonly DemoQuickAccessMenuItem[] =>
  DEMO_QUICK_ACCESS_COMMANDS.map((id) => ({
    id,
    label: ui.quickAccess.menu[id],
    checked: shown.includes(id),
  }));

/** Add or remove `id`, keeping the toolbar in customize-menu order. */
export const toggleDemoQuickAccess = (
  shown: readonly DemoQuickAccessId[],
  id: DemoQuickAccessId,
): readonly DemoQuickAccessId[] => {
  const next = new Set(shown);
  if (next.has(id)) next.delete(id);
  else next.add(id);
  return DEMO_QUICK_ACCESS_COMMANDS.filter((command) => next.has(command));
};

export interface DemoQuickAccessDeps {
  readonly instance: SpreadsheetInstance | null;
  readonly newWorkbook: () => void;
  readonly saveWorkbook: () => void;
  readonly saveWorkbookAs: () => void;
  readonly setRibbonTab: (tab: RibbonTab) => void;
  readonly applyCommand: (commandId: string) => boolean;
}

/** Whether a toolbar button can run now: Undo needs history, and the demo has
 *  no mail service to send a workbook through. */
export const demoQuickAccessEnabled = (
  id: DemoQuickAccessId,
  instance: SpreadsheetInstance | null,
): boolean => {
  if (id === 'sendWorkbook') return false;
  if (id === 'undo') return instance?.history.canUndo ?? false;
  if (id === 'redo') return instance?.history.canRepeat ?? false;
  return true;
};

/** The history flags the title bar's Undo and Redo buttons read, as one value for a change check. */
export const demoHistorySnapshot = (instance: SpreadsheetInstance | null): string =>
  instance
    ? `${instance.history.canUndo}/${instance.history.canRedo}/${instance.history.canRepeat}`
    : '';

/** Run a toolbar button. AutoSave is a local switch the caller owns. */
export const runDemoQuickAccess = (id: DemoQuickAccessId, deps: DemoQuickAccessDeps): void => {
  switch (id) {
    case 'autosave':
      return;
    case 'new':
      deps.newWorkbook();
      return;
    case 'home':
      deps.setRibbonTab('home');
      return;
    case 'save':
      deps.saveWorkbook();
      return;
    case 'saveAs':
      deps.saveWorkbookAs();
      return;
    case 'sendWorkbook':
      return;
    case 'print':
      deps.instance?.print('print');
      return;
    case 'spelling':
      deps.applyCommand('spellingReview');
      return;
    case 'undo':
      deps.instance?.history.undo();
      return;
    case 'redo':
      // The Mac button repeats the last action rather than redoing it.
      deps.instance?.history.repeat();
      return;
    case 'sortAsc':
      deps.applyCommand('sortAsc');
      return;
    case 'sortDesc':
      deps.applyCommand('sortDesc');
      return;
  }
};
