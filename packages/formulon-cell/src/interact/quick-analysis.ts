import { aggregateSelection } from '../commands/aggregate.js';
import type { History } from '../commands/history.js';
import {
  buildQuickAnalysisActions,
  executeQuickAnalysisAction,
  groupQuickAnalysisActions,
  type QuickAnalysisAction,
  type QuickAnalysisActionId,
  type QuickAnalysisGroup,
} from '../commands/quick-analysis.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import { rangeRects } from '../render/geometry.js';
import type { SpreadsheetStore } from '../store/store.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { createInteractionButton } from './chip-button.js';
import { inheritHostTokens } from './inherit-host-tokens.js';

export interface QuickAnalysisDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  strings: Strings;
  history?: History | null;
  onAfterCommit?: () => void;
  invalidate?: () => void;
  onOpenPivotTable?: () => void;
  canOpenPivotTable?: () => boolean;
  canCreateChart?: () => boolean;
}

export interface QuickAnalysisHandle {
  open(): void;
  close(): void;
  setStrings(next: Strings): void;
  bindWorkbook(next: WorkbookHandle): void;
  detach(): void;
}

const GROUP_ORDER: QuickAnalysisGroup[] = [
  'formatting',
  'charts',
  'totals',
  'tables',
  'sparklines',
];

const ACTION_LABELS: Record<string, keyof Strings['quickAnalysis']['actions']> = {
  dataBar: 'dataBar',
  colorScale: 'colorScale',
  iconSet: 'iconSet',
  greaterThan: 'greaterThan',
  top10: 'top10',
  clearFormat: 'clearFormat',
  sumRow: 'sumRow',
  sumCol: 'sumCol',
  avgRow: 'avgRow',
  countRow: 'countRow',
  formatAsTable: 'formatAsTable',
  pivotTable: 'pivotTable',
  sparkLine: 'sparkLine',
  sparkColumn: 'sparkColumn',
  sparkWinLoss: 'sparkWinLoss',
  chartColumn: 'chartColumn',
  chartLine: 'chartLine',
};

function actionLabel(strings: Strings, action: QuickAnalysisAction): string {
  const key = ACTION_LABELS[action.labelKey];
  return key ? strings.quickAnalysis.actions[key] : action.labelKey;
}

function actionDisabledReason(strings: Strings, action: QuickAnalysisAction): string | null {
  const key = action.disabledReason;
  return key ? strings.quickAnalysis.disabledReasons[key] : null;
}

function positionPanel(host: HTMLElement, root: HTMLElement, store: SpreadsheetStore): void {
  const s = store.getState();
  const rects = rangeRects(s.layout, s.viewport, s.selection.range);
  const anchor = rects[rects.length - 1];
  const panelW = root.offsetWidth || 280;
  const panelH = root.offsetHeight || 240;
  const left = anchor ? anchor.x + anchor.w + 8 : host.clientWidth / 2 - panelW / 2;
  const top = anchor ? anchor.y + anchor.h + 8 : host.clientHeight / 2 - panelH / 2;
  root.style.left = `${Math.max(8, Math.min(host.clientWidth - panelW - 8, left))}px`;
  root.style.top = `${Math.max(8, Math.min(host.clientHeight - panelH - 8, top))}px`;
}

function createQuickAnalysisLauncher(): HTMLButtonElement {
  const button = createInteractionButton({ className: 'fc-quick__button' });
  button.hidden = true;
  button.setAttribute('aria-haspopup', 'dialog');
  return button;
}

function createQuickAnalysisTab(
  group: QuickAnalysisGroup,
  label: string,
  active: boolean,
): HTMLButtonElement {
  const tab = createInteractionButton({
    className: 'fc-quick__tab',
    role: 'tab',
    selected: active,
    tabIndex: active ? 0 : -1,
    dataset: { group },
    text: label,
  });
  tab.id = `fc-quick-tab-${group}`;
  tab.setAttribute('aria-controls', `fc-quick-panel-${group}`);
  return tab;
}

function createQuickAnalysisActionButton(
  strings: Strings,
  action: QuickAnalysisAction,
): HTMLButtonElement {
  const disabled = action.disabled === true;
  const label = actionLabel(strings, action);
  const button = createInteractionButton({
    className: 'fc-quick__action',
    dataset: { action: action.id },
    text: label,
  });
  projectDisabledState(button, disabled, actionDisabledReason(strings, action), {
    datasetKey: 'disabledReason',
    titlePrefix: label,
  });
  return button;
}

export function attachQuickAnalysis(deps: QuickAnalysisDeps): QuickAnalysisHandle {
  const { host, store } = deps;
  let wb = deps.wb;
  let strings = deps.strings;
  let open = false;
  let activeGroup: QuickAnalysisGroup = 'formatting';
  let restoreFocusEl: HTMLElement | null = null;

  const button = createQuickAnalysisLauncher();
  host.appendChild(button);
  inheritHostTokens(host, button);

  const root = document.createElement('div');
  root.className = 'fc-quick';
  root.setAttribute('role', 'dialog');
  root.setAttribute('aria-modal', 'false');
  root.hidden = true;
  root.tabIndex = -1;
  host.appendChild(root);
  inheritHostTokens(host, root);

  const isMultiSelection = (): boolean => {
    const r = store.getState().selection.range;
    return r.r1 > r.r0 || r.c1 > r.c0;
  };

  const positionButton = (): void => {
    const state = store.getState();
    const rects = rangeRects(state.layout, state.viewport, state.selection.range);
    const anchor = rects[rects.length - 1];
    if (!anchor || !isMultiSelection() || state.ui.editor.kind !== 'idle' || open) {
      button.hidden = true;
      return;
    }
    button.hidden = false;
    button.title = strings.quickAnalysis.title;
    button.setAttribute('aria-label', strings.quickAnalysis.title);
    button.style.left = `${Math.max(4, Math.min(host.clientWidth - 26, anchor.x + anchor.w + 3))}px`;
    button.style.top = `${Math.max(4, Math.min(host.clientHeight - 26, anchor.y + anchor.h + 3))}px`;
  };

  const close = (restoreFocus = false): void => {
    if (!open) return;
    open = false;
    const focusTarget = restoreFocusEl;
    restoreFocusEl = null;
    root.hidden = true;
    positionButton();
    if (
      restoreFocus &&
      focusTarget &&
      (root.contains(document.activeElement) || document.activeElement === document.body)
    ) {
      focusTarget.focus({ preventScroll: true });
    }
  };

  const execute = (actionId: QuickAnalysisActionId): void => {
    const state = store.getState();
    const stats = aggregateSelection(state);
    const range = state.selection.range;
    if (actionId === 'tables-pivot' && deps.onOpenPivotTable) {
      deps.onOpenPivotTable();
      close(false);
      return;
    }
    const result = executeQuickAnalysisAction({
      store,
      wb,
      actionId,
      range,
      stats,
      history: deps.history ?? null,
    });
    if (!result.ok) return;
    if (result.kind === 'formula') deps.onAfterCommit?.();
    deps.invalidate?.();
    close(false);
  };

  const selectGroup = (group: QuickAnalysisGroup): void => {
    activeGroup = group;
    render();
    positionPanel(host, root, store);
  };

  const focusActiveTab = (): void => {
    root.querySelector<HTMLButtonElement>('.fc-quick__tab[aria-selected="true"]')?.focus({
      preventScroll: true,
    });
  };

  const moveTab = (from: QuickAnalysisGroup, delta: number): void => {
    const tabs = Array.from(root.querySelectorAll<HTMLButtonElement>('.fc-quick__tab'));
    const groups = tabs
      .map((tab) => tab.dataset.group as QuickAnalysisGroup | undefined)
      .filter((group): group is QuickAnalysisGroup => group != null);
    if (groups.length === 0) return;
    const index = Math.max(0, groups.indexOf(from));
    const next = groups[(index + delta + groups.length) % groups.length];
    if (!next) return;
    selectGroup(next);
    focusActiveTab();
  };

  const onTabKeyDown = (event: KeyboardEvent, group: QuickAnalysisGroup): void => {
    if (event.key === 'ArrowRight' || event.key === 'ArrowDown') {
      event.preventDefault();
      moveTab(group, 1);
    } else if (event.key === 'ArrowLeft' || event.key === 'ArrowUp') {
      event.preventDefault();
      moveTab(group, -1);
    } else if (event.key === 'Home') {
      event.preventDefault();
      const first = root.querySelector<HTMLButtonElement>('.fc-quick__tab');
      const next = first?.dataset.group as QuickAnalysisGroup | undefined;
      if (next) {
        selectGroup(next);
        focusActiveTab();
      }
    } else if (event.key === 'End') {
      event.preventDefault();
      const tabs = Array.from(root.querySelectorAll<HTMLButtonElement>('.fc-quick__tab'));
      const last = tabs[tabs.length - 1];
      const next = last?.dataset.group as QuickAnalysisGroup | undefined;
      if (next) {
        selectGroup(next);
        focusActiveTab();
      }
    }
  };

  const render = (): void => {
    const state = store.getState();
    const stats = aggregateSelection(state);
    const actions = buildQuickAnalysisActions({
      range: state.selection.range,
      stats,
      pivotTableAvailable:
        wb.capabilities.pivotTableMutate &&
        !!deps.onOpenPivotTable &&
        (deps.canOpenPivotTable?.() ?? true),
      chartAvailable: deps.canCreateChart?.() ?? true,
    });
    const grouped = groupQuickAnalysisActions(actions);
    const availableGroups = GROUP_ORDER.filter((group) => grouped[group].length > 0);
    if (!availableGroups.includes(activeGroup)) activeGroup = availableGroups[0] ?? 'formatting';

    root.replaceChildren();
    const title = document.createElement('div');
    title.className = 'fc-quick__title';
    title.textContent = strings.quickAnalysis.title;
    root.appendChild(title);

    const tabs = document.createElement('div');
    tabs.className = 'fc-quick__tabs';
    tabs.setAttribute('role', 'tablist');
    tabs.setAttribute('aria-label', strings.quickAnalysis.title);
    root.appendChild(tabs);

    for (const group of GROUP_ORDER) {
      const groupActions = grouped[group];
      if (groupActions.length === 0) continue;
      const tab = createQuickAnalysisTab(
        group,
        strings.quickAnalysis.groups[group],
        group === activeGroup,
      );
      tab.addEventListener('click', () => selectGroup(group));
      tab.addEventListener('keydown', (event) => onTabKeyDown(event, group));
      tabs.appendChild(tab);

      const section = document.createElement('section');
      section.className = 'fc-quick__section';
      section.id = `fc-quick-panel-${group}`;
      section.setAttribute('role', 'tabpanel');
      section.setAttribute('aria-labelledby', tab.id);
      section.hidden = group !== activeGroup;
      const grid = document.createElement('div');
      grid.className = 'fc-quick__actions';
      for (const action of groupActions) {
        const btn = createQuickAnalysisActionButton(strings, action);
        btn.addEventListener('click', () => execute(action.id));
        grid.appendChild(btn);
      }
      section.appendChild(grid);
      root.appendChild(section);
    }
  };

  const openPanel = (): void => {
    activeGroup = 'formatting';
    render();
    restoreFocusEl = document.activeElement instanceof HTMLElement ? document.activeElement : host;
    root.hidden = false;
    button.hidden = true;
    open = true;
    positionPanel(host, root, store);
    root.focus({ preventScroll: true });
  };

  const onKey = (e: KeyboardEvent): void => {
    if (e.key === 'Escape') close(true);
  };
  const onHostPointerDown = (e: PointerEvent): void => {
    if (!open) return;
    if (root.contains(e.target as Node | null)) return;
    close(true);
  };
  root.addEventListener('keydown', onKey);
  host.addEventListener('pointerdown', onHostPointerDown, true);
  button.addEventListener('click', openPanel);
  const unsub = store.subscribe(positionButton);
  positionButton();

  return {
    open: openPanel,
    close,
    setStrings(next) {
      strings = next;
      positionButton();
      if (open) {
        render();
        positionPanel(host, root, store);
      }
    },
    bindWorkbook(next) {
      wb = next;
    },
    detach() {
      root.removeEventListener('keydown', onKey);
      host.removeEventListener('pointerdown', onHostPointerDown, true);
      button.removeEventListener('click', openPanel);
      unsub();
      button.remove();
      root.remove();
    },
  };
}
