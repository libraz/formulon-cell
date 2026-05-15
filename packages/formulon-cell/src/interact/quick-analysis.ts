import { aggregateSelection } from '../commands/aggregate.js';
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
import { inheritHostTokens } from './inherit-host-tokens.js';

export interface QuickAnalysisDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  strings: Strings;
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
  pivotStub: 'pivotStub',
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

export function attachQuickAnalysis(deps: QuickAnalysisDeps): QuickAnalysisHandle {
  const { host, store } = deps;
  let wb = deps.wb;
  let strings = deps.strings;
  let open = false;
  let restoreFocusEl: HTMLElement | null = null;

  const root = document.createElement('div');
  root.className = 'fc-quick';
  root.setAttribute('role', 'dialog');
  root.setAttribute('aria-modal', 'false');
  root.hidden = true;
  root.tabIndex = -1;
  host.appendChild(root);
  inheritHostTokens(host, root);

  const close = (restoreFocus = false): void => {
    if (!open) return;
    open = false;
    const focusTarget = restoreFocusEl;
    restoreFocusEl = null;
    root.hidden = true;
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
    const result = executeQuickAnalysisAction({ store, wb, actionId, range, stats });
    if (!result.ok) return;
    if (result.kind === 'formula') deps.onAfterCommit?.();
    deps.invalidate?.();
    close(false);
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

    root.replaceChildren();
    const title = document.createElement('div');
    title.className = 'fc-quick__title';
    title.textContent = strings.quickAnalysis.title;
    root.appendChild(title);

    for (const group of GROUP_ORDER) {
      const groupActions = grouped[group];
      if (groupActions.length === 0) continue;
      const section = document.createElement('section');
      section.className = 'fc-quick__section';
      const heading = document.createElement('div');
      heading.className = 'fc-quick__heading';
      heading.textContent = strings.quickAnalysis.groups[group];
      section.appendChild(heading);
      const grid = document.createElement('div');
      grid.className = 'fc-quick__actions';
      for (const action of groupActions) {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'fc-quick__action';
        btn.dataset.action = action.id;
        btn.disabled = action.disabled === true;
        btn.textContent = actionLabel(strings, action);
        btn.addEventListener('click', () => execute(action.id));
        grid.appendChild(btn);
      }
      section.appendChild(grid);
      root.appendChild(section);
    }
  };

  const openPanel = (): void => {
    render();
    restoreFocusEl = document.activeElement instanceof HTMLElement ? document.activeElement : host;
    root.hidden = false;
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

  return {
    open: openPanel,
    close,
    setStrings(next) {
      strings = next;
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
      root.remove();
    },
  };
}
