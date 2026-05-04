import { aggregateSelection } from '../commands/aggregate.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { mutators, type SpreadsheetStore, type StatusAggKey } from '../store/store.js';

export interface StatusBarDeps {
  /** The status bar element built by mount.ts. We take it over and lay
   *  out three sections: left (state), center (aggregates), right (engine). */
  statusbar: HTMLElement;
  store: SpreadsheetStore;
  strings?: Strings;
  /** Engine label rendered on the far right. Recomputed on every chrome
   *  update — pass a function rather than a string. */
  getEngineLabel: () => string;
}

export interface StatusBarHandle {
  /** Force a re-render of the status bar (useful after engine swap). */
  refresh(): void;
  detach(): void;
}

const ALL_KEYS: StatusAggKey[] = ['average', 'count', 'countNumbers', 'min', 'max', 'sum'];
const VIEWPORT_PAD = 4;

const fmt = (n: number): string => {
  if (!Number.isFinite(n)) return '—';
  const abs = Math.abs(n);
  if (abs !== 0 && (abs < 0.01 || abs >= 1e9)) return n.toExponential(3);
  return n.toLocaleString('en-US', { maximumFractionDigits: 4 });
};

export function attachStatusBar(deps: StatusBarDeps): StatusBarHandle {
  const { statusbar, store, getEngineLabel } = deps;
  const strings = deps.strings ?? defaultStrings;
  const t = strings.statusBar;

  statusbar.replaceChildren();

  const left = document.createElement('span');
  left.className = 'fc-host__statusbar-left';
  const dot = document.createElement('span');
  dot.className = 'fc-host__statusbar-dot';
  left.appendChild(dot);
  left.appendChild(document.createTextNode('Ready'));

  const center = document.createElement('span');
  center.className = 'fc-host__statusbar-aggs';
  center.setAttribute('role', 'status');

  const right = document.createElement('span');
  right.className = 'fc-host__statusbar-right';
  right.textContent = '—';

  statusbar.append(left, center, right);

  const labelFor = (key: StatusAggKey): string => {
    switch (key) {
      case 'sum':
        return t.sum;
      case 'average':
        return t.average;
      case 'count':
        return t.count;
      case 'countNumbers':
        return t.countNumbers;
      case 'min':
        return t.min;
      case 'max':
        return t.max;
    }
  };

  const valueFor = (
    key: StatusAggKey,
    stats: ReturnType<typeof aggregateSelection>,
  ): string | null => {
    if (key === 'count') return stats.nonBlankCount > 0 ? String(stats.nonBlankCount) : null;
    if (key === 'countNumbers') return stats.numericCount > 0 ? String(stats.numericCount) : null;
    if (stats.numericCount === 0) return null;
    switch (key) {
      case 'sum':
        return fmt(stats.sum);
      case 'average':
        return fmt(stats.avg);
      case 'min':
        return fmt(stats.min);
      case 'max':
        return fmt(stats.max);
      default:
        return null;
    }
  };

  const refresh = (): void => {
    const s = store.getState();
    const stats = aggregateSelection(s);
    const keys = s.ui.statusAggs;
    const pieces: string[] = [];
    for (const key of keys) {
      const v = valueFor(key, stats);
      if (v != null) pieces.push(`${labelFor(key)}: ${v}`);
    }
    center.textContent = pieces.join(' · ');

    const sel = s.selection.range;
    const cells = (sel.r1 - sel.r0 + 1) * (sel.c1 - sel.c0 + 1);
    const engine = getEngineLabel();
    right.textContent = cells === 1 ? `1 cell · ${engine}` : `${cells} cells · ${engine}`;
  };

  // Chooser popover. Lives in document.body so it escapes any clipping
  // ancestor and survives statusbar layout changes.
  const popover = document.createElement('div');
  popover.className = 'fc-statusbar__chooser';
  popover.setAttribute('role', 'menu');
  popover.style.display = 'none';
  document.body.appendChild(popover);

  let popoverVisible = false;

  const buildChooser = (): void => {
    popover.replaceChildren();
    const heading = document.createElement('div');
    heading.className = 'fc-statusbar__chooser-heading';
    heading.textContent = t.aggregatesHeading;
    popover.appendChild(heading);
    const active = new Set(store.getState().ui.statusAggs);
    for (const key of ALL_KEYS) {
      const row = document.createElement('button');
      row.type = 'button';
      row.className = 'fc-statusbar__chooser-item';
      row.setAttribute('role', 'menuitemcheckbox');
      const check = document.createElement('span');
      check.className = 'fc-statusbar__chooser-check';
      check.textContent = active.has(key) ? '✓' : '';
      const label = document.createElement('span');
      label.textContent = labelFor(key);
      row.append(check, label);
      row.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        mutators.toggleStatusAgg(store, key);
        check.textContent = store.getState().ui.statusAggs.includes(key) ? '✓' : '';
      });
      popover.appendChild(row);
    }
  };

  const placeChooser = (clientX: number, clientY: number): void => {
    const w = popover.offsetWidth;
    const h = popover.offsetHeight;
    const vw = window.innerWidth;
    const vh = window.innerHeight;
    const x = Math.max(VIEWPORT_PAD, Math.min(clientX, vw - w - VIEWPORT_PAD));
    const y = Math.max(VIEWPORT_PAD, Math.min(clientY - h - 8, vh - h - VIEWPORT_PAD));
    popover.style.left = `${x}px`;
    popover.style.top = `${y}px`;
  };

  const showChooser = (clientX: number, clientY: number): void => {
    buildChooser();
    popover.style.display = 'block';
    popover.style.left = '-9999px';
    popover.style.top = '-9999px';
    popoverVisible = true;
    placeChooser(clientX, clientY);
  };

  const hideChooser = (): void => {
    if (!popoverVisible) return;
    popoverVisible = false;
    popover.style.display = 'none';
  };

  const onContextMenu = (e: MouseEvent): void => {
    e.preventDefault();
    showChooser(e.clientX, e.clientY);
  };

  // Left-click on the center aggregate strip also opens the chooser.
  const onCenterClick = (e: MouseEvent): void => {
    e.preventDefault();
    showChooser(e.clientX, e.clientY);
  };

  const onDocPointerDown = (e: MouseEvent): void => {
    if (!popoverVisible) return;
    if (e.target instanceof Node && popover.contains(e.target)) return;
    hideChooser();
  };

  const onDocKey = (e: KeyboardEvent): void => {
    if (!popoverVisible) return;
    if (e.key === 'Escape') {
      e.preventDefault();
      hideChooser();
    }
  };

  statusbar.addEventListener('contextmenu', onContextMenu);
  center.addEventListener('click', onCenterClick);
  document.addEventListener('mousedown', onDocPointerDown, true);
  document.addEventListener('keydown', onDocKey, true);

  const unsub = store.subscribe(refresh);
  refresh();

  return {
    refresh,
    detach() {
      statusbar.removeEventListener('contextmenu', onContextMenu);
      center.removeEventListener('click', onCenterClick);
      document.removeEventListener('mousedown', onDocPointerDown, true);
      document.removeEventListener('keydown', onDocKey, true);
      popover.remove();
      unsub();
    },
  };
}
