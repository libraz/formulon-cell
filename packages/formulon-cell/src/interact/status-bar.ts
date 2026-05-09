import {
  aggregateSelection,
  STATUS_AGGREGATE_KEYS,
  statusAggregateValue,
} from '../commands/aggregate.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { mutators, type SpreadsheetStore, type StatusAggKey } from '../store/store.js';
import { inheritHostTokens } from './inherit-host-tokens.js';

export interface StatusBarDeps {
  /** The status bar element built by mount.ts. We take it over and lay
   *  out three sections: left (state), center (aggregates), right (engine). */
  statusbar: HTMLElement;
  store: SpreadsheetStore;
  strings?: Strings;
  /** Engine label rendered on the far right. Recomputed on every chrome
   *  update — pass a function rather than a string. */
  getEngineLabel: () => string;
  /** Optional calc-mode badge driver. Returns `null` when the engine
   *  doesn't expose `calcMode` — the badge is hidden in that case. The
   *  badge is clickable: `onClickCalcMode` cycles the mode (Auto →
   *  Manual → AutoNoTable → Auto) when present, otherwise the badge is
   *  inert. `onRecalc` fires F9 / Ctrl+Alt+F9. */
  getCalcMode?: () => 0 | 1 | 2 | null;
  onCycleCalcMode?: () => void;
  onRecalc?: () => void;
  /** Optional Excel-style zoom control driver. `zoom` is a multiplier
   *  (1.0 = 100%). The status bar clamps UI input to the store's supported
   *  [0.5, 4] range before calling this. */
  onZoomChange?: (zoom: number) => void;
}

export interface StatusBarHandle {
  /** Force a re-render of the status bar (useful after engine swap). */
  refresh(): void;
  /** Swap the active dictionary; live-updates labels in place. */
  setStrings(next: Strings): void;
  detach(): void;
}

const ALL_KEYS: readonly StatusAggKey[] = STATUS_AGGREGATE_KEYS;
const VIEWPORT_PAD = 4;

const fmt = (n: number): string => {
  if (!Number.isFinite(n)) return '—';
  const abs = Math.abs(n);
  if (abs !== 0 && (abs < 0.01 || abs >= 1e9)) return n.toExponential(3);
  return n.toLocaleString('en-US', { maximumFractionDigits: 4 });
};

const clampZoom = (zoom: number): number => Math.max(0.5, Math.min(4, zoom));

export function attachStatusBar(deps: StatusBarDeps): StatusBarHandle {
  const { statusbar, store, getEngineLabel } = deps;
  let strings = deps.strings ?? defaultStrings;

  statusbar.replaceChildren();

  const left = document.createElement('span');
  left.className = 'fc-host__statusbar-left';
  const dot = document.createElement('span');
  dot.className = 'fc-host__statusbar-dot';
  left.appendChild(dot);
  const readyText = document.createTextNode('');
  left.appendChild(readyText);

  const center = document.createElement('span');
  center.className = 'fc-host__statusbar-aggs';
  center.setAttribute('role', 'status');

  const calcBadge = document.createElement('button');
  calcBadge.type = 'button';
  calcBadge.className = 'fc-host__statusbar-calcmode';
  calcBadge.style.display = 'none';
  calcBadge.addEventListener('click', () => {
    // Click on the badge itself toggles the mode; double-click recalcs.
    deps.onCycleCalcMode?.();
  });
  calcBadge.addEventListener('dblclick', () => {
    deps.onRecalc?.();
  });

  const right = document.createElement('span');
  right.className = 'fc-host__statusbar-right';
  right.textContent = '—';

  const zoom = document.createElement('div');
  zoom.className = 'fc-host__statusbar-zoom';
  const zoomOut = document.createElement('button');
  zoomOut.type = 'button';
  zoomOut.className = 'fc-host__statusbar-zoom-btn';
  zoomOut.textContent = '−';
  const zoomSlider = document.createElement('input');
  zoomSlider.type = 'range';
  zoomSlider.className = 'fc-host__statusbar-zoom-slider';
  zoomSlider.min = '50';
  zoomSlider.max = '400';
  zoomSlider.step = '10';
  const zoomIn = document.createElement('button');
  zoomIn.type = 'button';
  zoomIn.className = 'fc-host__statusbar-zoom-btn';
  zoomIn.textContent = '+';
  const zoomLabel = document.createElement('span');
  zoomLabel.className = 'fc-host__statusbar-zoom-label';
  zoom.append(zoomOut, zoomSlider, zoomIn, zoomLabel);

  statusbar.append(left, center, calcBadge, right, zoom);

  const calcLabelFor = (mode: 0 | 1 | 2): string => {
    const t = strings.statusBar;
    switch (mode) {
      case 0:
        return t.calcAuto;
      case 1:
        return t.calcManual;
      case 2:
        return t.calcAutoNoTable;
    }
  };

  const refreshCalcBadge = (): void => {
    const mode = deps.getCalcMode?.();
    if (mode === undefined || mode === null) {
      calcBadge.style.display = 'none';
      return;
    }
    const t = strings.statusBar;
    calcBadge.style.display = '';
    calcBadge.textContent = `${t.calcLabel}: ${calcLabelFor(mode)}`;
    calcBadge.title = t.calcRecalcHint;
    calcBadge.setAttribute('aria-label', `${calcBadge.textContent}. ${t.calcRecalcHint}`);
    calcBadge.dataset.calcMode = String(mode);
  };

  const refreshStaticLabels = (): void => {
    const t = strings.statusBar;
    zoomOut.setAttribute('aria-label', t.zoomOut);
    zoomSlider.setAttribute('aria-label', t.zoom);
    zoomIn.setAttribute('aria-label', t.zoomIn);
    readyText.nodeValue = t.ready;
  };

  const labelFor = (key: StatusAggKey): string => {
    const t = strings.statusBar;
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

  const refresh = (): void => {
    const s = store.getState();
    refreshStaticLabels();
    const stats = aggregateSelection(s);
    const keys = s.ui.statusAggs;
    const pieces: string[] = [];
    for (const key of keys) {
      const v = statusAggregateValue(key, stats);
      if (v != null) pieces.push(`${labelFor(key)}: ${fmt(v)}`);
    }
    center.textContent = pieces.join(' · ');

    const sel = s.selection.range;
    const cells = (sel.r1 - sel.r0 + 1) * (sel.c1 - sel.c0 + 1);
    const engine = getEngineLabel();
    right.textContent =
      cells === 1
        ? `1 ${strings.statusBar.cell} · ${engine}`
        : `${cells} ${strings.statusBar.cells} · ${engine}`;
    const zoomPct = Math.round(s.viewport.zoom * 100);
    zoomSlider.value = String(zoomPct);
    zoomLabel.textContent = `${zoomPct}%`;
    zoomOut.disabled = s.viewport.zoom <= 0.5;
    zoomIn.disabled = s.viewport.zoom >= 4;
    refreshCalcBadge();
  };

  const setZoom = (next: number): void => {
    const z = clampZoom(next);
    if (deps.onZoomChange) deps.onZoomChange(z);
    else mutators.setZoom(store, z);
    refresh();
  };

  zoomSlider.addEventListener('input', () => {
    setZoom(Number(zoomSlider.value) / 100);
  });
  zoomOut.addEventListener('click', () => {
    setZoom(store.getState().viewport.zoom - 0.1);
  });
  zoomIn.addEventListener('click', () => {
    setZoom(store.getState().viewport.zoom + 0.1);
  });

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
    heading.textContent = strings.statusBar.aggregatesHeading;
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
    inheritHostTokens(statusbar, popover);
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
    setStrings(next: Strings): void {
      strings = next;
      // Already-open popover rebuilds with fresh strings on next show; for
      // already-rendered chooser repaint the heading label so an open menu
      // doesn't keep stale text.
      if (popoverVisible) buildChooser();
      refresh();
    },
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
