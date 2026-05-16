import { aggregateSelection, type SpreadsheetInstance } from '@libraz/formulon-cell';
import { focusMenuItem, handleMenuKeydown, prepareMenu } from './menu-a11y.js';

export const ACTIVE_CLASS = 'demo__rb--active';

type StatKey = 'sum' | 'avg' | 'count' | 'min' | 'max';
const STAT_KEYS: StatKey[] = ['sum', 'avg', 'count', 'min', 'max'];

export interface StatusProjectionCtx {
  getInst: () => SpreadsheetInstance | null;
  ribbonLang: 'ja' | 'en';
  statusSelection: HTMLElement | null;
  statusMetric: HTMLElement | null;
  statusObjects: HTMLElement | null;
  legacyCommandIds: Record<string, string>;
  getFormulaBarVisible: () => boolean;
  currentRibbonControlValue: (id: string) => string;
  ribbonSelectLabel: (wrap: HTMLElement, current: string) => string;
}

export interface StatusProjectionApi {
  readonly ACTIVE_CLASS: string;
  projectStatus: () => void;
  projectFormatToolbar: () => void;
  refreshObjectsBadge: (
    source: 'passthroughs' | 'tables',
    detail: { count: number; byCategory?: Record<string, number> },
  ) => void;
  setActive: (id: string, on: boolean) => void;
  setRibbonCommandActive: (command: string, on: boolean) => void;
  markCurrentLegacyRibbonBindings: () => void;
  persistStats: () => void;
  colLabel: (n: number) => string;
  fmt: (n: number) => string;
}

export const createStatusProjection = (ctx: StatusProjectionCtx): StatusProjectionApi => {
  const {
    getInst,
    statusSelection,
    statusMetric,
    statusObjects,
    legacyCommandIds,
    getFormulaBarVisible,
    currentRibbonControlValue,
    ribbonSelectLabel,
  } = ctx;

  const colLabel = (n: number): string => {
    let out = '';
    let v = n;
    do {
      out = String.fromCharCode(65 + (v % 26)) + out;
      v = Math.floor(v / 26) - 1;
    } while (v >= 0);
    return out;
  };

  const fmt = (n: number): string => {
    if (!Number.isFinite(n)) return '—';
    const abs = Math.abs(n);
    if (abs !== 0 && (abs < 0.01 || abs >= 1e9)) return n.toExponential(3);
    return n.toLocaleString('en-US', { maximumFractionDigits: 4 });
  };

  const activeStats: Set<StatKey> = (() => {
    try {
      const saved = localStorage.getItem('fc-status-stats');
      if (saved) return new Set(JSON.parse(saved) as StatKey[]);
    } catch {}
    return new Set<StatKey>(['sum', 'avg', 'count']);
  })();

  const persistStats = (): void => {
    try {
      localStorage.setItem('fc-status-stats', JSON.stringify(Array.from(activeStats)));
    } catch {}
  };

  // Composite badge showing both passthrough OOXML parts and spreadsheet Tables.
  // We accumulate the latest snapshot from each event and render together so
  // switching workbooks doesn't leak stale numbers from the previous one.
  const objectCounts = { passthroughs: 0, tables: 0, passByCat: {} as Record<string, number> };
  const refreshObjectsBadge = (
    source: 'passthroughs' | 'tables',
    detail: { count: number; byCategory?: Record<string, number> },
  ): void => {
    if (source === 'passthroughs') {
      objectCounts.passthroughs = detail.count;
      objectCounts.passByCat = detail.byCategory ?? {};
    } else {
      objectCounts.tables = detail.count;
    }
    if (!statusObjects) return;
    const parts: string[] = [];
    if (objectCounts.tables > 0)
      parts.push(`${objectCounts.tables} table${objectCounts.tables === 1 ? '' : 's'}`);
    const charts = objectCounts.passByCat.charts ?? 0;
    const drawings = objectCounts.passByCat.drawings ?? 0;
    const pivots = objectCounts.passByCat.pivotTables ?? 0;
    if (charts > 0) parts.push(`${charts} chart${charts === 1 ? '' : 's'}`);
    if (drawings > 0) parts.push(`${drawings} drawing${drawings === 1 ? '' : 's'}`);
    if (pivots > 0) parts.push(`${pivots} pivot${pivots === 1 ? '' : 's'}`);
    if (parts.length === 0) {
      statusObjects.hidden = true;
      statusObjects.textContent = '';
      return;
    }
    statusObjects.hidden = false;
    statusObjects.textContent = `objects · ${parts.join(', ')}`;
    statusObjects.title = 'Read-only — loaded from .xlsx but not editable in formulon-cell';
  };

  const projectStatus = (): void => {
    const inst = getInst();
    if (!inst) return;
    const s = inst.store.getState();
    const a = s.selection.active;
    const r = s.selection.range;

    if (statusSelection) {
      if (r.r0 === r.r1 && r.c0 === r.c1) {
        statusSelection.textContent = `${colLabel(a.col)}${a.row + 1}`;
      } else {
        const tl = `${colLabel(r.c0)}${r.r0 + 1}`;
        const br = `${colLabel(r.c1)}${r.r1 + 1}`;
        const cells = (r.r1 - r.r0 + 1) * (r.c1 - r.c0 + 1);
        statusSelection.textContent = `${tl}:${br} · ${cells} cells`;
      }
    }

    if (statusMetric) {
      const stats = aggregateSelection(s);
      if (stats.numericCount === 0) {
        statusMetric.textContent = '';
      } else {
        const parts: string[] = [];
        if (activeStats.has('sum')) parts.push(`Sum ${fmt(stats.sum)}`);
        if (activeStats.has('avg')) parts.push(`Avg ${fmt(stats.avg)}`);
        if (activeStats.has('count')) parts.push(`Count ${stats.numericCount}`);
        if (activeStats.has('min')) parts.push(`Min ${fmt(stats.min)}`);
        if (activeStats.has('max')) parts.push(`Max ${fmt(stats.max)}`);
        statusMetric.textContent = parts.join(' · ');
      }
    }
  };

  // Right-click on the status metric → checkbox menu to toggle stats.
  statusMetric?.addEventListener('contextmenu', (e) => {
    e.preventDefault();
    const opener =
      document.activeElement instanceof HTMLElement ? document.activeElement : statusMetric;
    const menu = document.createElement('div');
    menu.className = 'app__dropdown';
    prepareMenu(menu, 'Selection summary');
    menu.style.position = 'fixed';
    menu.style.left = `${e.clientX}px`;
    menu.style.bottom = `${window.innerHeight - e.clientY + 4}px`;
    menu.style.top = '';
    let cleanupMenuListeners = (): void => {};
    const closeMenu = (restoreFocus = false): void => {
      menu.remove();
      cleanupMenuListeners();
      if (restoreFocus) opener?.focus();
    };
    for (const key of STAT_KEYS) {
      const item = document.createElement('button');
      item.type = 'button';
      item.className = 'app__menu-item';
      item.setAttribute('role', 'menuitemcheckbox');
      item.setAttribute('aria-checked', activeStats.has(key) ? 'true' : 'false');
      item.tabIndex = -1;
      item.textContent = `${activeStats.has(key) ? '✓ ' : '  '}${key.toUpperCase()}`;
      item.addEventListener('click', () => {
        if (activeStats.has(key)) activeStats.delete(key);
        else activeStats.add(key);
        persistStats();
        projectStatus();
        const checked = activeStats.has(key);
        item.setAttribute('aria-checked', checked ? 'true' : 'false');
        item.textContent = `${checked ? '✓ ' : '  '}${key.toUpperCase()}`;
      });
      menu.appendChild(item);
    }
    const close = (ev: MouseEvent): void => {
      if (!menu.contains(ev.target as Node)) {
        closeMenu();
      }
    };
    menu.addEventListener('keydown', (event) => {
      handleMenuKeydown(event, menu, { close: closeMenu, restoreFocusTo: opener });
    });
    cleanupMenuListeners = () => document.removeEventListener('mousedown', close, true);
    document.body.appendChild(menu);
    focusMenuItem(menu);
    setTimeout(() => document.addEventListener('mousedown', close, true), 0);
  });

  const setActive = (id: string, on: boolean): void => {
    const el = document.getElementById(id);
    if (!el) return;
    el.classList.toggle(ACTIVE_CLASS, on);
  };

  const markCurrentLegacyRibbonBindings = (): void => {
    for (const command of Object.keys(legacyCommandIds)) {
      document
        .querySelector<HTMLButtonElement>(`button[data-ribbon-command="${command}"]`)
        ?.setAttribute('data-legacy-bound', '1');
    }
  };

  const setRibbonCommandActive = (command: string, on: boolean): void => {
    const el = document.querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`);
    if (!el) return;
    el.classList.toggle(ACTIVE_CLASS, on);
    el.setAttribute('aria-pressed', on ? 'true' : 'false');
  };

  const projectFormatToolbar = (): void => {
    const inst = getInst();
    if (!inst) return;
    const s = inst.store.getState();
    const a = s.selection.active;
    const key = `${a.sheet}:${a.row}:${a.col}`;
    const f = s.format.formats.get(key);
    setActive('btn-bold', !!f?.bold);
    setActive('btn-italic', !!f?.italic);
    setActive('btn-underline', !!f?.underline);
    setActive('btn-strike', !!f?.strike);
    setActive('btn-align-left', f?.align === 'left');
    setActive('btn-align-center', f?.align === 'center');
    setActive('btn-align-right', f?.align === 'right');
    setActive('btn-currency', f?.numFmt?.kind === 'currency');
    setActive('btn-percent', f?.numFmt?.kind === 'percent');
    setRibbonCommandActive('viewGridlines', s.ui.showGridLines !== false);
    setRibbonCommandActive('viewHeadings', s.ui.showHeaders !== false);
    setRibbonCommandActive('viewFormulas', !!s.ui.showFormulas);
    setRibbonCommandActive('showFormulasFormula', !!s.ui.showFormulas);
    setRibbonCommandActive('viewFormulaBar', getFormulaBarVisible());
    setRibbonCommandActive('viewR1C1', !!s.ui.r1c1);
    setRibbonCommandActive('viewNormal', s.ui.workbookView === 'normal');
    setRibbonCommandActive('viewPageLayout', s.ui.workbookView === 'pageLayout');
    setRibbonCommandActive('viewPageBreakPreview', s.ui.workbookView === 'pageBreakPreview');
    for (const wrap of document.querySelectorAll<HTMLElement>('[data-ribbon-select]')) {
      const id = wrap.dataset.ribbonSelect;
      if (!id) continue;
      const current = currentRibbonControlValue(id);
      const value = wrap.querySelector<HTMLElement>('.demo__rb-dd__value');
      if (value) value.textContent = ribbonSelectLabel(wrap, current);
      for (const option of wrap.querySelectorAll<HTMLElement>('.demo__rb-dd__opt')) {
        const selected = option.dataset.value === current;
        option.classList.toggle('demo__rb-dd__opt--selected', selected);
        option.setAttribute('aria-selected', selected ? 'true' : 'false');
      }
    }
    const fontColorSwatch = document.querySelector<HTMLElement>(
      '[data-ribbon-command="fontColor"] .demo__rb-color__swatch',
    );
    if (fontColorSwatch) fontColorSwatch.style.background = f?.color ?? '#201f1e';
    const fillColorSwatch = document.querySelector<HTMLElement>(
      '[data-ribbon-command="fillColor"] .demo__rb-color__swatch',
    );
    if (fillColorSwatch) fillColorSwatch.style.background = f?.fill ?? '#ffffff';
  };

  return {
    ACTIVE_CLASS,
    projectStatus,
    projectFormatToolbar,
    refreshObjectsBadge,
    setActive,
    setRibbonCommandActive,
    markCurrentLegacyRibbonBindings,
    persistStats,
    colLabel,
    fmt,
  };
};
