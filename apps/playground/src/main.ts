import {
  Spreadsheet,
  type SpreadsheetInstance,
  WorkbookHandle,
  aggregateSelection,
  applyMerge,
  applyUnmerge,
  attachFilterDropdown,
  autoSum,
  bumpDecimals,
  clearFilter,
  clearFormat,
  cycleBorders,
  cycleCurrency,
  cyclePercent,
  moveSheet,
  mutators,
  recordFormatChange,
  removeDuplicates,
  removeSheet,
  renameSheet,
  setAlign,
  setFreezePanes,
  setSheetHidden,
  setSheetZoom,
  sortRange,
  toggleBold,
  toggleItalic,
  toggleStrike,
  toggleUnderline,
  toggleWrap,
} from '@libraz/formulon-cell';

const sheetEl = document.getElementById('sheet');
const themeToggle = document.getElementById('theme-toggle') as HTMLButtonElement | null;
const themeLabel = document.getElementById('theme-label');
const docState = document.getElementById('doc-state');
const enginePill = document.getElementById('engine-pill');
const statusState = document.getElementById('status-state');
const statusSelection = document.getElementById('status-selection');
const statusMetric = document.getElementById('status-metric');
const statusEngine = document.getElementById('status-engine');
const statusObjects = document.getElementById('status-objects');

if (!sheetEl) throw new Error('#sheet missing');

// `paper` / `ink` are the core's theme names; the UI labels them Light / Dark.
type CoreTheme = 'paper' | 'ink';
type UiTheme = 'light' | 'dark';

const html = document.documentElement;
let uiTheme: UiTheme = (html.dataset.theme as UiTheme | undefined) ?? 'light';
const toCore = (t: UiTheme): CoreTheme => (t === 'dark' ? 'ink' : 'paper');

let inst: SpreadsheetInstance | null = null;

const seed = (wb: WorkbookHandle): void => {
  // Small editorial demo. Headers in row 0, items below, formulas in col D + E.
  wb.setText({ sheet: 0, row: 0, col: 0 }, 'item');
  wb.setText({ sheet: 0, row: 0, col: 1 }, 'qty');
  wb.setText({ sheet: 0, row: 0, col: 2 }, 'unit');
  wb.setText({ sheet: 0, row: 0, col: 3 }, 'subtotal');
  wb.setText({ sheet: 0, row: 0, col: 4 }, 'tax (8%)');

  const rows = [
    ['paper', 24, 0.42],
    ['vermillion ink', 6, 12.5],
    ['rule pen', 2, 8.9],
    ['draftsman pad', 1, 24.0],
    ['eraser', 3, 1.25],
  ] as const;

  rows.forEach(([name, qty, unit], i) => {
    const r = i + 1;
    wb.setText({ sheet: 0, row: r, col: 0 }, name);
    wb.setNumber({ sheet: 0, row: r, col: 1 }, qty);
    wb.setNumber({ sheet: 0, row: r, col: 2 }, unit);
    wb.setFormula({ sheet: 0, row: r, col: 3 }, `=B${r + 1}*C${r + 1}`);
    wb.setFormula({ sheet: 0, row: r, col: 4 }, `=D${r + 1}*0.08`);
  });

  wb.setText({ sheet: 0, row: 7, col: 0 }, 'total');
  wb.setFormula({ sheet: 0, row: 7, col: 3 }, '=SUM(D2:D6)');
  wb.setFormula({ sheet: 0, row: 7, col: 4 }, '=SUM(E2:E6)');
  wb.setFormula({ sheet: 0, row: 8, col: 3 }, '=D8+E8');
  wb.setText({ sheet: 0, row: 8, col: 0 }, 'with tax');

  wb.recalc();
};

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

type StatKey = 'sum' | 'avg' | 'count' | 'min' | 'max';
const STAT_KEYS: StatKey[] = ['sum', 'avg', 'count', 'min', 'max'];
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

// Composite badge showing both passthrough OOXML parts and Excel Tables.
// We accumulate the latest snapshot from each event and render together so
// switching workbooks doesn't leak stale numbers from the previous one.
const objectCounts = { passthroughs: 0, tables: 0, passByCat: {} as Record<string, number> };
function refreshObjectsBadge(
  source: 'passthroughs' | 'tables',
  detail: { count: number; byCategory?: Record<string, number> },
): void {
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
}

function projectStatus(): void {
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
}

// Right-click on the status metric → checkbox menu to toggle stats.
statusMetric?.addEventListener('contextmenu', (e) => {
  e.preventDefault();
  const menu = document.createElement('div');
  menu.className = 'app__dropdown';
  menu.style.position = 'fixed';
  menu.style.left = `${e.clientX}px`;
  menu.style.bottom = `${window.innerHeight - e.clientY + 4}px`;
  menu.style.top = '';
  for (const key of STAT_KEYS) {
    const item = document.createElement('button');
    item.type = 'button';
    item.className = 'app__menu-item';
    item.textContent = `${activeStats.has(key) ? '✓ ' : '  '}${key.toUpperCase()}`;
    item.addEventListener('click', () => {
      if (activeStats.has(key)) activeStats.delete(key);
      else activeStats.add(key);
      persistStats();
      projectStatus();
      menu.remove();
    });
    menu.appendChild(item);
  }
  const close = (ev: MouseEvent): void => {
    if (!menu.contains(ev.target as Node)) {
      menu.remove();
      document.removeEventListener('mousedown', close);
    }
  };
  document.body.appendChild(menu);
  setTimeout(() => document.addEventListener('mousedown', close), 0);
});

const ACTIVE_CLASS = 'app__tool--active';
const setActive = (id: string, on: boolean): void => {
  const el = document.getElementById(id);
  if (!el) return;
  el.classList.toggle(ACTIVE_CLASS, on);
};

function projectFormatToolbar(): void {
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
}

async function boot(): Promise<void> {
  // Default to the real WASM engine. Pass ?engine=stub to force the JS fallback
  // (useful for environments without crossOriginIsolated, or for diffing behavior).
  const params = new URLSearchParams(window.location.search);
  const preferStub = params.get('engine') === 'stub';
  const wb = await WorkbookHandle.createDefault({
    preferStub,
    onFallback: (reason) => {
      // eslint-disable-next-line no-console
      console.info('[formulon-cell]', reason);
    },
  });

  inst = await Spreadsheet.mount(sheetEl as HTMLElement, {
    theme: toCore(uiTheme),
    seed,
    workbook: wb,
    locale: 'en',
  });
  // Debug-only: expose for browser console / e2e poking. Safe to leave on the
  // playground build; the core package never references this global.
  (window as unknown as { __fcInst?: SpreadsheetInstance }).__fcInst = inst;

  filterDropdown = attachFilterDropdown({ store: inst.store });

  // Read-only badge — chart/drawing/pivot counts and Excel-Tables. Hidden
  //  until the loaded workbook actually carries any of these objects.
  inst.host.addEventListener('fc:passthroughs', (ev) => {
    const e = ev as CustomEvent<{ count: number; byCategory: Record<string, number> }>;
    refreshObjectsBadge('passthroughs', e.detail);
  });
  inst.host.addEventListener('fc:tables', (ev) => {
    const e = ev as CustomEvent<{ count: number }>;
    refreshObjectsBadge('tables', e.detail);
  });
  // Header chevron click → open the filter dropdown anchored under the header.
  inst.host.addEventListener('fc:openfilter', (ev) => {
    const e = ev as CustomEvent<{
      range: { sheet: number; r0: number; c0: number; r1: number; c1: number };
      col: number;
      anchor: { clientX: number; clientY: number; h: number };
    }>;
    const { range, col, anchor } = e.detail;
    filterDropdown?.open(range, col, {
      x: anchor.clientX,
      y: anchor.clientY - 4,
      h: anchor.h,
    });
  });

  const engineLabel = wb.isStub ? 'stub engine' : `formulon ${wb.version}`;
  if (enginePill) enginePill.textContent = `engine · ${engineLabel}`;
  if (statusEngine) statusEngine.textContent = engineLabel;
  if (docState) docState.textContent = 'Saved';
  if (statusState) statusState.textContent = 'Ready';

  inst.store.subscribe(() => {
    projectStatus();
    projectFormatToolbar();
    markDirty();
    refreshZoom();
  });
  projectStatus();
  projectFormatToolbar();
  renderSheetTabs();
  refreshZoom();

  // Reflect Format Painter state on the toolbar button (any path can deactivate
  // it — Esc, post-paint, or programmatic).
  inst.formatPainter?.subscribe((active, sticky) => {
    formatPainterBtn?.classList.toggle(ACTIVE_CLASS, active);
    formatPainterBtn?.classList.toggle('app__tool--sticky', active && sticky);
  });
}

document.getElementById('btn-autosum')?.addEventListener('click', () => {
  if (!inst) return;
  const result = autoSum(inst.store.getState(), inst.workbook);
  if (!result) return;
  mutators.replaceCells(inst.store, inst.workbook.cells(result.addr.sheet));
  mutators.setActive(inst.store, result.addr);
  (sheetEl as HTMLElement).focus();
});

document.getElementById('btn-undo')?.addEventListener('click', () => {
  if (!inst) return;
  if (!inst.undo()) return;
  (sheetEl as HTMLElement).focus();
});

document.getElementById('btn-redo')?.addEventListener('click', () => {
  if (!inst) return;
  if (!inst.redo()) return;
  (sheetEl as HTMLElement).focus();
});

// Format Painter — single click arms one-shot, double click arms sticky mode.
// Re-clicking the active button deactivates.
const formatPainterBtn = document.getElementById('btn-format-painter');
let painterStickyTimer: number | null = null;
formatPainterBtn?.addEventListener('click', () => {
  if (!inst) return;
  // Defer one-shot activation briefly so a follow-up click within the
  // dblclick window can promote it to sticky without painting twice.
  if (painterStickyTimer != null) return;
  painterStickyTimer = window.setTimeout(() => {
    painterStickyTimer = null;
    if (!inst) return;
    const fp = inst.formatPainter;
    if (!fp) return;
    if (fp.isActive()) fp.deactivate();
    else fp.activate(false);
    (sheetEl as HTMLElement).focus();
    formatPainterBtn?.classList.toggle(ACTIVE_CLASS, fp.isActive());
  }, 220);
});
formatPainterBtn?.addEventListener('dblclick', () => {
  if (!inst) return;
  if (painterStickyTimer != null) {
    clearTimeout(painterStickyTimer);
    painterStickyTimer = null;
  }
  const fp = inst.formatPainter;
  if (!fp) return;
  fp.activate(true);
  (sheetEl as HTMLElement).focus();
  formatPainterBtn?.classList.toggle(ACTIVE_CLASS, fp.isActive());
});

const wireFormat = (
  id: string,
  fn: (
    state: ReturnType<SpreadsheetInstance['store']['getState']>,
    store: SpreadsheetInstance['store'],
  ) => void,
): void => {
  document.getElementById(id)?.addEventListener('click', () => {
    const i = inst;
    if (!i) return;
    // Wrap each toolbar mutation so Cmd+Z reverts the format change.
    recordFormatChange(i.history, i.store, () => {
      fn(i.store.getState(), i.store);
    });
    (sheetEl as HTMLElement).focus();
  });
};

wireFormat('btn-bold', toggleBold);
wireFormat('btn-italic', toggleItalic);
wireFormat('btn-underline', toggleUnderline);
wireFormat('btn-strike', toggleStrike);
wireFormat('btn-currency', cycleCurrency);
wireFormat('btn-percent', cyclePercent);
wireFormat('btn-borders', cycleBorders);
wireFormat('btn-align-left', (state, store) => setAlign(state, store, 'left'));
wireFormat('btn-align-center', (state, store) => setAlign(state, store, 'center'));
wireFormat('btn-align-right', (state, store) => setAlign(state, store, 'right'));
wireFormat('btn-decimals-up', (state, store) => bumpDecimals(state, store, 1));
wireFormat('btn-decimals-down', (state, store) => bumpDecimals(state, store, -1));

void clearFormat; // Reserved for a "Clear formatting" menu item; keep the import live.

// ── Freeze Panes menu ─────────────────────────────────────────────────────
const freezeBtn = document.getElementById('btn-freeze');
const freezeMenu = document.getElementById('menu-freeze');

const closeFreezeMenu = (): void => {
  if (!freezeMenu) return;
  freezeMenu.hidden = true;
  freezeBtn?.setAttribute('aria-expanded', 'false');
};
const openFreezeMenu = (): void => {
  if (!freezeMenu) return;
  freezeMenu.hidden = false;
  freezeBtn?.setAttribute('aria-expanded', 'true');
  // Focus first item for keyboard nav
  (freezeMenu.querySelector('button') as HTMLButtonElement | null)?.focus();
};

freezeBtn?.addEventListener('click', (e) => {
  e.stopPropagation();
  if (!freezeMenu) return;
  if (freezeMenu.hidden) openFreezeMenu();
  else closeFreezeMenu();
});

document.addEventListener('mousedown', (e) => {
  if (!freezeMenu || freezeMenu.hidden) return;
  if (freezeMenu.contains(e.target as Node)) return;
  if (freezeBtn?.contains(e.target as Node)) return;
  closeFreezeMenu();
});

document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && !freezeMenu?.hidden) closeFreezeMenu();
});

freezeMenu?.querySelectorAll<HTMLButtonElement>('[data-freeze]').forEach((btn) => {
  btn.addEventListener('click', () => {
    const i = inst;
    if (!i) return;
    const action = btn.dataset.freeze;
    const s = i.store.getState();

    let rows = s.layout.freezeRows;
    let cols = s.layout.freezeCols;
    if (action === 'row') {
      rows = 1;
      cols = 0;
    } else if (action === 'col') {
      rows = 0;
      cols = 1;
    } else if (action === 'selection') {
      // Excel: freeze rows above and cols left of the active cell.
      rows = s.selection.active.row;
      cols = s.selection.active.col;
    } else if (action === 'off') {
      rows = 0;
      cols = 0;
    }

    setFreezePanes(i.store, i.history, rows, cols, i.workbook);
    closeFreezeMenu();
    (sheetEl as HTMLElement).focus();
  });
});

themeToggle?.addEventListener('click', () => {
  uiTheme = uiTheme === 'light' ? 'dark' : 'light';
  html.dataset.theme = uiTheme;
  if (themeLabel) themeLabel.textContent = uiTheme === 'light' ? 'Light' : 'Dark';
  themeToggle.setAttribute('aria-pressed', uiTheme === 'dark' ? 'true' : 'false');
  // Theme is a UI-only preference; don't let the resulting store update mark the workbook as edited.
  suppressDirty = true;
  inst?.setTheme(toCore(uiTheme));
  suppressDirty = false;
});

// ── File menu (New / Open / Save / Save As) ───────────────────────────────
const fileMenuBtn = document.getElementById('menu-file');
const fileMenuDrop = document.getElementById('menu-file-dropdown');
const fileInput = document.getElementById('file-input') as HTMLInputElement | null;

let docName = 'Untitled';

const setDocName = (name: string): void => {
  docName = name;
  const el = document.getElementById('doc-name');
  if (el) el.textContent = name;
};

const openFileMenu = (): void => {
  if (!fileMenuDrop) return;
  fileMenuDrop.hidden = false;
  fileMenuBtn?.setAttribute('aria-expanded', 'true');
};
const closeFileMenu = (): void => {
  if (!fileMenuDrop) return;
  fileMenuDrop.hidden = true;
  fileMenuBtn?.setAttribute('aria-expanded', 'false');
};

fileMenuBtn?.addEventListener('click', (e) => {
  e.stopPropagation();
  if (!fileMenuDrop) return;
  if (fileMenuDrop.hidden) openFileMenu();
  else closeFileMenu();
});

document.addEventListener('mousedown', (e) => {
  if (!fileMenuDrop || fileMenuDrop.hidden) return;
  if (fileMenuDrop.contains(e.target as Node)) return;
  if (fileMenuBtn?.contains(e.target as Node)) return;
  closeFileMenu();
});

document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && !fileMenuDrop?.hidden) closeFileMenu();
});

const triggerOpen = (): void => fileInput?.click();

const downloadBytes = (bytes: Uint8Array, filename: string): void => {
  const blob = new Blob([bytes as BlobPart], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1_000);
};

const triggerSave = (filename = `${docName.replace(/\.xlsx$/i, '')}.xlsx`): void => {
  if (!inst) return;
  try {
    const bytes = inst.workbook.save();
    downloadBytes(bytes, filename);
    if (docState) docState.textContent = 'Saved';
  } catch (err) {
    // eslint-disable-next-line no-console
    console.error('save failed', err);
    if (docState) docState.textContent = 'Save failed';
  }
};

const triggerSaveAs = (): void => {
  const name = prompt('File name', docName) ?? null;
  if (!name) return;
  setDocName(name);
  triggerSave(name.endsWith('.xlsx') ? name : `${name}.xlsx`);
};

const loadXlsxFile = async (file: File): Promise<void> => {
  if (!inst) return;
  if (docState) docState.textContent = 'Loading…';
  try {
    const buf = await file.arrayBuffer();
    const next = await WorkbookHandle.loadBytes(new Uint8Array(buf));
    await inst.setWorkbook(next);
    setDocName(file.name);
    if (docState) docState.textContent = 'Saved';
    renderSheetTabs();
  } catch (err) {
    // eslint-disable-next-line no-console
    console.error('open failed', err);
    if (docState) docState.textContent = 'Open failed';
    alert(err instanceof Error ? err.message : String(err));
  }
};

fileInput?.addEventListener('change', () => {
  const f = fileInput.files?.[0];
  if (f) void loadXlsxFile(f);
  fileInput.value = ''; // allow same-file re-open
});

fileMenuDrop?.querySelectorAll<HTMLButtonElement>('[data-file]').forEach((btn) => {
  btn.addEventListener('click', () => {
    const action = btn.dataset.file;
    closeFileMenu();
    if (!inst) return;
    if (action === 'new') {
      void (async () => {
        const next = await WorkbookHandle.createDefault();
        await inst?.setWorkbook(next);
        setDocName('Untitled');
        if (docState) docState.textContent = 'Saved';
        renderSheetTabs();
      })();
    } else if (action === 'open') {
      triggerOpen();
    } else if (action === 'save') {
      triggerSave();
    } else if (action === 'save-as') {
      triggerSaveAs();
    }
  });
});

// Drag & drop xlsx onto the page.
window.addEventListener('dragover', (e) => {
  if (!e.dataTransfer) return;
  e.preventDefault();
  e.dataTransfer.dropEffect = 'copy';
});
window.addEventListener('drop', (e) => {
  e.preventDefault();
  const f = e.dataTransfer?.files?.[0];
  if (!f) return;
  if (!/\.xlsx?$/i.test(f.name)) return;
  void loadXlsxFile(f);
});

// Ctrl/Cmd-O / Ctrl/Cmd-S / Ctrl/Cmd-N for file actions.
window.addEventListener('keydown', (e) => {
  if (!(e.ctrlKey || e.metaKey)) return;
  const k = e.key.toLowerCase();
  if (k === 'o') {
    e.preventDefault();
    triggerOpen();
  } else if (k === 's') {
    e.preventDefault();
    if (e.shiftKey) triggerSaveAs();
    else triggerSave();
  } else if (k === 'n' && !e.shiftKey) {
    // Ctrl+N — create a fresh workbook in place.
    e.preventDefault();
    void (async () => {
      const next = await WorkbookHandle.createDefault();
      await inst?.setWorkbook(next);
      setDocName('Untitled');
      renderSheetTabs();
    })();
  }
});

// Mark the document dirty whenever any cell change flows through.
let dirtyTimer: number | null = null;
let suppressDirty = false;
const markDirty = (): void => {
  if (suppressDirty) return;
  if (dirtyTimer != null) return;
  dirtyTimer = window.setTimeout(() => {
    dirtyTimer = null;
    if (docState) docState.textContent = 'Edited';
  }, 200);
};
// Subscribe once boot completes — see end of boot().

// ── View menu (Show Formulas / R1C1 / Grid / Headers toggles) ────────────
const viewBtn = document.getElementById('menu-view');
const viewDrop = document.getElementById('menu-view-dropdown');
const closeViewMenu = (): void => {
  if (!viewDrop) return;
  viewDrop.hidden = true;
  viewBtn?.setAttribute('aria-expanded', 'false');
};
const refreshViewMenu = (): void => {
  if (!inst || !viewDrop) return;
  const ui = inst.store.getState().ui;
  const update = (action: string, on: boolean): void => {
    const item = viewDrop.querySelector<HTMLElement>(`[data-view="${action}"] [data-fc-check]`);
    if (item) item.textContent = on ? '✓' : '';
  };
  update('show-formulas', !!ui.showFormulas);
  update('r1c1', !!ui.r1c1);
  update('grid', ui.showGridLines !== false);
  update('headers', ui.showHeaders !== false);
};
viewBtn?.addEventListener('click', (e) => {
  e.stopPropagation();
  if (!viewDrop) return;
  refreshViewMenu();
  viewDrop.hidden = !viewDrop.hidden;
  viewBtn.setAttribute('aria-expanded', viewDrop.hidden ? 'false' : 'true');
});
document.addEventListener('mousedown', (e) => {
  if (!viewDrop || viewDrop.hidden) return;
  if (viewDrop.contains(e.target as Node) || viewBtn?.contains(e.target as Node)) return;
  closeViewMenu();
});
viewDrop?.querySelectorAll<HTMLButtonElement>('[data-view]').forEach((btn) => {
  btn.addEventListener('click', () => {
    if (!inst) return;
    const action = btn.dataset.view;
    const ui = inst.store.getState().ui;
    if (action === 'show-formulas') mutators.setShowFormulas(inst.store, !ui.showFormulas);
    else if (action === 'r1c1') mutators.setR1C1(inst.store, !ui.r1c1);
    else if (action === 'grid') mutators.setShowGridLines(inst.store, !ui.showGridLines);
    else if (action === 'headers') mutators.setShowHeaders(inst.store, !ui.showHeaders);
    refreshViewMenu();
  });
});

// ── Tools menu (Iterative / Names / Conditional) ─────────────────────────
const toolsBtn = document.getElementById('menu-tools');
const toolsDrop = document.getElementById('menu-tools-dropdown');
const closeToolsMenu = (): void => {
  if (!toolsDrop) return;
  toolsDrop.hidden = true;
  toolsBtn?.setAttribute('aria-expanded', 'false');
};
toolsBtn?.addEventListener('click', (e) => {
  e.stopPropagation();
  if (!toolsDrop) return;
  toolsDrop.hidden = !toolsDrop.hidden;
  toolsBtn.setAttribute('aria-expanded', toolsDrop.hidden ? 'false' : 'true');
});
document.addEventListener('mousedown', (e) => {
  if (!toolsDrop || toolsDrop.hidden) return;
  if (toolsDrop.contains(e.target as Node) || toolsBtn?.contains(e.target as Node)) return;
  closeToolsMenu();
});
toolsDrop?.querySelectorAll<HTMLButtonElement>('[data-tools]').forEach((btn) => {
  btn.addEventListener('click', () => {
    if (!inst) return;
    const action = btn.dataset.tools;
    closeToolsMenu();
    if (action === 'iterative') inst.openIterativeDialog();
    else if (action === 'named') inst.openNamedRangeDialog();
    else if (action === 'conditional') inst.openConditionalDialog();
  });
});

// ── Sheet tabs ───────────────────────────────────────────────────────────
const tabsList = document.getElementById('sheet-tabs');
const tabAddBtn = document.getElementById('btn-sheet-add');
const tabPrevBtn = document.getElementById('btn-sheet-prev');
const tabNextBtn = document.getElementById('btn-sheet-next');

const renderSheetTabs = (): void => {
  if (!inst || !tabsList) return;
  const wb = inst.workbook;
  const state = inst.store.getState();
  const activeIdx = state.data.sheetIndex;
  const hidden = state.layout.hiddenSheets;
  const n = wb.sheetCount;
  tabsList.replaceChildren();
  for (let i = 0; i < n; i += 1) {
    if (hidden.has(i)) continue;
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'app__tab';
    if (i === activeIdx) btn.classList.add('app__tab--active');
    btn.setAttribute('role', 'tab');
    btn.setAttribute('aria-selected', i === activeIdx ? 'true' : 'false');
    const label = document.createElement('span');
    label.className = 'app__tab-label';
    label.textContent = wb.sheetName(i);
    btn.appendChild(label);
    btn.addEventListener('click', () => switchSheet(i));
    btn.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      openTabMenu(i, e.clientX, e.clientY);
    });
    tabsList.appendChild(btn);
  }
  // "Unhide…" affordance — surfaced as an extra tab pill when at least one
  // sheet is hidden. Click opens a list of hidden sheets to restore.
  if (hidden.size > 0) {
    const unhide = document.createElement('button');
    unhide.type = 'button';
    unhide.className = 'app__tab app__tab--unhide';
    unhide.textContent = `Unhide… (${hidden.size})`;
    unhide.addEventListener('click', (e) => {
      const r = (e.currentTarget as HTMLElement).getBoundingClientRect();
      openUnhideMenu(r.left, r.bottom);
    });
    tabsList.appendChild(unhide);
  }
};

const openUnhideMenu = (x: number, y: number): void => {
  if (!inst) return;
  closeTabMenu();
  const wb = inst.workbook;
  const store = inst.store;
  const hidden = store.getState().layout.hiddenSheets;
  if (hidden.size === 0) return;

  const menu = document.createElement('div');
  menu.className = 'app__menu';
  menu.style.position = 'fixed';
  menu.style.left = `${x}px`;
  menu.style.top = `${y}px`;
  menu.style.zIndex = '90';

  for (const i of Array.from(hidden).sort((a, b) => a - b)) {
    const it = document.createElement('button');
    it.type = 'button';
    it.className = 'app__menu-item';
    it.textContent = wb.sheetName(i);
    it.addEventListener('click', () => {
      closeTabMenu();
      if (setSheetHidden(store, wb, inst?.history ?? null, i, false)) {
        renderSheetTabs();
      }
    });
    menu.appendChild(it);
  }

  document.body.appendChild(menu);
  tabMenuEl = menu;

  const rect = menu.getBoundingClientRect();
  if (rect.right > window.innerWidth) {
    menu.style.left = `${Math.max(0, window.innerWidth - rect.width - 4)}px`;
  }
  if (rect.bottom > window.innerHeight) {
    menu.style.top = `${Math.max(0, window.innerHeight - rect.height - 4)}px`;
  }

  const onDocDown = (ev: MouseEvent): void => {
    if (!tabMenuEl) return;
    if (ev.target instanceof Node && tabMenuEl.contains(ev.target)) return;
    closeTabMenu();
    document.removeEventListener('mousedown', onDocDown, true);
  };
  document.addEventListener('mousedown', onDocDown, true);
};

let tabMenuEl: HTMLDivElement | null = null;
const closeTabMenu = (): void => {
  if (!tabMenuEl) return;
  tabMenuEl.remove();
  tabMenuEl = null;
};
const openTabMenu = (idx: number, x: number, y: number): void => {
  if (!inst) return;
  closeTabMenu();
  const wb = inst.workbook;
  const store = inst.store;
  const n = wb.sheetCount;

  const menu = document.createElement('div');
  menu.className = 'app__menu';
  menu.style.position = 'fixed';
  menu.style.left = `${x}px`;
  menu.style.top = `${y}px`;
  menu.style.zIndex = '90';

  const addItem = (text: string, disabled: boolean, onClick: () => void): void => {
    const it = document.createElement('button');
    it.type = 'button';
    it.className = 'app__menu-item';
    it.textContent = text;
    it.disabled = disabled;
    it.style.opacity = disabled ? '0.45' : '1';
    it.style.cursor = disabled ? 'not-allowed' : 'pointer';
    it.addEventListener('click', () => {
      if (disabled) return;
      closeTabMenu();
      onClick();
    });
    menu.appendChild(it);
  };

  addItem('Rename…', false, () => {
    if (!inst) return;
    const cur = wb.sheetName(idx);
    const next = window.prompt('Sheet name:', cur);
    if (next == null || next === cur || next.length === 0) return;
    if (renameSheet(wb, idx, next)) renderSheetTabs();
  });
  addItem('Delete', n <= 1, () => {
    if (!inst) return;
    if (!window.confirm(`Delete "${wb.sheetName(idx)}"?`)) return;
    if (removeSheet(store, wb, idx)) {
      const newActive = store.getState().data.sheetIndex;
      mutators.replaceCells(store, wb.cells(newActive));
      renderSheetTabs();
    }
  });
  // Hide tab — disabled when this is the last visible sheet and when the
  // engine doesn't expose `setSheetTabHidden`.
  const visibleCount = n - store.getState().layout.hiddenSheets.size;
  const hideDisabled = !wb.capabilities.sheetTabHidden || visibleCount <= 1;
  addItem('Hide tab', hideDisabled, () => {
    if (!inst) return;
    if (setSheetHidden(store, wb, inst.history, idx, true)) {
      const newActive = store.getState().data.sheetIndex;
      mutators.replaceCells(store, wb.cells(newActive));
      renderSheetTabs();
    }
  });
  const sep = document.createElement('div');
  sep.className = 'app__menu-sep';
  menu.appendChild(sep);
  addItem('Move left', idx === 0, () => {
    if (!inst) return;
    if (moveSheet(store, wb, idx, idx - 1)) renderSheetTabs();
  });
  addItem('Move right', idx >= n - 1, () => {
    if (!inst) return;
    if (moveSheet(store, wb, idx, idx + 1)) renderSheetTabs();
  });

  document.body.appendChild(menu);
  tabMenuEl = menu;

  // Clamp into viewport.
  const rect = menu.getBoundingClientRect();
  const vw = window.innerWidth;
  const vh = window.innerHeight;
  if (rect.right > vw) menu.style.left = `${Math.max(0, vw - rect.width - 4)}px`;
  if (rect.bottom > vh) menu.style.top = `${Math.max(0, vh - rect.height - 4)}px`;

  const onDocClick = (ev: MouseEvent): void => {
    if (!tabMenuEl) return;
    if (ev.target instanceof Node && tabMenuEl.contains(ev.target)) return;
    closeTabMenu();
    document.removeEventListener('mousedown', onDocClick, true);
    document.removeEventListener('keydown', onDocKey, true);
  };
  const onDocKey = (ev: KeyboardEvent): void => {
    if (ev.key === 'Escape') {
      closeTabMenu();
      document.removeEventListener('mousedown', onDocClick, true);
      document.removeEventListener('keydown', onDocKey, true);
    }
  };
  document.addEventListener('mousedown', onDocClick, true);
  document.addEventListener('keydown', onDocKey, true);
};

const switchSheet = (idx: number): void => {
  if (!inst) return;
  const n = inst.workbook.sheetCount;
  if (idx < 0 || idx >= n) return;
  if (inst.store.getState().data.sheetIndex === idx) return;
  mutators.setSheetIndex(inst.store, idx);
  mutators.replaceCells(inst.store, inst.workbook.cells(idx));
  renderSheetTabs();
  (sheetEl as HTMLElement).focus();
};

tabAddBtn?.addEventListener('click', () => {
  if (!inst) return;
  const idx = inst.workbook.addSheet();
  if (idx < 0) return;
  // The wb.subscribe handler in mount.ts will pick up sheet-add as a no-op for cells,
  // but we re-render tabs and switch to the new sheet here.
  renderSheetTabs();
  switchSheet(idx);
});

// Zoom display + rail. Excel uses a 10–400% range; we clamp to 25–300% which
// matches what the engine accepts comfortably.
const zoomDisplay = document.getElementById('zoom-display');
const zoomRailFill = document.getElementById('zoom-rail-fill');
const zoomRailThumb = document.getElementById('zoom-rail-thumb');
const Z_MIN = 0.25;
const Z_MAX = 3.0;
const refreshZoom = (): void => {
  if (!inst) return;
  const z = inst.store.getState().viewport.zoom;
  if (zoomDisplay) zoomDisplay.textContent = `${Math.round(z * 100)}%`;
  // Map [Z_MIN..Z_MAX] → [0..100%] on the rail. The thumb tracks the fill.
  const pct = Math.max(0, Math.min(1, (z - Z_MIN) / (Z_MAX - Z_MIN))) * 100;
  if (zoomRailFill) zoomRailFill.style.width = `${pct}%`;
  if (zoomRailThumb) zoomRailThumb.style.left = `${pct}%`;
};
const stepZoom = (delta: number): void => {
  if (!inst) return;
  const z = inst.store.getState().viewport.zoom;
  const next = Math.max(Z_MIN, Math.min(Z_MAX, Math.round((z + delta) * 100) / 100));
  if (next === z) return;
  setSheetZoom(inst.store, next, inst.workbook);
  refreshZoom();
};
zoomDisplay?.addEventListener('click', () => {
  if (!inst) return;
  // Cycle 75 → 100 → 125 → 150 → 75 …
  const z = inst.store.getState().viewport.zoom;
  const next = z >= 1.5 ? 0.75 : Math.round((z + 0.25) * 100) / 100;
  setSheetZoom(inst.store, next, inst.workbook);
  refreshZoom();
});
document.getElementById('btn-zoom-out')?.addEventListener('click', () => stepZoom(-0.1));
document.getElementById('btn-zoom-in')?.addEventListener('click', () => stepZoom(0.1));

tabPrevBtn?.addEventListener('click', () => {
  if (!inst) return;
  switchSheet(inst.store.getState().data.sheetIndex - 1);
});
tabNextBtn?.addEventListener('click', () => {
  if (!inst) return;
  switchSheet(inst.store.getState().data.sheetIndex + 1);
});

// ── Merge / Wrap / Sort buttons ───────────────────────────────────────────
document.getElementById('btn-merge')?.addEventListener('click', () => {
  if (!inst) return;
  const s = inst.store.getState();
  const r = s.selection.range;
  // If the range is already merged, unmerge; otherwise merge.
  const anchorAt0 = s.merges.byAnchor.get(`${r.sheet}:${r.r0}:${r.c0}`);
  const isExactMerge =
    anchorAt0 &&
    r.r0 === anchorAt0.r0 &&
    r.c0 === anchorAt0.c0 &&
    r.r1 === anchorAt0.r1 &&
    r.c1 === anchorAt0.c1;
  if (isExactMerge) {
    applyUnmerge(inst.store, inst.workbook, inst.history, r);
  } else {
    applyMerge(inst.store, inst.workbook, inst.history, r);
  }
  (sheetEl as HTMLElement).focus();
});

document.getElementById('btn-wrap')?.addEventListener('click', () => {
  if (!inst) return;
  recordFormatChange(inst.history, inst.store, () => {
    toggleWrap(inst!.store.getState(), inst!.store);
  });
  (sheetEl as HTMLElement).focus();
});

const sortBtn = document.getElementById('btn-sort');
const sortMenu = document.getElementById('menu-sort');
const closeSortMenu = (): void => {
  if (!sortMenu) return;
  sortMenu.hidden = true;
  sortBtn?.setAttribute('aria-expanded', 'false');
};
const openSortMenu = (): void => {
  if (!sortMenu) return;
  sortMenu.hidden = false;
  sortBtn?.setAttribute('aria-expanded', 'true');
};
sortBtn?.addEventListener('click', (e) => {
  e.stopPropagation();
  if (!sortMenu) return;
  if (sortMenu.hidden) openSortMenu();
  else closeSortMenu();
});
document.addEventListener('mousedown', (e) => {
  if (!sortMenu || sortMenu.hidden) return;
  if (sortMenu.contains(e.target as Node)) return;
  if (sortBtn?.contains(e.target as Node)) return;
  closeSortMenu();
});
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && !sortMenu?.hidden) closeSortMenu();
});

sortMenu?.querySelectorAll<HTMLButtonElement>('[data-sort]').forEach((btn) => {
  btn.addEventListener('click', () => {
    if (!inst) return;
    const action = btn.dataset.sort;
    closeSortMenu();
    const state = inst.store.getState();
    const r = state.selection.range;
    if (r.r0 === r.r1 && r.c0 === r.c1) return; // single cell — nothing to sort
    if (action === 'asc' || action === 'desc') {
      sortRange(state, inst.store, inst.workbook, r, {
        byCol: r.c0,
        direction: action,
      });
      mutators.replaceCells(inst.store, inst.workbook.cells(state.data.sheetIndex));
    } else if (action === 'dedupe') {
      const removed = removeDuplicates(state, inst.store, inst.workbook, r);
      mutators.replaceCells(inst.store, inst.workbook.cells(state.data.sheetIndex));
      if (statusMetric)
        statusMetric.textContent = `Removed ${removed} duplicate row${removed === 1 ? '' : 's'}`;
    } else if (action === 'filter') {
      // Stamp the autofilter range so column headers paint the chevron, even
      // before the user picks any filter values.
      mutators.setFilterRange(inst.store, r);
      const sheetRect = sheetEl?.getBoundingClientRect() ?? { left: 0, top: 0, height: 0 };
      filterDropdown?.open(r, r.c0, {
        x: sheetRect.left + 80,
        y: sheetRect.top,
        h: 32,
      });
    } else if (action === 'filter-clear') {
      clearFilter(state, inst.store, r);
    } else if (action === 'conditional') {
      inst.openConditionalDialog();
    } else if (action === 'named') {
      inst.openNamedRangeDialog();
    }
    (sheetEl as HTMLElement).focus();
  });
});

let filterDropdown: ReturnType<typeof attachFilterDropdown> | null = null;

boot().catch((err) => {
  // eslint-disable-next-line no-console
  console.error('formulon-cell boot failed', err);
  if (sheetEl) {
    sheetEl.innerHTML = `<pre style="padding:24px;color:#d24545;font-family:'IBM Plex Mono',monospace;white-space:pre-wrap">${
      err instanceof Error ? (err.stack ?? err.message) : String(err)
    }</pre>`;
  }
});
