import { applyFilter, clearFilter, distinctValues } from '../commands/filter.js';
import { type History, recordSlicersChange } from '../commands/history.js';
import { parseRangeRef } from '../engine/range-resolver.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { mutators, type SlicerSpec, type SpreadsheetStore } from '../store/store.js';

export interface SlicerDeps {
  /** Element the floating panels attach to. Each panel is appended as a
   *  child of `host` so theme variables and focus scope are inherited. */
  host: HTMLElement;
  store: SpreadsheetStore;
  /** Lazy workbook getter — keeps a fresh reference even when `setWorkbook`
   *  swaps the engine. */
  getWb: () => WorkbookHandle;
  /** Optional unified history. When supplied, every chip toggle and
   *  add/remove is wrapped in `recordSlicersChange` so undo/redo restores
   *  prior selections. */
  history?: History | null;
  strings?: Strings;
}

export interface SlicerHandle {
  /** Add a new slicer for `tableName` + `column`. Returns the freshly-built
   *  spec (including the auto-assigned id) for caller convenience. Throws
   *  when the table or column can't be resolved against the workbook. */
  addSlicer(input: {
    tableName: string;
    column: string;
    selected?: readonly string[];
    x?: number;
    y?: number;
  }): SlicerSpec;
  /** Remove the slicer with `id`. No-op when absent. */
  removeSlicer(id: string): void;
  /** Re-pull distinct values + repaint chips. Call after a recalc batch so
   *  the panel reflects the freshest data. */
  refresh(): void;
  detach(): void;
  setStrings(next: Strings): void;
}

/** Lookup an Excel Table on the workbook by case-insensitive name. Falls
 *  back to display-name match so callers can use either the tab label or
 *  the OOXML name. */
function findTable(
  wb: WorkbookHandle,
  tableName: string,
): {
  name: string;
  displayName: string;
  ref: string;
  sheetIndex: number;
  columns: string[];
} | null {
  const target = tableName.toLowerCase();
  for (const t of wb.getTables()) {
    if (t.name.toLowerCase() === target || t.displayName.toLowerCase() === target) return t;
  }
  return null;
}

/** Resolve a SlicerSpec to its (range, byCol) tuple. Returns null when the
 *  table or column has gone away (e.g. workbook swap). */
function resolveSpec(wb: WorkbookHandle, spec: SlicerSpec): { range: Range; byCol: number } | null {
  const table = findTable(wb, spec.tableName);
  if (!table) return null;
  const parsed = parseRangeRef(table.ref);
  if (!parsed) return null;
  const colIdx = table.columns.indexOf(spec.column);
  if (colIdx < 0) return null;
  return {
    range: {
      sheet: table.sheetIndex,
      r0: parsed.r0,
      c0: parsed.c0,
      r1: parsed.r1,
      c1: parsed.c1,
    },
    byCol: parsed.c0 + colIdx,
  };
}

/** Recompute the autoFilter state for every active slicer. Each slicer
 *  contributes a value-set predicate against its column; v1 applies them
 *  sequentially so the last one wins (Excel's slicer semantics with a
 *  single connected table). Empty selection on a slicer means
 *  "include all" — that slicer doesn't constrain the column. */
function recomputeFilters(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  slicers: readonly SlicerSpec[],
): void {
  // Group slicers by their resolved table range; clear that range first so
  //  toggling a chip widens the visible rows when appropriate.
  const ranges: { range: Range; byCol: number; selected: readonly string[] }[] = [];
  for (const sp of slicers) {
    const r = resolveSpec(wb, sp);
    if (!r) continue;
    ranges.push({ range: r.range, byCol: r.byCol, selected: sp.selected });
  }
  const cleared = new Set<string>();
  for (const entry of ranges) {
    const key = rangeKey(entry.range);
    if (cleared.has(key)) continue;
    clearFilter(store.getState(), store, entry.range);
    cleared.add(key);
  }
  for (const entry of ranges) {
    if (entry.selected.length === 0) continue;
    const want = new Set(entry.selected);
    applyFilter(store.getState(), store, entry.range, entry.byCol, (cell) => {
      const key = cellToKey(cell?.value);
      return want.has(key);
    });
  }
}

const rangeKey = (r: Range): string => `${r.sheet}:${r.r0}:${r.c0}:${r.r1}:${r.c1}`;

const cellToKey = (v: unknown): string => {
  if (!v || typeof v !== 'object') return '';
  const cv = v as { kind: string; value?: unknown };
  if (cv.kind === 'number') return String(cv.value);
  if (cv.kind === 'text') return String(cv.value ?? '');
  if (cv.kind === 'bool') return cv.value ? 'TRUE' : 'FALSE';
  return '';
};

/** Default offset (px) for fresh panels, applied when the spec doesn't
 *  carry explicit `x`/`y` coordinates. Matches `.fc-watch` panel chrome. */
const DEFAULT_OFFSET_X = 16;
const DEFAULT_OFFSET_Y = 16;

interface PanelEntry {
  root: HTMLDivElement;
  body: HTMLDivElement;
  title: HTMLSpanElement;
  closeBtn: HTMLButtonElement;
  /** Detach DOM listeners specific to this panel. */
  dispose(): void;
}

/**
 * Excel-style table slicer manager. Renders one floating panel per
 * `SlicerSpec` in `state.slicers.slicers`; chip clicks pipe through
 * `setSlicerSelected` and recompute the table's autoFilter from the union
 * of every slicer's selection. Subscribes to store changes so chips repaint
 * when the underlying spec list mutates (history undo/redo, external API).
 */
export function attachSlicer(deps: SlicerDeps): SlicerHandle {
  const { host, store, getWb, history } = deps;
  let strings = deps.strings ?? defaultStrings;
  const panels = new Map<string, PanelEntry>();

  let idCounter = 0;
  const nextId = (): string => {
    idCounter += 1;
    // Combine a monotonic counter with a short random suffix so concurrent
    //  attaches across mounts don't collide.
    return `slicer-${idCounter}-${Math.random().toString(36).slice(2, 6)}`;
  };

  const buildPanel = (spec: SlicerSpec): PanelEntry => {
    const root = document.createElement('div');
    root.className = 'fc-slicer';
    root.dataset.fcSlicer = spec.id;
    root.setAttribute('role', 'region');
    root.setAttribute('aria-label', `${strings.slicer.title}: ${spec.column}`);
    root.style.position = 'absolute';
    root.style.left = `${spec.x ?? DEFAULT_OFFSET_X}px`;
    root.style.top = `${spec.y ?? DEFAULT_OFFSET_Y}px`;

    const header = document.createElement('div');
    header.className = 'fc-slicer__header';
    const title = document.createElement('span');
    title.className = 'fc-slicer__title';
    title.textContent = spec.column;
    const actions = document.createElement('span');
    actions.className = 'fc-slicer__actions';
    const clearBtn = document.createElement('button');
    clearBtn.type = 'button';
    clearBtn.className = 'fc-slicer__btn fc-slicer__clear';
    clearBtn.textContent = strings.slicer.clear;
    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.className = 'fc-slicer__btn fc-slicer__close';
    closeBtn.setAttribute('aria-label', strings.slicer.close);
    closeBtn.textContent = '×';
    actions.append(clearBtn, closeBtn);
    header.append(title, actions);

    const body = document.createElement('div');
    body.className = 'fc-slicer__body';

    root.append(header, body);
    host.appendChild(root);

    const onClear = (): void => {
      withHistory(() => mutators.setSlicerSelected(store, spec.id, []));
      recomputeAndRender();
    };
    const onClose = (): void => {
      withHistory(() => mutators.removeSlicer(store, spec.id));
      recomputeAndRender();
    };
    clearBtn.addEventListener('click', onClear);
    closeBtn.addEventListener('click', onClose);

    return {
      root,
      body,
      title,
      closeBtn,
      dispose(): void {
        clearBtn.removeEventListener('click', onClear);
        closeBtn.removeEventListener('click', onClose);
        root.remove();
      },
    };
  };

  const renderChips = (entry: PanelEntry, spec: SlicerSpec): void => {
    entry.body.replaceChildren();
    const wb = getWb();
    const resolved = resolveSpec(wb, spec);
    if (!resolved) {
      const empty = document.createElement('div');
      empty.className = 'fc-slicer__empty';
      empty.textContent = strings.slicer.tablePlaceholder;
      entry.body.appendChild(empty);
      return;
    }
    const distinct = distinctValues(store.getState(), resolved.range, resolved.byCol);
    const selected = new Set(spec.selected);
    for (const value of distinct) {
      const chip = document.createElement('button');
      chip.type = 'button';
      chip.className = 'fc-slicer__chip';
      chip.dataset.fcValue = value;
      const isOn = selected.size === 0 || selected.has(value);
      // "Selected" visual state — when no chip is selected (empty array)
      //  every chip reads as on (include-all). When at least one is on,
      //  the unselected ones dim.
      chip.classList.toggle('fc-slicer__chip--on', isOn);
      chip.setAttribute('aria-pressed', String(isOn));
      chip.textContent = value === '' ? '(blank)' : value;

      chip.addEventListener('click', () => {
        const current = new Set(spec.selected);
        if (current.size === 0) {
          // First chip toggle: start with the clicked value as the only
          //  enabled item.
          current.add(value);
        } else if (current.has(value)) {
          current.delete(value);
        } else {
          current.add(value);
        }
        // If the new set covers every distinct value, collapse back to
        //  "all" so the next click can re-narrow without going through
        //  every chip.
        const next = current.size === distinct.length ? [] : Array.from(current).sort();
        withHistory(() => mutators.setSlicerSelected(store, spec.id, next));
        recomputeAndRender();
      });
      entry.body.appendChild(chip);
    }
  };

  const recomputeAndRender = (): void => {
    const wb = getWb();
    const slicers = store.getState().slicers.slicers;
    recomputeFilters(wb, store, slicers);
    renderAll();
  };

  const renderAll = (): void => {
    const slicers = store.getState().slicers.slicers;
    const liveIds = new Set(slicers.map((sp) => sp.id));
    // Drop panels for specs that vanished (history undo, external remove).
    for (const [id, entry] of panels) {
      if (!liveIds.has(id)) {
        entry.dispose();
        panels.delete(id);
      }
    }
    for (const spec of slicers) {
      let entry = panels.get(spec.id);
      if (!entry) {
        entry = buildPanel(spec);
        panels.set(spec.id, entry);
      } else {
        // Keep the title in sync — column may change via updateSlicer.
        entry.title.textContent = spec.column;
        entry.root.setAttribute('aria-label', `${strings.slicer.title}: ${spec.column}`);
        entry.root.style.left = `${spec.x ?? DEFAULT_OFFSET_X}px`;
        entry.root.style.top = `${spec.y ?? DEFAULT_OFFSET_Y}px`;
      }
      renderChips(entry, spec);
    }
  };

  const withHistory = (mutate: () => void): void => {
    if (history) recordSlicersChange(history, store, mutate);
    else mutate();
  };

  // Subscribe to store changes — chip repaint when slicers list changes,
  //  underlying cell map mutates (recalc), or the active sheet swaps.
  let lastSlicers = store.getState().slicers.slicers;
  let lastCells = store.getState().data.cells;
  let lastSheet = store.getState().data.sheetIndex;
  const unsub = store.subscribe(() => {
    const s = store.getState();
    const slicersChanged = s.slicers.slicers !== lastSlicers;
    const cellsChanged = s.data.cells !== lastCells;
    const sheetChanged = s.data.sheetIndex !== lastSheet;
    if (slicersChanged) lastSlicers = s.slicers.slicers;
    if (cellsChanged) lastCells = s.data.cells;
    if (sheetChanged) lastSheet = s.data.sheetIndex;
    if (slicersChanged || cellsChanged || sheetChanged) renderAll();
  });

  // Initial paint for any pre-existing slicers (e.g. restored from a
  //  serialized session before mount).
  renderAll();

  const handle: SlicerHandle = {
    addSlicer(input): SlicerSpec {
      const wb = getWb();
      const table = findTable(wb, input.tableName);
      if (!table) {
        throw new Error(`Slicer: table "${input.tableName}" not found`);
      }
      if (!table.columns.includes(input.column)) {
        throw new Error(`Slicer: column "${input.column}" not in table "${input.tableName}"`);
      }
      const spec: SlicerSpec = {
        id: nextId(),
        tableName: table.name,
        column: input.column,
        selected: input.selected ? [...input.selected] : [],
        x: input.x,
        y: input.y,
      };
      withHistory(() => mutators.addSlicer(store, spec));
      recomputeAndRender();
      return spec;
    },
    removeSlicer(id): void {
      withHistory(() => mutators.removeSlicer(store, id));
      recomputeAndRender();
    },
    refresh(): void {
      renderAll();
    },
    detach(): void {
      unsub();
      for (const entry of panels.values()) entry.dispose();
      panels.clear();
    },
    setStrings(next): void {
      strings = next;
      // Re-render to pick up new aria labels + clear button text.
      renderAll();
    },
  };
  return handle;
}
