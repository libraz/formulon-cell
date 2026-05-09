import type { History } from '../commands/history.js';
import { activateSheetView, deleteSheetView, saveSheetView } from '../commands/sheet-views.js';
import { setFreezePanes, setSheetZoom } from '../commands/structure.js';
import {
  setGridlinesVisible,
  setHeadingsVisible,
  setR1C1ReferenceStyle,
  setShowFormulas,
} from '../commands/view.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';

export interface ViewToolbarDeps {
  toolbar: HTMLElement;
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  history: History;
  strings?: Strings;
  onOpenObjects?: () => void;
  onChanged?: () => void;
}

export interface ViewToolbarHandle {
  refresh(): void;
  setStrings(next: Strings): void;
  bindWorkbook(next: WorkbookHandle): void;
  detach(): void;
}

const ZOOM_PRESETS = [75, 100, 125, 150, 200];
const CURRENT_VIEW_VALUE = '';

function makeButton(className: string): HTMLButtonElement {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = className;
  return btn;
}

export function attachViewToolbar(deps: ViewToolbarDeps): ViewToolbarHandle {
  const { toolbar, store, history } = deps;
  let strings = deps.strings ?? defaultStrings;
  let wb = deps.wb;

  toolbar.replaceChildren();

  const title = document.createElement('span');
  title.className = 'fc-viewbar__title';

  const toggles = document.createElement('div');
  toggles.className = 'fc-viewbar__group';

  const gridlines = makeButton('fc-viewbar__toggle');
  const headings = makeButton('fc-viewbar__toggle');
  const formulas = makeButton('fc-viewbar__toggle');
  const r1c1 = makeButton('fc-viewbar__toggle');
  toggles.append(gridlines, headings, formulas, r1c1);

  const freeze = document.createElement('div');
  freeze.className = 'fc-viewbar__group';
  const freezeNone = makeButton('fc-viewbar__button');
  const freezeTop = makeButton('fc-viewbar__button');
  const freezeFirst = makeButton('fc-viewbar__button');
  const freezePanes = makeButton('fc-viewbar__button');
  freeze.append(freezeNone, freezeTop, freezeFirst, freezePanes);

  const zoom = document.createElement('div');
  zoom.className = 'fc-viewbar__group fc-viewbar__group--zoom';
  const zoomLabel = document.createElement('span');
  zoomLabel.className = 'fc-viewbar__label';
  const zoomSelect = document.createElement('select');
  zoomSelect.className = 'fc-viewbar__select';
  for (const pct of ZOOM_PRESETS) {
    const opt = document.createElement('option');
    opt.value = String(pct);
    opt.textContent = `${pct}%`;
    zoomSelect.appendChild(opt);
  }
  const zoomFit = makeButton('fc-viewbar__button');
  zoom.append(zoomLabel, zoomSelect, zoomFit);

  const sheetViews = document.createElement('div');
  sheetViews.className = 'fc-viewbar__group fc-viewbar__group--sheetviews';
  const sheetViewsLabel = document.createElement('span');
  sheetViewsLabel.className = 'fc-viewbar__label';
  const sheetViewsSelect = document.createElement('select');
  sheetViewsSelect.className = 'fc-viewbar__select';
  const saveView = makeButton('fc-viewbar__button');
  const deleteView = makeButton('fc-viewbar__button');
  sheetViews.append(sheetViewsLabel, sheetViewsSelect, saveView, deleteView);

  const objects = document.createElement('div');
  objects.className = 'fc-viewbar__group';
  const objectsBtn = makeButton('fc-viewbar__button');
  if (deps.onOpenObjects) objects.appendChild(objectsBtn);

  toolbar.append(title, toggles, freeze, zoom, sheetViews, objects);

  const applyChanged = (): void => {
    deps.onChanged?.();
    refresh();
  };

  gridlines.addEventListener('click', () => {
    setGridlinesVisible(store, !store.getState().ui.showGridLines);
    applyChanged();
  });
  headings.addEventListener('click', () => {
    setHeadingsVisible(store, !store.getState().ui.showHeaders);
    applyChanged();
  });
  formulas.addEventListener('click', () => {
    setShowFormulas(store, !store.getState().ui.showFormulas);
    applyChanged();
  });
  r1c1.addEventListener('click', () => {
    setR1C1ReferenceStyle(store, !store.getState().ui.r1c1);
    applyChanged();
  });
  freezeNone.addEventListener('click', () => {
    setFreezePanes(store, history, 0, 0, wb);
    applyChanged();
  });
  freezeTop.addEventListener('click', () => {
    setFreezePanes(store, history, 1, 0, wb);
    applyChanged();
  });
  freezeFirst.addEventListener('click', () => {
    setFreezePanes(store, history, 0, 1, wb);
    applyChanged();
  });
  freezePanes.addEventListener('click', () => {
    const active = store.getState().selection.active;
    setFreezePanes(store, history, active.row, active.col, wb);
    applyChanged();
  });
  zoomSelect.addEventListener('change', () => {
    setSheetZoom(store, Number(zoomSelect.value) / 100, wb);
    applyChanged();
  });
  zoomFit.addEventListener('click', () => {
    setSheetZoom(store, 1, wb);
    applyChanged();
  });
  sheetViewsSelect.addEventListener('change', () => {
    if (sheetViewsSelect.value === CURRENT_VIEW_VALUE) {
      store.setState((s) => ({ ...s, sheetViews: { ...s.sheetViews, activeViewId: null } }));
      applyChanged();
      return;
    }
    const result = activateSheetView(store, sheetViewsSelect.value);
    if (result.ok) applyChanged();
    else refresh();
  });
  saveView.addEventListener('click', () => {
    const count = store.getState().sheetViews.views.length + 1;
    const id = `view-${Date.now().toString(36)}-${count}`;
    saveSheetView(store, id, `${strings.viewToolbar.views} ${count}`);
    applyChanged();
  });
  deleteView.addEventListener('click', () => {
    const id = store.getState().sheetViews.activeViewId;
    if (!id) return;
    deleteSheetView(store, id);
    applyChanged();
  });
  objectsBtn.addEventListener('click', () => deps.onOpenObjects?.());

  function refreshLabels(): void {
    const t = strings.viewToolbar;
    title.textContent = t.title;
    gridlines.textContent = t.gridlines;
    headings.textContent = t.headings;
    formulas.textContent = t.formulas;
    r1c1.textContent = t.r1c1;
    freezeNone.textContent = t.freezeNone;
    freezeTop.textContent = t.freezeTopRow;
    freezeFirst.textContent = t.freezeFirstColumn;
    freezePanes.textContent = t.freezePanes;
    zoomLabel.textContent = t.zoom;
    zoomFit.textContent = t.zoom100;
    sheetViewsLabel.textContent = t.views;
    saveView.textContent = t.saveView;
    deleteView.textContent = t.deleteView;
    objectsBtn.textContent = t.objects;
    for (const btn of [gridlines, headings, formulas, r1c1, freezeNone, freezeTop, freezeFirst]) {
      btn.setAttribute('aria-label', btn.textContent ?? '');
    }
    freezePanes.setAttribute('aria-label', t.freezePanes);
    zoomSelect.setAttribute('aria-label', t.zoom);
    zoomFit.setAttribute('aria-label', t.zoom100);
    sheetViewsSelect.setAttribute('aria-label', t.views);
    saveView.setAttribute('aria-label', t.saveView);
    deleteView.setAttribute('aria-label', t.deleteView);
    objectsBtn.setAttribute('aria-label', t.objects);
  }

  function refresh(): void {
    const s = store.getState();
    refreshLabels();
    gridlines.setAttribute('aria-pressed', String(s.ui.showGridLines));
    headings.setAttribute('aria-pressed', String(s.ui.showHeaders));
    formulas.setAttribute('aria-pressed', String(s.ui.showFormulas));
    r1c1.setAttribute('aria-pressed', String(s.ui.r1c1));
    freezeNone.setAttribute(
      'aria-pressed',
      String(s.layout.freezeRows === 0 && s.layout.freezeCols === 0),
    );
    freezeTop.setAttribute('aria-pressed', String(s.layout.freezeRows === 1));
    freezeFirst.setAttribute('aria-pressed', String(s.layout.freezeCols === 1));
    freezePanes.setAttribute(
      'aria-pressed',
      String(s.layout.freezeRows > 0 || s.layout.freezeCols > 0),
    );
    const pct = Math.round(s.viewport.zoom * 100);
    const known = ZOOM_PRESETS.includes(pct);
    zoomSelect.value = known ? String(pct) : '100';
    zoomSelect.title = `${pct}%`;

    sheetViewsSelect.replaceChildren();
    const current = document.createElement('option');
    current.value = CURRENT_VIEW_VALUE;
    current.textContent = strings.viewToolbar.currentView;
    sheetViewsSelect.appendChild(current);
    for (const view of s.sheetViews.views.filter((v) => v.sheet === s.data.sheetIndex)) {
      const opt = document.createElement('option');
      opt.value = view.id;
      opt.textContent = view.name;
      sheetViewsSelect.appendChild(opt);
    }
    sheetViewsSelect.value = s.sheetViews.activeViewId ?? CURRENT_VIEW_VALUE;
    deleteView.disabled = !s.sheetViews.activeViewId;
  }

  const unsub = store.subscribe(refresh);
  refresh();

  return {
    refresh,
    setStrings(next) {
      strings = next;
      refresh();
    },
    bindWorkbook(next) {
      wb = next;
    },
    detach() {
      unsub();
      toolbar.replaceChildren();
    },
  };
}
