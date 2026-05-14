import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { type SpreadsheetEmitter, selectionEquals } from '../events.js';
import type { SpreadsheetStore } from '../store/store.js';
import { mutators } from '../store/store.js';
import { formatSelectionRef, lookupDefinedName, parseCellRef, parseRangeRef } from './ref-utils.js';
import type { SheetTabsController } from './sheet-tabs-controller.js';

interface AttachChromeSyncInput {
  a11y: HTMLElement;
  fxInput: HTMLTextAreaElement;
  getFormulaEditing: () => boolean;
  getSheetTabs: () => SheetTabsController | null;
  getWb: () => WorkbookHandle;
  host: HTMLElement;
  invalidate: () => void;
  store: SpreadsheetStore;
  tag: HTMLInputElement;
  emitter: SpreadsheetEmitter;
}

export interface ChromeSyncController {
  detach(): void;
  updateChrome(): void;
}

export function attachChromeSync(input: AttachChromeSyncInput): ChromeSyncController {
  const {
    a11y,
    emitter,
    fxInput,
    getFormulaEditing,
    getSheetTabs,
    getWb,
    host,
    invalidate,
    store,
    tag,
  } = input;

  const updateChrome = (): void => {
    const wb = getWb();
    const s = store.getState();
    const a = s.selection.active;
    const ref = formatSelectionRef(s.selection.range, a, s.ui.r1c1 === true);
    if (document.activeElement !== tag) tag.value = ref;
    const cell = s.data.cells.get(`${a.sheet}:${a.row}:${a.col}`);
    const formula = cell?.formula ?? '';
    let display = '';
    if (formula) display = formula;
    else if (cell) {
      const v = cell.value;
      switch (v.kind) {
        case 'number':
          display = String(v.value);
          break;
        case 'bool':
          display = v.value ? 'TRUE' : 'FALSE';
          break;
        case 'text':
          display = v.value;
          break;
        case 'error':
          display = v.text;
          break;
        default: {
          const lambda = wb.getLambdaText(a);
          display = lambda ? `=${lambda}` : '';
          break;
        }
      }
    }
    if (!getFormulaEditing()) fxInput.value = display;
    a11y.textContent = `${ref} ${display}`;
  };

  const onTagFocus = (): void => tag.select();
  const onTagKey = (e: KeyboardEvent): void => {
    if (e.key === 'Enter') {
      e.preventDefault();
      e.stopPropagation();
      const sheetIdx = store.getState().data.sheetIndex;
      const range = parseRangeRef(tag.value);
      if (range) {
        store.setState((s) => ({
          ...s,
          selection: {
            active: { sheet: sheetIdx, row: range.r0, col: range.c0 },
            anchor: { sheet: sheetIdx, row: range.r0, col: range.c0 },
            range: { sheet: sheetIdx, ...range },
          },
        }));
        host.focus();
        return;
      }
      const parsed = parseCellRef(tag.value);
      if (parsed) {
        mutators.setActive(store, {
          sheet: sheetIdx,
          row: parsed.row,
          col: parsed.col,
        });
        host.focus();
        return;
      }
      const dn = lookupDefinedName(getWb(), tag.value.trim());
      if (dn) {
        const sub = parseRangeRef(dn) ?? parseCellRef(dn);
        if (sub) {
          if ('r0' in sub) {
            store.setState((s) => ({
              ...s,
              selection: {
                active: { sheet: sheetIdx, row: sub.r0, col: sub.c0 },
                anchor: { sheet: sheetIdx, row: sub.r0, col: sub.c0 },
                range: { sheet: sheetIdx, ...sub },
              },
            }));
          } else {
            mutators.setActive(store, { sheet: sheetIdx, row: sub.row, col: sub.col });
          }
          host.focus();
          return;
        }
      }
    } else if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      host.focus();
      updateChrome();
    }
  };
  const onTagBlur = (): void => updateChrome();

  tag.addEventListener('focus', onTagFocus);
  tag.addEventListener('keydown', onTagKey);
  tag.addEventListener('blur', onTagBlur);

  let lastSheetIdx = store.getState().data.sheetIndex;
  let lastHiddenSheets = store.getState().layout.hiddenSheets;
  let lastSelection = store.getState().selection;
  const unsub = store.subscribe(() => {
    const s = store.getState();
    const sheetChanged = s.data.sheetIndex !== lastSheetIdx;
    if (sheetChanged) {
      getWb().clearViewportHint();
      lastSheetIdx = s.data.sheetIndex;
    }
    if (sheetChanged || s.layout.hiddenSheets !== lastHiddenSheets) {
      lastHiddenSheets = s.layout.hiddenSheets;
      getSheetTabs()?.update();
    }
    if (!selectionEquals(lastSelection, s.selection)) {
      lastSelection = s.selection;
      emitter.emit('selectionChange', {
        active: s.selection.active,
        anchor: s.selection.anchor,
        range: s.selection.range,
      });
    }
    invalidate();
    updateChrome();
  });

  updateChrome();

  return {
    detach(): void {
      tag.removeEventListener('focus', onTagFocus);
      tag.removeEventListener('keydown', onTagKey);
      tag.removeEventListener('blur', onTagBlur);
      unsub();
    },
    updateChrome,
  };
}
