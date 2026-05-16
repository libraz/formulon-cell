// Ribbon control dispatch extracted from main.ts. Owns the read/write side of
// the ribbon: reading the current value of a control from the active cell or
// page setup, and applying a control change (font, fill, number format,
// merge, page-setup presets, sheet views) back into the workbook. Also hosts
// the small `createRibbonIcon` SVG helper because both this module and the
// select/color factory need it.

import {
  activateSheetView,
  applyMerge,
  applyUnmerge,
  fluentIconPaths,
  getPageSetup,
  type MarginPreset,
  marginPresetOf,
  mutators,
  type NumberFormatAction,
  type NumFmt,
  type PageOrientation,
  type PageScaleMenuText,
  type PaperSize,
  projectActiveState,
  recordFormatChange,
  recordPageSetupChange,
  type SpreadsheetInstance,
  setAlign,
  setFillColor,
  setFont,
  setFontColor,
  setMarginPreset,
  setNumFmt,
  setPageOrientation,
  setPaperSize,
  type ToolbarText,
  numberFormatForAction as toolbarNumberFormatForAction,
} from '@libraz/formulon-cell';

import { showPrompt } from '../dialogs.js';

export type RibbonFormatMutator = (
  state: ReturnType<SpreadsheetInstance['store']['getState']>,
  store: SpreadsheetInstance['store'],
) => void;

export interface ControlDispatchCtx {
  getInst: () => SpreadsheetInstance | null;
  ribbonLang: 'ja' | 'en';
  ribbonText: ToolbarText;
  pageScaleText: PageScaleMenuText;
  sheetEl: HTMLElement;
  focusSheet: () => void;
  refreshWorkbookCells: () => void;
  projectFormatToolbar: () => void;
}

export interface ControlDispatchApi {
  createRibbonIcon: (name: string) => SVGSVGElement | null;
  currentRibbonControlValue: (id: string) => string;
  applyRibbonFormat: (fn: RibbonFormatMutator) => void;
  applyRibbonControl: (id: string, value: string) => void;
  applyMergeControl: (value: string) => void;
}

export const createControlDispatch = (ctx: ControlDispatchCtx): ControlDispatchApi => {
  const {
    getInst,
    ribbonLang,
    ribbonText,
    pageScaleText,
    sheetEl,
    focusSheet,
    refreshWorkbookCells,
    projectFormatToolbar,
  } = ctx;

  const createRibbonIcon = (name: string): SVGSVGElement | null => {
    const paths = fluentIconPaths(name);
    if (!paths) return null;
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.classList.add('demo__rb-icon');
    svg.setAttribute('viewBox', '0 0 24 24');
    svg.setAttribute('fill', 'currentColor');
    svg.setAttribute('focusable', 'false');
    svg.setAttribute('aria-hidden', 'true');
    for (const d of paths) {
      const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      path.setAttribute('d', d);
      svg.appendChild(path);
    }
    return svg;
  };

  const activeCellFormat = () => {
    const inst = getInst();
    if (!inst) return null;
    const s = inst.store.getState();
    const a = s.selection.active;
    return s.format.formats.get(`${a.sheet}:${a.row}:${a.col}`) ?? null;
  };

  const currentRibbonControlValue = (id: string): string => {
    const inst = getInst();
    const f = activeCellFormat();
    const pageSetup =
      inst &&
      (id === 'marginsPreset' ||
        id === 'orientationPreset' ||
        id === 'paperSizePreset' ||
        id === 'scaleWidth' ||
        id === 'scaleHeight' ||
        id === 'scalePercent')
        ? getPageSetup(inst.store.getState(), inst.store.getState().data.sheetIndex)
        : null;
    if (id === 'fontFamily')
      return f?.fontFamily ?? (ribbonLang === 'ja' ? '游ゴシック Regular' : 'Aptos');
    if (id === 'fontSize') return String(f?.fontSize ?? (ribbonLang === 'ja' ? 12 : 11));
    if (id === 'fontColor') return f?.color ?? '#201f1e';
    if (id === 'fillColor') return f?.fill ?? '#ffffff';
    if (id === 'numberFormat') return inst ? projectActiveState(inst).numberFormat : 'general';
    if (id === 'merge') {
      if (!inst) return 'mergeCenter';
      const state = inst.store.getState();
      const r = state.selection.range;
      const anchor = state.merges.byAnchor.get(`${r.sheet}:${r.r0}:${r.c0}`);
      return anchor &&
        anchor.r0 === r.r0 &&
        anchor.c0 === r.c0 &&
        anchor.r1 === r.r1 &&
        anchor.c1 === r.c1
        ? 'unmergeCells'
        : 'mergeCenter';
    }
    if (id === 'marginsPreset')
      return pageSetup ? (marginPresetOf(pageSetup.margins) ?? 'custom') : 'normal';
    if (id === 'orientationPreset') return pageSetup?.orientation ?? 'portrait';
    if (id === 'paperSizePreset') return pageSetup?.paperSize ?? 'A4';
    if (id === 'scaleWidth') return String(pageSetup?.fitWidth ?? 0);
    if (id === 'scaleHeight') return String(pageSetup?.fitHeight ?? 0);
    if (id === 'scalePercent') return String(Math.round((pageSetup?.scale ?? 1) * 100));
    if (id === 'sheetViewSelect')
      return inst?.store.getState().sheetViews.activeViewId ?? 'current';
    return '';
  };

  const numberFormatForAction = (action: string): NumFmt | null =>
    toolbarNumberFormatForAction(action as NumberFormatAction, ribbonLang);

  const applyRibbonFormat = (fn: RibbonFormatMutator): void => {
    const i = getInst();
    if (!i) return;
    recordFormatChange(i.history, i.store, () => {
      fn(i.store.getState(), i.store);
    });
    sheetEl.focus();
  };

  const applyCustomPageScaleControl = async (
    id: 'scaleWidth' | 'scaleHeight' | 'scalePercent',
  ): Promise<void> => {
    const i = getInst();
    if (!i) return;
    const sheet = i.store.getState().data.sheetIndex;
    const setup = getPageSetup(i.store.getState(), sheet);
    const isScale = id === 'scalePercent';
    const initial = isScale
      ? String(Math.round((setup.scale ?? 1) * 100))
      : String(id === 'scaleWidth' ? (setup.fitWidth ?? 1) : (setup.fitHeight ?? 1));
    const value = await showPrompt({
      title: isScale
        ? ribbonText.scale
        : id === 'scaleWidth'
          ? pageScaleText.width
          : pageScaleText.height,
      label: isScale ? pageScaleText.customScalePrompt : pageScaleText.customPagesPrompt,
      initial,
      okLabel: 'OK',
      cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
      validate: (raw) => {
        const n = Number.parseInt(raw.trim(), 10);
        if (isScale)
          return Number.isInteger(n) && n >= 10 && n <= 400 ? null : pageScaleText.invalidScale;
        return Number.isInteger(n) && n >= 1 && n <= 99 ? null : pageScaleText.invalidPages;
      },
    });
    if (value === null) {
      focusSheet();
      return;
    }
    const n = Number.parseInt(value.trim(), 10);
    recordPageSetupChange(i.history, i.store, () => {
      if (isScale)
        mutators.setPageSetup(i.store, sheet, { scale: n / 100, fitWidth: 0, fitHeight: 0 });
      else
        mutators.setPageSetup(
          i.store,
          sheet,
          id === 'scaleWidth' ? { fitWidth: n } : { fitHeight: n },
        );
    });
    projectFormatToolbar();
    focusSheet();
  };

  const applyMergeControl = (value: string): void => {
    const i = getInst();
    if (!i) return;
    const range = i.store.getState().selection.range;
    if (value === 'unmergeCells') {
      applyUnmerge(i.store, i.workbook, i.history, range);
    } else if (value === 'mergeAcross') {
      i.history.begin();
      try {
        for (let row = range.r0; row <= range.r1; row += 1) {
          applyMerge(i.store, i.workbook, i.history, {
            sheet: range.sheet,
            r0: row,
            c0: range.c0,
            r1: row,
            c1: range.c1,
          });
        }
      } finally {
        i.history.end();
      }
    } else {
      const merged = applyMerge(i.store, i.workbook, i.history, range);
      if (merged && value === 'mergeCenter') {
        applyRibbonFormat((state, store) => setAlign(state, store, 'center'));
      }
    }
    refreshWorkbookCells();
    projectFormatToolbar();
    sheetEl.focus();
  };

  const applyRibbonControl = (id: string, value: string): void => {
    if (id === 'fontFamily') {
      applyRibbonFormat((state, store) => setFont(state, store, { fontFamily: value }));
    } else if (id === 'fontSize') {
      applyRibbonFormat((state, store) => setFont(state, store, { fontSize: Number(value) }));
    } else if (id === 'fontColor') {
      applyRibbonFormat((state, store) => setFontColor(state, store, value));
    } else if (id === 'fillColor') {
      applyRibbonFormat((state, store) => setFillColor(state, store, value));
    } else if (id === 'numberFormat') {
      if (value === 'more') {
        getInst()?.openFormatDialog();
        return;
      }
      const fmt = numberFormatForAction(value);
      if (fmt) applyRibbonFormat((state, store) => setNumFmt(state, store, fmt));
    } else if (id === 'merge') {
      applyMergeControl(value);
    } else if (id === 'marginsPreset') {
      const i = getInst();
      if (!i) return;
      if (value === 'custom') {
        i.openPageSetup();
        return;
      }
      const sheet = i.store.getState().data.sheetIndex;
      recordPageSetupChange(i.history, i.store, () =>
        setMarginPreset(i.store, sheet, value as MarginPreset),
      );
      projectFormatToolbar();
      sheetEl.focus();
    } else if (id === 'orientationPreset') {
      const i = getInst();
      if (!i) return;
      const sheet = i.store.getState().data.sheetIndex;
      recordPageSetupChange(i.history, i.store, () =>
        setPageOrientation(i.store, sheet, value as PageOrientation),
      );
      projectFormatToolbar();
      sheetEl.focus();
    } else if (id === 'paperSizePreset') {
      const i = getInst();
      if (!i) return;
      const sheet = i.store.getState().data.sheetIndex;
      recordPageSetupChange(i.history, i.store, () =>
        setPaperSize(i.store, sheet, value as PaperSize),
      );
      projectFormatToolbar();
      sheetEl.focus();
    } else if (id === 'scaleWidth' || id === 'scaleHeight') {
      const i = getInst();
      if (!i) return;
      if (value === 'custom') {
        void applyCustomPageScaleControl(id);
        return;
      }
      const sheet = i.store.getState().data.sheetIndex;
      const n = Math.max(0, Math.min(99, Number.parseInt(value, 10) || 0));
      recordPageSetupChange(i.history, i.store, () => {
        mutators.setPageSetup(
          i.store,
          sheet,
          id === 'scaleWidth' ? { fitWidth: n } : { fitHeight: n },
        );
      });
      projectFormatToolbar();
      sheetEl.focus();
    } else if (id === 'scalePercent') {
      const i = getInst();
      if (!i) return;
      if (value === 'custom') {
        void applyCustomPageScaleControl('scalePercent');
        return;
      }
      const sheet = i.store.getState().data.sheetIndex;
      const pct = Math.max(10, Math.min(400, Number.parseInt(value, 10) || 100));
      recordPageSetupChange(i.history, i.store, () => {
        mutators.setPageSetup(i.store, sheet, { scale: pct / 100, fitWidth: 0, fitHeight: 0 });
      });
      projectFormatToolbar();
      sheetEl.focus();
    } else if (id === 'sheetViewSelect') {
      const i = getInst();
      if (!i) return;
      if (value === 'current') {
        i.store.setState((s) => ({ ...s, sheetViews: { ...s.sheetViews, activeViewId: null } }));
        projectFormatToolbar();
        focusSheet();
        return;
      }
      const result = activateSheetView(i.store, value);
      if (result.ok) {
        refreshWorkbookCells();
        projectFormatToolbar();
        focusSheet();
      }
    }
  };

  return {
    createRibbonIcon,
    currentRibbonControlValue,
    applyRibbonFormat,
    applyRibbonControl,
    applyMergeControl,
  };
};
