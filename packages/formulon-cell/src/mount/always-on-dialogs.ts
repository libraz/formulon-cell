// Host dialogs that are always mounted — not toggleable via
// `MountOptions.features`. Dialogs without a `setStrings` hook are rebuilt on
// locale change so their captured strings stay fresh.

import type { History } from '../commands/history.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import { attachCellStylesGallery } from '../interact/cell-styles-gallery.js';
import { attachCfRulesDialog } from '../interact/cf-rules-dialog.js';
import type { ConditionalDialogOpenOptions } from '../interact/conditional-dialog.js';
import { attachEvaluateFormulaDialog } from '../interact/evaluate-formula-dialog.js';
import { attachExternalLinksDialog } from '../interact/external-links-dialog.js';
import { attachFilterDropdown } from '../interact/filter-dropdown.js';
import { cellRect, gridOriginY, layoutForView } from '../render/geometry.js';
import type { SpreadsheetStore } from '../store/store.js';

export interface AlwaysOnDialogsDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  history: History;
  getWb: () => WorkbookHandle;
  getStrings: () => Strings;
  getLocale: () => string;
  /** Called after the CF rules dialog changed the rule set. */
  onConditionalRulesChanged: () => void;
  openConditionalDialog: (options: ConditionalDialogOpenOptions) => void;
  /** True while an interaction policy blocks the filter dropdown. */
  isRestricted: () => boolean;
}

interface OpenFilterDetail {
  range: Range;
  col: number;
  anchor: { x: number; y: number; h: number; clientX: number; clientY: number };
}

export interface AlwaysOnDialogs {
  openExternalLinks: () => void;
  openCfRules: () => void;
  openCellStyles: () => void;
  openEvaluateFormula: () => void;
  /** Open the filter dropdown under the header chevron of `col` (default: active column). */
  openFilterAtHeader: (range?: Range, col?: number) => void;
  setStrings: (next: Strings) => void;
  detach: () => void;
}

export function attachAlwaysOnDialogs(deps: AlwaysOnDialogsDeps): AlwaysOnDialogs {
  const { host, store, history, getWb, getStrings, getLocale } = deps;
  const attachExternalLinks = () =>
    attachExternalLinksDialog({ host, getWb, strings: getStrings() });
  const attachCfRules = () =>
    attachCfRulesDialog({
      host,
      getWb,
      getActiveSheet: () => store.getState().data.sheetIndex,
      getSelectionRange: () => store.getState().selection.range,
      onChanged: deps.onConditionalRulesChanged,
      onNewRule: () => deps.openConditionalDialog({ mode: 'new' }),
      onEditRule: (editIndex) => deps.openConditionalDialog({ mode: 'edit', editIndex }),
      store,
      history,
      strings: getStrings(),
    });
  const attachEvaluateFormula = () =>
    attachEvaluateFormulaDialog({ host, store, getWb, strings: getStrings() });
  const attachFilter = () =>
    attachFilterDropdown({ host, store, history, strings: getStrings(), locale: getLocale() });

  let externalLinksDialog = attachExternalLinks();
  let cfRulesDialog = attachCfRules();
  let evaluateFormulaDialog = attachEvaluateFormula();
  const cellStylesGallery = attachCellStylesGallery({
    host,
    store,
    history,
    getWb,
    strings: getStrings(),
  });
  let filterDropdown = attachFilter();

  // Filter dropdown — opens when the pointer dispatches `fc:openfilter`
  // from a clicked column-filter chevron; no public toggle.
  const onOpenFilter = (e: Event): void => {
    if (deps.isRestricted()) return;
    const detail = (e as CustomEvent<OpenFilterDetail>).detail;
    if (!detail) return;
    // The dropdown is positioned with `position: fixed`, so it expects
    // viewport-relative coords. The pointer payload's `x/y` are host-relative;
    // use `clientX/clientY` instead. `- 4` matches the chevron offset.
    filterDropdown.open(detail.range, detail.col, {
      x: detail.anchor.clientX,
      y: detail.anchor.clientY - 4,
      h: detail.anchor.h,
    });
  };
  host.addEventListener('fc:openfilter', onOpenFilter);

  return {
    openExternalLinks: () => externalLinksDialog.open(),
    openCfRules: () => cfRulesDialog.open(),
    openCellStyles: () => cellStylesGallery.open(),
    openEvaluateFormula: () => evaluateFormulaDialog.open(),
    openFilterAtHeader(range, col) {
      const s = store.getState();
      const layout = layoutForView(s);
      const targetRange = range ?? s.ui.filterRange ?? s.selection.range;
      const targetCol = Math.min(
        Math.max(col ?? s.selection.active.col, targetRange.c0),
        targetRange.c1,
      );
      const hostRect = host.getBoundingClientRect();
      // The chevron hangs off the header cell's trailing edge, which the
      // mirror puts on the left of the cell for a right-to-left sheet.
      const cell = cellRect(layout, s.viewport, targetRange.r0, targetCol);
      const x = hostRect.left + (layout.rtl ? cell.x + 4 : cell.x + cell.w - 4);
      const y = hostRect.top + gridOriginY(layout) - 4;
      filterDropdown.open(targetRange, targetCol, { x, y, h: layout.headerRowHeight });
    },
    setStrings(next) {
      cellStylesGallery.setStrings(next);
      externalLinksDialog.detach();
      externalLinksDialog = attachExternalLinks();
      cfRulesDialog.detach();
      cfRulesDialog = attachCfRules();
      evaluateFormulaDialog.detach();
      evaluateFormulaDialog = attachEvaluateFormula();
      filterDropdown.detach();
      filterDropdown = attachFilter();
    },
    detach() {
      host.removeEventListener('fc:openfilter', onOpenFilter);
      evaluateFormulaDialog.detach();
      externalLinksDialog.detach();
      cfRulesDialog.detach();
      cellStylesGallery.detach();
      filterDropdown.detach();
    },
  };
}
