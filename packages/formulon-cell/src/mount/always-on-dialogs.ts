// Host dialogs that are always mounted — not toggleable via
// `MountOptions.features`. Dialogs without a `setStrings` hook are rebuilt on
// locale change so their captured strings stay fresh.

import { activeFilterRange } from '../commands/filter-ranges.js';
import type { History } from '../commands/history.js';
import type { InteractionOperation } from '../commands/interaction-policy.js';
import { refreshPivotReportsAfterEdit } from '../commands/pivot-table.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { isMacHost } from '../extensions/ui-options.js';
import type { Strings } from '../i18n/strings.js';
import { attachCellStylesGallery } from '../interact/cell-styles-gallery.js';
import { attachCfRulesDialog } from '../interact/cf-rules-dialog.js';
import type { ConditionalDialogOpenOptions } from '../interact/conditional-dialog.js';
import { attachEvaluateFormulaDialog } from '../interact/evaluate-formula-dialog.js';
import { attachExternalLinksDialog } from '../interact/external-links-dialog.js';
import { attachFilterDropdown } from '../interact/filter-dropdown.js';
import { openPivotFieldFilter } from '../interact/pivot-filter-panel.js';
import { cellRect, gridOriginY, layoutForView, pointPx } from '../render/geometry.js';
import type { SpreadsheetStore } from '../store/store.js';
import { resolveTheme } from '../theme/resolve.js';

export interface AlwaysOnDialogsDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  history: History;
  getWb: () => WorkbookHandle;
  getStrings: () => Strings;
  getLocale: () => string;
  /** Called after the CF rules dialog changed the rule set. */
  onConditionalRulesChanged: () => void;
  /** Identifies engine mirrors of session CF rules for list de-duplication. */
  getSessionConditionalFormatIndexes?: (
    sheet: number,
    currentFormats: ReturnType<WorkbookHandle['getConditionalFormats']>,
  ) => ReadonlySet<number>;
  openConditionalDialog: (options: ConditionalDialogOpenOptions) => void;
  /** Whether the host policy allows `operation`; `null` names a dialog that edits nothing. */
  permits: (operation: InteractionOperation | null) => boolean;
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
      getSessionConditionalFormatIndexes: deps.getSessionConditionalFormatIndexes,
      onNewRule: () => deps.openConditionalDialog({ mode: 'new' }),
      onEditRule: (editIndex) => deps.openConditionalDialog({ mode: 'edit', editIndex }),
      store,
      history,
      strings: getStrings(),
    });
  const attachEvaluateFormula = () =>
    attachEvaluateFormulaDialog({ host, store, getWb, getLocale, strings: getStrings() });
  const attachFilter = () =>
    attachFilterDropdown({
      host,
      store,
      history,
      getWb,
      strings: getStrings(),
      locale: getLocale(),
    });

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

  // The Mac panel hangs from the header cell: left at the cell's right edge
  // less 1 pt, top at its bottom edge, wherever the button was clicked.
  const macFilterAnchor = (
    range: Range,
    col: number,
  ): { x: number; y: number; h: number } | null => {
    const s = store.getState();
    const layout = layoutForView(s);
    if (!isMacHost(host) || layout.rtl) return null;
    const cell = cellRect(layout, s.viewport, range.r0, col);
    const hostRect = host.getBoundingClientRect();
    return {
      x: hostRect.left + cell.x + cell.w - pointPx(layout, s.viewport),
      y: hostRect.top + cell.y + cell.h,
      h: 0,
    };
  };

  // Filter dropdown — opens when the pointer dispatches `fc:openfilter`
  // from a clicked column-filter chevron; no public toggle.
  const onOpenFilter = (e: Event): void => {
    if (!deps.permits('filter')) return;
    const detail = (e as CustomEvent<OpenFilterDetail>).detail;
    if (!detail) return;
    // The dropdown is positioned with `position: fixed`, so it expects
    // viewport-relative coords. The pointer payload's `x/y` are host-relative;
    // use `clientX/clientY` instead. `- 4` matches the chevron offset.
    filterDropdown.open(
      detail.range,
      detail.col,
      macFilterAnchor(detail.range, detail.col) ?? {
        x: detail.anchor.clientX,
        y: detail.anchor.clientY - 4,
        h: detail.anchor.h,
      },
    );
  };
  host.addEventListener('fc:openfilter', onOpenFilter);

  // A report's drop-down button dispatches `fc:openpivotfilter` with the
  // button cell and the button's bottom-left corner in viewport coordinates.
  const onOpenPivotFilter = (e: Event): void => {
    if (!deps.permits('pivot')) return;
    const detail = (e as CustomEvent<{ addr: Addr; anchor: { x: number; y: number } }>).detail;
    if (!detail) return;
    const wb = getWb();
    openPivotFieldFilter(
      {
        host,
        store,
        origin: 'contextMenu',
        commandId: 'pivotFilterValue',
        wb,
        history,
        strings: getStrings(),
        onChanged: () =>
          refreshPivotReportsAfterEdit(store, getWb(), { theme: resolveTheme(host) }),
      },
      detail.addr,
      detail.anchor,
    );
  };
  host.addEventListener('fc:openpivotfilter', onOpenPivotFilter);

  return {
    openExternalLinks: () => externalLinksDialog.open(),
    openCfRules: () => cfRulesDialog.open(),
    openCellStyles: () => cellStylesGallery.open(),
    openEvaluateFormula: () => evaluateFormulaDialog.open(),
    openFilterAtHeader(range, col) {
      const s = store.getState();
      const layout = layoutForView(s);
      const targetRange = range ?? activeFilterRange(s) ?? s.selection.range;
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
      filterDropdown.open(
        targetRange,
        targetCol,
        macFilterAnchor(targetRange, targetCol) ?? { x, y, h: layout.headerRowHeight },
      );
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
      host.removeEventListener('fc:openpivotfilter', onOpenPivotFilter);
      evaluateFormulaDialog.detach();
      externalLinksDialog.detach();
      cfRulesDialog.detach();
      cellStylesGallery.detach();
      filterDropdown.detach();
    },
  };
}
