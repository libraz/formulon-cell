import {
  applyPivotGrandTotals,
  applyPivotReportLayout,
  applyPivotSubtotalPosition,
  clearPivotFields,
  clearPivotFilters,
  type PivotRef,
  pivotLayoutState,
} from '../../commands/pivot-layout.js';
import { refreshPivotReportsAfterEdit } from '../../commands/pivot-table.js';
import { findPivotTableAtCell } from '../../engine/passthrough-sync.js';
import { PivotReportLayout } from '../../engine/types.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import { commitPivotChange } from '../../interact/pivot-field-settings-dialog-apply.js';
import { showPivotSpillCallout } from '../../interact/pivot-spill-callout.js';
import type { SpreadsheetInstance } from '../../mount.js';
import { mutators } from '../../store/store.js';
import { resolveTheme } from '../../theme/resolve.js';
import { projectDisabledState } from '../../toolbar/menu-a11y.js';

export interface PivotDropdownDefaults {
  /** Runs one `data-pivot-layout` action of the Design / Analyze menus. */
  applyPivotLayoutAction: (action: string) => void;
  /** Projects the active report's state onto an open Design / Analyze menu. */
  updatePivotLayoutMenu: (menu: HTMLElement) => void;
}

/** Edits that rewrite the report: each is one undo step and re-lays out once. */
const REPORT_EDITS: Readonly<Record<string, (wb: WorkbookHandle, ref: PivotRef) => boolean>> = {
  'totals-off': (wb, ref) => applyPivotGrandTotals(wb, ref, { rows: false, cols: false }),
  'totals-on': (wb, ref) => applyPivotGrandTotals(wb, ref, { rows: true, cols: true }),
  'totals-rows': (wb, ref) => applyPivotGrandTotals(wb, ref, { rows: true, cols: false }),
  'totals-cols': (wb, ref) => applyPivotGrandTotals(wb, ref, { rows: false, cols: true }),
  'layout-compact': (wb, ref) => applyPivotReportLayout(wb, ref, PivotReportLayout.Compact),
  'layout-outline': (wb, ref) => applyPivotReportLayout(wb, ref, PivotReportLayout.Outline),
  'layout-tabular': (wb, ref) => applyPivotReportLayout(wb, ref, PivotReportLayout.Tabular),
  'subtotals-bottom': (wb, ref) => applyPivotSubtotalPosition(wb, ref, false),
  'subtotals-top': (wb, ref) => applyPivotSubtotalPosition(wb, ref, true),
  'clear-all': clearPivotFields,
  'clear-filters': clearPivotFilters,
};

/** The Design / Analyze command each report edit runs under. */
const REPORT_EDIT_COMMAND: Readonly<Record<string, string>> = {
  'totals-off': 'pivotGrandTotals',
  'totals-on': 'pivotGrandTotals',
  'totals-rows': 'pivotGrandTotals',
  'totals-cols': 'pivotGrandTotals',
  'layout-compact': 'pivotReportLayout',
  'layout-outline': 'pivotReportLayout',
  'layout-tabular': 'pivotReportLayout',
  'subtotals-bottom': 'pivotSubtotals',
  'subtotals-top': 'pivotSubtotals',
  'clear-all': 'pivotClear',
  'clear-filters': 'pivotClear',
};

/** Re-reads the reports after a pivot edit and points out a report that no longer fits. */
export function refreshAfterPivotEdit(instance: SpreadsheetInstance): void {
  const { spilled } = refreshPivotReportsAfterEdit(instance.store, instance.workbook, {
    theme: resolveTheme(instance.host),
  });
  const spill = spilled[0];
  if (spill) showPivotSpillCallout(instance.host, instance.store, instance.i18n.strings, spill);
}

export function createPivotDropdownDefaults(instance: SpreadsheetInstance): PivotDropdownDefaults {
  const activePivot = (): PivotRef | null => {
    const pivot = findPivotTableAtCell(
      instance.workbook as unknown as Parameters<typeof findPivotTableAtCell>[0],
      instance.store.getState().selection.active,
    );
    return pivot ? { sheetIndex: pivot.sheetIndex, pivotIndex: pivot.pivotIndex } : null;
  };

  const onChanged = (): void => refreshAfterPivotEdit(instance);

  return {
    applyPivotLayoutAction(action) {
      const ref = activePivot();
      if (!ref) return;
      const wb = instance.workbook;
      if (action === 'select-table') {
        const range = wb
          .pivotBounds(ref.sheetIndex)
          .find((b) => b.pivotIndex === ref.pivotIndex)?.range;
        if (range) mutators.setRange(instance.store, range);
        return;
      }
      const edit = REPORT_EDITS[action];
      if (!edit) return;
      commitPivotChange(
        {
          wb,
          history: instance.history,
          store: instance.store,
          origin: 'ribbon',
          commandId: REPORT_EDIT_COMMAND[action] ?? 'pivotReportLayout',
          onChanged,
        },
        ref.sheetIndex,
        ref.pivotIndex,
        () => edit(wb, ref),
      );
    },
    updatePivotLayoutMenu(menu) {
      const ref = activePivot();
      const state = ref ? pivotLayoutState(instance.workbook, ref) : null;
      const t = instance.i18n.strings.ribbonMenu;
      const project = (action: string, disabled: boolean): void => {
        const button = menu.querySelector<HTMLButtonElement>(`[data-pivot-layout="${action}"]`);
        if (button) {
          projectDisabledState(button, disabled, disabled ? t.pivotUnavailable : null, {
            datasetKey: 'disabledReason',
          });
        }
      };
      project('clear-all', state?.hasFields !== true);
      project('clear-filters', state?.hasFilter !== true);
    },
  };
}
