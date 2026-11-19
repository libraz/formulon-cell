import { pivotReportRange } from '../../commands/pivot-actions.js';
import {
  applyPivotGrandTotals,
  applyPivotReportLayout,
  applyPivotSubtotalPosition,
  clearPivotFields,
  clearPivotFilters,
  type PivotRef,
  pivotLayoutState,
} from '../../commands/pivot-layout.js';
import { refreshPivotReportsAfterEdit } from '../../commands/pivot-report-update.js';
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

/** Edits that rewrite the report: each is one undo step under its ribbon command and re-lays out once. */
const REPORT_EDITS: Readonly<
  Record<string, { run: (wb: WorkbookHandle, ref: PivotRef) => boolean; command: string }>
> = {
  'totals-off': {
    run: (wb, ref) => applyPivotGrandTotals(wb, ref, { rows: false, cols: false }),
    command: 'pivotGrandTotals',
  },
  'totals-on': {
    run: (wb, ref) => applyPivotGrandTotals(wb, ref, { rows: true, cols: true }),
    command: 'pivotGrandTotals',
  },
  'totals-rows': {
    run: (wb, ref) => applyPivotGrandTotals(wb, ref, { rows: true, cols: false }),
    command: 'pivotGrandTotals',
  },
  'totals-cols': {
    run: (wb, ref) => applyPivotGrandTotals(wb, ref, { rows: false, cols: true }),
    command: 'pivotGrandTotals',
  },
  'layout-compact': {
    run: (wb, ref) => applyPivotReportLayout(wb, ref, PivotReportLayout.Compact),
    command: 'pivotReportLayout',
  },
  'layout-outline': {
    run: (wb, ref) => applyPivotReportLayout(wb, ref, PivotReportLayout.Outline),
    command: 'pivotReportLayout',
  },
  'layout-tabular': {
    run: (wb, ref) => applyPivotReportLayout(wb, ref, PivotReportLayout.Tabular),
    command: 'pivotReportLayout',
  },
  'subtotals-bottom': {
    run: (wb, ref) => applyPivotSubtotalPosition(wb, ref, false),
    command: 'pivotSubtotals',
  },
  'subtotals-top': {
    run: (wb, ref) => applyPivotSubtotalPosition(wb, ref, true),
    command: 'pivotSubtotals',
  },
  'clear-all': { run: clearPivotFields, command: 'pivotClear' },
  'clear-filters': { run: clearPivotFilters, command: 'pivotClear' },
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
      instance.workbook,
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
        const range = pivotReportRange(wb, ref.sheetIndex, ref.pivotIndex);
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
          commandId: edit.command,
          onChanged,
        },
        ref.sheetIndex,
        ref.pivotIndex,
        () => edit.run(wb, ref),
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
