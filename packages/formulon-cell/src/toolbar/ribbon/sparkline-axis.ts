import { SparklineAxisType } from '@libraz/formulon';
import { isSheetProtected } from '../../commands/protection.js';
import {
  activeSparklineGroup,
  patchSelectedSparklineGroupsAction,
} from '../../commands/sparkline-design.js';
import { formatA1Range } from '../../engine/address.js';
import type { SpreadsheetInstance } from '../../mount/types.js';
import type { SparklineGroupSettings } from '../../store/store.js';
import { showSparklineAxisDialog } from '../dialogs/sparkline.js';
import { SPARKLINE_COMMANDS } from './sparkline-tabs.js';

/** Open the Axis dialog on the group of the active cell and apply what it returns as one undo
 *  step. On a protected sheet the button stays enabled and does nothing, as on the desktop. */
export const openSparklineAxisDialog = async (instance: SpreadsheetInstance): Promise<void> => {
  const state = instance.store.getState();
  const group = activeSparklineGroup(state);
  if (!group || isSheetProtected(state, state.selection.active.sheet)) return;
  const settings = group.settings;
  const m = instance.i18n.strings.mac.sparkline;
  const range = settings.dateAxisRange?.rect;
  const result = await showSparklineAxisDialog({
    host: instance.host,
    store: instance.store,
    getWb: () => instance.workbook,
    text: {
      title: m.axisTitle,
      horizontal: m.axisHorizontal,
      vertical: m.axisVertical,
      general: m.axisGeneral,
      date: m.dateTitle,
      dateRange: m.dateRangeLabel,
      showAxis: m.axisShow,
      plotRtl: m.axisPlotRtl,
      rtlUnsupported: m.axisRtlUnsupported,
      minimum: m.axisMinimum,
      maximum: m.axisMaximum,
      autoEach: m.axisAutoEach,
      sameAll: m.axisSameAll,
      customValue: m.customValueLabel,
      invalidRange: m.invalidSource,
      dateRangeNotLine: m.dateAxisNotLine,
      pickerLabel: instance.i18n.strings.pivotTableDialog.rangePickerSelect,
      ok: m.ok,
      cancel: m.cancel,
    },
    initial: {
      dateAxis: settings.dateAxis,
      dateRange: range
        ? formatA1Range({
            r0: range.firstRow,
            c0: range.firstCol,
            r1: range.lastRow,
            c1: range.lastCol,
          })
        : '',
      displayXAxis: settings.displayXAxis,
      min: { type: settings.minAxisType, value: settings.manualMin ?? null },
      max: { type: settings.maxAxisType, value: settings.manualMax ?? null },
    },
  });
  if (!result) return;
  const r = result.dateRange;
  const patch: Partial<SparklineGroupSettings> = {
    dateAxis: result.dateAxis,
    displayXAxis: result.displayXAxis,
    minAxisType: result.min.type,
    maxAxisType: result.max.type,
    ...(r
      ? {
          dateAxisRange: {
            sheet: r.sheet,
            rect: { firstRow: r.r0, firstCol: r.c0, lastRow: r.r1, lastCol: r.c1 },
          },
        }
      : {}),
    ...(result.min.type === SparklineAxisType.Custom && result.min.value !== null
      ? { manualMin: result.min.value }
      : {}),
    ...(result.max.type === SparklineAxisType.Custom && result.max.value !== null
      ? { manualMax: result.max.value }
      : {}),
  };
  patchSelectedSparklineGroupsAction(
    {
      store: instance.store,
      history: instance.history,
      origin: 'ribbon',
      commandId: SPARKLINE_COMMANDS.axis,
    },
    patch,
  );
};
