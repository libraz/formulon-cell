import type { FunctionCategory } from '../../../commands/function-categories.js';
import { objectChangeAllowed, objectWrite } from '../../../commands/object-actions.js';
import { membersOfGroup } from '../../../commands/sparkline.js';
import { activeSparklineGroup } from '../../../commands/sparkline-design.js';
import { unionSparklineSources } from '../../../commands/sparkline-layout.js';
import { resolveSparklineRange } from '../../../commands/sparkline-ranges.js';
import { formatA1Range, quoteSheetName } from '../../../engine/address.js';
import {
  disposeDataTableDialog,
  openDataTableDialogFrom,
} from '../../../interact/data-table-dialog.js';
import { openPivotTableDialogFrom } from '../../../interact/dialog-entry.js';
import {
  disposeMacConsolidateDialog,
  openMacConsolidateDialog,
} from '../../../interact/mac-consolidate-dialog.js';
import {
  disposeMacGoalSeekDialog,
  openMacGoalSeekDialog,
} from '../../../interact/mac-goal-seek-dialog.js';
import {
  attachMacSlicerDialog,
  type MacSlicerDialogHandle,
  type MacSlicerOpenOptions,
} from '../../../interact/mac-slicer-dialog.js';
import {
  attachMacSparklineDialog,
  type MacSparklineDialogHandle,
  type MacSparklineEditTarget,
} from '../../../interact/mac-sparkline-dialog.js';
import {
  disposeMacSubtotalDialog,
  openMacSubtotalDialog,
} from '../../../interact/mac-subtotal-dialog.js';
import { attachMacWorkbookStatistics } from '../../../interact/mac-workbook-statistics.js';
import type { SpreadsheetInstance } from '../../../mount/types.js';
import type { SlicerSpec } from '../../../store/store.js';
import {
  disposeMacAutomationDialogs,
  openMacAutomationGallery,
  runMacAutomationScript,
} from './automation-gallery.js';

export const MAC_RIBBON_ACTION_IDS = [
  'mac.insert.pivotTable',
  'mac.insert.sparkline',
  'mac.insert.sparkline.line',
  'mac.insert.sparkline.column',
  'mac.insert.sparkline.winLoss',
  'mac.insert.slicer',
  'mac.formulas.insertFunction',
  'mac.formulas.recent',
  'mac.formulas.financial',
  'mac.formulas.logical',
  'mac.formulas.text',
  'mac.formulas.dateTime',
  'mac.formulas.lookup',
  'mac.formulas.math',
  'mac.formulas.more',
  'mac.formulas.category.recent',
  'mac.formulas.category.all',
  'mac.review.stats',
  'mac.data.goalSeek',
  'mac.data.dataTable',
  'mac.data.consolidate',
  'mac.data.subtotal',
  'mac.automate.showScripts',
  'mac.automate.showScripts.samples',
  'mac.automate.allRowsColumns',
  'mac.automate.freezeSelection',
  'mac.automate.makeSubtable',
  'mac.automate.removeHyperlinks',
  'mac.automate.countEmptyRows',
  'mac.automate.tableToJson',
  'mac.automate.newPivotTable',
] as const;

export type MacRibbonActionId = (typeof MAC_RIBBON_ACTION_IDS)[number];
export type MacRibbonAction = () => void | Promise<void>;
export type MacRibbonActions = Readonly<Record<string, MacRibbonAction>>;

const ACTIONS_BY_INSTANCE = new WeakMap<object, MacRibbonActions>();
const SPARKLINE_DIALOGS = new WeakMap<object, MacSparklineDialogHandle>();
const SLICER_DIALOGS = new WeakMap<object, MacSlicerDialogHandle>();

/** Open the instance's shared sparkline dialog: Insert Sparkline, or Edit Sparklines for `edit`. */
export const openMacSparklineDialog = (
  instance: SpreadsheetInstance,
  edit?: MacSparklineEditTarget,
): void => {
  createMacRibbonActions(instance);
  SPARKLINE_DIALOGS.get(instance)?.open(edit);
};

/** Open Edit Sparklines on the group of the active cell. Both ranges name their sheet, and the
 *  location is pinned with `$`, as the desktop dialog fills them. */
export const openSparklineGroupEditor = (instance: SpreadsheetInstance): void => {
  const state = instance.store.getState();
  const group = activeSparklineGroup(state);
  if (!group) return;
  const members = membersOfGroup(state, group.id);
  const first = members[0];
  if (!first) return;
  const wb = instance.workbook;
  const sheet = first.addr.sheet;
  const qualified = (on: number, a1: string): string => `${quoteSheetName(wb.sheetName(on))}!${a1}`;
  const union = unionSparklineSources(members.map(({ spec }) => spec.source));
  const data = resolveSparklineRange(union, sheet, wb);
  const rows = members.map(({ addr }) => addr.row);
  const cols = members.map(({ addr }) => addr.col);
  openMacSparklineDialog(instance, {
    groupId: group.id,
    sheet,
    data: data ? qualified(data.sheet, formatA1Range(data)) : union,
    location: qualified(
      sheet,
      formatA1Range(
        {
          r0: Math.min(...rows),
          c0: Math.min(...cols),
          r1: Math.max(...rows),
          c1: Math.max(...cols),
        },
        { absolute: true },
      ),
    ),
  });
};

/** Open the instance's shared Insert Slicer dialog for the ribbon command that invoked it. */
export const openMacSlicerDialog = (
  instance: SpreadsheetInstance,
  options?: MacSlicerOpenOptions,
): void => {
  createMacRibbonActions(instance);
  SLICER_DIALOGS.get(instance)?.open(options);
};

const formulaCategory = (instance: SpreadsheetInstance, category: FunctionCategory): void => {
  instance.openFunctionArguments(undefined, { category });
};

/** Build the core actions consumed by the Mac ribbon integration. The factory
 *  keeps dialog handles tied to this instance and resolves the workbook lazily
 *  so a later `setWorkbook()` call is reflected by every chooser. */
export function createMacRibbonActions(instance: SpreadsheetInstance): MacRibbonActions {
  const existing = ACTIONS_BY_INSTANCE.get(instance);
  if (existing) return existing;

  const sparkline = attachMacSparklineDialog({
    host: instance.host,
    store: instance.store,
    history: instance.history,
    getWb: () => instance.workbook,
    strings: instance.i18n.strings,
    canCommit: (_source, destination) =>
      objectChangeAllowed(
        { store: instance.store, origin: 'ribbon', commandId: 'mac.insert.sparkline' },
        destination.sheet,
        [{ kind: 'cells', cells: [destination] }],
      ),
  });
  const slicer = attachMacSlicerDialog({
    host: instance.host,
    store: instance.store,
    getWb: () => instance.workbook,
    strings: instance.i18n.strings,
    onAdd: (inputs, commandId) => {
      const ctx = {
        store: instance.store,
        history: instance.history,
        origin: 'ribbon',
        commandId,
      } as const;
      const sheet = instance.store.getState().data.sheetIndex;
      // One undo step; a refused slicer rolls back the ones already added.
      const token = instance.history.begin();
      let added: SlicerSpec[] | null = null;
      try {
        added = objectWrite(ctx, sheet, null, () =>
          inputs.map((input) => instance.addSlicer(input)),
        );
      } catch {
        added = null;
      }
      if (added) instance.history.end(token);
      else instance.history.abort(token);
      return added;
    },
  });
  SPARKLINE_DIALOGS.set(instance, sparkline);
  SLICER_DIALOGS.set(instance, slicer);
  const statistics = attachMacWorkbookStatistics({
    host: instance.host,
    getWb: () => instance.workbook,
    store: instance.store,
    strings: instance.i18n.strings,
  });

  const unsubscribeLocale = instance.i18n.on('change', (next) => {
    // Dialogs stay mounted across a locale switch so an open picker does not
    // lose its current values. Each handle updates its labels in place.
    sparkline.setStrings(next);
    slicer.setStrings(next);
    statistics.setStrings(next);
  });

  const actions: Record<string, MacRibbonAction> = {
    'mac.insert.pivotTable': () =>
      openPivotTableDialogFrom(instance.store, undefined, {
        origin: 'ribbon',
        commandId: 'mac.insert.pivotTable',
      }),
    'mac.insert.sparkline': () => sparkline.open(),
    'mac.insert.sparkline.line': () => sparkline.open(undefined, 'line'),
    'mac.insert.sparkline.column': () => sparkline.open(undefined, 'column'),
    'mac.insert.sparkline.winLoss': () => sparkline.open(undefined, 'win-loss'),
    'mac.insert.slicer': () => slicer.open(),
    'mac.formulas.insertFunction': () => instance.openFunctionArguments(),
    'mac.formulas.recent': () => formulaCategory(instance, 'recent'),
    'mac.formulas.financial': () => formulaCategory(instance, 'financial'),
    'mac.formulas.logical': () => formulaCategory(instance, 'logical'),
    'mac.formulas.text': () => formulaCategory(instance, 'text'),
    'mac.formulas.dateTime': () => formulaCategory(instance, 'datetime'),
    'mac.formulas.lookup': () => formulaCategory(instance, 'lookup'),
    'mac.formulas.math': () => formulaCategory(instance, 'math'),
    'mac.formulas.more': () => formulaCategory(instance, 'all'),
    'mac.formulas.category.recent': () => formulaCategory(instance, 'recent'),
    'mac.formulas.category.all': () => formulaCategory(instance, 'all'),
    'mac.review.stats': () => statistics.open(),
    'mac.data.goalSeek': () => openMacGoalSeekDialog(instance),
    'mac.data.dataTable': () => openDataTableDialogFrom(instance, 'ribbon'),
    'mac.data.consolidate': () => openMacConsolidateDialog(instance),
    'mac.data.subtotal': () => openMacSubtotalDialog(instance),
    'mac.automate.showScripts': () => openMacAutomationGallery(instance),
    'mac.automate.showScripts.samples': () => openMacAutomationGallery(instance),
    'mac.automate.allRowsColumns': () => runMacAutomationScript(instance, 'allRowsColumns'),
    'mac.automate.freezeSelection': () => runMacAutomationScript(instance, 'freezeSelection'),
    'mac.automate.makeSubtable': () => runMacAutomationScript(instance, 'makeSubtable'),
    'mac.automate.removeHyperlinks': () => runMacAutomationScript(instance, 'removeHyperlinks'),
    'mac.automate.countEmptyRows': () => runMacAutomationScript(instance, 'countEmptyRows'),
    'mac.automate.tableToJson': () => runMacAutomationScript(instance, 'tableToJson'),
    'mac.automate.newPivotTable': () => runMacAutomationScript(instance, 'newPivotTable'),
  };

  // The provider is a plain command map by contract. Expose cleanup as a
  // non-enumerable helper for hosts that keep providers longer than the DOM
  // mount; ribbon dispatchers that iterate keys never see it as a command.
  Object.defineProperty(actions, 'dispose', {
    enumerable: false,
    value: (): void => {
      disposeMacAutomationDialogs(instance);
      disposeMacGoalSeekDialog(instance);
      disposeDataTableDialog(instance);
      disposeMacConsolidateDialog(instance);
      disposeMacSubtotalDialog(instance);
      sparkline.detach();
      slicer.detach();
      statistics.detach();
      unsubscribeLocale();
      ACTIONS_BY_INSTANCE.delete(instance);
    },
  });
  const frozen = Object.freeze(actions);
  ACTIONS_BY_INSTANCE.set(instance, frozen);
  return frozen;
}

/** Dispose the memoized provider when a mount goes away. Safe when the
 *  toolbar never requested Mac actions. */
export function disposeMacRibbonActions(instance: SpreadsheetInstance): void {
  const actions = ACTIONS_BY_INSTANCE.get(instance) as
    | (MacRibbonActions & { dispose?: () => void })
    | undefined;
  actions?.dispose?.();
}
