import type { FunctionCategory } from '../../../commands/function-categories.js';
import type { InteractionOperation } from '../../../commands/interaction-policy.js';
import {
  disposeMacConsolidateDialog,
  openMacConsolidateDialog,
} from '../../../interact/mac-consolidate-dialog.js';
import {
  disposeMacGoalSeekDialog,
  openMacGoalSeekDialog,
} from '../../../interact/mac-goal-seek-dialog.js';
import { attachMacSlicerDialog } from '../../../interact/mac-slicer-dialog.js';
import { attachMacSparklineDialog } from '../../../interact/mac-sparkline-dialog.js';
import {
  disposeMacSubtotalDialog,
  openMacSubtotalDialog,
} from '../../../interact/mac-subtotal-dialog.js';
import { attachMacWorkbookStatistics } from '../../../interact/mac-workbook-statistics.js';
import type { SpreadsheetInstance } from '../../../mount/types.js';
import {
  disposeMacAutomationDialogs,
  openMacAutomationGallery,
  runMacAutomationScript,
} from './automation-gallery.js';

export const MAC_RIBBON_ACTION_IDS = [
  'mac.insert.pivotTable',
  'mac.insert.sparkline',
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
  'mac.data.consolidate',
  'mac.data.subtotal',
  'mac.automate.showScripts',
  'mac.automate.gallery',
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

const policyAllows = (
  instance: SpreadsheetInstance,
  operation: InteractionOperation,
  commandId: string,
): boolean => {
  if (!instance.commands.policy) return true;
  const range = instance.store.getState().selection.range;
  return instance.commands.canExecute({
    operation,
    origin: 'ribbon',
    commandId,
    effects: [{ kind: 'range', range }],
  }).allowed;
};

const ACTIONS_BY_INSTANCE = new WeakMap<object, MacRibbonActions>();

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
    canCommit: (_source, destination) => {
      if (!instance.commands.policy) return true;
      return instance.commands.canExecute({
        operation: 'format',
        origin: 'ribbon',
        commandId: 'mac.insert.sparkline',
        effects: [{ kind: 'cells', cells: [destination] }],
      }).allowed;
    },
  });
  const slicer = attachMacSlicerDialog({
    host: instance.host,
    store: instance.store,
    getWb: () => instance.workbook,
    strings: instance.i18n.strings,
    onAdd: (input) => {
      if (!policyAllows(instance, 'table', 'mac.insert.slicer')) return null;
      try {
        return instance.addSlicer(input);
      } catch {
        return null;
      }
    },
  });
  const statistics = attachMacWorkbookStatistics({
    host: instance.host,
    getWb: () => instance.workbook,
    store: instance.store,
    strings: instance.i18n.strings,
  });

  const unsubscribeLocale = instance.i18n.subscribe((next) => {
    // Dialogs stay mounted across a locale switch so an open picker does not
    // lose its current values. Each handle updates its labels in place.
    sparkline.setStrings(next);
    slicer.setStrings(next);
    statistics.setStrings(next);
  });

  const actions: Record<string, MacRibbonAction> = {
    'mac.insert.pivotTable': () => instance.openPivotTableDialog(),
    'mac.insert.sparkline': () => sparkline.open(),
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
    'mac.data.consolidate': () => openMacConsolidateDialog(instance),
    'mac.data.subtotal': () => openMacSubtotalDialog(instance),
    'mac.automate.showScripts': () => openMacAutomationGallery(instance),
    'mac.automate.gallery': () => openMacAutomationGallery(instance),
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
