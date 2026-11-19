import { authorizedBatch } from './authorized-write.js';
import { hyperlinkIntents, hyperlinkOperation } from './hyperlink-actions.js';
import type { RefusedDecision } from './permission-reasons.js';
import { sameJsonValue } from './prepared-change.js';
// Shared "Clear" ribbon split-button action — invoked by every host wrapper.
// The command owns selection planning, policy preflight, sparse metadata
// mutation, and the history boundary so all ribbon hosts have the same
// behavior for disjoint selections.

import { addrKey, parseAddrKey, uniqueAddrs } from '../engine/address.js';
import { flushFormatToEngine } from '../engine/cell-format-sync.js';
import { allowedEditRangeAreas } from '../engine/protection-sync.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { subtractFromRangeValidations } from '../store/range-validations.js';
import {
  isValidSheetRange,
  normalizeRange,
  rangesIntersect,
  selectionContainsAddr,
  selectionCoversRange,
  selectionRangeList,
} from '../store/selection-geometry.js';
import { type CellFormat, mutators, type SpreadsheetStore, type State } from '../store/store.js';
import {
  type ClearContentsCommands,
  clearSelectedContents,
  collectSelectedContentAddresses,
  selectionBounds,
} from './clear-contents.js';
import {
  clearComments,
  listThreadedComments,
  recordCommentChange,
  recordThreadedChange,
  removeThreadedComment,
} from './comment.js';
import { type CommentActionContext, commentIntents, threadWrite } from './comment-actions.js';
import { withSelectionFormatOrigin } from './format.js';
import type { History, HistoryEntry } from './history.js';
import type { InteractionController } from './interaction-controller.js';
import {
  InteractionController as FullInteractionController,
  interactionControllerFor,
} from './interaction-controller.js';
import {
  appliedBatch,
  type ChangeBatchResult,
  type InteractionOperation,
  type InteractionOrigin,
  type OperationEffect,
  refusedBatch,
  refusedByDecision,
} from './interaction-policy.js';
import {
  isAllowedEditRangeOpen,
  isCellWritable,
  isSheetProtected,
  MAX_MATERIALIZED_LOCK_CELLS,
  noteChangePermitted,
} from './protection.js';
import { recordConditionalRulesChange } from './slice-history.js';
import { clearSparklinesAction } from './sparkline.js';

export type RibbonClearAction =
  | 'all'
  | 'formats'
  | 'contents'
  | 'comments'
  | 'hyperlinks'
  | 'conditional';

export interface ExecuteRibbonClearActionDeps {
  store: SpreadsheetStore;
  workbook: WorkbookHandle;
  history: History;
  action: RibbonClearAction;
  /** Optional mounted controller. Contents uses it to share policy/history. */
  commands?: Pick<InteractionController, 'execute' | 'policy'>;
  /** The entry point the clear runs as; the Home ribbon's Clear menu item when omitted. */
  entry?: { readonly origin: InteractionOrigin; readonly commandId: string };
}

const visualFormatKeys: readonly (keyof CellFormat)[] = [
  'cellStyle',
  'numFmt',
  'bold',
  'italic',
  'underline',
  'strike',
  'align',
  'vAlign',
  'wrap',
  'shrinkToFit',
  'indent',
  'rotation',
  'textDirection',
  'borders',
  'color',
  'fill',
  'fillPattern',
  'fillPatternColor',
  'fillGradient',
  'fontOutline',
  'fontShadow',
  'fontCondense',
  'fontExtend',
  'fontFamily',
  'fontScheme',
  'fontSize',
];

type SelectionClearPlan = {
  readonly sheet: number;
  readonly ranges: readonly Range[];
  readonly content: readonly Addr[];
  readonly formatEntries: readonly Addr[];
  readonly visualEntries: readonly Addr[];
  readonly comments: readonly Addr[];
  readonly hyperlinks: readonly Addr[];
  readonly validations: readonly Addr[];
  readonly conditionalRuleIndexes: readonly number[];
  readonly writableConditionalRuleIndexes: readonly number[];
};

type ScopedFormatSnapshot = ReadonlyMap<string, CellFormat | undefined>;

const cloneValue = <T>(value: T): T => {
  if (value === undefined || value === null) return value;
  return structuredClone(value);
};

const parseNonNegativeAddrKey = (key: string): Addr | null => {
  const addr = parseAddrKey(key);
  return addr && addr.sheet >= 0 && addr.row >= 0 && addr.col >= 0 ? addr : null;
};

const hasComment = (format: CellFormat): boolean =>
  typeof format.comment === 'string' && format.comment.length > 0;

const hasHyperlink = (format: CellFormat): boolean =>
  typeof format.hyperlink === 'string' && format.hyperlink.length > 0;

const hasVisualFormat = (format: CellFormat): boolean =>
  visualFormatKeys.some((key) => format[key] !== undefined);

const buildPlan = (
  store: SpreadsheetStore,
  workbook: WorkbookHandle,
): SelectionClearPlan | null => {
  const state = store.getState();
  const primary = state.selection.range;
  if (!primary) return null;
  const ranges = [primary, ...(state.selection.extraRanges ?? [])].map((range) => ({ ...range }));
  if (!ranges.every((range) => isValidSheetRange(range, primary.sheet))) return null;
  const selection = { range: ranges[0] as Range, extraRanges: ranges.slice(1) };

  // A selection touching only part of a merge is unsafe to clear. This check
  // is shared by content and format paths, so malformed plans are complete
  // no-ops instead of partially clearing the rectangle around a merge.
  for (const merge of state.merges.byAnchor.values()) {
    if (
      ranges.some((range) => rangesIntersect(range, merge)) &&
      !selectionCoversRange(selection, merge)
    ) {
      return null;
    }
  }

  const formatEntries: Addr[] = [];
  const visualEntries: Addr[] = [];
  const comments: Addr[] = [];
  const hyperlinks: Addr[] = [];
  const validations: Addr[] = [];
  for (const [key, format] of state.format.formats) {
    const addr = parseNonNegativeAddrKey(key);
    if (!addr || !selectionContainsAddr(selection, addr)) continue;
    formatEntries.push(addr);
    if (hasVisualFormat(format)) visualEntries.push(addr);
    if (hasComment(format)) comments.push(addr);
    if (hasHyperlink(format)) hyperlinks.push(addr);
    if (format.validation !== undefined) validations.push(addr);
  }

  const conditionalRuleIndexes: number[] = [];
  const writableConditionalRuleIndexes: number[] = [];
  state.conditional.rules.forEach((rule, index) => {
    if (ranges.some((range) => rangesIntersect(range, rule.range))) {
      conditionalRuleIndexes.push(index);
      if (conditionalRuleWritable(state, rule.range)) writableConditionalRuleIndexes.push(index);
    }
  });

  return {
    sheet: primary.sheet,
    ranges,
    content: collectSelectedContentAddresses(store, workbook),
    formatEntries: uniqueAddrs(formatEntries),
    visualEntries: uniqueAddrs(visualEntries),
    comments: uniqueAddrs(comments),
    hyperlinks: uniqueAddrs(hyperlinks),
    validations: uniqueAddrs(validations),
    conditionalRuleIndexes,
    writableConditionalRuleIndexes,
  };
};

const snapshotScopedFormats = (state: State, addrs: readonly Addr[]): ScopedFormatSnapshot => {
  const snapshot = new Map<string, CellFormat | undefined>();
  for (const addr of uniqueAddrs(addrs)) {
    const key = addrKey(addr);
    snapshot.set(key, cloneValue(state.format.formats.get(key)));
  }
  return snapshot;
};

const sameScopedFormats = (before: ScopedFormatSnapshot, after: ScopedFormatSnapshot): boolean => {
  if (before.size !== after.size) return false;
  for (const [key, value] of before) {
    if (!sameJsonValue(value, after.get(key))) return false;
  }
  return true;
};

const applyScopedFormats = (
  store: SpreadsheetStore,
  workbook: WorkbookHandle,
  sheet: number,
  snapshot: ScopedFormatSnapshot,
): void => {
  store.setState((state) => {
    const formats = new Map(state.format.formats);
    for (const [key, format] of snapshot) {
      if (format === undefined) formats.delete(key);
      else formats.set(key, cloneValue(format));
    }
    return { ...state, format: { ...state.format, formats } };
  });
  flushFormatToEngine(workbook, store, sheet);
};

interface ScopedFormatOptions {
  readonly repeat?: () => void;
}

/** Record only selected format entries and keep engine state synchronized on
 * forward, undo, and redo. Whole-map format snapshots would overwrite edits
 * made to unrelated cells after Clear Hyperlinks/Formats. */
const recordClearFormatEntries = (
  history: History | null,
  store: SpreadsheetStore,
  workbook: WorkbookHandle,
  sheet: number,
  addrs: readonly Addr[],
  mutate: () => void,
  options: ScopedFormatOptions = {},
): void => {
  const tracked = uniqueAddrs(addrs);
  const before = snapshotScopedFormats(store.getState(), tracked);
  try {
    mutate();
    flushFormatToEngine(workbook, store, sheet);
  } catch (error) {
    try {
      applyScopedFormats(store, workbook, sheet, before);
    } catch {
      // Preserve the original engine/mutation error; the outer transaction
      // still restores its material children.
    }
    throw error;
  }
  const after = snapshotScopedFormats(store.getState(), tracked);
  if (options.repeat) history?.setRepeat(options.repeat);
  if (!history || history.isReplaying() || sameScopedFormats(before, after)) return;
  const changedBefore = new Map<string, CellFormat | undefined>();
  const changedAfter = new Map<string, CellFormat | undefined>();
  for (const [key, value] of before) {
    const next = after.get(key);
    if (sameJsonValue(value, next)) continue;
    changedBefore.set(key, value);
    changedAfter.set(key, next);
  }
  const entry: HistoryEntry = {
    undo: () => applyScopedFormats(store, workbook, sheet, changedBefore),
    redo: () => applyScopedFormats(store, workbook, sheet, changedAfter),
    repeat: options.repeat,
  };
  history.push(entry);
};

const clearPlannedFormats = (
  store: SpreadsheetStore,
  plan: SelectionClearPlan,
  visualOnly: boolean,
): void => {
  const selected = new Set(plan.formatEntries.map(addrKey));
  store.setState((state) => {
    const formats = new Map(state.format.formats);
    for (const key of selected) {
      const current = formats.get(key);
      const addr = parseNonNegativeAddrKey(key);
      if (!current || !addr || !isCellWritable(state, addr)) continue;
      if (!visualOnly) {
        formats.delete(key);
        continue;
      }
      const next = { ...current };
      for (const formatKey of visualFormatKeys) delete next[formatKey];
      if (Object.keys(next).length === 0) formats.delete(key);
      else formats.set(key, next);
    }
    let rangeValidations = state.format.rangeValidations;
    if (!visualOnly) {
      for (const range of plan.ranges) {
        rangeValidations = subtractFromRangeValidations(rangeValidations, range);
      }
    }
    return {
      ...state,
      format: { ...state.format, formats, rangeValidations },
    };
  });
};

const clearPlannedHyperlinks = (store: SpreadsheetStore, plan: SelectionClearPlan): void => {
  const selected = new Set(plan.hyperlinks.map(addrKey));
  store.setState((state) => {
    const formats = new Map(state.format.formats);
    for (const key of selected) {
      const current = formats.get(key);
      const addr = parseNonNegativeAddrKey(key);
      if (!current || !addr || !isCellWritable(state, addr)) continue;
      const {
        hyperlink: _hyperlink,
        hyperlinkLocation: _location,
        hyperlinkSpan: _span,
        hyperlinkDisplay: _display,
        hyperlinkTooltip: _tooltip,
        ...next
      } = current;
      if (Object.keys(next).length === 0) formats.delete(key);
      else formats.set(key, next);
    }
    return { ...state, format: { ...state.format, formats } };
  });
};

const rangesForRuleIndexes = (state: State, indexes: readonly number[]): Range[] =>
  indexes.flatMap((index) => {
    const rule = state.conditional.rules[index];
    return rule ? [{ ...rule.range }] : [];
  });

function conditionalRuleWritable(state: State, range: Range): boolean {
  if (!isSheetProtected(state, range.sheet)) return true;
  const area = (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);
  if (Number.isSafeInteger(area) && area <= MAX_MATERIALIZED_LOCK_CELLS) {
    for (let row = range.r0; row <= range.r1; row += 1) {
      for (let col = range.c0; col <= range.c1; col += 1) {
        if (!isCellWritable(state, { sheet: range.sheet, row, col })) return false;
      }
    }
    return true;
  }
  const allowed = state.protection.allowedEditRanges
    .filter((entry) => entry.range.sheet === range.sheet && isAllowedEditRangeOpen(state, entry))
    .flatMap((entry) => allowedEditRangeAreas(entry));
  const first = allowed[0];
  if (!first) return false;
  return selectionCoversRange({ range: first, extraRanges: allowed.slice(1) }, range);
}

const clearPlannedConditionalRules = (store: SpreadsheetStore, plan: SelectionClearPlan): void => {
  const indexes = new Set(plan.conditionalRuleIndexes);
  const writableIndexes = new Set(plan.writableConditionalRuleIndexes);
  store.setState((state) => {
    const rules = state.conditional.rules.filter(
      (_rule, index) => !indexes.has(index) || !writableIndexes.has(index),
    );
    return rules.length === state.conditional.rules.length
      ? state
      : { ...state, conditional: { rules } };
  });
};

const effectsForCells = (
  cells: readonly Addr[],
  includesFormula = false,
): readonly OperationEffect[] =>
  cells.length === 0
    ? []
    : [{ kind: 'cells', cells, ...(includesFormula ? { includesFormula: true } : {}) }];

const effectsForRanges = (ranges: readonly Range[]): readonly OperationEffect[] =>
  ranges.map((range) => ({ kind: 'range' as const, range }));

/** The decision that refuses one category of a planned clear, or null. */
const preflight = (
  controller: FullInteractionController | undefined,
  origin: InteractionOrigin,
  operation: InteractionOperation,
  commandId: string,
  effects: readonly OperationEffect[],
): RefusedDecision | null => {
  if (!controller || effects.length === 0 || !controller.policy) return null;
  const decision = controller.canExecute({ operation, origin, commandId, effects });
  return decision.allowed ? null : decision;
};

const resolveTransactionController = (
  deps: ExecuteRibbonClearActionDeps,
  alwaysCreate: boolean,
): { controller?: FullInteractionController; mismatch: boolean; ephemeral: boolean } => {
  const registered = interactionControllerFor(deps.store);
  if (
    registered &&
    deps.commands &&
    deps.commands !== registered &&
    deps.commands.policy !== registered.policy
  ) {
    return { mismatch: true, ephemeral: false };
  }
  // A registered controller always owns the mounted write; the ephemeral
  // transaction controller serves only hosts without one.
  if (registered) return { controller: registered, mismatch: false, ephemeral: false };
  const policy = deps.commands?.policy;
  if (!alwaysCreate && !policy) return { mismatch: false, ephemeral: false };
  const controller = new FullInteractionController({
    store: deps.store,
    getWb: () => deps.workbook,
    history: deps.history,
    manageHistoryGuard: false,
  });
  if (policy) controller.setPolicy(policy);
  return { controller, mismatch: false, ephemeral: true };
};

/** The decision that refuses the planned clear under the host policy, or null. */
const preflightPlan = (
  deps: ExecuteRibbonClearActionDeps,
  plan: SelectionClearPlan,
  controller: FullInteractionController | undefined,
  origin: InteractionOrigin,
  commandId: string,
): RefusedDecision | null => {
  if (!controller?.policy) return null;
  const check = (operation: InteractionOperation, effects: readonly OperationEffect[]) =>
    preflight(controller, origin, operation, commandId, effects);
  return (
    check('clear', effectsForCells(plan.content)) ??
    check('comment', effectsForCells(plan.comments)) ??
    check(hyperlinkOperation(true), effectsForCells(plan.hyperlinks)) ??
    check('validation', effectsForCells(plan.validations)) ??
    check('format', effectsForCells(plan.formatEntries)) ??
    check(
      'format',
      effectsForRanges(rangesForRuleIndexes(deps.store.getState(), plan.conditionalRuleIndexes)),
    )
  );
};

/** Remove the threads of the planned selection as one step, replies before their opener. */
const clearSelectedThreads = (
  ctx: CommentActionContext,
  workbook: WorkbookHandle,
  plan: SelectionClearPlan,
): boolean => {
  if (!workbook.capabilities.threadedComments) return false;
  const { sheet } = plan;
  const threads = listThreadedComments(ctx.store, sheet, workbook).filter((entry) =>
    plan.ranges.some(
      (range) =>
        entry.cell.row >= Math.min(range.r0, range.r1) &&
        entry.cell.row <= Math.max(range.r0, range.r1) &&
        entry.cell.col >= Math.min(range.c0, range.c1) &&
        entry.cell.col <= Math.max(range.c0, range.c1),
    ),
  );
  if (threads.length === 0) return false;
  const cells = threads.map((entry) => ({ sheet, row: entry.cell.row, col: entry.cell.col }));
  return threadWrite(ctx, cells, false, () => {
    recordThreadedChange(
      ctx.history,
      ctx.store,
      workbook,
      sheet,
      () => {
        for (const entry of [...threads].sort(
          (a, b) => Number(a.parentId === null) - Number(b.parentId === null),
        ))
          removeThreadedComment(ctx.store, sheet, entry.id, workbook, ctx);
      },
      ctx,
    );
    return true;
  });
};

const refreshCells = (store: SpreadsheetStore, workbook: WorkbookHandle, sheet: number): void => {
  mutators.replaceCells(store, workbook.cells(sheet));
};

/** A Home ribbon Clear action, as one batch result: `range` is the selection, `applied` the cells
 *  whose content it changed (none for the clears that leave content as it is). */
export const executeRibbonClearAction = (deps: ExecuteRibbonClearActionDeps): ChangeBatchResult => {
  const { store, workbook, history, action } = deps;
  const origin = deps.entry?.origin ?? 'ribbon';
  const selected = selectionBounds(store.getState());

  // Contents keeps its public command adapter behavior. Clear All runs its
  // content child through the resolved controller inside one History step.
  if (action === 'contents') {
    const range = store.getState().selection.range;
    const result = clearSelectedContents({
      store,
      workbook,
      history,
      commands: deps.commands as ClearContentsCommands | undefined,
      origin,
      ...(deps.entry ? { commandId: deps.entry.commandId } : {}),
    });
    if (result.status !== 'refused') refreshCells(store, workbook, range.sheet);
    return result;
  }

  const resolved = resolveTransactionController(deps, action === 'all');
  if (resolved.mismatch)
    return refusedBatch({ code: 'invalid', reason: 'the commands do not belong to this store' });
  const plan = buildPlan(store, workbook);
  if (!plan) {
    if (resolved.ephemeral) resolved.controller?.dispose();
    return refusedBatch({ code: 'invalid', reason: 'the selection cannot be cleared' });
  }

  const controller = resolved.controller;
  const policyPlan: SelectionClearPlan =
    action === 'formats'
      ? {
          ...plan,
          comments: [],
          hyperlinks: [],
          validations: [],
          content: [],
          formatEntries: plan.visualEntries,
          conditionalRuleIndexes: [],
        }
      : action === 'comments'
        ? {
            ...plan,
            // A protected sheet keeps its notes unless Edit objects is on.
            comments: plan.comments.filter((addr) => noteChangePermitted(store.getState(), addr)),
            hyperlinks: [],
            validations: [],
            content: [],
            formatEntries: [],
            visualEntries: [],
            conditionalRuleIndexes: [],
          }
        : action === 'hyperlinks'
          ? {
              ...plan,
              comments: [],
              validations: [],
              content: [],
              formatEntries: [],
              visualEntries: [],
              conditionalRuleIndexes: [],
            }
          : action === 'conditional'
            ? {
                ...plan,
                comments: [],
                hyperlinks: [],
                validations: [],
                content: [],
                formatEntries: [],
                visualEntries: [],
                conditionalRuleIndexes: plan.conditionalRuleIndexes,
              }
            : plan;
  const commandId =
    deps.entry?.commandId ??
    (action === 'all'
      ? 'clear-all'
      : action === 'formats'
        ? 'clear-formats'
        : action === 'comments'
          ? 'clear-comments'
          : action === 'hyperlinks'
            ? 'clear-hyperlinks'
            : 'clear-conditional');
  const refusal = preflightPlan(deps, policyPlan, controller, origin, commandId);
  if (refusal) {
    if (resolved.ephemeral) controller?.dispose();
    return refusedByDecision(refusal);
  }
  const unchangedContent = appliedBatch([], selected);

  try {
    if (action === 'formats') {
      const targets = plan.visualEntries;
      recordClearFormatEntries(
        history,
        store,
        workbook,
        plan.sheet,
        targets,
        () =>
          withSelectionFormatOrigin(
            store,
            origin,
            () => clearPlannedFormats(store, plan, true),
            commandId,
          ),
        { repeat: () => executeRibbonClearAction(deps) },
      );
      return unchangedContent;
    }
    if (action === 'comments') {
      // A protected sheet keeps its notes unless Edit objects is on; it never keeps threads.
      const notes = plan.comments.filter((addr) => noteChangePermitted(store.getState(), addr));
      const ctx = { store, history, origin, commandId };
      const noteResult =
        notes.length === 0
          ? unchangedContent
          : authorizedBatch({
              store,
              history,
              intents: commentIntents(ctx, notes, false),
              origin,
              write: () => {
                recordCommentChange(history, store, workbook, notes, () => {
                  clearComments(store, notes, workbook);
                });
                return unchangedContent;
              },
            });
      return clearSelectedThreads(ctx, workbook, plan) ? unchangedContent : noteResult;
    }
    if (action === 'hyperlinks') {
      if (plan.hyperlinks.length === 0) return unchangedContent;
      return authorizedBatch({
        store,
        history,
        intents: hyperlinkIntents({ commandId }, plan.hyperlinks, true),
        origin,
        write: () => {
          recordClearFormatEntries(history, store, workbook, plan.sheet, plan.hyperlinks, () => {
            clearPlannedHyperlinks(store, plan);
          });
          return unchangedContent;
        },
      });
    }
    if (action === 'conditional') {
      recordConditionalRulesChange(history, store, () => {
        clearPlannedConditionalRules(store, plan);
      });
      return unchangedContent;
    }

    let cleared: ChangeBatchResult;
    try {
      cleared = history.transact(
        () => {
          const result = clearSelectedContents({
            store,
            workbook,
            history,
            commands: controller,
            origin,
            commandId,
          });
          if (result.status === 'refused') return result;

          // Clear All removes the sparklines of the cleared, unlocked cells (read before the formats
          // reset their lock); the rest of a group stays.
          const { selection } = store.getState();
          for (const range of selectionRangeList(selection)) {
            clearSparklinesAction({ store, history, origin, commandId }, normalizeRange(range));
          }

          if (plan.comments.length > 0) {
            recordCommentChange(history, store, workbook, plan.comments, () => {
              clearComments(store, plan.comments, workbook);
            });
          }

          const formatTargets = plan.formatEntries;
          if (formatTargets.length > 0) {
            recordClearFormatEntries(history, store, workbook, plan.sheet, formatTargets, () =>
              withSelectionFormatOrigin(
                store,
                origin,
                () => clearPlannedFormats(store, plan, false),
                commandId,
              ),
            );
          }

          if (plan.conditionalRuleIndexes.length > 0) {
            recordConditionalRulesChange(history, store, () => {
              clearPlannedConditionalRules(store, plan);
            });
          }
          refreshCells(store, workbook, plan.sheet);
          return result;
        },
        { keep: (result) => result.status !== 'refused' },
      );
    } catch (error) {
      refreshCells(store, workbook, plan.sheet);
      throw error;
    }
    if (cleared.status === 'refused') refreshCells(store, workbook, plan.sheet);
    return cleared;
  } finally {
    if (resolved.ephemeral) resolved.controller?.dispose();
  }
};
