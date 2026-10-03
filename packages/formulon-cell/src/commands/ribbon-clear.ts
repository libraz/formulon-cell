// Shared "Clear" ribbon split-button action — invoked by every host wrapper.
// The command owns selection planning, policy preflight, sparse metadata
// mutation, and the history boundary so all ribbon hosts have the same
// behavior for disjoint selections.

import { addrKey, MAX_COL, MAX_ROW, parseAddrKey } from '../engine/address.js';
import { flushFormatToEngine } from '../engine/cell-format-sync.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import {
  rangesIntersect,
  selectionContainsAddr,
  selectionCoversRange,
} from '../store/selection-geometry.js';
import { type CellFormat, mutators, type SpreadsheetStore, type State } from '../store/store.js';
import {
  type ClearContentsCommands,
  clearSelectedContents,
  collectSelectedContentAddresses,
} from './clear-contents.js';
import { clearComments, recordCommentChange } from './comment.js';
import { withSelectionFormatOrigin } from './format.js';
import type { History, HistoryEntry } from './history.js';
import type { InteractionController } from './interaction-controller.js';
import {
  InteractionController as FullInteractionController,
  interactionControllerFor,
} from './interaction-controller.js';
import type { InteractionOperation, OperationEffect } from './interaction-policy.js';
import { isCellWritable, isSheetProtected } from './protection.js';
import { recordConditionalRulesChange } from './slice-history.js';

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
}

const MAX_MATERIALIZED_LOCK_CELLS = 100_000;

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
  'fontFamily',
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
  readonly pendingTarget: Addr | null;
  readonly pendingBefore: State['ui']['pendingFormat'];
  readonly conditionalRuleIndexes: readonly number[];
  readonly writableConditionalRuleIndexes: readonly number[];
};

type ScopedFormatSnapshot = ReadonlyMap<string, CellFormat | undefined>;

const cloneValue = <T>(value: T): T => {
  if (value === undefined || value === null) return value;
  return structuredClone(value);
};

const cloneAddr = (addr: Addr): Addr => ({ sheet: addr.sheet, row: addr.row, col: addr.col });

const validRange = (range: Range, sheet: number): boolean =>
  range.sheet === sheet &&
  Number.isInteger(range.sheet) &&
  Number.isInteger(range.r0) &&
  Number.isInteger(range.c0) &&
  Number.isInteger(range.r1) &&
  Number.isInteger(range.c1) &&
  range.sheet >= 0 &&
  range.r0 >= 0 &&
  range.c0 >= 0 &&
  range.r0 <= range.r1 &&
  range.c0 <= range.c1 &&
  range.r1 <= MAX_ROW &&
  range.c1 <= MAX_COL;

const parseNonNegativeAddrKey = (key: string): Addr | null => {
  const addr = parseAddrKey(key);
  return addr && addr.sheet >= 0 && addr.row >= 0 && addr.col >= 0 ? addr : null;
};

const uniqueAddrs = (addrs: readonly Addr[]): Addr[] => {
  const seen = new Set<string>();
  const result: Addr[] = [];
  for (const addr of addrs) {
    const key = addrKey(addr);
    if (seen.has(key)) continue;
    seen.add(key);
    result.push(cloneAddr(addr));
  }
  return result;
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
  if (!ranges.every((range) => validRange(range, primary.sheet))) return null;
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

  const pending = state.ui.pendingFormat;
  const pendingTarget =
    pending && selectionContainsAddr(selection, pending.addr) ? cloneAddr(pending.addr) : null;
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
    pendingTarget,
    pendingBefore: cloneValue(pending),
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

const sameValue = (left: unknown, right: unknown): boolean => {
  if (left === right) return true;
  return JSON.stringify(left) === JSON.stringify(right);
};

const sameScopedFormats = (before: ScopedFormatSnapshot, after: ScopedFormatSnapshot): boolean => {
  if (before.size !== after.size) return false;
  for (const [key, value] of before) {
    if (!sameValue(value, after.get(key))) return false;
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
  const pendingBefore = cloneValue(store.getState().ui.pendingFormat);
  try {
    mutate();
    flushFormatToEngine(workbook, store, sheet);
  } catch (error) {
    try {
      applyScopedFormats(store, workbook, sheet, before);
    } catch {
      // Preserve the original engine/mutation error; the outer transaction
      // still restores its material children and pending state.
    }
    mutators.setPendingFormat(store, pendingBefore);
    throw error;
  }
  const after = snapshotScopedFormats(store.getState(), tracked);
  if (options.repeat) history?.setRepeat(options.repeat);
  if (!history || history.isReplaying() || sameScopedFormats(before, after)) return;
  const changedBefore = new Map<string, CellFormat | undefined>();
  const changedAfter = new Map<string, CellFormat | undefined>();
  for (const [key, value] of before) {
    const next = after.get(key);
    if (sameValue(value, next)) continue;
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

const selectionForPlan = (plan: SelectionClearPlan): State['selection'] => ({
  active: {
    sheet: plan.sheet,
    row: plan.ranges[0]?.r0 ?? 0,
    col: plan.ranges[0]?.c0 ?? 0,
  },
  anchor: {
    sheet: plan.sheet,
    row: plan.ranges[0]?.r0 ?? 0,
    col: plan.ranges[0]?.c0 ?? 0,
  },
  range: plan.ranges[0] ?? { sheet: plan.sheet, r0: 0, c0: 0, r1: 0, c1: 0 },
  extraRanges: plan.ranges.slice(1),
});

const clearPlannedFormats = (
  store: SpreadsheetStore,
  plan: SelectionClearPlan,
  visualOnly: boolean,
): void => {
  const selected = new Set(plan.formatEntries.map(addrKey));
  const selection = selectionForPlan(plan);
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
    const pending = state.ui.pendingFormat;
    const pendingSelected =
      pending !== null &&
      pending !== undefined &&
      selectionContainsAddr(selection, pending.addr) &&
      isCellWritable(state, pending.addr);
    return {
      ...state,
      format: { ...state.format, formats },
      ui: { ...state.ui, ...(pendingSelected ? { pendingFormat: null } : {}) },
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
    .filter((entry) => entry.range.sheet === range.sheet)
    .map((entry) => entry.range);
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

const preflight = (
  controller: FullInteractionController | undefined,
  operation: InteractionOperation,
  commandId: string,
  effects: readonly OperationEffect[],
): boolean => {
  if (!controller || effects.length === 0 || !controller.policy) return true;
  return controller.canExecute({ operation, origin: 'ribbon', commandId, effects }).allowed;
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

const preflightPlan = (
  deps: ExecuteRibbonClearActionDeps,
  plan: SelectionClearPlan,
  controller: FullInteractionController | undefined,
  commandId: string,
): boolean => {
  if (!controller?.policy) return true;
  if (!preflight(controller, 'clear', commandId, effectsForCells(plan.content))) {
    return false;
  }
  if (!preflight(controller, 'comment', commandId, effectsForCells(plan.comments))) return false;
  if (!preflight(controller, 'hyperlink', commandId, effectsForCells(plan.hyperlinks))) {
    return false;
  }
  if (!preflight(controller, 'validation', commandId, effectsForCells(plan.validations)))
    return false;
  const fullFormat = uniqueAddrs([
    ...plan.formatEntries,
    ...(plan.pendingTarget ? [plan.pendingTarget] : []),
  ]);
  if (!preflight(controller, 'format', commandId, effectsForCells(fullFormat))) return false;
  return preflight(
    controller,
    'format',
    commandId,
    effectsForRanges(rangesForRuleIndexes(deps.store.getState(), plan.conditionalRuleIndexes)),
  );
};

const restorePending = (store: SpreadsheetStore, pending: State['ui']['pendingFormat']): void => {
  mutators.setPendingFormat(store, cloneValue(pending));
};

const refreshCells = (store: SpreadsheetStore, workbook: WorkbookHandle, sheet: number): void => {
  mutators.replaceCells(store, workbook.cells(sheet));
};

export const executeRibbonClearAction = (deps: ExecuteRibbonClearActionDeps): void => {
  const { store, workbook, history, action } = deps;

  // Contents keeps its public command adapter behavior. Clear All runs its
  // content child through the resolved controller inside one History step.
  if (action === 'contents') {
    const range = store.getState().selection.range;
    const result = clearSelectedContents({
      store,
      workbook,
      history,
      commands: deps.commands as ClearContentsCommands | undefined,
      origin: 'ribbon',
    });
    if (result.status !== 'rejected') refreshCells(store, workbook, range.sheet);
    return;
  }

  const resolved = resolveTransactionController(deps, action === 'all');
  if (resolved.mismatch) return;
  const plan = buildPlan(store, workbook);
  if (!plan) {
    if (resolved.ephemeral) resolved.controller?.dispose();
    return;
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
            hyperlinks: [],
            validations: [],
            content: [],
            formatEntries: [],
            visualEntries: [],
            pendingTarget: null,
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
              pendingTarget: null,
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
                pendingTarget: null,
                conditionalRuleIndexes: plan.conditionalRuleIndexes,
              }
            : plan;
  if (
    !preflightPlan(
      deps,
      policyPlan,
      controller,
      action === 'all'
        ? 'clear-all'
        : action === 'formats'
          ? 'clear-formats'
          : action === 'comments'
            ? 'clear-comments'
            : action === 'hyperlinks'
              ? 'clear-hyperlinks'
              : 'clear-conditional',
    )
  ) {
    if (resolved.ephemeral) controller?.dispose();
    return;
  }

  try {
    if (action === 'formats') {
      const targets = uniqueAddrs([
        ...plan.visualEntries,
        ...(plan.pendingTarget ? [plan.pendingTarget] : []),
      ]);
      recordClearFormatEntries(
        history,
        store,
        workbook,
        plan.sheet,
        targets,
        () =>
          withSelectionFormatOrigin(
            store,
            'ribbon',
            () => clearPlannedFormats(store, plan, true),
            'clear-formats',
          ),
        { repeat: () => executeRibbonClearAction(deps) },
      );
      return;
    }
    if (action === 'comments') {
      recordCommentChange(history, store, workbook, plan.comments, () => {
        clearComments(store, plan.comments, workbook);
      });
      return;
    }
    if (action === 'hyperlinks') {
      recordClearFormatEntries(history, store, workbook, plan.sheet, plan.hyperlinks, () => {
        clearPlannedHyperlinks(store, plan);
      });
      return;
    }
    if (action === 'conditional') {
      recordConditionalRulesChange(history, store, () => {
        clearPlannedConditionalRules(store, plan);
      });
      return;
    }

    const pendingBefore = cloneValue(plan.pendingBefore);
    const token = history.begin();
    try {
      const result = clearSelectedContents({
        store,
        workbook,
        history,
        commands: controller,
        origin: 'ribbon',
        commandId: 'clear-all',
      });
      if (result.status === 'rejected') {
        history.abort(token);
        refreshCells(store, workbook, plan.sheet);
        return;
      }

      if (plan.comments.length > 0) {
        recordCommentChange(history, store, workbook, plan.comments, () => {
          clearComments(store, plan.comments, workbook);
        });
      }

      const formatTargets = uniqueAddrs([
        ...plan.formatEntries,
        ...(plan.pendingTarget ? [plan.pendingTarget] : []),
      ]);
      if (formatTargets.length > 0) {
        recordClearFormatEntries(history, store, workbook, plan.sheet, formatTargets, () =>
          withSelectionFormatOrigin(
            store,
            'ribbon',
            () => clearPlannedFormats(store, plan, false),
            'clear-all',
          ),
        );
      }

      if (plan.conditionalRuleIndexes.length > 0) {
        recordConditionalRulesChange(history, store, () => {
          clearPlannedConditionalRules(store, plan);
        });
      }

      refreshCells(store, workbook, plan.sheet);
      history.end(token);
    } catch (error) {
      try {
        history.abort(token);
      } finally {
        restorePending(store, pendingBefore);
        refreshCells(store, workbook, plan.sheet);
      }
      throw error;
    }
  } finally {
    if (resolved.ephemeral) resolved.controller?.dispose();
  }
};
