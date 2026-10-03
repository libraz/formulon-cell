import { formatAsTable, inferTableHasHeaders } from '../../../commands/format-as-table.js';
import type {
  InteractionOperation,
  OperationIntent,
} from '../../../commands/interaction-policy.js';
import { setFreezePanes, showCols, showRows } from '../../../commands/row-col-layout.js';
import { recordTablesChange } from '../../../commands/slice-history.js';
import { addrKey, MAX_COL, MAX_ROW, parseAddrKey } from '../../../engine/address.js';
import { parseRangeRef } from '../../../engine/range-resolver.js';
import type { Addr, Range } from '../../../engine/types.js';
import type { EngineHyperlinkRecord } from '../../../engine/workbook-handle-annotations.js';
import {
  appendDialogFrame,
  createDialogButton,
  createDialogShell,
} from '../../../interact/dialog-shell.js';
import type { SpreadsheetInstance } from '../../../mount/types.js';
import type { CellFormat } from '../../../store/types.js';
import { toolbarLangForLocale } from './locale.js';
import { reportMacRibbonError } from './report-error.js';

export type MacAutomationScriptId =
  | 'allRowsColumns'
  | 'freezeSelection'
  | 'makeSubtable'
  | 'removeHyperlinks'
  | 'countEmptyRows'
  | 'tableToJson'
  | 'newPivotTable';

export interface MacAutomationScriptDescriptor {
  readonly id: MacAutomationScriptId;
  readonly commandId: `mac.automate.${string}`;
  readonly label: string;
  readonly detail: string;
  readonly labelJa: string;
  readonly detailJa: string;
}

/** The sample gallery shipped with Excel for the web/mac parity surface.
 *  Each item is backed by a local workbook operation below; no item claims to
 *  call an Office service or an external script host. */
export const MAC_AUTOMATION_GALLERY: readonly MacAutomationScriptDescriptor[] = [
  {
    id: 'allRowsColumns',
    commandId: 'mac.automate.allRowsColumns',
    label: 'Show all rows and columns',
    detail: 'Unhide rows and columns in the used area.',
    labelJa: 'すべての行と列を表示',
    detailJa: '使用範囲の行と列の非表示を解除します。',
  },
  {
    id: 'freezeSelection',
    commandId: 'mac.automate.freezeSelection',
    label: 'Freeze selection',
    detail: 'Freeze rows above and columns left of the active selection.',
    labelJa: '選択範囲を固定',
    detailJa: '選択範囲の上の行と左の列を固定します。',
  },
  {
    id: 'makeSubtable',
    commandId: 'mac.automate.makeSubtable',
    label: 'Create a subtable from selection',
    detail: 'Create a real table/overlay from the current selection.',
    labelJa: '選択範囲からサブテーブルを作成',
    detailJa: '現在の選択範囲から実際のテーブルを作成します。',
  },
  {
    id: 'removeHyperlinks',
    commandId: 'mac.automate.removeHyperlinks',
    label: 'Remove hyperlinks from sheet',
    detail: 'Clear every hyperlink on the active sheet.',
    labelJa: 'シートのハイパーリンクを削除',
    detailJa: 'アクティブなシートのハイパーリンクをすべて削除します。',
  },
  {
    id: 'countEmptyRows',
    commandId: 'mac.automate.countEmptyRows',
    label: 'Count blank rows',
    detail: 'Count blank rows inside the used area.',
    labelJa: '空白行を数える',
    detailJa: '使用範囲内の空白行を数えます。',
  },
  {
    id: 'tableToJson',
    commandId: 'mac.automate.tableToJson',
    label: 'Return table data as JSON',
    detail: 'Serialize the active table or selection and expose it to the host.',
    labelJa: 'テーブルデータを JSON として取得',
    detailJa: 'アクティブなテーブルまたは選択範囲を JSON に変換して表示します。',
  },
  {
    id: 'newPivotTable',
    commandId: 'mac.automate.newPivotTable',
    label: 'Create new PivotTable from table',
    detail: 'Open the native PivotTable creation dialog.',
    labelJa: 'テーブルから新しいピボットテーブルを作成',
    detailJa: '標準のピボットテーブル作成ダイアログを開きます。',
  },
];

/** Gallery dialog title, shared with the ribbon command that opens it. */
export const MAC_AUTOMATION_GALLERY_TITLE = {
  ja: 'スクリプト ギャラリー',
  en: 'Script Gallery',
} as const;

const MAX_AUTOMATION_MATERIALIZED_CELLS = 100_000;

type AutomationDisposer = () => void;
const AUTOMATION_DISPOSERS = new WeakMap<object, Set<AutomationDisposer>>();

const registerAutomationDisposer = (
  instance: SpreadsheetInstance,
  disposer: AutomationDisposer,
): (() => void) => {
  let disposers = AUTOMATION_DISPOSERS.get(instance);
  if (!disposers) {
    disposers = new Set<AutomationDisposer>();
    AUTOMATION_DISPOSERS.set(instance, disposers);
  }
  disposers.add(disposer);
  return (): void => {
    disposers?.delete(disposer);
    if (disposers?.size === 0) AUTOMATION_DISPOSERS.delete(instance);
  };
};

/** Remove any open Automate gallery/result overlays owned by an instance. */
export function disposeMacAutomationDialogs(instance: SpreadsheetInstance): void {
  const disposers = AUTOMATION_DISPOSERS.get(instance);
  if (!disposers) return;
  for (const dispose of [...disposers]) dispose();
  AUTOMATION_DISPOSERS.delete(instance);
}

const normalizedRange = (range: Range): Range => ({
  sheet: range.sheet,
  r0: Math.max(0, Math.min(range.r0, range.r1)),
  c0: Math.max(0, Math.min(range.c0, range.c1)),
  r1: Math.min(MAX_ROW, Math.max(range.r0, range.r1)),
  c1: Math.min(MAX_COL, Math.max(range.c0, range.c1)),
});

const isMeaningful = (cell: { value: { kind: string }; formula: string | null }): boolean =>
  cell.formula !== null || cell.value.kind !== 'blank';

/** Derive a finite used range. Empty sheets fall back to the current selection
 *  so the scripts remain deterministic and never scan the full Excel grid. */
export function usedRangeForMacAutomation(instance: SpreadsheetInstance): Range {
  const state = instance.store.getState();
  const sheet = state.data.sheetIndex;
  let r0 = MAX_ROW;
  let c0 = MAX_COL;
  let r1 = 0;
  let c1 = 0;
  let found = false;
  const cells =
    typeof instance.workbook.physicalCells === 'function'
      ? instance.workbook.physicalCells(sheet)
      : instance.workbook.cells(sheet);
  for (const cell of cells) {
    if (!isMeaningful(cell)) continue;
    found = true;
    r0 = Math.min(r0, cell.addr.row);
    c0 = Math.min(c0, cell.addr.col);
    r1 = Math.max(r1, cell.addr.row);
    c1 = Math.max(c1, cell.addr.col);
  }
  for (const [key, cell] of state.data.cells) {
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== sheet || !isMeaningful(cell)) continue;
    found = true;
    r0 = Math.min(r0, addr.row);
    c0 = Math.min(c0, addr.col);
    r1 = Math.max(r1, addr.row);
    c1 = Math.max(c1, addr.col);
  }
  for (const key of state.format.formats.keys()) {
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== sheet) continue;
    found = true;
    r0 = Math.min(r0, addr.row);
    c0 = Math.min(c0, addr.col);
    r1 = Math.max(r1, addr.row);
    c1 = Math.max(c1, addr.col);
  }
  return found ? { sheet, r0, c0, r1, c1 } : normalizedRange(state.selection.range);
}

const policyAllows = (
  instance: SpreadsheetInstance,
  operation: InteractionOperation,
  range: Range,
  commandId: string,
): boolean => {
  if (!instance.commands.policy) return true;
  return instance.commands.canExecute({
    operation,
    origin: 'ribbon',
    commandId,
    effects: [{ kind: 'range', range }],
  }).allowed;
};

const markResult = (instance: SpreadsheetInstance, commandId: string, result: unknown): void => {
  instance.host.dataset.fcMacAutomation = commandId;
  instance.host.dataset.fcMacAutomationResult =
    typeof result === 'string' ? result : JSON.stringify(result);
  instance.host.dispatchEvent(
    new CustomEvent('fc:mac-automation', { detail: { commandId, result } }),
  );
};

interface MacAutomationResultOptions {
  title: string;
  summary: string;
  value?: string;
  copyLabel?: string;
  closeLabel: string;
}

/** Keep script output visible after the command returns. The host event and
 * dataset remain useful for integrations, while this small result window
 * gives a user an Excel-like result even when clipboard permission is denied.
 */
const openMacAutomationResult = (
  instance: SpreadsheetInstance,
  options: MacAutomationResultOptions,
): void => {
  let closed = false;
  let shell: ReturnType<typeof createDialogShell>;
  let unregister: (() => void) | null = null;
  const finish = (): void => {
    if (closed) return;
    closed = true;
    unregister?.();
    shell.dispose();
    instance.host.focus();
  };
  shell = createDialogShell({
    host: instance.host,
    className: 'fc-macautomationresult fc-fmtdlg',
    ariaLabel: options.title,
    onDismiss: finish,
  });
  unregister = registerAutomationDisposer(instance, finish);
  const { header, body, footer } = appendDialogFrame(shell, {
    title: options.title,
    panelClasses: ['fc-fmtdlg__panel', 'fc-macautomationresult__panel'],
    bodyClass: 'fc-fmtdlg__body fc-macautomationresult__body',
  });
  header.textContent = options.title;
  const summary = document.createElement('p');
  summary.className = 'fc-macautomationresult__summary';
  summary.textContent = options.summary;
  body.appendChild(summary);
  if (options.value !== undefined) {
    const value = document.createElement('textarea');
    value.className = 'fc-macautomationresult__value';
    value.dataset.macAutomationResult = 'true';
    value.readOnly = true;
    value.spellcheck = false;
    value.value = options.value;
    value.rows = Math.min(18, Math.max(4, options.value.split('\n').length));
    body.appendChild(value);
    if (options.copyLabel) {
      const copy = createDialogButton({ label: options.copyLabel });
      shell.on(copy, 'click', () => {
        if (typeof navigator !== 'undefined' && navigator.clipboard?.writeText) {
          void navigator.clipboard.writeText(options.value ?? '').then(
            () => {
              copy.dataset.copied = 'true';
            },
            () => {
              value.focus();
              value.select();
            },
          );
        } else {
          value.focus();
          value.select();
        }
      });
      footer.appendChild(copy);
    }
  }
  const close = createDialogButton({ label: options.closeLabel, variant: 'primary' });
  shell.on(close, 'click', finish);
  footer.appendChild(close);
  shell.open();
};

type AutomationIntentFactory = (origin: 'ribbon' | 'undo' | 'redo') => readonly OperationIntent[];

/** Authorize apply/undo/redo up front, then open a history transaction that replays under the same checks. */
const beginAuthorizedHistory = (
  instance: SpreadsheetInstance,
  makeIntents: AutomationIntentFactory,
): ReturnType<SpreadsheetInstance['history']['begin']> | null => {
  const undo = makeIntents('undo');
  const redo = makeIntents('redo');
  const intents = [...makeIntents('ribbon'), ...undo, ...redo];
  if (intents.some((intent) => !instance.commands.canExecute(intent).allowed)) return null;
  try {
    return instance.history.begin({ replayAuthorization: { undo, redo } });
  } catch {
    return null;
  }
};

/** Run `apply` inside the transaction; on failure abort and rethrow so the caller can report it. */
const commitOrAbort = (
  instance: SpreadsheetInstance,
  token: NonNullable<ReturnType<typeof beginAuthorizedHistory>>,
  rollbackMessage: string,
  apply: () => void,
): void => {
  try {
    apply();
    instance.history.end(token);
  } catch (error) {
    try {
      instance.history.abort(token);
    } catch (rollback) {
      throw new AggregateError([error, rollback], rollbackMessage);
    }
    throw error;
  }
};

const runAllRowsColumns = (instance: SpreadsheetInstance): void => {
  const range = usedRangeForMacAutomation(instance);
  const token = beginAuthorizedHistory(instance, (origin) =>
    (['resizeRows', 'resizeColumns'] as const).map((operation) => ({
      operation,
      origin,
      commandId: 'mac.automate.allRowsColumns',
      effects: [{ kind: 'range' as const, range: { ...range } }],
    })),
  );
  if (!token) return;
  commitOrAbort(instance, token, 'Show all rows and columns rollback failed.', () => {
    showRows(instance.store, instance.history, range.r0, range.r1, instance.workbook);
    showCols(instance.store, instance.history, range.c0, range.c1, instance.workbook);
  });
  markResult(instance, 'mac.automate.allRowsColumns', { range });
};

const runFreezeSelection = (instance: SpreadsheetInstance): void => {
  const selection = normalizedRange(instance.store.getState().selection.range);
  const token = beginAuthorizedHistory(instance, (origin) => [
    {
      operation: 'format',
      origin,
      commandId: 'mac.automate.freezeSelection',
      effects: [{ kind: 'range', range: { ...selection } }],
    },
  ]);
  if (!token) return;
  commitOrAbort(instance, token, 'Freeze selection rollback failed.', () => {
    setFreezePanes(instance.store, instance.history, selection.r0, selection.c0, instance.workbook);
  });
  markResult(instance, 'mac.automate.freezeSelection', {
    rows: selection.r0,
    cols: selection.c0,
  });
};

const runMakeSubtable = (instance: SpreadsheetInstance): void => {
  const range = normalizedRange(instance.store.getState().selection.range);
  if (!policyAllows(instance, 'table', range, 'mac.automate.makeSubtable')) return;
  let created = false;
  recordTablesChange(instance.history, instance.store, () => {
    created =
      formatAsTable(instance.store, range, {
        workbook: instance.workbook,
        showHeader: inferTableHasHeaders(instance.workbook, range),
      }) !== null;
  });
  if (created) markResult(instance, 'mac.automate.makeSubtable', { range });
};

interface HyperlinkRemovalPlan {
  readonly sheet: number;
  readonly cells: readonly Addr[];
  readonly storeBefore: ReadonlyMap<string, HyperlinkFormatFields>;
  readonly nativeBefore: readonly EngineHyperlinkRecord[];
  readonly nativeAnchorKeys: ReadonlySet<string>;
  readonly removedAnchorCount: number;
}

type HyperlinkFormatFields = Pick<
  CellFormat,
  'hyperlink' | 'hyperlinkDisplay' | 'hyperlinkTooltip'
>;

const hyperlinkCellInGrid = (addr: Addr): boolean =>
  Number.isInteger(addr.sheet) &&
  Number.isInteger(addr.row) &&
  Number.isInteger(addr.col) &&
  addr.sheet >= 0 &&
  addr.row >= 0 &&
  addr.col >= 0 &&
  addr.row <= MAX_ROW &&
  addr.col <= MAX_COL;

const hyperlinkAnchorKey = (sheet: number, row: number, col: number): string =>
  `${sheet}:${row}:${col}`;

const cloneHyperlinkRecord = (record: EngineHyperlinkRecord): EngineHyperlinkRecord => ({
  row: record.row,
  col: record.col,
  lastRow: record.lastRow,
  lastCol: record.lastCol,
  target: record.target,
  location: record.location,
  display: record.display,
  tooltip: record.tooltip,
});

const hyperlinkFormatFields = (format: CellFormat): HyperlinkFormatFields => ({
  hyperlink: format.hyperlink,
  hyperlinkDisplay: format.hyperlinkDisplay,
  hyperlinkTooltip: format.hyperlinkTooltip,
});

const removeStoreHyperlinkFields = (
  instance: SpreadsheetInstance,
  snapshot: ReadonlyMap<string, HyperlinkFormatFields>,
): void => {
  instance.store.setState((state) => {
    const formats = new Map(state.format.formats);
    for (const key of snapshot.keys()) {
      const current = formats.get(key);
      if (!current) continue;
      const next = { ...current };
      delete next.hyperlink;
      delete next.hyperlinkDisplay;
      delete next.hyperlinkTooltip;
      if (Object.keys(next).length === 0) formats.delete(key);
      else formats.set(key, next);
    }
    return { ...state, format: { ...state.format, formats } };
  });
};

const restoreStoreHyperlinkFields = (
  instance: SpreadsheetInstance,
  snapshot: ReadonlyMap<string, HyperlinkFormatFields>,
): void => {
  instance.store.setState((state) => {
    const formats = new Map(state.format.formats);
    for (const [key, before] of snapshot) {
      const current = formats.get(key);
      if (!current) formats.set(key, { ...before });
      else {
        const next = { ...current };
        if (before.hyperlink === undefined) delete next.hyperlink;
        else next.hyperlink = before.hyperlink;
        if (before.hyperlinkDisplay === undefined) delete next.hyperlinkDisplay;
        else next.hyperlinkDisplay = before.hyperlinkDisplay;
        if (before.hyperlinkTooltip === undefined) delete next.hyperlinkTooltip;
        else next.hyperlinkTooltip = before.hyperlinkTooltip;
        formats.set(key, next);
      }
    }
    return { ...state, format: { ...state.format, formats } };
  });
};

const removeNativeHyperlinkAnchors = (
  instance: SpreadsheetInstance,
  sheet: number,
  anchorKeys: ReadonlySet<string>,
): void => {
  for (const key of anchorKeys) {
    const addr = parseAddrKey(key);
    if (!addr) continue;
    if (!instance.workbook.removeHyperlink(sheet, addr.row, addr.col))
      throw new Error(`Could not remove hyperlink at ${addr.row}:${addr.col}.`);
  }
};

const restoreNativeHyperlinks = (
  instance: SpreadsheetInstance,
  sheet: number,
  anchorKeys: ReadonlySet<string>,
  records: readonly EngineHyperlinkRecord[],
): void => {
  removeNativeHyperlinkAnchors(instance, sheet, anchorKeys);
  for (const record of records) {
    if (
      !instance.workbook.addHyperlinkRange(
        sheet,
        record.row,
        record.col,
        record.lastRow,
        record.lastCol,
        record.target,
        record.display,
        record.tooltip,
        record.location,
      )
    )
      throw new Error(`Could not restore hyperlink at ${record.row}:${record.col}.`);
  }
};

const planHyperlinkRemoval = (instance: SpreadsheetInstance): HyperlinkRemovalPlan | null => {
  const state = instance.store.getState();
  const sheet = state.data.sheetIndex;
  const native = instance.workbook.getHyperlinksFull(sheet);
  if (native === null) return null;
  if (native.length > 0 && !instance.workbook.supportsHyperlinkRangeWrite()) return null;

  const cells = new Map<string, Addr>();
  const nativeAnchorKeys = new Set<string>();
  const storeAnchorKeys = new Set<string>();
  const addCell = (addr: Addr): boolean => {
    if (!hyperlinkCellInGrid(addr) || addr.sheet !== sheet) return false;
    if (!cells.has(addrKey(addr))) cells.set(addrKey(addr), { ...addr });
    return true;
  };

  for (const record of native) {
    if (
      !Number.isInteger(record.row) ||
      !Number.isInteger(record.col) ||
      !Number.isInteger(record.lastRow) ||
      !Number.isInteger(record.lastCol) ||
      record.row < 0 ||
      record.col < 0 ||
      record.lastRow < record.row ||
      record.lastCol < record.col ||
      record.lastRow > MAX_ROW ||
      record.lastCol > MAX_COL
    )
      return null;
    nativeAnchorKeys.add(hyperlinkAnchorKey(sheet, record.row, record.col));
    for (let row = record.row; row <= record.lastRow; row += 1) {
      for (let col = record.col; col <= record.lastCol; col += 1) {
        if (cells.size >= MAX_AUTOMATION_MATERIALIZED_CELLS && !cells.has(`${sheet}:${row}:${col}`))
          return null;
        if (!addCell({ sheet, row, col })) return null;
      }
    }
  }

  for (const [key, format] of state.format.formats) {
    if (!format.hyperlink) continue;
    const addr = parseAddrKey(key);
    if (!addr) return null;
    if (addr.sheet !== sheet) continue;
    if (!hyperlinkCellInGrid(addr)) return null;
    const cellKey = addrKey(addr);
    if (cells.size >= MAX_AUTOMATION_MATERIALIZED_CELLS && !cells.has(cellKey)) return null;
    if (!addCell(addr)) return null;
    storeAnchorKeys.add(cellKey);
  }

  const storeBefore = new Map<string, HyperlinkFormatFields>();
  for (const key of storeAnchorKeys) {
    const format = state.format.formats.get(key);
    if (!format?.hyperlink) return null;
    storeBefore.set(key, hyperlinkFormatFields(format));
  }
  const orderedCells = [...cells.values()].sort((a, b) => a.row - b.row || a.col - b.col);
  const anchorKeys = new Set([...nativeAnchorKeys, ...storeAnchorKeys]);
  return {
    sheet,
    cells: orderedCells,
    storeBefore,
    nativeBefore: native.map(cloneHyperlinkRecord),
    nativeAnchorKeys,
    removedAnchorCount: anchorKeys.size,
  };
};

const runRemoveHyperlinks = (instance: SpreadsheetInstance): void => {
  const plan = planHyperlinkRemoval(instance);
  if (!plan || plan.cells.length === 0) return;
  const token = beginAuthorizedHistory(instance, (origin) => [
    {
      operation: 'hyperlink',
      origin,
      commandId: 'mac.automate.removeHyperlinks',
      effects: [{ kind: 'cells', cells: plan.cells }],
    },
  ]);
  if (!token) return;

  const applyStoreRemoved = (): void => removeStoreHyperlinkFields(instance, plan.storeBefore);
  const restoreStoreBefore = (): void => restoreStoreHyperlinkFields(instance, plan.storeBefore);
  const applyRemoved = (): void => {
    try {
      instance.workbook.withEngineSyncMuted(() => {
        removeNativeHyperlinkAnchors(instance, plan.sheet, plan.nativeAnchorKeys);
        applyStoreRemoved();
      });
    } catch (error) {
      try {
        instance.workbook.withEngineSyncMuted(() => {
          restoreNativeHyperlinks(instance, plan.sheet, plan.nativeAnchorKeys, plan.nativeBefore);
          restoreStoreBefore();
        });
      } catch (rollback) {
        throw new AggregateError([error, rollback], 'Remove hyperlinks rollback failed.');
      }
      throw error;
    }
  };
  const applyRestored = (): void => {
    try {
      instance.workbook.withEngineSyncMuted(() => {
        restoreNativeHyperlinks(instance, plan.sheet, plan.nativeAnchorKeys, plan.nativeBefore);
        restoreStoreBefore();
      });
    } catch (error) {
      try {
        instance.workbook.withEngineSyncMuted(() => {
          removeNativeHyperlinkAnchors(instance, plan.sheet, plan.nativeAnchorKeys);
          applyStoreRemoved();
        });
      } catch (rollback) {
        throw new AggregateError([error, rollback], 'Restore hyperlinks rollback failed.');
      }
      throw error;
    }
  };

  commitOrAbort(instance, token, 'Remove hyperlinks history rollback failed.', () => {
    applyRemoved();
    instance.history.push({ undo: applyRestored, redo: applyRemoved });
  });
  markResult(instance, 'mac.automate.removeHyperlinks', {
    removed: plan.removedAnchorCount,
    sheet: plan.sheet,
  });
};

const runCountEmptyRows = (instance: SpreadsheetInstance): void => {
  const range = usedRangeForMacAutomation(instance);
  const rowsWithData = new Set<number>();
  const cells =
    typeof instance.workbook.physicalCells === 'function'
      ? instance.workbook.physicalCells(range.sheet)
      : instance.workbook.cells(range.sheet);
  for (const cell of cells) {
    if (isMeaningful(cell) && cell.addr.row >= range.r0 && cell.addr.row <= range.r1)
      rowsWithData.add(cell.addr.row);
  }
  const state = instance.store.getState();
  for (const [key, cell] of state.data.cells) {
    const addr = parseAddrKey(key);
    if (
      addr &&
      addr.sheet === range.sheet &&
      addr.row >= range.r0 &&
      addr.row <= range.r1 &&
      isMeaningful(cell)
    )
      rowsWithData.add(addr.row);
  }
  let blankRows = 0;
  for (let row = range.r0; row <= range.r1; row += 1) {
    if (!rowsWithData.has(row)) blankRows += 1;
  }
  markResult(instance, 'mac.automate.countEmptyRows', { blankRows, range });
  const japanese = toolbarLangForLocale(instance.i18n.locale) === 'ja';
  openMacAutomationResult(instance, {
    title: japanese ? '空白行を数える' : 'Count blank rows',
    summary: japanese
      ? `使用範囲内の空白行: ${blankRows} 行`
      : `Blank rows inside the used area: ${blankRows}`,
    closeLabel: japanese ? '閉じる' : 'Close',
  });
};

const jsonValue = (value: ReturnType<SpreadsheetInstance['workbook']['getValue']>): unknown => {
  switch (value.kind) {
    case 'blank':
      return null;
    case 'number':
    case 'text':
    case 'bool':
      return value.value;
    case 'error':
      return value.text || value.code;
  }
};

interface JsonPlan {
  readonly range: Range;
  readonly dataStartRow: number;
  readonly columns?: readonly string[];
}

const safeRangeArea = (range: Range): number | null => {
  if (
    !Number.isInteger(range.sheet) ||
    range.sheet < 0 ||
    !Number.isInteger(range.r0) ||
    !Number.isInteger(range.c0) ||
    !Number.isInteger(range.r1) ||
    !Number.isInteger(range.c1) ||
    range.r0 < 0 ||
    range.c0 < 0 ||
    range.r1 < range.r0 ||
    range.c1 < range.c0 ||
    range.r1 > MAX_ROW ||
    range.c1 > MAX_COL
  )
    return null;
  const rows = range.r1 - range.r0 + 1;
  const cols = range.c1 - range.c0 + 1;
  if (!Number.isSafeInteger(rows) || !Number.isSafeInteger(cols) || rows <= 0 || cols <= 0)
    return null;
  const area = rows * cols;
  return Number.isSafeInteger(area) && area <= MAX_AUTOMATION_MATERIALIZED_CELLS ? area : null;
};

const planJson = (instance: SpreadsheetInstance): JsonPlan | null => {
  const state = instance.store.getState();
  const active = state.selection.active;
  const table = instance.workbook.getTables().find((candidate) => {
    const parsed = parseRangeRef(candidate.ref);
    return (
      parsed !== null &&
      candidate.sheetIndex === active.sheet &&
      active.row >= parsed.r0 &&
      active.row <= parsed.r1 &&
      active.col >= parsed.c0 &&
      active.col <= parsed.c1
    );
  });
  if (table) {
    const parsed = parseRangeRef(table.ref);
    if (parsed) {
      const range: Range = {
        sheet: table.sheetIndex,
        r0: parsed.r0,
        c0: parsed.c0,
        r1: parsed.r1,
        c1: parsed.c1,
      };
      if (safeRangeArea(range) === null) return null;
      const columns = table.columns.length
        ? table.columns
        : Array.from({ length: range.c1 - range.c0 + 1 }, (_, i) => `Column${i + 1}`);
      return { range, dataStartRow: range.r0 + 1, columns };
    }
  }
  const range = normalizedRange(state.selection.range);
  if (safeRangeArea(range) === null) return null;
  return { range, dataStartRow: range.r0 };
};

const toJson = (instance: SpreadsheetInstance, plan: JsonPlan): string => {
  const { range } = plan;
  if (plan.columns) {
    const rows: Record<string, unknown>[] = [];
    for (let row = plan.dataStartRow; row <= range.r1; row += 1) {
      const item: Record<string, unknown> = {};
      for (let col = range.c0; col <= range.c1; col += 1) {
        const key = plan.columns[col - range.c0] ?? `Column${col - range.c0 + 1}`;
        item[key] = jsonValue(instance.workbook.getValue({ sheet: range.sheet, row, col }));
      }
      rows.push(item);
    }
    return JSON.stringify(rows, null, 2);
  }
  const rows: unknown[][] = [];
  for (let row = plan.dataStartRow; row <= range.r1; row += 1) {
    const values: unknown[] = [];
    for (let col = range.c0; col <= range.c1; col += 1)
      values.push(jsonValue(instance.workbook.getValue({ sheet: range.sheet, row, col })));
    rows.push(values);
  }
  return JSON.stringify(rows, null, 2);
};

const runTableToJson = async (instance: SpreadsheetInstance): Promise<void> => {
  const plan = planJson(instance);
  if (!plan) return;
  const { range } = plan;
  if (
    !instance.commands.canExecute({
      operation: 'export',
      origin: 'ribbon',
      commandId: 'mac.automate.tableToJson',
      effects: [{ kind: 'range', range }],
    }).allowed
  )
    return;
  const json = toJson(instance, plan);
  let copied = false;
  if (typeof navigator !== 'undefined' && navigator.clipboard?.writeText) {
    try {
      await navigator.clipboard.writeText(json);
      copied = true;
    } catch {
      // A denied clipboard permission still leaves the JSON available to the
      // host event below, just as Excel's script result remains inspectable.
    }
  }
  markResult(instance, 'mac.automate.tableToJson', { json, range, copied });
  const japanese = toolbarLangForLocale(instance.i18n.locale) === 'ja';
  openMacAutomationResult(instance, {
    title: japanese ? 'テーブルデータを JSON として取得' : 'Return table data as JSON',
    summary: japanese
      ? copied
        ? 'JSON をクリップボードにコピーしました。'
        : 'クリップボードにアクセスできないため、下の JSON をコピーしてください。'
      : copied
        ? 'The JSON was copied to the clipboard.'
        : 'Clipboard access was unavailable; copy the JSON below.',
    value: json,
    copyLabel: japanese ? 'コピー' : 'Copy',
    closeLabel: japanese ? '閉じる' : 'Close',
  });
};

/** Execute one gallery script. The return value is intentionally small and
 *  serializable so a host can expose it in a toast or test harness. */
export async function runMacAutomationScript(
  instance: SpreadsheetInstance,
  id: MacAutomationScriptId,
): Promise<void> {
  switch (id) {
    case 'allRowsColumns':
      runAllRowsColumns(instance);
      return undefined;
    case 'freezeSelection':
      runFreezeSelection(instance);
      return undefined;
    case 'makeSubtable':
      runMakeSubtable(instance);
      return undefined;
    case 'removeHyperlinks':
      runRemoveHyperlinks(instance);
      return undefined;
    case 'countEmptyRows':
      runCountEmptyRows(instance);
      return undefined;
    case 'tableToJson':
      await runTableToJson(instance);
      return;
    case 'newPivotTable':
      instance.openPivotTableDialog({ placement: 'new' });
      return;
  }
}

/** Show the native-looking sample script gallery used by the Automate tab. */
export async function openMacAutomationGallery(instance: SpreadsheetInstance): Promise<void> {
  await new Promise<void>((resolve) => {
    const japanese = toolbarLangForLocale(instance.i18n.locale) === 'ja';
    const title = japanese ? MAC_AUTOMATION_GALLERY_TITLE.ja : MAC_AUTOMATION_GALLERY_TITLE.en;
    let closed = false;
    let shell: ReturnType<typeof createDialogShell>;
    let unregister: (() => void) | null = null;
    const finish = (): void => {
      if (closed) return;
      closed = true;
      unregister?.();
      shell.dispose();
      instance.host.focus();
      resolve();
    };
    shell = createDialogShell({
      host: instance.host,
      className: 'fc-macautomationdlg fc-fmtdlg',
      ariaLabel: title,
      onDismiss: finish,
    });
    unregister = registerAutomationDisposer(instance, finish);
    const { header, body, footer } = appendDialogFrame(shell, {
      title,
      panelClasses: ['fc-fmtdlg__panel', 'fc-macautomationdlg__panel'],
      bodyClass: 'fc-fmtdlg__body fc-macautomationdlg__body',
    });
    header.textContent = title;
    for (const script of MAC_AUTOMATION_GALLERY) {
      const button = createDialogButton({
        label: japanese ? script.labelJa : script.label,
        baseClass: 'fc-macautomationdlg__script',
      });
      button.dataset.macAutomationScript = script.id;
      button.title = japanese ? script.detailJa : script.detail;
      shell.on(button, 'click', () => {
        void runMacAutomationScript(instance, script.id)
          .catch((error: unknown) => reportMacRibbonError(instance, error))
          .finally(finish);
      });
      body.appendChild(button);
    }
    const cancel = createDialogButton({ label: japanese ? 'キャンセル' : 'Cancel' });
    shell.on(cancel, 'click', finish);
    footer.appendChild(cancel);
    shell.open();
  });
}
