// Composes the built-in context-menu entry list for each menu kind, gated by host wiring and state.

import type { ClipboardSnapshot } from '../commands/clipboard/snapshot.js';
import { tableForCell } from '../commands/format-as-table.js';
import { noteChangePermitted } from '../commands/protection.js';
import { rowDeleteHitsTableHeader } from '../commands/table-rows.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';
import type { ContextMenuDeps } from './context-menu.js';
import { wholeBandAxisFor } from './context-menu-insert-copied.js';
import { pivotMenuTarget } from './context-menu-pivot.js';
import {
  buildCellEntries,
  buildColEntries,
  buildMacCellEntries,
  buildMacColEntries,
  buildMacCornerEntries,
  buildMacRowEntries,
  buildPivotEntries,
  buildRowEntries,
  compactMenuEntries,
  isDialogMenuItemId,
  type MenuEntry,
  type MenuKind,
  PLAIN_INSERT_IDS,
  toMacBandEntries,
  withoutMenuItems,
  withTableInsertMenu,
} from './context-menu-spec.js';

/** Inputs the built-in entry list reads; `deps` keeps the host callbacks that gate entries. */
export interface BuiltInEntriesInput {
  kind: MenuKind;
  corner: boolean;
  mac: boolean;
  pivot: NonNullable<ReturnType<WorkbookHandle['pivotCellInfo']>> | null;
  pivotAddr: Addr;
  activeAddr: Addr;
  hasNote: boolean;
  strings: Strings;
  wb: WorkbookHandle;
  store: SpreadsheetStore;
  clipboard: { snapshot(): ClipboardSnapshot | null };
  deps: Pick<
    ContextMenuDeps,
    | 'onOpenPivotFieldList'
    | 'isPivotFieldListOpen'
    | 'dialogAttached'
    | 'onInsertHyperlink'
    | 'onOpenHyperlink'
    | 'onEditComment'
    | 'onPickFromList'
    | 'onToggleWatch'
    | 'onDefineName'
    | 'isWatched'
  >;
}

/** Builds the built-in entry list for a menu kind, filtered by the host wiring and current state. */
export function composeBuiltInEntries(input: BuiltInEntriesInput): MenuEntry[] {
  const {
    kind,
    corner,
    mac,
    pivot,
    pivotAddr,
    activeAddr,
    hasNote,
    strings,
    wb,
    store,
    clipboard,
    deps,
  } = input;
  const raw = pivot
    ? buildPivotEntries(
        strings,
        pivotMenuTarget(
          wb,
          { addr: pivotAddr, info: pivot },
          !!deps.onOpenPivotFieldList,
          deps.isPivotFieldListOpen ? !deps.isPivotFieldListOpen() : undefined,
        ),
        mac ? 'mac' : 'default',
      )
    : kind === 'row'
      ? mac
        ? buildMacRowEntries(strings)
        : buildRowEntries(strings)
      : kind === 'col'
        ? mac
          ? buildMacColEntries(strings)
          : buildColEntries(strings)
        : mac
          ? (corner ? buildMacCornerEntries : buildMacCellEntries)(strings, hasNote)
          : buildCellEntries(strings, hasNote);
  const stateNow = store.getState();
  const bandRange = stateNow.selection.range;
  const inTable =
    kind === 'cell' && !pivot && !stateNow.ui.copyRange
      ? tableForCell(stateNow.tables.tables, activeAddr.sheet, activeAddr.row, activeAddr.col)
      : null;
  const baseEntries = mac && pivot ? toMacBandEntries(raw, strings) : raw;
  const platformEntries = inTable
    ? withTableInsertMenu(baseEntries, strings, {
        rows: !(inTable.showHeader && bandRange.r0 === inTable.range.r0),
      })
    : baseEntries;
  const hasCopiedCells = !!store.getState().ui.copyRange;
  const pendingSnapshot = hasCopiedCells ? clipboard.snapshot() : null;
  const pendingCutBand =
    pendingSnapshot?.mode === 'cut' && wholeBandAxisFor(pendingSnapshot) === kind;
  const watched = !!deps.isWatched?.(activeAddr);
  const entries = compactMenuEntries(
    withoutMenuItems(
      platformEntries,
      (id) => isDialogMenuItemId(id) && deps.dialogAttached?.(id) === false,
    )
      .filter((e) => !(e.kind === 'item' && e.id === 'insertHyperlink' && !deps.onInsertHyperlink))
      .filter((e) => !(e.kind === 'item' && e.id === 'openHyperlink' && !deps.onOpenHyperlink))
      // A pending copy swaps the plain insert entries for "Insert Copied
      // Cells" — the desktop menus never show both at once.
      .filter((e) => !(e.kind === 'item' && e.id === 'insertCopiedCells' && !hasCopiedCells))
      .filter((e) => !(e.kind === 'item' && hasCopiedCells && PLAIN_INSERT_IDS.includes(e.id)))
      .filter((e) => !(e.kind === 'item' && e.id === 'insertComment' && !deps.onEditComment))
      .filter((e) => !(e.kind === 'item' && e.id === 'pickFromList' && !deps.onPickFromList))
      .filter(
        (e) => !(e.kind === 'item' && e.id === 'newComment' && !wb.capabilities.threadedComments),
      )
      .filter((e) => !(e.kind === 'item' && e.id === 'showHideNote' && !hasNote))
      // A protected sheet without Edit objects drops a note's edit and delete rows entirely,
      // and on the Mac New Note as well.
      .filter(
        (e) =>
          !(
            e.kind === 'item' &&
            (e.id === 'insertComment' || e.id === 'deleteComment') &&
            (mac || hasNote) &&
            !noteChangePermitted(store.getState(), activeAddr)
          ),
      )
      .filter((e) => !(e.kind === 'item' && e.id === 'toggleWatch' && !deps.onToggleWatch))
      .filter((e) => !(e.kind === 'item' && e.id === 'defineName' && !deps.onDefineName))
      .map((e) => {
        if (e.kind === 'item' && e.id === 'toggleWatch') {
          return {
            ...e,
            label: watched ? strings.contextMenu.removeWatch : strings.contextMenu.addWatch,
          };
        }
        // No delete is offered over a table's header row.
        if (
          e.kind === 'item' &&
          e.id === 'rowDelete' &&
          rowDeleteHitsTableHeader(
            store,
            bandRange.sheet,
            bandRange.r0,
            bandRange.r1 - bandRange.r0 + 1,
          )
        ) {
          return { ...e, disabled: true };
        }
        if (e.kind === 'item' && e.id === 'insertCopiedCells' && pendingCutBand) {
          return { ...e, label: strings.contextMenu.insertCutCells };
        }
        return e;
      }),
  );
  return entries;
}
